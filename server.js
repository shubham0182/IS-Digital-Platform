const express = require("express");
const compression = require("compression");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const multer = require("multer");
const PDFDocument = require("pdfkit");
const docx = require("docx");
const { Document, Packer, Paragraph, TextRun, TableRow, TableCell, Table, WidthType, AlignmentType, BorderStyle, ShadingType, ImageRun } = docx;

const app = express();
const PORT = process.env.PORT || 8080;
const DATA_DIR = path.join(__dirname, "data");
const UPLOADS_DIR = path.join(__dirname, "data", "uploads");
const PUBLIC_DIR = path.join(__dirname, "public");
const ADMIN_KEY = process.env.ADMIN_KEY || "shubham31";

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname) || ".jpg";
    cb(null, Date.now() + "-" + crypto.randomUUID().slice(0, 8) + ext);
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith("image/")) cb(null, true);
    else cb(new Error("Only image files are allowed"));
  },
});

function readJSON(file) {
  const p = path.join(DATA_DIR, file);
  if (!fs.existsSync(p)) return [];
  return JSON.parse(fs.readFileSync(p, "utf8"));
}
function writeJSON(file, data) {
  fs.writeFileSync(path.join(DATA_DIR, file), JSON.stringify(data, null, 2), "utf8");
}

// ===================== NOTIFICATIONS =====================
// Notifications are stored in data/notifications.json (one record per
// recipient). Broadcasts fan out into one record per recipient at creation
// time, which keeps reads trivial and matches the project's JSON store style.
// A separate data/fcm_tokens.json maps userId -> [web push tokens].

const NOTIF_FILE = "notifications.json";
const FCM_TOKENS_FILE = "fcm_tokens.json";

function nowISO() {
  return new Date().toISOString();
}

function getNotifs() {
  return readJSON(NOTIF_FILE);
}
function saveNotifs(list) {
  writeJSON(NOTIF_FILE, list);
}
function getFcmTokens() {
  return readJSON(FCM_TOKENS_FILE);
}
function saveFcmTokens(map) {
  writeJSON(FCM_TOKENS_FILE, map);
}

function getApprovedUsers() {
  return readJSON("users.json").filter(u => u.status === "approved");
}
function getUserById(id) {
  return readJSON("users.json").find(u => u.id === id) || null;
}

/* Resolve the recipient ids for a notification.
   - userId set  -> direct to that single user
   - role "admin"-> all approved admin users
   - role "all"  -> every approved user
   - role "user" -> every approved regular user                        */
function resolveNotificationRecipients({ userId, role }) {
  if (userId) return [userId];
  const users = getApprovedUsers();
  if (role === "admin") return users.filter(u => u.role === "admin").map(u => u.id);
  if (role === "all") return users.map(u => u.id);
  return users.filter(u => (u.role || "user") === "user").map(u => u.id);
}

/* Create one notification per recipient. dedupeKey prevents the same event
   being fired twice for the same recipient (e.g. repeated logins on a new
   device, double order submissions). */
function createNotification({ type, title, message, link = "#notifications", icon = "bell", userId = null, role = "user", dedupeKey = null, meta = {} }) {
  if (!title || !message) return [];
  const notifs = getNotifs();
  const created = [];
  const recipients = resolveNotificationRecipients({ userId, role });

  recipients.forEach(uid => {
    if (dedupeKey && notifs.some(n => n.userId === uid && n.dedupeKey === dedupeKey && !n.deleted)) return;
    const notification = {
      id: crypto.randomUUID(),
      userId: uid,
      type: type || "system",
      title,
      message,
      link,
      icon: icon || "bell",
      read: false,
      readAt: null,
      deleted: false,
      createdAt: nowISO(),
      dedupeKey: dedupeKey || null,
      meta: meta || {},
    };
    notifs.push(notification);
    created.push(notification);
  });

  saveNotifs(notifs);
  if (created.length) pushToRecipients(recipients, created[0]);
  return created;
}

/* Convenience: notify every admin user of a system event. */
function notifyAdmins(opts) {
  return createNotification({ ...opts, role: "admin" });
}

/* A lightweight device fingerprint (UA + IP) used for new-device logins. */
function deviceSignature(req) {
  const ua = req.headers["user-agent"] || "unknown";
  return crypto.createHash("sha1").update(ua + "|" + (req.ip || "")).digest("hex").slice(0, 12);
}

// ===================== FCM (web push) =====================
// Server-side sending is optional: it activates only when a Firebase service
// account file exists (config/firebase-service-account.json) or
// FCM_SERVICE_ACCOUNT env var points to one. Without it, the in-app
// notification system (badge + center) still works fully; the web push
// client tokens are still registered so they are ready to use.

let fcmReady = false;
let firebaseAdmin = null;
try { firebaseAdmin = require("firebase-admin"); } catch (e) { firebaseAdmin = null; }

function initFcm() {
  if (fcmReady) return true;
  if (!firebaseAdmin) return false;
  const saPath = process.env.FCM_SERVICE_ACCOUNT || path.join(__dirname, "config", "firebase-service-account.json");
  if (!fs.existsSync(saPath)) return false;
  try {
    if (!firebaseAdmin.apps.length) {
      firebaseAdmin.initializeApp({ credential: firebaseAdmin.credential.cert(require(saPath)) });
    }
    fcmReady = true;
    return true;
  } catch (e) {
    console.warn("FCM init failed:", e.message);
    return false;
  }
}

function pushToRecipients(recipientIds, notif) {
  if (!recipientIds.length || !initFcm()) return;
  const tokensMap = getFcmTokens();
  const tokens = [];
  recipientIds.forEach(uid => {
    (tokensMap[uid] || []).forEach(t => { if (!tokens.includes(t)) tokens.push(t); });
  });
  if (!tokens.length) return;
  const link = (notif.link || "#notifications").replace(/^#/, "/");
  firebaseAdmin.messaging().sendEachForMulticast({
    tokens,
    notification: { title: notif.title, body: notif.message, icon: "/logo.png" },
    data: { notificationId: notif.id, type: notif.type || "", url: link },
    webpush: { fcm_options: { link } },
  }).catch(err => console.warn("FCM send error:", err.message));
}

app.use(helmet({ contentSecurityPolicy: false, crossOriginEmbedderPolicy: false, crossOriginOpenerPolicy: { policy: "same-origin-allow-popups" } }));
app.use(compression());
app.use(cors());
app.use(express.json({ limit: "25mb" }));
app.use(express.urlencoded({ extended: true, limit: "25mb" }));

const apiLimiter = rateLimit({ windowMs: 15 * 60 * 1000, max: 100, standardHeaders: true, legacyHeaders: false });
const orderLimiter = rateLimit({ windowMs: 60 * 60 * 1000, max: 20, standardHeaders: true, legacyHeaders: false });

app.post("/api/orders", orderLimiter, upload.array("photos", 6), (req, res) => {
  try {
    const body = req.body;
    const { eventType, eventDate, hostName, name1, venueName, venueCity, cname, email, mobile, package: pkg, payment } = body;
    if (!eventType || !eventDate || !hostName || !name1 || !venueName || !venueCity || !cname || !email || !mobile || !pkg || !payment) {
      return res.status(400).json({ success: false, error: "Please fill all required fields." });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ success: false, error: "Please enter a valid email address." });
    }
    const photoPaths = (req.files || []).map(f => "/api/uploads/" + f.filename);
    const order = {
      id: crypto.randomUUID(),
      ...body,
      photos: photoPaths,
      createdAt: new Date().toISOString(),
      status: "pending",
    };
    if (order.userId) {
      const orderingUser = getUserById(order.userId);
      if (orderingUser) {
        order.userName = orderingUser.name;
        order.userEmail = orderingUser.email;
        order.userPicture = orderingUser.picture || "";
      }
    }
    const orders = readJSON("orders.json");
    orders.push(order);
    writeJSON("orders.json", orders);

    // Notify the ordering user and every admin about the new order.
    if (order.userId) {
      createNotification({
        type: "order_submitted",
        title: "Order received",
        message: "We've received your " + (body.eventType || "invitation") + " order. We'll reach out within 2 hours.",
        link: "#order",
        icon: "order",
        userId: order.userId,
        dedupeKey: "order:" + order.id,
      });
    }
    notifyAdmins({
      type: "order_submitted",
      title: "New order received",
      message: (body.cname || "A client") + " placed a " + (body.package || "") + " order for " + (body.eventType || "an event") + ".",
      link: "/admin.html?section=orders",
      icon: "order",
      dedupeKey: "order-admin:" + order.id,
    });

    res.json({ success: true, message: "Order submitted successfully! We'll reach out within 2 hours.", orderId: order.id });
  } catch (err) {
    console.error("Order error:", err);
    res.status(500).json({ success: false, error: "Something went wrong. Please try again." });
  }
});

/* A signed-in user's own orders, newest first. */
app.get("/api/my-orders", requireUser, apiLimiter, (req, res) => {
  const orders = readJSON("orders.json");
  const mine = orders
    .filter(o => o.userId === req.authUser.id)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  res.json({ success: true, orders: mine });
});

app.post("/api/contact", apiLimiter, (req, res) => {
  try {
    const { name, email, message } = req.body;
    if (!name || !email || !message) {
      return res.status(400).json({ success: false, error: "Please fill all fields." });
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return res.status(400).json({ success: false, error: "Please enter a valid email address." });
    }
    const contacts = readJSON("contacts.json");
    const contact = { id: crypto.randomUUID(), name, email, message, createdAt: new Date().toISOString(), read: false };
    contacts.push(contact);
    writeJSON("contacts.json", contacts);
    notifyAdmins({
      type: "contact_message",
      title: "New contact message",
      message: name + " (" + email + ") sent a message.",
      link: "/admin.html?section=contacts",
      icon: "mail",
      dedupeKey: "contact:" + contact.id,
    });
    res.json({ success: true, message: "Message sent! We'll reply within 24 hours." });
  } catch (err) {
    console.error("Contact error:", err);
    res.status(500).json({ success: false, error: "Something went wrong. Please try again." });
  }
});

function requireAdmin(req, res, next) {
  const key = req.headers["x-admin-key"] || req.query.key;
  if (key !== ADMIN_KEY && key !== "ishu3") return res.status(403).json({ success: false, error: "Unauthorized" });
  next();
}

/* User-scoped requests are identified by the id stored in localStorage
   (ts-user) and sent as x-user-id header, matching how the client already
   sends userId with order/contact payloads. */
function requireUser(req, res, next) {
  const userId = (req.headers["x-user-id"] || req.query.userId || (req.body && req.body.userId) || "").trim();
  if (!userId) return res.status(400).json({ success: false, error: "Missing user id" });
  const user = getUserById(userId);
  if (!user) return res.status(403).json({ success: false, error: "Unauthorized" });
  req.authUser = user;
  next();
}

/* Admin APIs accept either the legacy admin key or an approved admin-role user. */
function requireAdminOrKey(req, res, next) {
  const key = req.headers["x-admin-key"] || req.query.key;
  if (key === ADMIN_KEY || key === "ishu3") return next();
  const userId = req.headers["x-user-id"] || req.query.userId;
  const user = userId ? getUserById(userId) : null;
  if (user && user.role === "admin" && user.status === "approved") return next();
  return res.status(403).json({ success: false, error: "Unauthorized" });
}

// ===================== NOTIFICATION APIs =====================

/* Get the signed-in user's notifications (broadcasts included) with
   optional search + filter + unread count. */
app.get("/api/notifications", requireUser, apiLimiter, (req, res) => {
  try {
    const userId = req.authUser.id;
    const q = (req.query.q || "").toLowerCase().trim();
    const filter = req.query.filter || "all";
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 300);

    let list = getNotifs().filter(n => n.userId === userId && !n.deleted);
    list.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));

    const unreadCount = list.filter(n => !n.read).length;
    if (filter === "unread") list = list.filter(n => !n.read);
    else if (filter === "read") list = list.filter(n => n.read);
    if (q) list = list.filter(n => (n.title + " " + n.message).toLowerCase().includes(q));

    res.json({ success: true, notifications: list.slice(0, limit), total: list.length, unreadCount });
  } catch (err) {
    console.error("Notifications error:", err);
    res.status(500).json({ success: false, error: "Failed to load notifications" });
  }
});

/* Lightweight poll used by the bell badge. */
app.get("/api/notifications/unread-count", requireUser, apiLimiter, (req, res) => {
  try {
    const userId = req.authUser.id;
    const unreadCount = getNotifs().filter(n => n.userId === userId && !n.deleted && !n.read).length;
    res.json({ success: true, unreadCount });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to load unread count" });
  }
});

/* Mark a single notification as read. */
app.put("/api/notifications/:id/read", requireUser, apiLimiter, (req, res) => {
  try {
    const notifs = getNotifs();
    const n = notifs.find(x => x.id === req.params.id && x.userId === req.authUser.id && !x.deleted);
    if (!n) return res.status(404).json({ success: false, error: "Notification not found" });
    if (!n.read) { n.read = true; n.readAt = nowISO(); saveNotifs(notifs); }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to update notification" });
  }
});

/* Mark every notification for the user as read. */
app.put("/api/notifications/read-all", requireUser, apiLimiter, (req, res) => {
  try {
    const notifs = getNotifs();
    let changed = false;
    notifs.forEach(n => {
      if (n.userId === req.authUser.id && !n.deleted && !n.read) { n.read = true; n.readAt = nowISO(); changed = true; }
    });
    if (changed) saveNotifs(notifs);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to update notifications" });
  }
});

/* Delete (soft) a notification so the user stops seeing it while the
   server keeps a history record for admins. */
app.delete("/api/notifications/:id", requireUser, apiLimiter, (req, res) => {
  try {
    const notifs = getNotifs();
    const n = notifs.find(x => x.id === req.params.id && x.userId === req.authUser.id && !x.deleted);
    if (!n) return res.status(404).json({ success: false, error: "Notification not found" });
    n.deleted = true;
    n.deletedAt = nowISO();
    saveNotifs(notifs);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to delete notification" });
  }
});

/* Create a notification / broadcast (admin key or admin-role user).
   Used by the admin panel announcement composer and any integrations. */
app.post("/api/notifications", requireAdminOrKey, apiLimiter, (req, res) => {
  try {
    const { type, title, message, link, icon, role, userId, dedupeKey } = req.body || {};
    if (!title || !message) return res.status(400).json({ success: false, error: "Title and message are required" });
    const created = createNotification({ type, title, message, link, icon, role: role || "user", userId: userId || null, dedupeKey });
    res.json({ success: true, created: created.length, message: "Notification sent to " + created.length + " recipient(s)" });
  } catch (err) {
    console.error("Create notification error:", err);
    res.status(500).json({ success: false, error: "Failed to create notification" });
  }
});

/* Admin view: recent activity across all users. */
app.get("/api/notifications/admin", requireAdmin, apiLimiter, (req, res) => {
  try {
    const q = (req.query.q || "").toLowerCase().trim();
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 500);
    let list = getNotifs().filter(n => !n.deleted);
    list.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    if (q) list = list.filter(n => (n.title + " " + n.message + " " + (n.userId || "")).toLowerCase().includes(q));
    const unreadCount = list.filter(n => !n.read).length;
    res.json({ success: true, notifications: list.slice(0, limit), total: list.length, unreadCount });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to load notifications" });
  }
});

/* Register / refresh an FCM web-push token for a user. */
app.post("/api/notifications/fcm-token", requireUser, apiLimiter, (req, res) => {
  try {
    const { token } = req.body || {};
    if (!token || typeof token !== "string" || token.length > 4096) {
      return res.status(400).json({ success: false, error: "Invalid push token" });
    }
    const map = getFcmTokens();
    const existing = (map[req.authUser.id] || []).filter(t => t.token !== token);
    existing.push({ token, createdAt: nowISO() });
    map[req.authUser.id] = existing.slice(-5);
    saveFcmTokens(map);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to save push token" });
  }
});

// ===================== USER PROFILE =====================

app.get("/api/users/me", requireUser, apiLimiter, (req, res) => {
  const u = req.authUser;
  res.json({ success: true, user: { id: u.id, name: u.name, email: u.email, picture: u.picture, status: u.status, role: u.role || "user", phone: u.phone || "", city: u.city || "" } });
});

app.put("/api/users/me", requireUser, apiLimiter, (req, res) => {
  try {
    const { name, phone, city } = req.body || {};
    const users = readJSON("users.json");
    const u = users.find(x => x.id === req.authUser.id);
    if (!u) return res.status(404).json({ success: false, error: "User not found" });
    if (typeof name === "string" && name.trim()) u.name = name.trim().slice(0, 80);
    if (typeof phone === "string") u.phone = phone.trim().slice(0, 20);
    if (typeof city === "string") u.city = city.trim().slice(0, 60);
    u.updatedAt = nowISO();
    writeJSON("users.json", users);
    createNotification({
      type: "profile_updated",
      title: "Profile updated",
      message: "Your profile details were updated successfully.",
      link: "#notifications",
      icon: "user",
      userId: u.id,
      dedupeKey: "profile:" + u.id + ":" + nowISO().slice(11, 16),
    });
    res.json({ success: true, user: { id: u.id, name: u.name, email: u.email, picture: u.picture, status: u.status, role: u.role || "user", phone: u.phone || "", city: u.city || "" } });
  } catch (err) {
    console.error("Profile update error:", err);
    res.status(500).json({ success: false, error: "Failed to update profile" });
  }
});

/* Promote / demote a user's role between "user" and "admin". */
app.put("/api/admin/users/:id/role", requireAdmin, (req, res) => {
  try {
    const { role } = req.body || {};
    if (!["user", "admin"].includes(role)) return res.status(400).json({ success: false, error: "Invalid role" });
    const users = readJSON("users.json");
    const u = users.find(x => x.id === req.params.id);
    if (!u) return res.status(404).json({ success: false, error: "User not found" });
    u.role = role;
    u.roleUpdatedAt = nowISO();
    writeJSON("users.json", users);
    if (role === "admin") {
      createNotification({
        type: "role_changed", title: "You are now an admin",
        message: "Admin access has been granted to your account.",
        link: "/admin.html", icon: "shield", userId: u.id, dedupeKey: "role-admin:" + u.id,
      });
    }
    res.json({ success: true, message: "Role updated to " + role });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to update role" });
  }
});

app.get("/api/admin/orders", requireAdmin, (req, res) => {
  const orders = readJSON("orders.json");
  res.json({ success: true, orders: orders.reverse(), total: orders.length });
});

app.get("/api/admin/contacts", requireAdmin, (req, res) => {
  const contacts = readJSON("contacts.json");
  res.json({ success: true, contacts: contacts.reverse(), total: contacts.length });
});

app.get("/api/admin/stats", requireAdmin, (req, res) => {
  const orders = readJSON("orders.json");
  const contacts = readJSON("contacts.json");
  const users = readJSON("users.json");
  const notifs = getNotifs();
  const today = new Date().toISOString().slice(0, 10);
  res.json({
    success: true,
    stats: {
      totalOrders: orders.length,
      todayOrders: orders.filter(o => o.createdAt.startsWith(today)).length,
      totalContacts: contacts.length,
      pendingOrders: orders.filter(o => o.status === "pending").length,
      totalUsers: users.length,
      pendingUsers: users.filter(u => u.status === "pending").length,
      totalNotifications: notifs.filter(n => !n.deleted).length,
      unreadNotifications: notifs.filter(n => !n.deleted && !n.read).length,
    },
  });
});

app.put("/api/admin/orders/:id/status", requireAdmin, (req, res) => {
  try {
    const { status } = req.body;
    const valid = ["pending", "confirmed", "in-progress", "delivered", "cancelled"];
    if (!valid.includes(status)) return res.status(400).json({ success: false, error: "Invalid status" });
    const orders = readJSON("orders.json");
    const idx = orders.findIndex(o => o.id === req.params.id);
    if (idx === -1) return res.status(404).json({ success: false, error: "Order not found" });
    orders[idx].status = status;
    orders[idx].updatedAt = new Date().toISOString();
    writeJSON("orders.json", orders);

    // Notify the customer when their order status changes.
    if (orders[idx].userId) {
      const statusLabels = { confirmed: "confirmed", "in-progress": "in progress", delivered: "delivered and ready", cancelled: "cancelled" };
      createNotification({
        type: "order_status",
        title: "Order " + status.replace("-", " "),
        message: "Your order is now " + (statusLabels[status] || status) + ". Thank you for choosing IS Digital Platform!",
        link: "#order",
        icon: status === "delivered" ? "check" : "order",
        userId: orders[idx].userId,
        dedupeKey: "order-status:" + orders[idx].id + ":" + status,
      });
    }

    res.json({ success: true, message: "Status updated" });
  } catch (err) {
    console.error("Status error:", err);
    res.status(500).json({ success: false, error: "Failed to update status" });
  }
});

app.delete("/api/admin/orders/:id", requireAdmin, (req, res) => {
  try {
    let orders = readJSON("orders.json");
    const len = orders.length;
    orders = orders.filter(o => o.id !== req.params.id);
    if (orders.length === len) return res.status(404).json({ success: false, error: "Order not found" });
    writeJSON("orders.json", orders);
    res.json({ success: true, message: "Order deleted" });
  } catch (err) {
    console.error("Delete order error:", err);
    res.status(500).json({ success: false, error: "Failed to delete" });
  }
});

app.delete("/api/admin/contacts/:id", requireAdmin, (req, res) => {
  try {
    let contacts = readJSON("contacts.json");
    const len = contacts.length;
    contacts = contacts.filter(c => c.id !== req.params.id);
    if (contacts.length === len) return res.status(404).json({ success: false, error: "Contact not found" });
    writeJSON("contacts.json", contacts);
    res.json({ success: true, message: "Contact deleted" });
  } catch (err) {
    console.error("Delete contact error:", err);
    res.status(500).json({ success: false, error: "Failed to delete" });
  }
});

const PLAN_MAP = {
  basic: { name: "Basic Plan", price: 249, items: ["One Digital Invitation", "JPG & PNG format", "One Theme", "One Revision", "Delivery within 24 Hours"] },
  premium: { name: "Premium Plan", price: 499, items: ["Custom Design", "Multiple Theme Options", "Unlimited Text Changes", "QR Code", "WhatsApp & Instagram Sizes", "Two Revisions"] },
  animated: { name: "Animated Invitation", price: 999, items: ["15-30 Second Video", "Background Music", "Animation", "HD Quality", "WhatsApp Ready"] },
  website: { name: "Luxury Website Invitation", price: 2999, items: ["Personal Invitation Website", "Countdown Timer", "RSVP Form", "Photo Gallery", "Google Maps", "Mobile Friendly"] },
};

function getStudioSettings() {
  const defaults = {
    name: "IS Digital Platform",
    owner: "Shubham Thakor",
    tagline: "Invitations, Carved in Light",
    address: "Ahmedabad, Gujarat, India",
    email: "hello@isdigitalplatform.com",
    phone: "+91 00000 00000",
    website: "isdigitalplatform.com",
  };
  const p = path.join(DATA_DIR, "settings.json");
  if (!fs.existsSync(p)) { writeJSON("settings.json", defaults); return defaults; }
  try { return { ...defaults, ...JSON.parse(fs.readFileSync(p, "utf8")) }; }
  catch(e) { return defaults; }
}
function STUDIO() { return getStudioSettings(); }

function findOrder(id) {
  const orders = readJSON("orders.json");
  return orders.find(o => o.id === id) || null;
}

function fmtDate(d) {
  return new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}

function drawSectionHeader(doc, y, title) {
  const x = 50;
  doc.roundedRect(x, y, 495, 28, 4).fill("#7A2440");
  doc.fontSize(11).fillColor("#F6EDE0").font("Helvetica-Bold").text(title, x + 12, y + 7, { width: 470 });
  return y + 36;
}

function drawFieldRow(doc, y, label, value) {
  const x = 55;
  doc.font("Helvetica-Bold").fontSize(9).fillColor("#7A2440").text(label + ":", x, y, { continued: true, width: 160 });
  doc.font("Helvetica").fontSize(9).fillColor("#333").text("  " + (value || "-"), { width: 310 });
  return y + 18;
}

app.get("/api/admin/orders/:id/pdf", requireAdmin, (req, res) => {
  try {
    const order = findOrder(req.params.id);
    if (!order) return res.status(404).json({ success: false, error: "Order not found" });
    const plan = PLAN_MAP[order.package] || { name: order.package, price: 0, items: [] };

    const doc = new PDFDocument({ size: "A4", margin: 50, bufferPages: true });
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="Order-${order.id.slice(0,8)}.pdf"`);
    doc.pipe(res);

    const w = doc.page.width;
    const right = w - 50;

    doc.rect(0, 0, w, 120).fill("#150B18");
    const studio = STUDIO();
    doc.fontSize(24).fillColor("#C89B3C").font("Helvetica-Bold").text(studio.name, 50, 28, { align: "center", width: w - 100 });
    doc.fontSize(10).fillColor("#E8C77E").font("Helvetica").text(studio.tagline, 50, 58, { align: "center", width: w - 100 });

    doc.rect(50, 90, w - 100, 1.5).fill("#C89B3C");

    doc.fontSize(18).fillColor("#C89B3C").font("Helvetica-Bold").text("ORDER CONFIRMATION", 50, 130, { align: "center", width: w - 100 });
    doc.fontSize(9).fillColor("#666").font("Helvetica").text(`Order ID: ${order.id}`, 50, 155, { align: "center", width: w - 100 });
    doc.text(`Date: ${fmtDate(order.createdAt)}   |   Status: ${(order.status || "pending").toUpperCase()}`, 50, 168, { align: "center", width: w - 100 });

    doc.rect(50, 192, w - 100, 1).fill("#C89B3C");

    let y = 210;
    y = drawSectionHeader(doc, y, "Event Details");
    y = drawFieldRow(doc, y, "Event Type", order.eventType);
    y = drawFieldRow(doc, y, "Event Date", order.eventDate);
    y = drawFieldRow(doc, y, "Event Time", order.eventTime || "Not specified");
    y = drawFieldRow(doc, y, "Host Name", order.hostName);
    y = drawFieldRow(doc, y, "Host Relation", order.hostRelation);
    y = drawFieldRow(doc, y, "Celebrant 1", order.name1);
    y = drawFieldRow(doc, y, "Celebrant 2", order.name2);
    y = drawFieldRow(doc, y, "Venue", order.venueName);
    y = drawFieldRow(doc, y, "City", order.venueCity);
    y = drawFieldRow(doc, y, "Address", order.venueAddress);
    y += 12;

    y = drawSectionHeader(doc, y, "Package & Pricing");
    y = drawFieldRow(doc, y, "Package", plan.name);
    y = drawFieldRow(doc, y, "Price", "Rs. " + plan.price);
    plan.items.forEach(item => {
      doc.font("Helvetica").fontSize(9).fillColor("#444").text("   \u2726  " + item, 68, y, { width: 400 });
      y += 16;
    });
    y += 6;

    y = drawSectionHeader(doc, y, "Payment");
    y = drawFieldRow(doc, y, "Method", (order.payment || "-").toUpperCase());
    y = drawFieldRow(doc, y, "Coupon", order.coupon || "None");
    y += 12;

    y = drawSectionHeader(doc, y, "Contact Information");
    y = drawFieldRow(doc, y, "Name", order.cname);
    y = drawFieldRow(doc, y, "Email", order.email);
    y = drawFieldRow(doc, y, "Mobile", order.mobile);
    y += 12;

    y = drawSectionHeader(doc, y, "Theme & Colours");
    y = drawFieldRow(doc, y, "Theme", order.theme);
    y = drawFieldRow(doc, y, "Font Style", order.fontStyle);
    y = drawFieldRow(doc, y, "Colours", order.colors || "Not selected");

    const photos = order.photos || [];
    const validPhotos = photos.filter(p => {
      if (p.startsWith("/api/uploads/")) return fs.existsSync(path.join(UPLOADS_DIR, path.basename(p)));
      return false;
    });
    if (validPhotos.length) {
      y += 6;
      y = drawSectionHeader(doc, y, "Uploaded Photos (" + validPhotos.length + ")");
      const photoSize = 120;
      const photosPerRow = 3;
      const startX = 55;
      let px = startX;
      for (let i = 0; i < Math.min(validPhotos.length, 6); i++) {
        const filePath = path.join(UPLOADS_DIR, path.basename(validPhotos[i]));
        try {
          doc.image(filePath, px, y, { width: photoSize, height: photoSize, fit: [photoSize, photoSize] });
        } catch (e) { /* skip unreadable images */ }
        px += photoSize + 15;
        if ((i + 1) % photosPerRow === 0) { px = startX; y += photoSize + 10; }
      }
      if (validPhotos.length % photosPerRow !== 0) y += photoSize + 10;
    }

    const footerY = doc.page.height - 50;
    doc.rect(0, footerY - 10, w, 60).fill("#150B18");
    doc.rect(50, footerY - 10, w - 100, 1).fill("#C89B3C");
    doc.fontSize(8).fillColor("#E8C77E").font("Helvetica").text(`${studio.name}  |  ${studio.address}`, 50, footerY, { align: "center", width: w - 100 });
    doc.text(`${studio.email}  |  ${studio.phone}  |  ${studio.website}`, 50, footerY + 12, { align: "center", width: w - 100 });
    doc.fontSize(7).fillColor("#888").text("Thank you for choosing IS Digital Platform!", 50, footerY + 26, { align: "center", width: w - 100 });

    doc.end();
  } catch (err) {
    console.error("PDF error:", err);
    res.status(500).json({ success: false, error: "Failed to generate PDF" });
  }
});

app.get("/api/admin/orders/:id/word", requireAdmin, async (req, res) => {
  try {
    const order = findOrder(req.params.id);
    if (!order) return res.status(404).json({ success: false, error: "Order not found" });
    const plan = PLAN_MAP[order.package] || { name: order.package, price: 0, items: [] };
    const studio = STUDIO();

    const borderDef = { style: BorderStyle.SINGLE, size: 1, color: "C89B3C" };
    const cellBorder = { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef };

    function headerRow(label) {
      return new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, color: "F6EDE0", size: 24, font: "Calibri" })] })], columnSpan: 2, shading: { type: ShadingType.SOLID, color: "7A2440" }, borders: cellBorder }),
        ],
      });
    }
    function dataRow(label, value) {
      return new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 21, font: "Calibri", color: "7A2440" })] })], width: { size: 30, type: WidthType.PERCENTAGE }, borders: cellBorder }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: value || "-", size: 21, font: "Calibri" })] })], width: { size: 70, type: WidthType.PERCENTAGE }, borders: cellBorder }),
        ],
      });
    }

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: studio.name, bold: true, size: 36, color: "C89B3C", font: "Calibri" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: studio.tagline, size: 20, color: "888888", font: "Calibri", italics: true })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "ORDER CONFIRMATION", bold: true, size: 40, color: "7A2440", font: "Calibri" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: "Order ID: " + order.id, size: 18, color: "555555", font: "Calibri" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: "Date: " + fmtDate(order.createdAt), size: 18, color: "555555", font: "Calibri" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "Status: " + (order.status || "pending").toUpperCase(), size: 20, color: "D9772A", bold: true, font: "Calibri" })] }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
            headerRow("Event Details"),
            dataRow("Event Type", order.eventType), dataRow("Event Date", order.eventDate),
            dataRow("Event Time", order.eventTime), dataRow("Host Name", order.hostName),
            dataRow("Host Relation", order.hostRelation), dataRow("Celebrant 1", order.name1),
            dataRow("Celebrant 2", order.name2), dataRow("Venue", order.venueName),
            dataRow("City", order.venueCity), dataRow("Address", order.venueAddress),
          ] }),
          new Paragraph({ spacing: { after: 200 } }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
            headerRow("Package & Pricing"),
            dataRow("Package", plan.name), dataRow("Price", "Rs. " + plan.price),
            ...plan.items.map(i => dataRow("  ", "\u2726 " + i)),
          ] }),
          new Paragraph({ spacing: { after: 200 } }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
            headerRow("Payment"),
            dataRow("Method", (order.payment || "-").toUpperCase()), dataRow("Coupon", order.coupon || "None"),
          ] }),
          new Paragraph({ spacing: { after: 200 } }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
            headerRow("Contact Info"),
            dataRow("Name", order.cname), dataRow("Email", order.email), dataRow("Mobile", order.mobile),
          ] }),
          new Paragraph({ spacing: { after: 200 } }),
          new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
            headerRow("Theme & Colours"),
            dataRow("Theme", order.theme), dataRow("Font Style", order.fontStyle), dataRow("Colours", order.colors || "Not selected"),
          ] }),
          new Paragraph({ spacing: { after: 200 } }),
          ...(function() {
            const photos = order.photos || [];
            const validPhotos = photos.filter(p => {
              if (p.startsWith("/api/uploads/")) return fs.existsSync(path.join(UPLOADS_DIR, path.basename(p)));
              return false;
            });
            if (!validPhotos.length) return [];
            const children = [
              new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
                headerRow("Uploaded Photos (" + validPhotos.length + ")"),
              ] }),
            ];
            const photoElements = [];
            for (let i = 0; i < Math.min(validPhotos.length, 6); i++) {
              try {
                const buf = fs.readFileSync(path.join(UPLOADS_DIR, path.basename(validPhotos[i])));
                photoElements.push(new Paragraph({ spacing: { before: 100 }, children: [
                  new ImageRun({ data: buf, transformation: { width: 280, height: 280 }, type: "jpg" }),
                ] }));
              } catch (e) { /* skip */ }
            }
            if (photoElements.length) {
              children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
                new TableRow({ children: photoElements.map(img => new TableCell({
                  children: [img],
                  borders: { top: borderDef, bottom: borderDef, left: borderDef, right: borderDef },
                })) }),
              ] }));
            }
            return children;
          })(),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${studio.name}  |  ${studio.address}  |  ${studio.email}  |  ${studio.phone}`, size: 16, color: "888888", font: "Calibri" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100 }, children: [new TextRun({ text: "Thank you for choosing IS Digital Platform!", size: 20, color: "C89B3C", bold: true, font: "Calibri" })] }),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="Order-${order.id.slice(0,8)}.docx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("Word error:", err);
    res.status(500).json({ success: false, error: "Failed to generate Word document" });
  }
});

app.get("/api/uploads/:filename", (req, res) => {
  const filePath = path.join(UPLOADS_DIR, req.params.filename);
  if (!fs.existsSync(filePath)) return res.status(404).json({ success: false, error: "File not found" });
  res.sendFile(filePath);
});

app.get("/api/admin/orders/:id/download-photo/:index", requireAdmin, (req, res) => {
  try {
    const order = findOrder(req.params.id);
    if (!order) return res.status(404).json({ success: false, error: "Order not found" });
    const photos = order.photos || [];
    const idx = parseInt(req.params.index, 10);
    if (idx < 0 || idx >= photos.length) return res.status(404).json({ success: false, error: "Photo not found" });
    const photoUrl = photos[idx];
    if (photoUrl.startsWith("/api/uploads/")) {
      const filePath = path.join(UPLOADS_DIR, path.basename(photoUrl));
      if (!fs.existsSync(filePath)) return res.status(404).json({ success: false, error: "File not found" });
      const ext = path.extname(filePath);
      res.setHeader("Content-Disposition", "attachment; filename=\"photo-" + (idx + 1) + ext + "\"");
      res.sendFile(filePath);
    } else {
      return res.status(404).json({ success: false, error: "Invalid photo path" });
    }
  } catch (err) {
    console.error("Download photo error:", err);
    res.status(500).json({ success: false, error: "Failed to download photo" });
  }
});

app.get("/api/admin/orders/:id/download-all-photos", requireAdmin, (req, res) => {
  try {
    const order = findOrder(req.params.id);
    if (!order) return res.status(404).json({ success: false, error: "Order not found" });
    const photos = order.photos || [];
    if (!photos.length) return res.status(404).json({ success: false, error: "No photos to download" });
    const available = photos.filter(p => {
      if (p.startsWith("/api/uploads/")) return fs.existsSync(path.join(UPLOADS_DIR, path.basename(p)));
      return false;
    });
    if (!available.length) return res.status(404).json({ success: false, error: "No photo files found on disk" });
    if (available.length === 1) {
      const filePath = path.join(UPLOADS_DIR, path.basename(available[0]));
      res.setHeader("Content-Disposition", "attachment; filename=\"photo-1" + path.extname(filePath) + "\"");
      return res.sendFile(filePath);
    }
    res.json({ success: true, photos: available.map((p, i) => ({
      url: "/api/admin/orders/" + order.id + "/download-photo/" + i,
      name: "photo-" + (i + 1) + path.extname(p),
    }))});
  } catch (err) {
    console.error("Download all photos error:", err);
    res.status(500).json({ success: false, error: "Failed to get photos" });
  }
});

// ===================== USER AUTH (Google Login Simulation) =====================

app.post("/api/auth/google", apiLimiter, (req, res) => {
  try {
    const { googleId, name, email, picture } = req.body;
    if (!googleId || !name || !email) {
      return res.status(400).json({ success: false, error: "Missing Google account info" });
    }
    const users = readJSON("users.json");
    const devSig = deviceSignature(req);
    let user = users.find(u => u.googleId === googleId);
    const isNew = !user;

    if (!user) {
      user = {
        id: crypto.randomUUID(),
        googleId,
        name,
        email,
        picture: picture || "",
        role: "user",
        phone: "",
        city: "",
        status: "pending",
        deviceHash: devSig,
        loginHistory: [],
        createdAt: new Date().toISOString(),
      };
      users.push(user);
    }

    // For returning users detect sign-in from a brand new device/network.
    let newDevice = false;
    if (!isNew) {
      const history = Array.isArray(user.loginHistory) ? user.loginHistory : [];
      newDevice = !user.deviceHash || user.deviceHash !== devSig;
      user.deviceHash = devSig;
      history.push({ deviceHash: devSig, ip: req.ip, at: nowISO() });
      user.loginHistory = history.slice(-20);
      user.lastLoginAt = nowISO();
    }
    writeJSON("users.json", users);

    if (isNew) {
      // New user welcome + notify all admins of a pending registration.
      createNotification({
        type: "user_registered",
        title: "Welcome to IS Digital Platform!",
        message: "Thanks for joining. Your account is pending admin approval — you'll be notified once it's live.",
        link: "#notifications",
        icon: "user",
        userId: user.id,
        dedupeKey: "welcome:" + user.id,
      });
      notifyAdmins({
        type: "user_registered",
        title: "New user registration",
        message: name + " (" + email + ") signed up and is waiting for approval.",
        link: "/admin.html?section=users",
        icon: "user",
        dedupeKey: "reg:" + user.id,
      });
    } else if (newDevice) {
      createNotification({
        type: "new_login",
        title: "New device sign-in detected",
        message: "Your account was signed in from a new device (" + (req.headers["user-agent"] || "Unknown device") + ").",
        link: "#notifications",
        icon: "shield",
        userId: user.id,
        dedupeKey: "login:" + user.id + ":" + devSig,
      });
    }

    res.json({
      success: true,
      user: { id: user.id, name: user.name, email: user.email, picture: user.picture, status: user.status, role: user.role || "user", phone: user.phone || "", city: user.city || "" },
    });
  } catch (err) {
    console.error("Auth error:", err);
    res.status(500).json({ success: false, error: "Authentication failed" });
  }
});

app.get("/api/auth/user", apiLimiter, (req, res) => {
  try {
    const userId = req.query.userId;
    if (!userId) return res.status(400).json({ success: false, error: "Missing userId" });
    const users = readJSON("users.json");
    const user = users.find(u => u.id === userId);
    if (!user) return res.status(404).json({ success: false, error: "User not found" });
    res.json({
      success: true,
      user: { id: user.id, name: user.name, email: user.email, picture: user.picture, status: user.status, role: user.role || "user", phone: user.phone || "", city: user.city || "" },
    });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed" });
  }
});

// ===================== ADMIN: USER MANAGEMENT =====================

app.get("/api/admin/users", requireAdmin, (req, res) => {
  const users = readJSON("users.json");
  res.json({ success: true, users: users.reverse(), total: users.length });
});

app.put("/api/admin/users/:id/approve", requireAdmin, (req, res) => {
  try {
    const users = readJSON("users.json");
    const idx = users.findIndex(u => u.id === req.params.id);
    if (idx === -1) return res.status(404).json({ success: false, error: "User not found" });
    users[idx].status = "approved";
    users[idx].approvedAt = new Date().toISOString();
    writeJSON("users.json", users);
    createNotification({
      type: "account_approved",
      title: "Account approved",
      message: "Congratulations! Your account is approved and you can now place orders.",
      link: "#order",
      icon: "check",
      userId: users[idx].id,
      dedupeKey: "approve:" + users[idx].id,
    });
    res.json({ success: true, message: "User approved" });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to approve" });
  }
});

app.put("/api/admin/users/:id/reject", requireAdmin, (req, res) => {
  try {
    const users = readJSON("users.json");
    const idx = users.findIndex(u => u.id === req.params.id);
    if (idx === -1) return res.status(404).json({ success: false, error: "User not found" });
    users[idx].status = "rejected";
    users[idx].rejectedAt = new Date().toISOString();
    writeJSON("users.json", users);
    createNotification({
      type: "account_rejected",
      title: "Account rejected",
      message: "Your account was rejected. Please contact support for assistance.",
      link: "#contact",
      icon: "alert",
      userId: users[idx].id,
      dedupeKey: "reject:" + users[idx].id,
    });
    res.json({ success: true, message: "User rejected" });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to reject" });
  }
});

app.delete("/api/admin/users/:id", requireAdmin, (req, res) => {
  try {
    let users = readJSON("users.json");
    const len = users.length;
    users = users.filter(u => u.id !== req.params.id);
    if (users.length === len) return res.status(404).json({ success: false, error: "User not found" });
    writeJSON("users.json", users);
    res.json({ success: true, message: "User deleted" });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to delete" });
  }
});

// ===================== ADMIN: STUDIO SETTINGS =====================

app.get("/api/settings", (req, res) => {
  const studio = getStudioSettings();
  res.json({ success: true, name: studio.name, owner: studio.owner, tagline: studio.tagline });
});

app.get("/api/admin/settings", requireAdmin, (req, res) => {
  const studio = getStudioSettings();
  res.json({ success: true, settings: studio });
});

app.put("/api/admin/settings", requireAdmin, (req, res) => {
  try {
    const current = getStudioSettings();
    const updates = req.body;
    const allowed = ["name", "owner", "tagline", "address", "email", "phone", "website"];
    allowed.forEach(k => { if (updates[k] !== undefined) current[k] = updates[k]; });
    writeJSON("settings.json", current);
    res.json({ success: true, message: "Settings updated", settings: current });
  } catch (err) {
    res.status(500).json({ success: false, error: "Failed to update settings" });
  }
});

app.use(express.static(PUBLIC_DIR, {
  maxAge: "7d",
  etag: true,
  lastModified: true,
  setHeaders(res, filePath) {
    if (filePath.endsWith(".html")) res.setHeader("Cache-Control", "no-cache");
  },
}));

app.get("/admin", (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, "admin.html"));
});

app.get("*", (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`IS Digital Platform running on http://localhost:${PORT}`);
  console.log(`Admin panel: http://localhost:${PORT}/admin?key=${ADMIN_KEY}`);
});
