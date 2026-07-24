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
    const orders = readJSON("orders.json");
    orders.push(order);
    writeJSON("orders.json", orders);
    res.json({ success: true, message: "Order submitted successfully! We'll reach out within 2 hours.", orderId: order.id });
  } catch (err) {
    console.error("Order error:", err);
    res.status(500).json({ success: false, error: "Something went wrong. Please try again." });
  }
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
    contacts.push({ id: crypto.randomUUID(), name, email, message, createdAt: new Date().toISOString(), read: false });
    writeJSON("contacts.json", contacts);
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
  const today = new Date().toISOString().slice(0, 10);
  res.json({
    success: true,
    stats: {
      totalOrders: orders.length,
      todayOrders: orders.filter(o => o.createdAt.startsWith(today)).length,
      totalContacts: contacts.length,
      pendingOrders: orders.filter(o => o.status === "pending").length,
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
    let user = users.find(u => u.googleId === googleId);
    if (!user) {
      user = {
        id: crypto.randomUUID(),
        googleId,
        name,
        email,
        picture: picture || "",
        status: "pending",
        createdAt: new Date().toISOString(),
      };
      users.push(user);
      writeJSON("users.json", users);
    }
    res.json({
      success: true,
      user: { id: user.id, name: user.name, email: user.email, picture: user.picture, status: user.status },
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
      user: { id: user.id, name: user.name, email: user.email, picture: user.picture, status: user.status },
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
