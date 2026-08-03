/* ============================================================
   IS Digital Platform — Notification Center module
   ------------------------------------------------------------
   Self-contained vanilla-JS module that powers:
     - the notification bell + unread badge in the navbar
     - the Notifications page (search / filter / mark read /
       mark all / delete / navigate-on-click)
     - the logged-in user dropdown (Notifications, My Profile)
     - the profile edit modal
     - Firebase Cloud Messaging (web push) registration + handling
   It talks to the Express JSON-store backend via the REST APIs.
   It is loaded AFTER the main inline script and reuses the
   window.showToast / window.navigateTo hooks exposed there.
   ============================================================ */
(function () {
  "use strict";

  // ---- Config -------------------------------------------------
  // Set your Firebase Cloud Messaging web-push VAPID key here
  // (Firebase Console > Project settings > Cloud Messaging) to
  // enable web push. Leave empty to keep in-app notifications only.
  var FCM_VAPID_KEY = "";

  var POLL_INTERVAL_MS = 30000; // badge refresh cadence
  var STORE_KEY = "ts-user";

  var state = {
    user: null,
    filter: "all",
    query: "",
    total: 0,
    loading: false,
    pollTimer: null,
  };

  // ---- Small DOM / storage helpers ----------------------------
  function $(id) { return document.getElementById(id); }
  function esc(str) {
    return String(str == null ? "" : str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function getUser() {
    try { return JSON.parse(localStorage.getItem(STORE_KEY)); } catch (e) { return null; }
  }

  // Relative timestamp: "Just now", "5 minutes ago", "Yesterday", ...
  function timeAgo(iso) {
    if (!iso) return "";
    var diff = Date.now() - new Date(iso).getTime();
    if (diff < 0) diff = 0;
    var s = Math.floor(diff / 1000);
    if (s < 60) return "Just now";
    var m = Math.floor(s / 60);
    if (m < 60) return m + " minute" + (m > 1 ? "s" : "") + " ago";
    var h = Math.floor(m / 60);
    if (h < 24) return h + " hour" + (h > 1 ? "s" : "") + " ago";
    var d = Math.floor(h / 24);
    if (d === 1) return "Yesterday";
    if (d < 7) return d + " days ago";
    if (d < 30) return Math.floor(d / 7) + " week" + (Math.floor(d / 7) > 1 ? "s" : "") + " ago";
    var mo = Math.floor(d / 30);
    if (mo < 12) return mo + " month" + (mo > 1 ? "s" : "") + " ago";
    return new Date(iso).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  }

  // ---- Notification type -> icon -------------------------------
  var ICONS = {
    bell: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
    order: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2h12l2 6-8 14L4 8l2-6z"/><path d="M4 8h16"/><path d="M12 22V8"/></svg>',
    user: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
    shield: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8 3v6c0 5-3.5 8.5-8 11-4.5-2.5-8-6-8-11V5z"/><path d="M9 12l2 2 4-4"/></svg>',
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="8 12 11 15 16 9"/></svg>',
    alert: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9L1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    mail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2z"/><polyline points="22 6 12 13 2 6"/></svg>',
    megaphone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 11v3a1 1 0 0 0 1 1h2l3 4V4L6 8H4a1 1 0 0 0-1 3z"/><path d="M14 6a5 5 0 0 1 0 12"/><path d="M18 3.5a10 10 0 0 1 0 17"/></svg>',
    star: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>',
  };
  function iconFor(type) { return ICONS[type] || ICONS.bell; }

  // ---- Navigation ----------------------------------------------
  function goToLink(link) {
    if (!link) return;
    if (link.indexOf("/admin") === 0) { window.location.href = link; return; }
    if (link.charAt(0) === "#") {
      var pageId = link.slice(1);
      if (typeof window.navigateTo === "function") { window.navigateTo(pageId); }
      else { window.location.hash = pageId; }
    } else {
      window.location.href = link;
    }
  }

  // ---- API helpers ---------------------------------------------
  function apiGet(url) {
    var user = getUser();
    var headers = user && user.id ? { "x-user-id": user.id } : {};
    return fetch(url, { headers: headers }).then(function (r) { return r.json(); });
  }
  function apiSend(url, method, body) {
    var user = getUser();
    var headers = { "Content-Type": "application/json" };
    if (user && user.id) headers["x-user-id"] = user.id;
    return fetch(url, { method: method, headers: headers, body: JSON.stringify(body || {}) })
      .then(function (r) { return r.json(); });
  }
  function toast(msg, type) {
    if (typeof window.showToast === "function") { window.showToast(msg, type); return; }
    console.log(msg);
  }

  // ---- Badge + bell visibility ---------------------------------
  function refreshBadge() {
    var bell = $("notifBell");
    var badge = $("notifBadge");
    var user = getUser();
    if (!bell || !badge) return;
    if (!user || !user.id) { bell.hidden = true; badge.hidden = true; return; }
    bell.hidden = false;
    apiGet("/api/notifications/unread-count").then(function (d) {
      var n = (d && d.success) ? d.unreadCount : 0;
      badge.hidden = n === 0;
      badge.textContent = n > 99 ? "99+" : String(n);
      var dropdown = $("userDropdownUnread");
      if (dropdown) { dropdown.hidden = n === 0; dropdown.textContent = n > 99 ? "99+" : String(n); }
    }).catch(function () { badge.hidden = true; });
  }

  // ---- Notification list rendering -----------------------------
  function loadNotifications() {
    var listEl = $("notifList");
    var emptyEl = $("notifEmpty");
    var signinEl = $("notifSignin");
    var user = getUser();

    if (!listEl) return;
    if (!user || !user.id) {
      if (signinEl) signinEl.hidden = false;
      if (listEl) listEl.innerHTML = "";
      if (emptyEl) emptyEl.hidden = true;
      return;
    }
    if (signinEl) signinEl.hidden = true;
    if (state.loading) return;
    state.loading = true;
    listEl.classList.add("is-loading");

    var url = "/api/notifications?filter=" + encodeURIComponent(state.filter) +
      (state.query ? "&q=" + encodeURIComponent(state.query) : "") + "&limit=300";

    apiGet(url).then(function (d) {
      var items = (d && d.notifications) || [];
      state.total = (d && d.total) || items.length;
      listEl.innerHTML = items.map(renderCard).join("");
      emptyEl.hidden = items.length > 0;
      refreshBadge();
    }).catch(function () {
      listEl.innerHTML = '<div class="notif-error">Could not load notifications. Please try again.</div>';
    }).finally(function () {
      state.loading = false;
      listEl.classList.remove("is-loading");
    });
  }

  function renderCard(n) {
    var unread = !n.read;
    return '<div class="notif-card' + (unread ? " unread" : "") + '" data-id="' + esc(n.id) + '" data-link="' + esc(n.link || "") + '" role="button" tabindex="0" aria-label="' + esc(n.title) + '">'
      + '<div class="notif-icon">' + iconFor(n.type) + '</div>'
      + '<div class="notif-body">'
      + '<div class="notif-head"><span class="notif-title">' + esc(n.title) + '</span><span class="notif-time">' + timeAgo(n.createdAt) + '</span></div>'
      + '<p class="notif-msg">' + esc(n.message) + '</p>'
      + '</div>'
      + '<div class="notif-actions">'
      + (unread ? '<button class="notif-act" data-act="read" title="Mark as read">' + ICONS.check + '</button>' : '')
      + '<button class="notif-act" data-act="delete" title="Delete"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg></button>'
      + '</div>'
      + '</div>';
  }

  function markRead(id) {
    return apiSend("/api/notifications/" + encodeURIComponent(id) + "/read", "PUT").then(function (d) {
      if (d && d.success) { loadNotifications(); }
      return d;
    });
  }
  function markAllRead() {
    return apiSend("/api/notifications/read-all", "PUT").then(function (d) {
      if (d && d.success) { toast("All notifications marked as read"); loadNotifications(); }
    });
  }
  function deleteNotif(id) {
    return apiSend("/api/notifications/" + encodeURIComponent(id), "DELETE").then(function (d) {
      if (d && d.success) { toast("Notification deleted"); loadNotifications(); }
    });
  }

  function handleListClick(e) {
    var actBtn = e.target.closest(".notif-act");
    if (actBtn) {
      e.stopPropagation();
      var card = actBtn.closest(".notif-card");
      if (!card) return;
      var id = card.dataset.id;
      if (actBtn.dataset.act === "delete") deleteNotif(id);
      else markRead(id);
      return;
    }
    var card = e.target.closest(".notif-card");
    if (!card) return;
    var link = card.dataset.link || "";
    if (!card.classList.contains("unread")) { goToLink(link); return; }
    markRead(card.dataset.id).then(function (d) {
      if (d && d.success) goToLink(link);
    });
  }
  function handleListKey(e) {
    if (e.key !== "Enter") return;
    var card = e.target.closest(".notif-card");
    if (!card) return;
    var link = card.dataset.link || "";
    if (!card.classList.contains("unread")) { goToLink(link); return; }
    markRead(card.dataset.id).then(function (d) { if (d && d.success) goToLink(link); });
  }

  // ---- Profile modal -------------------------------------------
  function openProfile() {
    var overlay = $("profileModal");
    if (!overlay) return;
    var user = getUser();
    if (!user || !user.id) { toast("Please sign in first", "error"); return; }
    overlay.classList.add("open");
    $("profileName").value = user.name || "";
    $("profileEmail").value = user.email || "";
    $("profilePhone").value = user.phone || "";
    $("profileCity").value = user.city || "";
    $("profileRole").textContent = (user.role === "admin" ? "Admin" : "User");
  }
  function closeProfile() {
    var overlay = $("profileModal");
    if (overlay) overlay.classList.remove("open");
  }
  function saveProfile() {
    var btn = $("profileSave");
    var name = ($("profileName").value || "").trim();
    var phone = ($("profilePhone").value || "").trim();
    var city = ($("profileCity").value || "").trim();
    if (!name) { toast("Name cannot be empty", "error"); return; }
    var orig = btn.textContent;
    btn.textContent = "Saving...";
    btn.disabled = true;
    apiSend("/api/users/me", "PUT", { name: name, phone: phone, city: city }).then(function (d) {
      if (d && d.success) {
        var user = getUser();
        if (user) {
          user.name = d.user.name;
          user.phone = d.user.phone;
          user.city = d.user.city;
          try { localStorage.setItem(STORE_KEY, JSON.stringify(user)); } catch (e) {}
          if (typeof window.updateGoogleBtn === "function") window.updateGoogleBtn();
        }
        toast("Profile updated");
        closeProfile();
        refreshBadge();
      } else {
        toast((d && d.error) || "Failed to update profile", "error");
      }
    }).catch(function () { toast("Network error. Please try again.", "error"); })
      .finally(function () { btn.textContent = orig; btn.disabled = false; });
  }

  // ---- User dropdown -------------------------------------------
  function toggleUserMenu() {
    var menu = $("userDropdown");
    if (!menu) return;
    var open = menu.classList.contains("open");
    if (open) { closeUserMenu(); return; }
    var user = getUser();
    var head = $("userDropdownHead");
    if (head && user) {
      head.innerHTML = (user.picture ? '<img class="google-user-avatar" src="' + esc(user.picture) + '" alt="">' : '')
        + '<div class="user-dropdown-id"><strong>' + esc(user.name || "") + '</strong><span>' + esc(user.email || "") + '</span></div>';
    }
    menu.classList.add("open");
    document.addEventListener("click", outsideClick, true);
  }
  function closeUserMenu() {
    var menu = $("userDropdown");
    if (menu) menu.classList.remove("open");
    document.removeEventListener("click", outsideClick, true);
  }
  function outsideClick(e) {
    var wrap = $("userMenuWrap");
    if (wrap && !wrap.contains(e.target)) closeUserMenu();
  }
  function handleDropdownClick(e) {
    var item = e.target.closest("[data-user-menu]");
    if (!item) return;
    closeUserMenu();
    var action = item.dataset.userMenu;
    if (action === "notifications") { goToLink("#notifications"); }
    else if (action === "profile") { openProfile(); }
  }

  // ---- Firebase Cloud Messaging (web push) ----------------------
  function registerServiceWorker() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("/firebase-messaging-sw.js")
      .catch(function () { /* non-fatal: push just won't arrive */ });
  }

  function setupFcm() {
    if (!FCM_VAPID_KEY) return; // VAPID key not configured yet
    if (!("Notification" in window)) return;
    if (Notification.permission === "denied") return;
    if (typeof firebase === "undefined" || !firebase.messaging) return;

    var messaging = firebase.messaging();
    messaging.getToken({ vapidKey: FCM_VAPID_KEY }).then(function (token) {
      var user = getUser();
      if (token && user && user.id) {
        apiSend("/api/notifications/fcm-token", "POST", { token: token })
          .catch(function () {});
      }
    }).catch(function () { /* permission deferred / unavailable */ });

    messaging.onMessage(function (payload) {
      var n = payload.notification || {};
      if (n.title) toast(n.title);
      refreshBadge();
    });
  }

  // ---- Lifecycle ------------------------------------------------
  function sync() {
    refreshBadge();
    var page = $("notifications");
    if (page && page.classList.contains("active-page")) loadNotifications();
  }

  function init() {
    var bell = $("notifBell");
    var search = $("notifSearch");
    var list = $("notifList");

    if (bell) bell.addEventListener("click", function () { goToLink("#notifications"); });

    if (search) {
      var debounce = null;
      search.addEventListener("input", function () {
        clearTimeout(debounce);
        debounce = setTimeout(function () { state.query = search.value.trim(); loadNotifications(); }, 300);
      });
    }

    var filters = document.querySelectorAll(".notif-filter");
    filters.forEach(function (btn) {
      btn.addEventListener("click", function () {
        filters.forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        state.filter = btn.dataset.filter;
        loadNotifications();
      });
    });

    if ($("notifMarkAll")) $("notifMarkAll").addEventListener("click", markAllRead);
    if (list) { list.addEventListener("click", handleListClick); list.addEventListener("keydown", handleListKey); }

    // User dropdown
    if ($("userDropdown")) $("userDropdown").addEventListener("click", handleDropdownClick);

    // Profile modal
    if ($("profileSave")) $("profileSave").addEventListener("click", saveProfile);
    if ($("profileClose")) $("profileClose").addEventListener("click", closeProfile);
    var pm = $("profileModal");
    if (pm) pm.addEventListener("click", function (e) { if (e.target === pm) closeProfile(); });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") { closeUserMenu(); closeProfile(); }
    });

    // Poll for new notifications while signed in
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(refreshBadge, POLL_INTERVAL_MS);

    // React to auth changes fired by the main script
    window.addEventListener("ts-auth-changed", function () { sync(); });

    // Kick off
    registerServiceWorker();
    setupFcm();
    sync();
  }

  // Public API used by the main inline script.
  window.Notifications = {
    init: init,
    sync: sync,
    refresh: refreshBadge,
    toggleUserMenu: toggleUserMenu,
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
