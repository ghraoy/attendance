// Google Drive sync: one JSON file (and optional Google Sheets copies) inside a "سجل حضور الماجستير" folder.
// Uses the drive.file scope: the app can only see files it created itself.
window.Drive = (() => {
  const SCOPE = "https://www.googleapis.com/auth/drive.file";
  const FOLDER = "سجل حضور الماجستير";
  const FILE = "سجل-حضور-الماجستير.json";
  let token = null, expires = 0, client = null, pending = null;

  function loadGis() {
    if (window.google?.accounts?.oauth2) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement("script"); s.src = "https://accounts.google.com/gsi/client"; s.async = true;
      s.onload = res; s.onerror = () => rej(new Error("تعذّر تحميل خدمة Google — تحقق من الإنترنت"));
      document.head.appendChild(s);
    });
  }

  async function signIn(clientId, interactive = true) {
    if (token && Date.now() < expires - 60000) return token;
    if (!clientId) throw new Error("أدخل Client ID من Google Cloud في الإعدادات أولاً");
    await loadGis();
    if (!client || client._id !== clientId) {
      client = google.accounts.oauth2.initTokenClient({
        client_id: clientId, scope: SCOPE,
        callback: r => { if (!pending) return; if (r.error) pending.rej(new Error(r.error_description || r.error)); else { token = r.access_token; expires = Date.now() + (r.expires_in || 3600) * 1000; pending.res(token); } pending = null; },
        error_callback: e => { if (pending) { pending.rej(new Error(e.type === "popup_closed" ? "أُغلقت نافذة تسجيل الدخول" : (e.message || "تعذّر تسجيل الدخول"))); pending = null; } }
      });
      client._id = clientId;
    }
    return new Promise((res, rej) => { pending = { res, rej }; client.requestAccessToken({ prompt: interactive ? "" : "none" }); });
  }
  const connected = () => !!token && Date.now() < expires - 60000;
  function signOut() { if (token && window.google?.accounts?.oauth2) google.accounts.oauth2.revoke(token, () => {}); token = null; expires = 0; }

  async function api(path, opts = {}) {
    const r = await fetch(path.startsWith("http") ? path : "https://www.googleapis.com/drive/v3/" + path,
      { ...opts, headers: { Authorization: "Bearer " + token, ...(opts.headers || {}) } });
    if (r.status === 401) { token = null; throw new Error("انتهت جلسة Google — اضغط مزامنة مرة أخرى"); }
    if (!r.ok) { let m = r.status; try { m = (await r.json()).error.message; } catch (e) {} throw new Error("Drive: " + m); }
    return r;
  }
  const q = s => encodeURIComponent(s);

  async function folderId(cfg) {
    if (cfg.folderId) {
      try { const j = await (await api(`files/${cfg.folderId}?fields=id,trashed`)).json(); if (!j.trashed) return cfg.folderId; } catch (e) {}
    }
    const found = await (await api(`files?q=${q(`name='${FOLDER}' and mimeType='application/vnd.google-apps.folder' and trashed=false`)}&fields=files(id)`)).json();
    if (found.files.length) return cfg.folderId = found.files[0].id;
    const made = await (await api("files?fields=id", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: FOLDER, mimeType: "application/vnd.google-apps.folder" }) })).json();
    return cfg.folderId = made.id;
  }

  async function findFile(cfg) {
    const fid = await folderId(cfg);
    const j = await (await api(`files?q=${q(`name='${FILE}' and '${fid}' in parents and trashed=false`)}&fields=files(id,modifiedTime)&orderBy=modifiedTime desc`)).json();
    return j.files[0] || null;
  }

  function multipart(meta, body, type) {
    const b = "mtu" + Math.random().toString(36).slice(2);
    const blob = new Blob([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${b}\r\nContent-Type: ${type}\r\n\r\n`, body, `\r\n--${b}--`]);
    return { body: blob, headers: { "Content-Type": "multipart/related; boundary=" + b } };
  }

  async function pull(cfg) {
    const f = await findFile(cfg); if (!f) return null;
    cfg.fileId = f.id;
    return (await api(`files/${f.id}?alt=media`)).json();
  }
  async function push(cfg, state) {
    const body = JSON.stringify(state);
    if (!cfg.fileId) { const f = await findFile(cfg); if (f) cfg.fileId = f.id; }
    if (cfg.fileId) {
      await api(`https://www.googleapis.com/upload/drive/v3/files/${cfg.fileId}?uploadType=media`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body });
    } else {
      const fid = await folderId(cfg);
      const m = multipart({ name: FILE, parents: [fid], mimeType: "application/json" }, body, "application/json");
      const j = await (await api("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", { method: "POST", ...m })).json();
      cfg.fileId = j.id;
    }
  }
  // Upload an .xlsx and let Drive convert it to a Google Sheet.
  async function uploadSheet(cfg, name, arrayBuffer) {
    const fid = await folderId(cfg);
    const m = multipart({ name, parents: [fid], mimeType: "application/vnd.google-apps.spreadsheet" }, new Blob([arrayBuffer]), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const j = await (await api("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink", { method: "POST", ...m })).json();
    return j.webViewLink;
  }
  return { signIn, signOut, connected, pull, push, uploadSheet, FOLDER };
})();
