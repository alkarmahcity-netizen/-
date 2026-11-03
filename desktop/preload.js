// Local replacement for the claude.ai page runtime: a JSON file database,
// the Windows user name as identity, and a Save dialog for Excel reports.
const fs = require("fs");
const path = require("path");
const os = require("os");
const { ipcRenderer } = require("electron");

const FILE = ipcRenderer.sendSync("cfg").dataFile;
const clone = (x) => (x == null ? x : JSON.parse(JSON.stringify(x)));
const empty = () => ({ v: 1, c: {} });

function load() {
  try { const d = JSON.parse(fs.readFileSync(FILE, "utf8")); return d && d.c ? d : empty(); }
  catch (_) { return empty(); }
}
let data = load();
let mtime = 0;
try { mtime = fs.statSync(FILE).mtimeMs; } catch (_) {}

function backup() {
  try {
    const dir = path.join(path.dirname(FILE), "backups");
    const f = path.join(dir, "anbar-data-" + new Date().toISOString().slice(0, 10) + ".json");
    if (!fs.existsSync(f) && fs.existsSync(FILE)) { fs.mkdirSync(dir, { recursive: true }); fs.copyFileSync(FILE, f); }
  } catch (_) {}
}

// Re-read the file before every write so two computers sharing one network file do not overwrite each other's changes.
function mutate(fn) {
  backup();
  const cur = load();
  fn(cur);
  fs.mkdirSync(path.dirname(FILE), { recursive: true });
  const tmp = FILE + "." + process.pid + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(cur));
  try { fs.renameSync(tmp, FILE); } catch (e) { fs.copyFileSync(tmp, FILE); fs.unlinkSync(tmp); }
  data = cur;
  try { mtime = fs.statSync(FILE).mtimeMs; } catch (_) {}
  notify();
}

const subs = new Set();
function notify() { for (const s of subs) setTimeout(() => s.fire(), 0); }
setInterval(() => {
  try { const m = fs.statSync(FILE).mtimeMs; if (m !== mtime) { mtime = m; data = load(); notify(); } } catch (_) {}
}, 2000);

const coll = (name) => (data.c[name] = data.c[name] || {});
const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

function query(name, ord, lim) {
  const q = {
    orderBy: (f, dir) => query(name, { f, dir: dir || "asc" }, lim),
    limit: (n) => query(name, ord, n),
    where: () => q,
    get: async () => snap(),
    onSnapshot(next, err) {
      const s = { fire: () => { try { next(snap()); } catch (e) { if (err) err({ code: "unavailable", message: String(e) }); } } };
      subs.add(s); setTimeout(s.fire, 0); return () => subs.delete(s);
    },
    doc: (id) => docRef(name, id || newId()),
    async add(d) { const id = newId(); mutate((cur) => { (cur.c[name] = cur.c[name] || {})[id] = clone(d); }); return docRef(name, id); },
  };
  function snap() {
    let rows = Object.entries(data.c[name] || {}).map(([id, d]) => ({ id, d }));
    if (ord) rows.sort((a, b) => { const x = String(a.d[ord.f] ?? ""), y = String(b.d[ord.f] ?? ""); return ord.dir === "desc" ? y.localeCompare(x) : x.localeCompare(y); });
    if (lim) rows = rows.slice(0, lim);
    const docs = rows.map((r) => ({ id: r.id, exists: true, data: () => clone(r.d), metadata: { fromCache: false, hasPendingWrites: false } }));
    return { docs, size: docs.length, empty: !docs.length, docChanges: () => [], metadata: { fromCache: false, hasPendingWrites: false } };
  }
  return q;
}
function docRef(name, id) {
  const snapDoc = () => { const d = (data.c[name] || {})[id]; return { id, exists: !!d, data: () => clone(d), metadata: { fromCache: false, hasPendingWrites: false } }; };
  return {
    id, path: name + "/" + id,
    get: async () => snapDoc(),
    async set(d) { mutate((cur) => { (cur.c[name] = cur.c[name] || {})[id] = clone(d); }); },
    async update(d) { mutate((cur) => { const c = (cur.c[name] = cur.c[name] || {}); c[id] = { ...(c[id] || {}), ...clone(d) }; }); },
    async delete() { mutate((cur) => { if (cur.c[name]) delete cur.c[name][id]; }); },
    onSnapshot(next) { const s = { fire: () => next(snapDoc()) }; subs.add(s); setTimeout(s.fire, 0); return () => subs.delete(s); },
    collection: (sub) => query(name + "/" + id + "/" + sub),
  };
}

const userName = (() => { try { return os.userInfo().username; } catch (_) { return "مستخدم"; } })();
const host = os.hostname();
const uid = "local:" + userName + "@" + host;

const caps = {
  db: { collection: (n) => query(n), doc: (p) => { const i = p.lastIndexOf("/"); return docRef(p.slice(0, i), p.slice(i + 1)); } },
  user: {
    id: async () => uid,
    me: async () => ({ id: uid, name: userName }),
    can: async () => true, canEdit: () => true, isOwner: () => true,
    profiles: async (ids) => Object.fromEntries(ids.map((i) => [i, { id: i, name: String(i).replace(/^local:/, "").replace(/@.*/, "") + (String(i).includes("@") ? " (" + String(i).split("@")[1] + ")" : "") }])),
  },
  downloads: {
    async save({ filename, data }) {
      const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
      const r = await ipcRenderer.invoke("save", { filename, data: bytes });
      if (r.status !== "saved") { const e = new Error("declined"); e.code = "declined"; throw e; }
      return r;
    },
  },
};
window.claude = { use: async (n) => caps[n] || null };
window.ANBAR_DESKTOP = { dataFile: FILE };
