const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require("electron");
const path = require("path");
const fs = require("fs");

const DATA_NAME = "anbar-data.json";
const cfgPath = () => path.join(app.getPath("userData"), "config.json");
function readCfg() {
  try { return JSON.parse(fs.readFileSync(cfgPath(), "utf8")); } catch (_) { return {}; }
}
function writeCfg(c) {
  fs.mkdirSync(path.dirname(cfgPath()), { recursive: true });
  fs.writeFileSync(cfgPath(), JSON.stringify(c, null, 2));
}
function dataFile() {
  const c = readCfg();
  return c.dataFile || path.join(app.getPath("documents"), "AnbarFinance", DATA_NAME);
}

let win;
if (!app.requestSingleInstanceLock()) app.quit();
app.on("second-instance", () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });

ipcMain.on("cfg", (e) => { e.returnValue = { dataFile: dataFile() }; });

ipcMain.handle("save", async (_e, { filename, data }) => {
  const r = await dialog.showSaveDialog(win, {
    defaultPath: path.join(app.getPath("documents"), filename),
    filters: [{ name: "Excel", extensions: ["xlsx"] }, { name: "All", extensions: ["*"] }],
  });
  if (r.canceled || !r.filePath) return { status: "declined" };
  fs.writeFileSync(r.filePath, Buffer.from(data));
  shell.showItemInFolder(r.filePath);
  return { status: "saved" };
});

async function chooseDataFolder() {
  const r = await dialog.showOpenDialog(win, {
    title: "اختر مجلد ملف البيانات (يمكن أن يكون مجلداً مشتركاً على الشبكة)",
    properties: ["openDirectory", "createDirectory"],
  });
  if (r.canceled || !r.filePaths[0]) return;
  const target = path.join(r.filePaths[0], DATA_NAME);
  const cur = dataFile();
  if (!fs.existsSync(target) && fs.existsSync(cur)) {
    const ans = await dialog.showMessageBox(win, {
      type: "question", buttons: ["نسخ البيانات الحالية إليه", "بدء ملف فارغ", "إلغاء"], defaultId: 0, cancelId: 2,
      message: "لا يوجد ملف بيانات في هذا المجلد.", detail: "هل تريد نسخ البيانات الحالية إليه؟",
    });
    if (ans.response === 2) return;
    if (ans.response === 0) fs.copyFileSync(cur, target);
  }
  writeCfg({ ...readCfg(), dataFile: target });
  win.reload();
}

function buildMenu() {
  const tpl = [
    {
      label: "ملف",
      submenu: [
        { label: "تغيير مكان ملف البيانات…", click: chooseDataFolder },
        { label: "فتح مجلد البيانات", click: () => shell.showItemInFolder(dataFile()) },
        { label: "نسخة احتياطية الآن…", click: async () => {
            const r = await dialog.showSaveDialog(win, { defaultPath: path.join(app.getPath("documents"), "anbar-backup-" + new Date().toISOString().slice(0, 10) + ".json") });
            if (!r.canceled && r.filePath && fs.existsSync(dataFile())) { fs.copyFileSync(dataFile(), r.filePath); shell.showItemInFolder(r.filePath); }
          } },
        { type: "separator" },
        { label: "خروج", role: "quit" },
      ],
    },
    {
      label: "عرض",
      submenu: [
        { label: "تحديث", role: "reload" },
        { label: "تكبير", role: "zoomIn" },
        { label: "تصغير", role: "zoomOut" },
        { label: "الحجم الطبيعي", role: "resetZoom" },
        { label: "ملء الشاشة", role: "togglefullscreen" },
      ],
    },
    {
      label: "مساعدة",
      submenu: [
        { label: "مكان ملف البيانات", click: () => dialog.showMessageBox(win, { message: "ملف البيانات الحالي", detail: dataFile() }) },
        { label: "حول البرنامج", click: () => dialog.showMessageBox(win, { message: "النظام المالي لديوان محافظة الأنبار", detail: "الإصدار " + app.getVersion() + "\nقسم الشؤون المالية" }) },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(tpl));
}

function createWindow() {
  win = new BrowserWindow({
    width: 1360, height: 900, minWidth: 900, minHeight: 600,
    title: "النظام المالي لديوان محافظة الأنبار",
    backgroundColor: "#f2f4f7",
    webPreferences: { preload: path.join(__dirname, "preload.js"), contextIsolation: false, nodeIntegration: false, sandbox: false, spellcheck: false },
  });
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) shell.openExternal(url); return { action: "deny" }; });
  win.webContents.on("will-navigate", (e, url) => { if (!url.startsWith("file:")) { e.preventDefault(); if (/^https?:/.test(url)) shell.openExternal(url); } });
  win.loadFile(path.join(__dirname, "app", "index.html"));
}

app.whenReady().then(() => { buildMenu(); createWindow(); });
app.on("window-all-closed", () => app.quit());
