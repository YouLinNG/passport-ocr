const { app, BrowserWindow, ipcMain, Menu } = require("electron");
const Tesseract = require("tesseract.js");
const XLSX = require("xlsx");

// 保持对window对象的全局引用，如果不这么做的话，当JavaScript对象被
// 垃圾回收的时候，window对象将会自动的关闭
let win;

function createWindow() {
  Menu.setApplicationMenu(null);
  // 创建浏览器窗口。
  win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  });

  // 加载index.html文件
  win.loadFile("index.html");

  // 打开开发者工具
  // win.webContents.openDevTools();

  // 当 window 被关闭，这个事件会被触发。
  win.on("closed", () => {
    // 取消引用 window 对象，如果你的应用支持多窗口的话，
    // 通常会把多个 window 对象存放在一个数组里面，
    // 与此同时，你应该删除相应的元素。
    win = null;
  });
}

// Electron 会在初始化后并准备
// 创建浏览器窗口时，调用这个函数。
// 部分 API 在 ready 事件触发后才能使用。
app.on("ready", createWindow);

// 当全部窗口关闭时退出。
app.on("window-all-closed", () => {
  // 在 macOS 上，除非用户用 Cmd + Q 确定地退出，
  // 否则绝大部分应用及其菜单栏会保持激活。
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  // 在macOS上，当单击dock图标并且没有其他窗口打开时，
  // 通常在应用程序中重新创建一个窗口。
  if (win === null) {
    createWindow();
  }
});

ipcMain.on("asynchronous-message", async (event, filePaths) => {
  for (const filePath of filePaths) {
    const {
      data: { text }
    } = await Tesseract.recognize(filePath, "eng");
    try {
      const result = {};
      const lines = text.split("\n");
      const infoLine1 = lines.find(line => line.indexOf("<<<") > 0);
      const splitedInfoLine1 = infoLine1.replace(/\s/g, "").split(/\<+/g);
      result.name = `${splitedInfoLine1[0].substr(5)} ${splitedInfoLine1[1]}`;
      const infoLine2 = lines[lines.indexOf(infoLine1) + 1].replace(/\s/g, "");
      result.id = infoLine2.substr(0, 9);
      result.country = infoLine2.substr(10, 3);
      result.birth = infoLine2.substr(13, 6);
      const year = result.birth.substr(0, 2);
      const month = result.birth.substr(2, 2);
      const day = result.birth.substr(4, 2);
      const fullYear = `${
        year >
        new Date()
          .getFullYear()
          .toString()
          .substr(2, 2)
          ? "19"
          : "20"
      }${year}`;
      result.birth = `${day}/${month}/${fullYear}`;
      result.gender = infoLine2.substr(20, 1);
      event.reply("asynchronous-reply", result);
    } catch (e) {
      console.log(e);
      event.reply("asynchronous-reply", "failed");
    }
  }
  event.reply("asynchronous-reply", "end");
});

ipcMain.on("save-message", async (event, { filePath, results }) => {
  try {
    const wb = XLSX.utils.book_new();
    const data = [['姓名', '性别', '出生日期', '国籍', '护照号码']].concat(results.map(i => [i.name, i.gender, i.birth, i.country, i.id]));
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'result');
    XLSX.writeFile(wb, filePath);
    event.reply("save-reply", "保存成功");
  } catch(e) {
    console.log(e)
    event.reply("save-reply", "保存失败");
  }
});
