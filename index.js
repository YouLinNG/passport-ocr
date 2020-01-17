const { app, BrowserWindow, ipcMain, Menu } = require("electron");
const Tesseract = require("tesseract.js");
const XLSX = require("xlsx");

const crypto = require('crypto')
const request = require('request')
const fs = require('fs')

function remoteOCR(base64) {
  /**
  * 第一步: 拼接规范请求串 CanonicalRequest
  * 注意: 可对照代码看参数含义及注意事项 。
  */
  // 说明: HTTP 请求方法（GET、POST ）。此示例取值为 POST
  var HTTPRequestMethod = 'POST';
  // 说明: URI 参数，API 3.0 固定为正斜杠（/）
  var CanonicalURI = '/';
  // 说明: POST请求时为空
  var CanonicalQueryString = "";
  /**  说明:
   * 参与签名的头部信息，content-type 和 host 为必选头部 ,
   * content-type 必须为小写 , 推荐 content-type 值 application/json , 对应方法为 TC3-HMAC-SHA256 签名方法 。
   * 其中 host 指接口请求域名 POST 请求支持的 Content-Type 类型有:
   * 1. application/json（推荐）, 必须使用 TC3-HMAC-SHA256 签名方法 ;
   * 2. application/x-www-form-urlencoded , 必须使用 HmacSHA1 或 HmacSHA256 签名方法 ;
   * 3. multipart/form-data（仅部分接口支持）, 必须使用 TC3-HMAC-SHA256 签名方法 。
   * 
   * 注意:
   * content-type 必须和实际发送的相符合 , 有些编程语言网络库即使未指定也会自动添加 charset 值 , 
   * 如果签名时和发送时不  一致，服务器会返回签名校验失败。
  */
  var CanonicalHeaders = "content-type:application/json\nhost:ocr.tencentcloudapi.com\n";
  /**  说明:
   * 参与签名的头部信息的 key，可以说明此次请求都有哪些头部参与了签名，和 CanonicalHeaders 包含的头部内容是一一对应的。
   * content-type 和 host 为必选头部 。 
   * 注意： 
   * 1. 头部 key 统一转成小写； 
   * 2. 多个头部 key（小写）按照 ASCII 升序进行拼接，并且以分号（;）分隔 。 
  */
  var SignedHeaders = "content-type;host";
  /**
   * 参与签名的头部信息的 key，可以说明此次请求都有哪些头部参与了签名，和 CanonicalHeaders 包含的头部内容是一一对应的。
   * content-type 和 host 为必选头部 。 
   * 注意： 
   * 1. 头部 key 统一转成小写； 
   * 2. 多个头部 key（小写）按照 ASCII 升序进行拼接，并且以分号（;）分隔 。 
   */
  // 传入需要做 HTTP 请求的正文 body
  var payload = {
    Type: "CN",
    ImageBase64: base64
  }
  /**  说明:
   * 对请求体加密后的字符串 , 每个语言加密加密最终结果一致 , 但加密方法不同 , 
   * 这里 nodejs 的加密方法为 crypto.createHash('sha256').update(JSON.stringify(payload)).digest('hex'); 
   * 选择加密函数需要能够满足对 HTTP 请求正文做 SHA256 哈希 , 然后十六进制编码 , 最后编码串转换成小写字母的功能即可 。
  */
  var HashedRequestPayload = crypto.createHash('sha256').update(JSON.stringify(payload)).digest('hex');
  // 最后拼接以上六个字段 , 注意中间用 '/n' 拼接 , 拼接格式一定要如下格式 , 否则会报错
  var CanonicalRequest = HTTPRequestMethod + '\n' +
    CanonicalURI + '\n' +
    CanonicalQueryString + '\n' +
    CanonicalHeaders + '\n' +
    SignedHeaders + '\n' +
    HashedRequestPayload;
  console.log('1. 拼接规范请求串' + '\n' + CanonicalRequest);

  // 2. 拼接待签名字符串
  // 签名算法，接口鉴权v3为固定值 TC3-HMAC-SHA256
  var Algorithm = "TC3-HMAC-SHA256";
  // 请求时间戳，即请求头部的公共参数 X-TC-Timestamp 取值，取当前时间 UNIX 时间戳，精确到秒
  var RequestTimestamp = Math.round(new Date().getTime() / 1000) + "";
  /**
   * Date 必须从时间戳 X-TC-Timestamp 计算得到，且时区为 UTC+0。
   * 如果加入系统本地时区信息，例如东八区，将导致白天和晚上调用成功，但是凌晨时调用必定失败。
   * 假设时间戳为 1551113065，在东八区的时间是 2019-02-26 00:44:25，但是计算得到的 Date 取 UTC+0 的日期应为 2019-02-25，而不是 2019-02-26。
   * Timestamp 必须是当前系统时间，且需确保系统时间和标准时间是同步的，如果相差超过五分钟则必定失败。
   * 如果长时间不和标准时间同步，可能导致运行一段时间后，请求必定失败，返回签名过期错误。
  */
  var t = new Date();
  var date = t.toISOString().substr(0, 10); // 计算 Date 日期   date = "2019-08-26"
  /**
   *  拼接 CredentialScope 凭证范围，格式为 Date/service/tc3_request ， 
   * service 为服务名，慧眼用 faceid ， OCR 文字识别用 ocr
  */
  var CredentialScope = date + "/ocr/tc3_request";
  // 将第一步拼接得到的 CanonicalRequest 再次进行哈希加密
  var HashedCanonicalRequest = crypto.createHash('sha256').update(CanonicalRequest).digest('hex');
  // 拼接 StringToSign
  var StringToSign = Algorithm + '\n' +
    RequestTimestamp + '\n' +
    CredentialScope + '\n' +
    HashedCanonicalRequest;
  console.log('2. 拼接待签名字符串' + '\n' + StringToSign);

  // 3. 计算签名 Signature
  var SecretKey = "Wx9F6g180cBVkjSu5zd8MTKx5x4vX7ev";
  var SecretDate = crypto.createHmac('sha256', "TC3" + SecretKey).update(date).digest();
  var SecretService = crypto.createHmac('sha256', SecretDate).update("ocr").digest();
  var SecretSigning = crypto.createHmac('sha256', SecretService).update("tc3_request").digest();
  var Signature = crypto.createHmac('sha256', SecretSigning).update(StringToSign).digest('hex');
  console.log('3. 计算签名' + Signature);

  // 4. 拼接签名 Authorization
  var SecretId = "AKIDsVw77rTktvhDUVkvsZ9YhbGfntlbO7xY"; // SecretId, 需要替换为自己的
  var Algorithm = "TC3-HMAC-SHA256";
  var Authorization =
    Algorithm + ' ' +
    'Credential=' + SecretId + '/' + CredentialScope + ', ' +
    'SignedHeaders=' + SignedHeaders + ', ' +
    'Signature=' + Signature
  console.log('4. 拼接Authorization' + '\n' + Authorization)

  // 5. 发送POST请求 options 配置
  var options = {
    url: 'https://ocr.tencentcloudapi.com/',
    method: 'POST',
    json: true,
    headers: {
      "Content-Type": "application/json",
      "Authorization": Authorization,
      "Host": "ocr.tencentcloudapi.com",
      "X-TC-Action": "PassportOCR",
      "X-TC-Version": "2018-11-19",
      "X-TC-Timestamp": RequestTimestamp,
      "X-TC-Region": "ap-guangzhou"
    },
    body: payload,
  };
  return new Promise(function(resolve, reject) {
    request(options, function (error, response, body) {
      if (error) reject(error);
      resolve(body);
    })
  });
}

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

    // const {
    //   data: { text }
    // } = await Tesseract.recognize(filePath, "eng");
    try {
      let bitmap = fs.readFileSync(filePath);
      const base64 = new Buffer(bitmap).toString('base64');
      const { Response: r } = await remoteOCR(base64);

      const result = {
        name: `${r.FamilyName} ${r.FirstName}`,
        id: r.PassportNo,
        birth: r.BirthDate,
        country: r.Nationality,
        gender: r.Sex
      };
      
      // console.log(text)
      // const lines = text.split("\n");
      // const infoLine1 = lines.find(line => line.indexOf("<<<") > 0);
      // const splitedInfoLine1 = infoLine1.replace(/\s/g, "").split(/\<+/g);
      // result.name = `${splitedInfoLine1[0].substr(5)} ${splitedInfoLine1[1]}`;
      // const infoLine2 = lines[lines.indexOf(infoLine1) + 1].replace(/\s/g, "");
      // result.id = infoLine2.substr(0, 9);
      // result.country = infoLine2.substr(10, 3);
      // result.birth = infoLine2.substr(13, 6);
      // const year = result.birth.substr(0, 2);
      // const month = result.birth.substr(2, 2);
      // const day = result.birth.substr(4, 2);
      // const fullYear = `${
      //   year >
      //     new Date()
      //       .getFullYear()
      //       .toString()
      //       .substr(2, 2)
      //     ? "19"
      //     : "20"
      //   }${year}`;
      // result.birth = `${day}/${month}/${fullYear}`;
      // result.gender = infoLine2.substr(20, 1);
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
  } catch (e) {
    console.log(e)
    event.reply("save-reply", "保存失败");
  }
});
