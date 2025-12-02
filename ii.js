const WebSocket = require("ws");
const http = require("http");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");
const _PORT = 8080;

let dataToDownload = {};

function serveFile(req, res, filePath) {
  //res.setHeader('Content-Type', mimeType);
  const fileStream = fs.createReadStream(path.join(__dirname, filePath));
  fileStream.on("error", (err) => {
    console.error("File Read Error:", err);
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("404 Not Found");
  });
  fileStream.pipe(res);
}

function sendProg(cod, mssg) {
  const ms = JSON.stringify({ cod: cod, mssg: mssg });
  //mssg = mssg.replace(" ", " <font color='green'>✔</font>&nbsp;");
  wss.clients.forEach((client) => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(ms);
    }
  });
}

var foi, foiAdd, modified;
var totSheets, remSheets, doneSheets;

function calcProgess() {
  let perSheet;
  if (totSheets < 0) {
    perSheet = 3;
    remSheets -= perSheet;
    doneSheets++;
  } else {
    perSheet = remSheets / (totSheets - doneSheets);
    remSheets -= perSheet;
    if (remSheets <= 0) {
      perSheet = 1;
      remSheets = 0;
    }
    doneSheets++;
  }
  return perSheet;
}

async function modify(Data) {
  const zip = await JSZip.loadAsync(Data);
  const rgx = /xl\/worksheets\/sheet\d+\.xml/;

  const modPromises = [];
  foi = [];
  foiAdd = [];
  modified = false;
  totSheets = -1;
  remSheets = 70;
  doneSheets = 0;

  zip.forEach(async (pth, file) => {
    if (pth === "xl/workbook.xml") {
      const modTask = (async (pth, file) => {
        let change = false;
        const buf = await file.async("text");
        const dom = new JSDOM(buf, {
          contentType: "application/xml",
        });
        const doc = dom.window.document;

        const wkProt = doc.querySelector("workbookProtection");
        if (wkProt) {
          wkProt.remove();
          change = true;
        }
        sendProg(
          11,
          `Workbook ${wkProt ? "<span>unlocked</span>" : "is not locked"}.`
        );

        const sheets = doc.querySelectorAll("sheets > sheet");
        for (const sheet of sheets) {
          foi.push({
            name: sheet.getAttribute("name"),
            state: sheet.getAttribute("state"),
            id: sheet.getAttribute("r:id"),
          });
        }

        let sCnt = 0,
          hid;
        totSheets = sheets.length * 2;
        for (const sheet of sheets) {
          hid = sheet.getAttribute("state") === "hidden";
          if (hid) {
            sheet.removeAttribute("state");
            change = true;
            sCnt++;
          }
          sendProg(
            calcProgess(),
            `Sheet <u>${sheet.getAttribute("name")}</u> ${
              hid ? "<span>unhidden</span>" : "not hidden"
            }.`
          );
        }

        if (change) {
          zip.file(pth, dom.serialize());
          modified = true;
        }
      })(pth, file);
      modPromises.push(modTask);
    } else if (rgx.test(pth)) {
      const modTask = (async (pth, file) => {
        const buf = await file.async("text");
        const dom = new JSDOM(buf, {
          contentType: "application/xml",
        });
        const doc = dom.window.document;
        const prt = doc.querySelector("sheetProtection");
        foiAdd.push({ pth: pth, prot: prt ? true : false });
        if (prt) {
          prt.remove();
          zip.file(pth, dom.serialize());
          modified = true;
        }
        sendProg(
          calcProgess(), `Sheet <u>${pth.slice(
            pth.lastIndexOf("/") + 1
          )}</u> ${prt ? "<span>unprotected</span>" : "not protected"}.`
        );
      })(pth, file);
      modPromises.push(modTask);
    }
  });

  await Promise.all(modPromises);
  sendProg(20, 'Processing complete.');
  if (!modified) {
    sendProg(-1,
      "<img src='public/logo.png'/>&nbsp;<font color='green'>This file is not locked.</font>"
    );
    return;
  }
  const newZipData = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  //fs.writeFileSync("unlocked.xlsx", newZipData);
  dataToDownload.data = newZipData;
  sendProg(-1, show());
}

function show() {
  for (el of foiAdd) {
    const num = parseInt(el.pth.match(/\d+\.xml$/)[0]) - 1;
    foi[num].prot = el.prot;
  }
  let num = 1;
  let str = "<br><font color='green'>FILE UNLOCKED:</font><br><table>";
  for (el of foi) {
    str += `<tr><td>${num}</td><td align="left">${el.name}</td><td>${
      el.state == "hidden" ? "<b lilac>hidden</b>" : "—"
    }</td><td>${el.prot ? "<b>protected</b>" : "—"}</td></tr>`;
    num++;
  }
  return (
    str +
    "</table><br><img src='public/logo.png'/>&nbsp;<a style='color:green' href='/download'>DOWNLOAD UNLOCKED</a>"
  );
}

const server = http.createServer((req, res) => {
  if (req.method === "GET") {
    if (req.url === "/") {
      serveFile(req, res, "public/index.html");
      return;
    }
    if (req.url.slice(0, 8) === "/public/") {
      serveFile(req, res, req.url);
      return;
    }
    if (req.url.slice(0, 9) === "/download") {
      download(res);
      return;
    }
  }
  return;
});

const wss = new WebSocket.Server({ server: server }); //{ port: _PORT });

wss.on("connection", (ws) => {
  console.log("A new client connected. Total: " + wss.clients.size);
  ws.on("message", (message, isBinary) => {
    if (isBinary) {
      console.log(`Receiving chunk of size: ${message.byteLength}`);
      sendProg(5, "Processing file...");
      modify(message);
      return;
    }
    /*const messageString = message.toString();
    console.log(`Received message from client: ${messageString}`);
    wss.clients.forEach((client) => {
      if (client.readyState === WebSocket.OPEN) {
        client.send(messageString);
      }
    });*/
  });
  ws.on("close", () => {
    console.log("Client has disconnected.");
  });
  ws.on("error", (error) => {
    console.error("WebSocket error:", error);
  });
});

function download(res) {
  res.setHeader("Content-Type", "application/octet-stream");
  res.setHeader("Content-Disposition", 'attachment; filename="unlocked.xlsx"');
  //res.setHeader('Content-Length', stat.size);
  res.end(dataToDownload.data);
}

const PORT = process.env.PORT || 5000;
server.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}/`);
});
