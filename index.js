const http = require("http");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");
const { send } = require("process");

//const uploadDir = path.join(__dirname, "uploads");

// Ensure the uploads directory exists
/*if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}*/

let dataToDownload = {};

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const contentType = req.headers["content-type"];
    if (!contentType || !contentType.includes("multipart/form-data")) {
      return reject(new Error("Invalid Content-Type"));
    }

    // 1. Extract the boundary string
    const boundaryMatch = contentType.match(/boundary=([^;]+)/);
    if (!boundaryMatch) {
      return reject(new Error("Boundary not found in Content-Type"));
    }
    const boundary = `--${boundaryMatch[1]}`;

    let fileSaved = false;
    let body = "";

    req.on("data", (chunk) => {
      // Collect chunks of the request body
      body += chunk.toString("latin1"); // Use 'latin1' for raw byte handling
    });

    req.on("end", () => {
      // 2. Split the body into parts based on the boundary
      const parts = body.split(boundary).slice(1, -1); // Remove preamble and epilogue
      let dataToWrite = {};

      for (const part of parts) {
        // 3. Find the headers and body separation
        const headersEnd = part.indexOf("\r\n\r\n");
        if (headersEnd === -1) continue;

        const headers = part.substring(0, headersEnd);
        let content = part.substring(headersEnd + 4); // +4 for \r\n\r\n

        // Extract Content-Disposition
        const dispositionMatch = headers.match(
          /Content-Disposition: form-data; name="([^"]+)"/
        );
        if (!dispositionMatch) continue;

        const name = dispositionMatch[1];

        // 4. Check for file data
        if (headers.includes("filename=")) {
          // This part is a file
          const filenameMatch = headers.match(/filename="([^"]+)"/);
          const mimeTypeMatch = headers.match(/Content-Type: ([^\r\n]+)/);

          const filename = filenameMatch ? filenameMatch[1] : `upload-${Date.now()}`;
          const mimeType = mimeTypeMatch
            ? mimeTypeMatch[1]
            : "application/octet-stream";

          console.log(
            `[FILE] Name: ${name}, Filename: ${filename}, Type: ${mimeType}`
          );

          const fileContent = Buffer.from(content, "latin1");
          dataToWrite.data = fileContent.slice(0, fileContent.length - 2);
          //const saveTo = path.join(uploadDir, filename);
          //fs.writeFileSync(saveTo, dataToWrite);
          //fs.writeFileSync("original.xlsx", dataToWrite);
          fileSaved = true;
          break; // Assuming only one file is uploaded
        } else {
          const value = content.trim().replace(/\r\n$/, "");
          console.log(`[FIELD] Name: ${name}, Value: ${value}`);
        }
      }

      if (fileSaved) {
        sendProg("11 File uploaded.");
        //resolve(`File "original.xlsx" uploaded and saved successfully.`);
        resolve(dataToWrite);
      } else {
        resolve("Form processed, but no file was saved.");
      }
    });

    req.on("error", (err) => {
      console.error("Request error:", err);
      reject(new Error("Request failed during stream processing."));
    });
  });
}

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

const clients = new Set();

function addClient(res) {
  clients.add(res);
  res.on("close", () => {
    clients.delete(res);
  });
}

function sendProg(mssg) {
  const ix = mssg.indexOf(" ");
  const message = `data: ${mssg.replace(
    " ",
    " <font color='green'>✔</font>&nbsp;"
  )}\n\n`;
  for (const client of clients) client.write(message);
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

async function modify(res, Data) {
  const archive = "original.xlsx";
  //const data = fs.readFileSync(archive);
  const zip = await JSZip.loadAsync(Data.data);
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
          `11 Workbook ${wkProt ? "<span>unlocked</span>" : "is not locked"}.`
        );

        const sheets = doc.querySelectorAll("sheets > sheet");
        for (const sheet of sheets) {
          foi.push({
            name: sheet.getAttribute("name"),
            state: sheet.getAttribute("state"),
            id: sheet.getAttribute("r:id"),
          });
        }

        let sCnt = 0, hid;
        totSheets = sheets.length * 2;
        for (const sheet of sheets) {
          hid = sheet.getAttribute("state") === "hidden";
          if (hid) {
            sheet.removeAttribute("state");
            change = true;
            sCnt++;
          }
          sendProg(
            `${calcProgess()} Sheet <u>${sheet.getAttribute("name")}</u> ${
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
        //getName(pth, prt ? true : false);
        foiAdd.push({ pth: pth, prot: prt ? true : false });
        if (prt) {
          prt.remove();
          zip.file(pth, dom.serialize());
          modified = true;
        }
        sendProg(
          `${calcProgess()} Sheet <u>${pth.slice(
            pth.lastIndexOf("/") + 1
          )}</u> ${prt ? "<span>unprotected</span>" : "not protected"}.`
        );
      })(pth, file);
      modPromises.push(modTask);
    }
  });

  await Promise.all(modPromises);
  sendProg(`+ Processing complete.`);
  if (!modified) {
    return res.end(
      "<img src='public/logo.png'/>&nbsp;<font color='green'>This file is not locked.</font>"
    );
  }
  const newZipData = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  //fs.writeFileSync("unlocked.xlsx", newZipData);
  dataToDownload.data = newZipData;
  return res.end(show());
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
    if (req.url.slice(0, 7) === "/events") {
      console.log("Client connected to SSE");
      res.setHeader("Content-Type", "text/event-stream");
      res.setHeader("Cache-Control", "no-cache");
      res.setHeader("Connection", "keep-alive");
      res.flushHeaders();
      addClient(res);
      sendProg("0 About to start...");
      return;
    }
    if (req.url.slice(0, 9) === "/download") {
      download(res);
      return;
    }
    return;
  }
  if (req.method === "POST" && req.url === "/upload") {
    sendProg("5 Processing started.");
    parseMultipart(req)
      .then((data) => {
        modify(res, data);
      })
      .catch((error) => {
        res.writeHead(500, { "Content-Type": "text/plain" });
        res.end(`Upload failed: ${error.message}`);
      });
  }
});

function download(res) {
        res.setHeader('Content-Type', 'application/octet-stream');
        res.setHeader('Content-Disposition', 'attachment; filename="unlocked.xlsx"');
        //res.setHeader('Content-Length', stat.size);
        res.end(dataToDownload.data)
}

const PORT = process.env.PORT || 5000;
server.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}/`);
});
