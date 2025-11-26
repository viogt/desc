const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");

const app = express();
const PORT = process.env.PORT || 5000;

const xlsxFileFilter = (req, file, cb) => {
    const isXlsxExt = path.extname(file.originalname).toLowerCase() === ".xlsx";
    const isXlsxMime =
        file.mimetype ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    if (isXlsxExt && isXlsxMime) {
        cb(null, true);
    } else {
        cb(
            new Error("Only Microsoft Excel (.xlsx) files are permitted."),
            false,
        );
    }
};

const storage = multer.diskStorage({
    destination: ".",
    filename: (req, file, cb) => {
        cb(null, "original.xlsx");
    },
});

const upload = multer({
    storage: storage,
    fileFilter: xlsxFileFilter,
});

const uploadFormHtml = `
<!doctype html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Upload an Excel File</title>
        <style>
body { background-color: #ccc;}
.container {
max-width: 600px;
margin: auto;
padding: 24px;
background-color: #fff;
box-shadow: 0 4px 8px #888;
border-radius: 16px;
font: normal 18px Arial, sans-serif;
text-align: center;
}
            h2 {
                font:
                    bold 32px Arial,
                    sans-serif;
                color: #aab;
            }
            input[type="file"] { font-size: 18px; }
b { font-weight:normal; background-color: orange; padding: 2px 8px; border-radius: 6px; }
b[lilac] { background-color: yellow; }
table {
  border-collapse: collapse;
  width: 100%;
  padding: 12px;
  border-top:.8px solid #ccc;
}
td {
  border-bottom: .8px solid #ccc;
  padding: 6px;
}
#progress {
    width: 100%; height: 20px; background: #ddd;
    margin-top: 20px; border-radius: 6px; overflow: hidden;
}
#bar {
    height: 100%; width: 0; background-color: green; transition: .2s;
    color: #fff; text-align: center; font: normal 14px/20px Arial, sans-serif;
}
        </style>
    </head>
    <body>
<div class="container">
<h2>Upload and Analyze</h2>
<input type="file" id="fileInput" onchange="uploadFile()" />
<div id="progress"><div id="bar"></div></div>
<div id="message"></div>
</div>
        <script>
const log = document.getElementById("message");
const bar = document.getElementById("bar");

const es = new EventSource("/events");
let progress;

es.onmessage = (event) => {
    let percent;
    log.innerHTML += event.data.substr(event.data.indexOf(' ')+1) + "<br>";
    if(event.data.charAt(0) == '0') return;
    if(event.data.charAt(0) == '+') percent = '100%';
    else percent = (progress += parseInt(event.data)) + '%';
    bar.style.width = percent;
    bar.textContent = percent;
};

es.onerror = () => {
    log.textContent += "[Error / connection lost]\\n";
};
            async function uploadFile() {
                const fileInput = document.getElementById("fileInput");
                const file = fileInput.files[0]; // Get the first selected file
                const messageElement = document.getElementById("message");
                messageElement.textContent = "";
                progress = 0;

                if (!file || !file.name.endsWith(".xlsx")) {
                    messageElement.innerHTML =
                        "<font color='red'>Please select an .XLSX file.</font>";
                    return;
                }
                const formData = new FormData();
                formData.append("myFile", file);

                try {
                    const response = await fetch("/upload", {
                        method: "POST",
                        body: formData,
                    });

                    if (response.ok)
                        messageElement.innerHTML += await response.text();
                    else {
                        const errorText = await response.text();
                        //messageElement.textContent = "Upload failed:" + response.status + " - " + errorText;
                    }
                } catch (error) {
                    messageElement.textContent = "Network error: " + error.message;
                    console.error("Fetch error:", error);
                }
            }
        </script>
    </body>
</html>
`;

const clients = new Set();

function addClient(res) {
    clients.add(res);
    res.on("close", () => {
        clients.delete(res);
    });
}

function sendProg(mssg) {
    const message = `data: ${mssg}\n\n`;
    for (const client of clients) client.write(message);
}

app.get("/", (req, res) => {
    res.send(uploadFormHtml);
    sendProg("0 Process started.");
});

app.get("/events", (req, res) => {
    console.log("Client connected to SSE");
    res.setHeader("Content-Type", "text/event-stream");
    res.setHeader("Cache-Control", "no-cache");
    res.setHeader("Connection", "keep-alive");
    res.flushHeaders();
    addClient(res);
});

app.post("/upload", async (req, res) => {
    sendProg("5 Processing started.");
    upload.single("myFile")(req, res, (err) => {
        if (err) {
            const errorMessage =
                err.message || "An unknown upload error occurred.";
            console.error("Upload Error:", errorMessage);
            return res.send(`ERROR: ${encodeURIComponent(errorMessage)}`);
        }

        if (!req.file) {
            return res.send("ERROR: No file selected");
        }

        sendProg("11 File uploaded.");
        modify(req.file.path, res);
    });
});

app.get("/download", (req, res) => {
    res.download("unlocked.xlsx");
});

var foi, foiAdd, modified;
var protSheets, totSheets;

async function modify(archive, res) {
    try {
        const data = fs.readFileSync(archive);
        const zip = await JSZip.loadAsync(data);
        const rgx = /xl\/worksheets\/sheet\d+\.xml/;
        sendProg("12 Iterating through elements.");

        const modPromises = [];
        (foi = []), (foiAdd = []);
        modified = false;

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
                        `7 Workbook ${wkProt ? "unlocked" : "is not locked"}.`,
                    );

                    const sheets = doc.querySelectorAll("sheets > sheet");
                    for (const sheet of sheets) {
                        foi.push({
                            name: sheet.getAttribute("name"),
                            state: sheet.getAttribute("state"),
                            id: sheet.getAttribute("r:id"),
                        });
                    }

                    let sCnt = 0;
                    totSheets = sheets.length;
                    for (const sheet of sheets) {
                        if (sheet.getAttribute("state") === "hidden") {
                            sheet.removeAttribute("state");
                            change = true;
                            sCnt++;
                        }
                    }
                    sendProg(`10 ${sCnt}/${totSheets} sheets unhidden.`);

                    if (change) {
                        zip.file(pth, dom.serialize());
                        modified = true;
                    }
                })(pth, file);
                modPromises.push(modTask);
            } else if (rgx.test(pth)) {
                protSheets = 0;
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
                        protSheets++;
                    }
                })(pth, file);
                modPromises.push(modTask);
            }
        });

        await Promise.all(modPromises);
        sendProg(`45 ${protSheets}/${totSheets} sheets unlocked.`);
        if (!modified) {
            sendProg("+ File is not locked.");
            return res.send(
                "<font color='blue'>This file is not locked.</font>",
            );
        }
        const newZipData = await zip.generateAsync({
            type: "nodebuffer",
            compression: "DEFLATE",
            compressionOptions: { level: 9 },
        });
        fs.writeFileSync("unlocked.xlsx", newZipData);
        sendProg("+ File prepared for download.");
        return res.send(show());
    } catch (err) {
        console.log(err);
        return res.send(`ERROR: ${err.message}`);
    }
}

function show() {
    for (el of foiAdd) {
        const num = parseInt(el.pth.match(/\d+\.xml$/)[0]) - 1;
        foi[num].prot = el.prot;
    }
    console.log(foi, foiAdd);

    console.log("\n-----------------sheets:", foi.length);
    let num = 1;
    for (el of foi) {
        console.log(
            num,
            el.name,
            el.state == "hidden" ? "HIDDEN" : "---",
            el.id,
            el.prot ? "(protected)" : "",
        );
        num++;
    }
    console.log("> Unlocked.xlsx created");
    num = 1;
    let str = "<br><font color='blue'>FILE UNLOCKED:</font><br><table>";
    for (el of foi) {
        str += `<tr><td>${num}</td><td align="left">${el.name}</td><td>${el.state == "hidden" ? "<b lilac>hidden</b>" : "—"}</td><td>${el.prot ? "<b>protected</b>" : "—"}</td></tr>`;
        num++;
    }
    return str + "</table><br>▾ <a href='/download'>Download unlocked</a>";
}

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
