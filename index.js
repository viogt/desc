const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");

const app = express();
const PORT = process.env.PORT || 5000;

const storage = multer.diskStorage({
    destination: ".",
    filename: (req, file, cb) => {
        cb(null, "original.xlsx");
    },
});

const upload = multer({ storage: storage });

const clients = new Set();

function addClient(res) {
    clients.add(res);
    res.on("close", () => {
        clients.delete(res);
    });
}

function sendProg(mssg) {
    //console.log(mssg);
    const ix = mssg.indexOf(" ");
    //mssg = `${mssg.substr(0,ix)} <span color='grren'>✔</span> ${mssg.slice(ix + 1)}`;
    //const message = `data: ${mssg.substr(0, ix)} <font color='green'>✔</font>&nbsp;${mssg.slice(ix + 1)}\n\n`;
    const message = `data: ${mssg.replace(" ", " <font color='green'>✔</font>&nbsp;")}\n\n`;
    for (const client of clients) client.write(message);
}

app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "index.html"));
    //res.send(uploadFormHtml);
    //sendProg("0 Process started.");
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
    console.log(
        "perSheet:",
        perSheet,
        "remSheets:",
        remSheets,
        "doneSheets:",
        doneSheets,
    );
    return perSheet;
}

async function modify(archive, res) {
    try {
        const data = fs.readFileSync(archive);
        const zip = await JSZip.loadAsync(data);
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
                        `11 Workbook ${wkProt ? "<span>unlocked</span>" : "is not locked"}.`,
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
                            `${calcProgess()} Sheet <u>${sheet.getAttribute("name")}</u> ${hid ? "<span>unhidden</span>" : "not hidden"}.`,
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
                        `${calcProgess()} Sheet <u>${pth.slice(pth.lastIndexOf("/") + 1)}</u> ${prt ? "<span>unprotected</span>" : "not protected"}.`,
                    );
                })(pth, file);
                modPromises.push(modTask);
            }
        });

        await Promise.all(modPromises);
        sendProg(`+ Processing complete.`);
        if (!modified) {
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
    return str + "</table><br>▾ <a href='/download'>DOWNLOAD UNLOCKED</a>";
}

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
