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
.loader {
width: 18px;
height: 18px;
border: 5px solid #00F;
border-bottom-color: transparent;
border-radius: 50%;
display: inline-block;
box-sizing: border-box;
animation: rotation 1s linear infinite;
}
@keyframes rotation {
0% { transform: rotate(0deg); }
100% { transform: rotate(360deg); }
}
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
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Upload and Analyze</h2>
            <input type="file" id="fileInput" onchange="uploadFile()" />
            <p id="message"></p>
        </div>

        <script>
            async function uploadFile() {
                const fileInput = document.getElementById("fileInput");
                const file = fileInput.files[0]; // Get the first selected file
                const messageElement = document.getElementById("message");

                if (!file || !file.name.endsWith(".xlsx")) {
                    messageElement.innerHTML =
                        "<font color='red'>Please select an .XLSX file.</font>";
                    return;
                }
                messageElement.innerHTML = "<span class='loader'></span> Processing...";

                const formData = new FormData();
                formData.append("myFile", file);

                try {
                    const response = await fetch("/upload", {
                        method: "POST",
                        body: formData,
                    });

                    if (response.ok)
                        messageElement.innerHTML = await response.text();
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

app.get("/", (req, res) => {
    res.send(uploadFormHtml);
});

app.post("/upload", async (req, res) => {
    console.log("File received");
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

        console.log(`Valid file received and saved: ${req.file.path}`);
        modify(req.file.path, res);
    });
});

app.get("/download", (req, res) => {
    res.download("unlocked.xlsx");
});

var foi, foiAdd, modified;

async function modify(archive, res) {
    try {
        const data = fs.readFileSync(archive);
        const zip = await JSZip.loadAsync(data);
        const rgx = /xl\/worksheets\/sheet\d+\.xml/;

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
                        console.log(pth, "protected");
                    }

                    const sheets = doc.querySelectorAll("sheets > sheet");
                    for (const sheet of sheets) {
                        foi.push({
                            name: sheet.getAttribute("name"),
                            state: sheet.getAttribute("state"),
                            id: sheet.getAttribute("r:id"),
                        });
                    }

                    for (const sheet of sheets) {
                        if (sheet.getAttribute("state") === "hidden") {
                            sheet.removeAttribute("state");
                            change = true;
                        }
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
                })(pth, file);
                modPromises.push(modTask);
            }
        });

        await Promise.all(modPromises);
        if (!modified)
            return res.send(
                "<font color='blue'>This file is not locked.</font>",
            );
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
    let str = "FILE UNLOCKED:<br><table>";
    for (el of foi) {
        str += `<tr><td>${num}</td><td align="left">${el.name}</td><td>${el.state == "hidden" ? "<b lilac>hidden</b>" : "—"}</td><td>${el.prot ? "<b>protected</b>" : "—"}</td></tr>`;
        num++;
    }
    return str + "</table><br>▾ <a href='/download'>Download unlocked</a>";
}

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
