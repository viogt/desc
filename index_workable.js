const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");

const app = express();
const PORT = 5000;
const UPLOAD_DIR = "uploads/";

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
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XLSX Only Uploader</title>
    <style>
        body { font-family: 'Arial', sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background-color: #f4f7f6; }
        .container { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 25px rgba(0, 0, 0, 0.1); width: 100%; max-width: 400px; text-align: center; }
        h1 { color: #333; margin-bottom: 25px; font-size: 1.8rem; }
        .allowed-text { color: #666; margin-bottom: 20px; font-style: italic; }
        .message { margin-top: 20px; padding: 15px; border-radius: 8px; font-weight: bold; }
        .success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        input[type="file"] { border: 1px solid #ccc; padding: 10px; width: 100%; border-radius: 6px; box-sizing: border-box; margin-bottom: 20px; }
        button { background-color: #007bff; color: white; padding: 12px 25px; border: none; border-radius: 6px; cursor: pointer; font-size: 1rem; transition: background-color 0.3s; width: 100%; }
        button:hover { background-color: #0056b3; }
        span { color: #fff; background-color: #c00; padding: 2px 8px; border-radius: 6px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>XLSX Upload Only</h1>
        <p class="allowed-text">Accepts only .xlsx files.</p>

        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="myFile" required>
            <button type="submit">Upload Spreadsheet</button>
        </form>

        <div id="statusMessage"></div>
    </div>

    <script>
        const urlParams = new URLSearchParams(window.location.search);
        const statusMessage = document.getElementById('statusMessage');

        if (urlParams.has('success')) {
            const fileName = urlParams.get('success');
            statusMessage.innerHTML = \`<div class="message success">Unlocked:\${fileName}</div>\`;
        } else if (urlParams.has('error')) {
            // Display error message from the Multer fileFilter or general error
            statusMessage.innerHTML = \`<div class="message error">Upload Failed: \${urlParams.get('error')}</div>\`;
        }
    </script>
</body>
</html>
`;

app.get("/", (req, res) => {
    res.send(uploadFormHtml);
});

app.post("/upload", async (req, res) => {
    upload.single("myFile")(req, res, (err) => {
        if (err) {
            const errorMessage =
                err.message || "An unknown upload error occurred.";
            console.error("Upload Error:", errorMessage);
            return res.redirect(`/?error=${encodeURIComponent(errorMessage)}`);
        }

        if (!req.file) {
            return res.redirect("/?error=No+file+selected");
        }

        console.log(`Valid file received and saved: ${req.file.path}`);
        modify(req.file.path, res);
    });
});

app.get("/download", (req, res) => {
    res.download("unlocked.xlsx");
});

var foi;

async function modify(archive, res) {
    try {
        const data = fs.readFileSync(archive);
        const zip = await JSZip.loadAsync(data);
        const rgx = /xl\/worksheets\/sheet\d+\.xml/;

        const modPromises = [];
        foi = [];

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

                    zip.file(pth, dom.serialize());

                    //if (change) outZip.addBuffer(dom.serialize(), file.path); //---!!! IF
                    //outZip.addBuffer(Buffer.from(newBuf), file.path);
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
                    getName(pth, prt ? true : false);
                    if (prt) {
                        prt.remove();
                        zip.file(pth, dom.serialize());
                    }
                })(pth, file);
                modPromises.push(modTask);
            }
        });

        await Promise.all(modPromises);
        console.log("Length = " + modPromises.length);
        const newZipData = await zip.generateAsync({ type: "nodebuffer" });
        fs.writeFileSync("unlocked.xlsx", newZipData);

        const fileNameForMessage = encodeURIComponent(show());
        return res.redirect(`/?success=${fileNameForMessage}`);
    } catch (err) {
        console.log("ERROR: " + err.message);
        return res.redirect(`/?error=${err.message}`);
    }
}

function getName(nm, flag) {
    const num = parseInt(nm.match(/\d+\.xml$/)[0]) - 1;
    foi[num].prot = flag;
}

function show() {
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
    let str =
        '<table cellPadding="4" cellSpacing="4" style="font-weight:normal;border:1px solid #888;color:#000;background-color:#fff;width:100%;">';
    for (el of foi) {
        str += `<tr><td>${num}</td><td align="left">${el.name}</td><td>${el.state == "hidden" ? "<span style='background-color:purple'>hidden</span>" : "—"}</td><td>${el.prot ? "<span>protected</span>" : "—"}</td></tr>`;
        num++;
    }
    return str + "</table><br>▾ <a href='/download'>Download unlocked</a>";
}

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
