const fs = require("fs");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");

const foi = [];

async function modify(archive) {
  try {
    const data = fs.readFileSync(archive);
    const zip = await JSZip.loadAsync(data);
    const rgx = /xl\/worksheets\/sheet\d+\.xml/;

    const modPromises = [];

    zip.forEach(async (pth, file) => {
      if (pth === "xl/workbook.xml") {
        const modTask = (async (pth, file) => {
          let change = false;
          const buf = await file.async("text");
          const dom = new JSDOM(buf, { contentType: "application/xml" });
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

          /*console.log(
          "workbookProtection: " +
            dom
              .serialize()
              .match(/<workbookProtection>.+?<\/workbookProtection>/)[0],
        );
        console.log(
          "Sheets: " + dom.serialize().match(/<sheets>.+?<\/sheets>/)[0],
        );
        buf = dom.serialize();*/
          zip.file(pth, dom.serialize());

          //if (change) outZip.addBuffer(dom.serialize(), file.path); //---!!! IF
          //outZip.addBuffer(Buffer.from(newBuf), file.path);
        })(pth, file);
        modPromises.push(modTask);
      } else if (rgx.test(pth)) {
        const modTask = (async (pth, file) => {
          const buf = await file.async("text");
          const dom = new JSDOM(buf, { contentType: "application/xml" });
          const doc = dom.window.document;
          const prt = doc.querySelector("sheetProtection");
          getName(pth, prt ? true : false);
          if (prt) {
            prt.remove();
            //console.log("From", pth, "removed protection:\n");
            zip.file(pth, dom.serialize());
          }
        })(pth, file);
        modPromises.push(modTask);
      }
    });

    await Promise.all(modPromises);
    const newZipData = await zip.generateAsync({ type: "nodebuffer" });
    fs.writeFileSync("rep.xlsx", newZipData);
    show();
  } catch (err) {
    console.log("ERROR: " + err.message);
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
  console.log("> rep.xlsx created");
}

modify("old.xlsx");
