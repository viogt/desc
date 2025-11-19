const fs = require("fs");
const JSZip = require("jszip");
const { JSDOM } = require("jsdom");

const archive = "old.xlsx";
const foi = [],
  promises = [];

async function modify() {
  const data = fs.readFileSync(archive);
  const zip = await JSZip.loadAsync(data);
  const rgx = /xl\/worksheets\/sheet\d+\.xml/;

  zip.forEach(async (pth, file) => {
    //console.log("Path:", relativePath, "Is directory:", file.dir);
    //const content = await zip.file(pth).async("string");
    //console.log(pth, file.name, "\n", content.slice(0, 100), "\n");

    if (pth === "xl/workbook.xml") {
      let change = false;

      const p = file.async("string").then((text) => {
        const dom = new JSDOM(text, { contentType: "application/xml" });
        const doc = dom.window.document;
        const wkProt = doc.querySelector("workbookProtection");
        if (wkProt) {
          wkProt.remove();
          change = true;
          console.log(pth, "protected");
          zip.file(pth, dom.serialize());
          promises.push[p];
        }
        //await Promise.all(promises);
        /*zip
          .generateAsync({ type: "nodebuffer" })
          .then((newZipData) => {
            fs.writeFileSync("rep.xlsx", newZipData);
            console.log("> rep.xlsx created");
          })
          .catch((e) => console.log(e.message));*/
        //zip.file(pth, dom.serialize()); // synchronous overwrite
      });
    }

    /*const sheets = doc.querySelectorAll("sheets > sheet");
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
      }*/

    //buf = dom.serialize();
    //zip.file(pth, dom.serialize());
    //if (change) outZip.addBuffer(dom.serialize(), file.path); //---!!! IF
    //outZip.addBuffer(Buffer.from(newBuf), file.path);
    /*    } else if (rgx.test(pth)) {
      const buf = await zip.file(pth).async("nodebuffer");
      const dom = new JSDOM(buf, { contentType: "application/xml" });
      const doc = dom.window.document;
      const prt = doc.querySelector("sheetProtection");
      getName(pth, prt ? true : false);
      if (prt) {
        prt.remove();
        console.log("From", pth, "removed protection:\n");
        zip.file(pth, dom.serialize());
      }
    }*/
  });

  // modify file
  //const newContent = "Hello modified!";
  //zip.file("path/to/file.txt", newContent);

  // await Promise.all(promises);
  zip
    .generateAsync({ type: "nodebuffer" })
    .then((newZipData) => {
      fs.writeFileSync("rep.xlsx", newZipData);
      console.log("> rep.xlsx created");
    })
    .catch((e) => console.log(e.message));

  console.log("rep.xlsx created");

  //show();
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

modify().catch((err) => console.log(err.message));
