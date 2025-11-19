const unzipper = require("unzipper");
const yazl = require("yazl");
const fs = require("fs");
const { JSDOM } = require("jsdom");

const archive = "old.xlsx";
const foi = [];

async function rebuildZip() {
  const directory = await unzipper.Open.file(archive);
  const outZip = new yazl.ZipFile();

  const rgx = /xl\/worksheets\/sheet\d+\.xml/;
  for (const file of directory.files) {
    if (file.path === "xl/workbook.xml") {
      let change = false;
      console.log(file.path, "found!");
      const buf = await file.buffer();
      const dom = new JSDOM(buf, { contentType: "application/xml" });
      const doc = dom.window.document;

      const wkProt = doc.querySelector("workbookProtection");
      if (wkProt) {
        wkProt.remove();
        change = true;
        console.log(file.path, "protected");
      }

      const sheets = doc.querySelectorAll("sheets > sheet");
      for (const sheet of sheets) {
        foi.push({
          name: sheet.getAttribute("name"),
          state: sheet.getAttribute("state"),
          id: sheet.getAttribute("r:id"),
        });
      }

      //console.log("sheets:", foi.length, JSON.stringify(foi));

      for (const sheet of sheets) {
        if (sheet.getAttribute("state") === "hidden") {
          sheet.removeAttribute("state");
          change = true;
          //console.log(sheet.getAttribute("name"), " hidden");
        }
      }

      //if (change) outZip.addBuffer(dom.serialize(), file.path); //---!!! IF
      //outZip.addBuffer(Buffer.from(newBuf), file.path);
    } else if (rgx.test(file.path)) {
      const buf = await file.buffer();
      const dom = new JSDOM(buf, { contentType: "application/xml" });
      const doc = dom.window.document;
      const prt = doc.querySelector("sheetProtection");
      /*console.log(
        getName(file.path, prt ? true : false),
        file.path,
        prt ? "protected" : "---",
      );*/
      getName(file.path, prt ? true : false);
      if (prt) {
        prt.remove();
        //outZip.addBuffer(dom.serialize(), file.path); //---!!! IF
      }
    }
  }

  show();
  outZip.end();
  outZip.outputStream.pipe(fs.createWriteStream("rep.xlsx"));
}

function getName(nm, flag) {
  const num = parseInt(nm.match(/\d+\.xml$/)[0]) - 1;
  //console.log(num + 1, foi[num].name);
  foi[num].prot = flag;
  //return foi[num].name;
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
  }
}

rebuildZip();
