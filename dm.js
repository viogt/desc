const { JSDOM } = require("jsdom");
const fs = require("fs");

const file = "output/xl/workbook.xml";
const xml = fs.readFileSync(file, "utf8");
const dom = new JSDOM(xml, { contentType: "application/xml" });
const doc = dom.window.document;

show();

const wkProt = doc.querySelector("workbookProtection");
if (wkProt) wkProt.remove();

const xsheets = doc.querySelectorAll("sheet");
for (const sheet of xsheets) {
  if (sheet.getAttribute("state") === "hidden") sheet.removeAttribute("state");
}

show();

fs.writeFileSync("rep.xml", dom.serialize());

async function show() {
  const sheets = doc.querySelectorAll("sheet");
  for (const sheet of sheets) console.log(sheet.getAttribute("state"));

  const sh = [];
  for (const sheet of sheets) {
    sh.push({
      name: sheet.getAttribute("name"),
      state: sheet.getAttribute("state"),
      id: sheet.getAttribute("r:id"),
    });
  }
  console.log("sheets:", sh.length);
  for (el of sh) console.log(JSON.stringify(el));
}
