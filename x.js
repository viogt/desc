

const promises = [];
zip.forEach((path, file) => {
  if (!file.dir && path.endsWith('.txt')) {
    const p = file.async('string').then(text => {
      const newText = text.replace(/foo/g, 'bar');
      zip.file(path, newText); // synchronous overwrite
    });
    promises.push(p);
  }
});
await Promise.all(promises);                // wait for all reads+writes
const out = await zip.generateAsync({ type: 'nodebuffer' });
fs.writeFileSync('archive.zip', out);














const fs = require("fs");
const JSZip = require("jszip");
//const { JSDOM } = require("jsdom");

const archive = "old.xlsx";
const foi = [];

async function modify() {
  const data = fs.readFileSync(archive);
  const zip = await JSZip.loadAsync(data);
  const rgx = /xl\/worksheets\/sheet\d+\.xml/;

  zip.forEach((pth, file) => {
    //console.log("Path:", relativePath, "Is directory:", file.dir);
    //const content = await zip.file(pth).async("string");
    //console.log(pth, file.name, "\n", content.slice(0, 100), "\n");

    if (pth === "xl/workbook.xml") {
      let change = false;
      zip
        .file(pth)
        .async("string")
        .then(function success(buf) {
          const rg = /<sheet name="(.*?)".+?\/>/gm;
          buf.replace(rg, (match, nm) => {
            let hd = match.includes('state="hidden"');
            foi.push({ name: nm, state: hd ? "hidden" : "---" });
            console.log(nm, ":", match);
            return match;
          });
          console.log(foi);
        });

      /*const rg = /<sheet name="(.*?)".+?\/>/gm;
      buf.replace(rg, (match, nm) => {
        let hd = match.includes('state="hidden"');
        foi.push({ name: nm, state: hd ? "hidden" : "---" });
        console.log(nm, ":", match);
        return match;
      });*/
    }
  });

  console.log(foi);

  //const newZipData = await zip.generateAsync({ type: "nodebuffer" });
  //fs.writeFileSync("rep.xlsx", newZipData);

  show();
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

modify();
