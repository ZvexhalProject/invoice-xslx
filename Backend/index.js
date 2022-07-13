const {
  exec,
  execSync
} = require("child_process");
const express = require('express');
const path = require('path');
var cors = require('cors')
var fs = require('fs');
const fse = require('fs-extra');
var JSZip = require("jszip");
var rmdir = require('rimraf');
const {
  v4: uuidv4
} = require('uuid');

const app = express();
app.use(cors())
app.use(express.text());
const port = process.env.PORT || 8080;

app.get('/', function (req, res) {
  res.sendFile(path.join(__dirname, '/index.html'));
});

app.post('/invoice', async function (req, res) {

  var id = uuidv4();
  fs.mkdirSync(`./invoice/${id}`);
  fs.writeFileSync(`./invoice/${id}/invoice.xml`, req.body);

  fse.copySync(`../Rechnung1-VorlageOriginal`, `./invoice/${id}/Rechnung1-VorlageOriginal`);

  execSync(`java -cp .\\saxon-he-10.3.jar net.sf.saxon.Transform -t -s:.\\invoice\\${id}\\invoice.xml  -xsl:.\\ExcelCreator.xsl -o:.\\invoice\\${id}\\Rechnung1-VorlageOriginal\\xl\\worksheets\\sheet1.xml`)

  var zip = new JSZip();
  var file = [];


  file.push("docProps/app.xml");
  file.push("xl/_rels/workbook.xml.rels");
  file.push("xl/printerSettings/printerSettings1.bin");
  file.push("xl/theme/theme1.xml");
  file.push("xl/workbook.xml");
  file.push("xl/styles.xml");
  file.push("xl/sharedStrings.xml");
  file.push("xl/calcChain.xml");
  file.push("xl/worksheets/_rels/sheet1.xml.rels");
  file.push("xl/worksheets/sheet1.xml");
  file.push("[Content_Types].xml");
  file.push("_rels/.rels");
  file.push("docProps/core.xml");


  for (var i = 0; i < file.length; i++) {
    zip.file(file[i], fs.readFileSync(`./invoice/${id}/Rechnung1-VorlageOriginal/` + file[i]));
  }


  let content = await zip.generateAsync({
    type: "nodebuffer"
  })

  fs.writeFileSync(`./invoice/${id}/Rechnung.xlsx`, content);
  var data = fs.readFileSync(`./invoice/${id}/Rechnung.xlsx`);
  res.contentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");


  rmdir(`./invoice/${id}/`, function (error) {
    if(error){
      console.log(error);
    }
  })

  return res.send(data);
})
app.listen(port);
console.log('Server started at http://localhost:' + port);