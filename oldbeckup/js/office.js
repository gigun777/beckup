// js/office.js
// Minimal offline DOCX/XLSX generator (Office Open XML) using our ZIP STORE writer.
// Goal: "real" .docx/.xlsx files that open in Word/Excel.
// Limitations: simple tables, no styling beyond basics.

import { makeZipStore } from "./zip.js";
import { safeName, nowStamp } from "./schema.js";
import { downloadBlob } from "./ui.js";

const te = new TextEncoder();

function xmlEsc(s){
  return String(s??"")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;").replace(/'/g,"&apos;");
}

export function exportDOCXTable({title, subtitle, columns, rows, filenameBase}){
  const stamp = nowStamp();
  const fname = `${safeName(filenameBase||"export")}_${stamp}.docx`;

  const tblRows = rows.map(r=>{
    const tds = columns.map(c=>{
      const v = xmlEsc(r[c] ?? "");
      return `<w:tc><w:tcPr/><w:p><w:r><w:t xml:space="preserve">${v}</w:t></w:r></w:p></w:tc>`;
    }).join("");
    return `<w:tr>${tds}</w:tr>`;
  }).join("");

  const headerRow = `<w:tr>` + columns.map(c=>{
    return `<w:tc><w:tcPr/><w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${xmlEsc(c)}</w:t></w:r></w:p></w:tc>`;
  }).join("") + `</w:tr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
  <w:body>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">${xmlEsc(title||"")}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">${xmlEsc(subtitle||"")}</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tblGrid>${columns.map(()=>'<w:gridCol w:w="2400"/>').join("")}</w:tblGrid>
      ${headerRow}
      ${tblRows}
    </w:tbl>
    <w:p/>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
  </w:body>
</w:document>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

  const zipBytes = makeZipStore([
    {name:"[Content_Types].xml", data: te.encode(contentTypes)},
    {name:"_rels/.rels", data: te.encode(rels)},
    {name:"word/document.xml", data: te.encode(documentXml)},
    {name:"word/_rels/document.xml.rels", data: te.encode(docRels)},
  ]);

  downloadBlob(new Blob([zipBytes], {type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}), fname);
}

function colLetter(n){
  let s="";
  while(n>0){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}
function isInt(v){ return /^-?\d+$/.test(String(v??"").trim()); }

export function exportXLSXTable({title, columns, rows, filenameBase}){
  const stamp = nowStamp();
  const fname = `${safeName(filenameBase||"export")}_${stamp}.xlsx`;

  const sheetName = "Sheet1";

  // Build worksheet with inline strings
  let sheetRows = "";
  const headerCells = columns.map((c,ci)=>{
    const addr = colLetter(ci+1)+"1";
    return `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(c)}</t></is></c>`;
  }).join("");
  sheetRows += `<row r="1">${headerCells}</row>`;

  for(let ri=0; ri<rows.length; ri++){
    const rIndex = ri+2;
    const row = rows[ri];
    const cells = columns.map((c,ci)=>{
      const addr = colLetter(ci+1)+String(rIndex);
      const v = row[c] ?? "";
      if(isInt(v)){
        return `<c r="${addr}" t="n"><v>${String(v).trim()}</v></c>`;
      }
      return `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
    }).join("");
    sheetRows += `<row r="${rIndex}">${cells}</row>`;
  }

  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    ${sheetRows}
  </sheetData>
</worksheet>`;

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

  const zipBytes = makeZipStore([
    {name:"[Content_Types].xml", data: te.encode(contentTypes)},
    {name:"_rels/.rels", data: te.encode(rels)},
    {name:"xl/workbook.xml", data: te.encode(workbookXml)},
    {name:"xl/_rels/workbook.xml.rels", data: te.encode(workbookRels)},
    {name:"xl/worksheets/sheet1.xml", data: te.encode(worksheetXml)},
  ]);

  downloadBlob(new Blob([zipBytes], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), fname);
}
