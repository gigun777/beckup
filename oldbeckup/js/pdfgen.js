// js/pdfgen.js
// Minimal offline PDF generator (no external libs).
// Produces a simple PDF with title + table text, using built-in Helvetica.
// Limitations: basic layout, no perfect wrapping.

import { downloadBlob } from "./ui.js";
import { safeName, nowStamp } from "./schema.js";

function pdfEsc(s){
  return String(s??"")
    .replace(/\\/g,"\\\\")
    .replace(/\(/g,"\\(")
    .replace(/\)/g,"\\)")
    .replace(/\r/g,"")
    .replace(/\n/g,"\\n");
}

function pageSizePts(pageSize, orientation){
  // points: 72 pt per inch; A4: 595x842, A3: 842x1191
  let w,h;
  if(pageSize==="A3"){ w=842; h=1191; }
  else { w=595; h=842; }
  if(orientation==="landscape"){ return {w:h,h:w}; }
  return {w,h};
}

export function exportPDFTable({title, subtitle, columns, rows, filenameBase, pageSize="A4", orientation="portrait"}){
  const stamp=nowStamp();
  const fname=`${safeName(filenameBase||"export")}_${stamp}.pdf`;

  const {w,h}=pageSizePts(pageSize, orientation);
  const margin=36;
  const lineH=12;
  const fontSize=10;

  const maxCharsPerLine = Math.floor((w - margin*2) / (fontSize*0.55)); // rough
  const wrap = (s)=>{
    const t=String(s??"");
    if(t.length<=maxCharsPerLine) return [t];
    const out=[];
    let cur="";
    for(const word of t.split(/\s+/)){
      if((cur+" "+word).trim().length>maxCharsPerLine){
        if(cur) out.push(cur);
        cur=word;
      } else cur=(cur?cur+" ":"")+word;
    }
    if(cur) out.push(cur);
    return out.slice(0,6); // cap
  };

  let y = h - margin;
  let content = "";
  const writeLine=(text)=>{
    content += `1 0 0 1 ${margin} ${y} Tm (${pdfEsc(text)}) Tj\n`;
    y -= lineH;
  };

  // Header
  content += "BT\n/F1 "+fontSize+" Tf\n";
  writeLine(title||"");
  if(subtitle){ writeLine(subtitle); }
  y -= 6;

  // Columns header
  writeLine(columns.join(" | "));
  writeLine("-".repeat(Math.min(maxCharsPerLine, 120)));

  for(const r of rows){
    // Flatten row into a single textual line per row (multi-line wrap)
    const parts = columns.map(c=>String(r[c]??""));
    const flat = parts.join(" | ");
    const lines = wrap(flat);
    for(const ln of lines){
      if(y < margin+lineH*2){
        // new page - simplistic: stop, won't paginate in MVP
        writeLine("... (обрізано, збільште експорт або зробимо пагінацію завтра)");
        y = margin; break;
      }
      writeLine(ln);
    }
  }

  content += "ET\n";

  // Build PDF objects
  const objects=[];
  const addObj=(s)=>{ objects.push(s); return objects.length; };

  const fontObj = addObj("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const contentStream = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  const contentObj = addObj(contentStream);
  const pageObj = addObj(`<< /Type /Page /Parent 4 0 R /MediaBox [0 0 ${w} ${h}] /Resources << /Font << /F1 ${fontObj} 0 R >> >> /Contents ${contentObj} 0 R >>`);
  const pagesObj = addObj(`<< /Type /Pages /Kids [${pageObj} 0 R] /Count 1 >>`);
  const catalogObj = addObj(`<< /Type /Catalog /Pages ${pagesObj} 0 R >>`);

  // xref
  let pdf="%PDF-1.4\n";
  const offsets=[0];
  for(let i=0;i<objects.length;i++){
    offsets.push(pdf.length);
    pdf += `${i+1} 0 obj\n${objects[i]}\nendobj\n`;
  }
  const xrefPos=pdf.length;
  pdf += "xref\n0 "+(objects.length+1)+"\n";
  pdf += "0000000000 65535 f \n";
  for(let i=1;i<offsets.length;i++){
    pdf += String(offsets[i]).padStart(10,"0")+" 00000 n \n";
  }
  pdf += `trailer\n<< /Size ${objects.length+1} /Root ${catalogObj} 0 R >>\nstartxref\n${xrefPos}\n%%EOF`;

  downloadBlob(new Blob([pdf],{type:"application/pdf"}), fname);
}
