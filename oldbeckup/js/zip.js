// js/zip.js
// Minimal ZIP store writer + reader (STORE only, no compression)
function u16(n){ return new Uint8Array([n & 255, (n>>>8)&255]); }
function u32(n){ return new Uint8Array([n & 255, (n>>>8)&255, (n>>>16)&255, (n>>>24)&255]); }
function concatBytes(chunks){
  const total = chunks.reduce((s,c)=>s+c.length,0);
  const out = new Uint8Array(total);
  let off=0;
  for(const c of chunks){ out.set(c, off); off+=c.length; }
  return out;
}
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i=0;i<256;i++){
    let c=i;
    for(let k=0;k<8;k++) c = (c & 1) ? (0xEDB88320 ^ (c>>>1)) : (c>>>1);
    table[i]=c>>>0;
  }
  return table;
})();
function crc32(bytes){
  let crc = 0 ^ (-1);
  for(let i=0;i<bytes.length;i++){
    crc = (crc>>>8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ (-1))>>>0;
}
function dosTimeDate(date){
  const d=date||new Date();
  let time=0;
  time |= ((Math.floor(d.getSeconds()/2)) & 31);
  time |= (d.getMinutes() & 63) << 5;
  time |= (d.getHours() & 31) << 11;
  let dt=0;
  dt |= (d.getDate() & 31);
  dt |= ((d.getMonth()+1) & 15) << 5;
  dt |= ((d.getFullYear()-1980) & 127) << 9;
  return {time:time&0xFFFF, date:dt&0xFFFF};
}
export function makeZipStore(files){
  const localParts=[], centralParts=[];
  let offset=0;
  const {time,date} = dosTimeDate(new Date());
  for(const f of files){
    const nameBytes=new TextEncoder().encode(f.name);
    const dataBytes=f.data;
    const c=crc32(dataBytes);
    const localHeader = concatBytes([
      u32(0x04034b50), u16(20), u16(0), u16(0),
      u16(time), u16(date),
      u32(c), u32(dataBytes.length), u32(dataBytes.length),
      u16(nameBytes.length), u16(0)
    ]);
    localParts.push(localHeader, nameBytes, dataBytes);
    const centralHeader = concatBytes([
      u32(0x02014b50),
      u16(20), u16(20),
      u16(0), u16(0),
      u16(time), u16(date),
      u32(c), u32(dataBytes.length), u32(dataBytes.length),
      u16(nameBytes.length),
      u16(0), u16(0),
      u16(0), u16(0),
      u32(0),
      u32(offset)
    ]);
    centralParts.push(centralHeader, nameBytes);
    offset += localHeader.length + nameBytes.length + dataBytes.length;
  }
  const centralDir = concatBytes(centralParts);
  const localData = concatBytes(localParts);
  const end = concatBytes([
    u32(0x06054b50),
    u16(0),u16(0),
    u16(files.length),u16(files.length),
    u32(centralDir.length),
    u32(localData.length),
    u16(0)
  ]);
  return concatBytes([localData, centralDir, end]);
}
// Reader
function readU16(dv,o){ return dv.getUint16(o,true); }
function readU32(dv,o){ return dv.getUint32(o,true); }
function findEOCD(dv){
  const sig=0x06054b50;
  const maxBack=Math.min(dv.byteLength, 22 + 0xFFFF);
  for(let i=dv.byteLength-22; i>=dv.byteLength-maxBack; i--){
    if(i<0) break;
    if(readU32(dv,i)===sig) return i;
  }
  return -1;
}
export function unzipStoreEntries(arrayBuffer){
  const dv=new DataView(arrayBuffer);
  const eocdOff=findEOCD(dv);
  if(eocdOff<0) throw new Error("ZIP: EOCD не знайдено");
  const cdSize=readU32(dv, eocdOff+12);
  const cdOff=readU32(dv, eocdOff+16);
  let p=cdOff;
  const files=[];
  while(p < cdOff + cdSize){
    const sig=readU32(dv,p);
    if(sig!==0x02014b50) throw new Error("ZIP: Central Directory пошкоджено");
    const compMethod=readU16(dv,p+10);
    const compSize=readU32(dv,p+20);
    const nameLen=readU16(dv,p+28);
    const extraLen=readU16(dv,p+30);
    const commentLen=readU16(dv,p+32);
    const localOff=readU32(dv,p+42);
    const nameBytes=new Uint8Array(arrayBuffer, p+46, nameLen);
    const name=new TextDecoder().decode(nameBytes);
    const lsig=readU32(dv, localOff);
    if(lsig!==0x04034b50) throw new Error("ZIP: Local Header пошкоджено");
    const lNameLen=readU16(dv, localOff+26);
    const lExtraLen=readU16(dv, localOff+28);
    const dataOff = localOff + 30 + lNameLen + lExtraLen;
    if(compMethod!==0) throw new Error(`ZIP: підтримується лише STORE. Файл: ${name}`);
    const data=new Uint8Array(arrayBuffer, dataOff, compSize);
    files.push({name, data});
    p += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

// Reader (STORE + DEFLATE). Returns Promise.
// NOTE: .xlsx files are ZIP with DEFLATE (method 8).
async function inflateRawBytes(u8){
  if(typeof DecompressionStream!=="function"){
    throw new Error("ZIP: потрібен DecompressionStream (Chrome/Edge) для DEFLATE");
  }
  // ZIP uses raw DEFLATE stream. Some runtimes only support "deflate".
  const tryAlg = async (alg)=>{
    const ds = new DecompressionStream(alg);
    const ab = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(ab);
  };
  try{ return await tryAlg("deflate-raw"); }
  catch(_e){ return await tryAlg("deflate"); }
}

export async function unzipEntries(arrayBuffer){
  const dv=new DataView(arrayBuffer);
  const eocdOff=findEOCD(dv);
  if(eocdOff<0) throw new Error("ZIP: EOCD не знайдено");
  const cdSize=readU32(dv, eocdOff+12);
  const cdOff=readU32(dv, eocdOff+16);
  let p=cdOff;
  const files=[];
  while(p < cdOff + cdSize){
    const sig=readU32(dv,p);
    if(sig!==0x02014b50) throw new Error("ZIP: Central Directory пошкоджено");
    const compMethod=readU16(dv,p+10);
    const compSize=readU32(dv,p+20);
    const uncompSize=readU32(dv,p+24);
    const nameLen=readU16(dv,p+28);
    const extraLen=readU16(dv,p+30);
    const commentLen=readU16(dv,p+32);
    const localOff=readU32(dv,p+42);
    const nameBytes=new Uint8Array(arrayBuffer, p+46, nameLen);
    const name=new TextDecoder().decode(nameBytes);

    const lsig=readU32(dv, localOff);
    if(lsig!==0x04034b50) throw new Error("ZIP: Local Header пошкоджено");
    const lNameLen=readU16(dv, localOff+26);
    const lExtraLen=readU16(dv, localOff+28);
    const dataOff = localOff + 30 + lNameLen + lExtraLen;
    const compData=new Uint8Array(arrayBuffer, dataOff, compSize);

    let data;
    if(compMethod===0){
      data = compData;
    } else if(compMethod===8){
      data = await inflateRawBytes(compData);
      // Some zips may report 0 here; don't hard-fail.
      if(uncompSize && data.length!==uncompSize){
        // keep going
      }
    } else {
      throw new Error(`ZIP: непідтримуваний метод ${compMethod}. Файл: ${name}`);
    }
    files.push({name, data});
    p += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}
