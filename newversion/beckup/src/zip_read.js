function readU16(dv, o) { return dv.getUint16(o, true); }
function readU32(dv, o) { return dv.getUint32(o, true); }

function findEOCD(dv) {
  const sig = 0x06054b50;
  const maxBack = Math.min(dv.byteLength, 22 + 0xFFFF);
  for (let i = dv.byteLength - 22; i >= dv.byteLength - maxBack; i -= 1) {
    if (i < 0) break;
    if (readU32(dv, i) === sig) return i;
  }
  return -1;
}

async function inflateRawBytes(u8) {
  if (typeof DecompressionStream !== 'function') {
    throw new Error('ZIP DEFLATE requires DecompressionStream support');
  }
  const tryAlg = async (alg) => {
    const ds = new DecompressionStream(alg);
    const ab = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(ab);
  };
  try { return await tryAlg('deflate-raw'); }
  catch (_) { return await tryAlg('deflate'); }
}

/**
 * Read ZIP entries (STORE + DEFLATE).
 * @param {ArrayBuffer} arrayBuffer
 * @returns {Promise<Array<{name:string,data:Uint8Array}>>}
 */
export async function unzipEntries(arrayBuffer) {
  const dv = new DataView(arrayBuffer);
  const eocdOff = findEOCD(dv);
  if (eocdOff < 0) throw new Error('ZIP EOCD not found');

  const cdSize = readU32(dv, eocdOff + 12);
  const cdOff = readU32(dv, eocdOff + 16);
  let p = cdOff;
  const files = [];

  while (p < cdOff + cdSize) {
    if (readU32(dv, p) !== 0x02014b50) throw new Error('Invalid central directory entry');

    const compMethod = readU16(dv, p + 10);
    const compSize = readU32(dv, p + 20);
    const nameLen = readU16(dv, p + 28);
    const extraLen = readU16(dv, p + 30);
    const commentLen = readU16(dv, p + 32);
    const localOff = readU32(dv, p + 42);

    const nameBytes = new Uint8Array(arrayBuffer, p + 46, nameLen);
    const name = new TextDecoder().decode(nameBytes);

    if (readU32(dv, localOff) !== 0x04034b50) throw new Error('Invalid local header entry');
    const lNameLen = readU16(dv, localOff + 26);
    const lExtraLen = readU16(dv, localOff + 28);
    const dataOff = localOff + 30 + lNameLen + lExtraLen;
    const compData = new Uint8Array(arrayBuffer, dataOff, compSize);

    let data;
    if (compMethod === 0) data = compData;
    else if (compMethod === 8) data = await inflateRawBytes(compData);
    else throw new Error(`Unsupported ZIP method: ${compMethod} (${name})`);

    files.push({ name, data });
    p += 46 + nameLen + extraLen + commentLen;
  }

  return files;
}

export function entriesMap(entries) {
  const m = new Map();
  for (const e of entries || []) m.set(e.name, e.data);
  return m;
}
