function u16(n) {
  return new Uint8Array([n & 255, (n >>> 8) & 255]);
}

function u32(n) {
  return new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
}

function concatBytes(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i += 1) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ (-1)) >>> 0;
}

function dosTimeDate(date) {
  const d = date || new Date();

  let time = 0;
  time |= (Math.floor(d.getSeconds() / 2) & 31);
  time |= (d.getMinutes() & 63) << 5;
  time |= (d.getHours() & 31) << 11;

  let dt = 0;
  dt |= (d.getDate() & 31);
  dt |= ((d.getMonth() + 1) & 15) << 5;
  dt |= ((d.getFullYear() - 1980) & 127) << 9;

  return { time: time & 0xFFFF, date: dt & 0xFFFF };
}

function ensureBytes(data) {
  if (data instanceof Uint8Array) return data;
  if (typeof data === 'string') return new TextEncoder().encode(data);
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  throw new Error('Zip file data must be Uint8Array, ArrayBuffer or string');
}

/**
 * @param {{name:string,data:Uint8Array|ArrayBuffer|string}[]} files
 * @returns {Uint8Array}
 */
export function makeZipStore(files) {
  if (!Array.isArray(files)) throw new Error('files must be an array');

  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const { time, date } = dosTimeDate(new Date());

  for (const file of files) {
    if (!file?.name) throw new Error('Each file must have a name');

    const nameBytes = new TextEncoder().encode(String(file.name));
    const dataBytes = ensureBytes(file.data);
    const crc = crc32(dataBytes);

    const localHeader = concatBytes([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(time),
      u16(date),
      u32(crc),
      u32(dataBytes.length),
      u32(dataBytes.length),
      u16(nameBytes.length),
      u16(0)
    ]);

    localParts.push(localHeader, nameBytes, dataBytes);

    const centralHeader = concatBytes([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(0),
      u16(time),
      u16(date),
      u32(crc),
      u32(dataBytes.length),
      u32(dataBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
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
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(centralDir.length),
    u32(localData.length),
    u16(0)
  ]);

  return concatBytes([localData, centralDir, end]);
}

export function makeZipBlob(files) {
  return new Blob([makeZipStore(files)], { type: 'application/zip' });
}
