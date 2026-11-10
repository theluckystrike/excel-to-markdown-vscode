'use strict';
// Zip bomb guard for .xlsx, .xlsm and .ods files, run before SheetJS sees the bytes. It reads the zip central
// directory, caps the entry count and the total unpacked size, and inflates every entry with a hard output cap of the
// size the directory claims, so an entry that unpacks to more than it says is refused instead of filling memory.
const zlib = require('zlib');

const ZIP_LIMITS = { maxEntries: 10000, maxTotal: 600 * 1024 * 1024, maxEntry: 400 * 1024 * 1024, maxRatio: 1000 };

class ZipError extends Error { constructor(m) { super(m); this.code = 'ZIP'; } }

function isZip(buf) { return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && (buf[2] === 3 || buf[2] === 5) ; }

/** Checks a zip buffer. Returns { entries, total } or throws ZipError with a plain message. */
function checkZip(buf, limits = {}) {
  const L = { ...ZIP_LIMITS, ...limits };
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf.buffer || buf, buf.byteOffset || 0, buf.byteLength);
  // end of central directory: in the last 64 KB + 22 bytes
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new ZipError('The file is not a complete zip archive (no central directory). It may be cut off.');
  const count = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOff = buf.readUInt32LE(eocd + 16);
  if (count === 0xffff || cdOff === 0xffffffff) throw new ZipError('This workbook uses ZIP64 (over 4 GB or 65,535 parts) and is not read.');
  if (count > L.maxEntries) throw new ZipError(`The workbook has ${count} parts; up to ${L.maxEntries} are read.`);
  if (cdOff + cdSize > buf.length) throw new ZipError('The zip directory points past the end of the file. The file may be cut off.');
  let p = cdOff, total = 0;
  const entries = [];
  for (let n = 0; n < count; n++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw new ZipError('The zip directory is damaged.');
    const method = buf.readUInt16LE(p + 10);
    const csize = buf.readUInt32LE(p + 20);
    const usize = buf.readUInt32LE(p + 24);
    const nlen = buf.readUInt16LE(p + 28), xlen = buf.readUInt16LE(p + 30), clen = buf.readUInt16LE(p + 32);
    const off = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, Math.min(buf.length, p + 46 + nlen));
    p += 46 + nlen + xlen + clen;
    if (usize > L.maxEntry) throw new ZipError(`The part ${name.slice(0, 80)} unpacks to ${Math.round(usize / 1048576)} MB; up to ${Math.round(L.maxEntry / 1048576)} MB per part is read.`);
    total += usize;
    if (total > L.maxTotal) throw new ZipError(`The workbook unpacks to more than ${Math.round(L.maxTotal / 1048576)} MB and is not read.`);
    if (csize > 0 && usize / csize > L.maxRatio && usize > 1048576) throw new ZipError(`The part ${name.slice(0, 80)} is packed ${Math.round(usize / csize)} times smaller than it unpacks. It looks like a zip bomb and is not read.`);
    entries.push({ name, method, csize, usize, off });
  }
  // inflate each entry with the claimed size as the cap
  for (const e of entries) {
    if (e.off + 30 > buf.length || buf.readUInt32LE(e.off) !== 0x04034b50) throw new ZipError(`The part ${e.name.slice(0, 80)} is damaged.`);
    const start = e.off + 30 + buf.readUInt16LE(e.off + 26) + buf.readUInt16LE(e.off + 28);
    if (start + e.csize > buf.length) throw new ZipError(`The part ${e.name.slice(0, 80)} is cut off.`);
    if (e.method === 0) { if (e.csize !== e.usize) throw new ZipError(`The part ${e.name.slice(0, 80)} is damaged.`); continue; }
    if (e.method !== 8) throw new ZipError(`The part ${e.name.slice(0, 80)} uses a zip method (${e.method}) that Excel does not write.`);
    let out;
    try { out = zlib.inflateRawSync(buf.subarray(start, start + e.csize), { maxOutputLength: Math.max(1, e.usize + 1) }); } catch (err) {
      if (err && (err.code === 'ERR_BUFFER_TOO_LARGE' || err instanceof RangeError)) throw new ZipError(`The part ${e.name.slice(0, 80)} unpacks to more than the ${e.usize} bytes it claims. It is not read.`);
      throw new ZipError(`The part ${e.name.slice(0, 80)} could not be unpacked: ${err && err.message}`);
    }
    if (out.length > e.usize) throw new ZipError(`The part ${e.name.slice(0, 80)} unpacks to more than it claims. It is not read.`);
  }
  return { entries: entries.length, total };
}

module.exports = { checkZip, isZip, ZipError, ZIP_LIMITS };
