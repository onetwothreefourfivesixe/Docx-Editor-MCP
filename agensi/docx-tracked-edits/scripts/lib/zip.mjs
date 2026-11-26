// Minimal ZIP reader and writer for .docx packages, built on Node's zlib.
// Uses no dependencies, so every line can be audited.
//
// Reading uses the central directory, so entries that use data descriptors
// work. Entries that aren't changed are copied byte for byte, still
// compressed. Encrypted archives and ZIP64 archives (over 4 GB) are rejected.

import zlib from 'node:zlib';

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_END = 0x06054b50;
const FLAG_ENCRYPTED = 0x0001;
const FLAG_UTF8 = 0x0800;

const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

export function crc32(buffer) {
    let crc = 0xffffffff;
    for (let i = 0; i < buffer.length; i++) crc = CRC_TABLE[(crc ^ buffer[i]) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
}

/**
 * @typedef {object} ZipEntry
 * @property {string} name
 * @property {Buffer} nameBytes
 * @property {number} flags
 * @property {number} method   0 = stored, 8 = deflate
 * @property {number} time     MS-DOS time
 * @property {number} date     MS-DOS date
 * @property {number} crc
 * @property {number} size     uncompressed size
 * @property {Buffer} data     compressed bytes
 */

/** @returns {ZipEntry[]} */
export function readZip(buffer) {
    // The end-of-central-directory record is in the last 22 + 65535 bytes.
    let end = -1;
    for (let i = buffer.length - 22; i >= Math.max(0, buffer.length - 22 - 0xffff); i--) {
        if (buffer.readUInt32LE(i) === SIG_END) {
            end = i;
            break;
        }
    }
    if (end < 0) throw new Error('Not a ZIP file (no end of central directory found)');

    const count = buffer.readUInt16LE(end + 10);
    const directoryOffset = buffer.readUInt32LE(end + 16);
    if (count === 0xffff || directoryOffset === 0xffffffff) throw new Error('ZIP64 archives are not supported');

    const entries = [];
    let p = directoryOffset;
    for (let n = 0; n < count; n++) {
        if (buffer.readUInt32LE(p) !== SIG_CENTRAL) throw new Error('Corrupt ZIP central directory');
        const flags = buffer.readUInt16LE(p + 8);
        const method = buffer.readUInt16LE(p + 10);
        const time = buffer.readUInt16LE(p + 12);
        const date = buffer.readUInt16LE(p + 14);
        const crc = buffer.readUInt32LE(p + 16);
        const compressedSize = buffer.readUInt32LE(p + 20);
        const size = buffer.readUInt32LE(p + 24);
        const nameLength = buffer.readUInt16LE(p + 28);
        const extraLength = buffer.readUInt16LE(p + 30);
        const commentLength = buffer.readUInt16LE(p + 32);
        const localOffset = buffer.readUInt32LE(p + 42);
        const nameBytes = Buffer.from(buffer.subarray(p + 46, p + 46 + nameLength));

        if (flags & FLAG_ENCRYPTED) throw new Error('Encrypted documents are not supported');
        if (buffer.readUInt32LE(localOffset) !== SIG_LOCAL) throw new Error('Corrupt ZIP local header');
        const dataStart =
            localOffset + 30 + buffer.readUInt16LE(localOffset + 26) + buffer.readUInt16LE(localOffset + 28);

        entries.push({
            name: nameBytes.toString(flags & FLAG_UTF8 ? 'utf8' : 'latin1'),
            nameBytes,
            flags,
            method,
            time,
            date,
            crc,
            size,
            data: Buffer.from(buffer.subarray(dataStart, dataStart + compressedSize)),
        });
        p += 46 + nameLength + extraLength + commentLength;
    }
    return entries;
}

/** Uncompressed contents of an entry. */
export function entryContent(entry) {
    if (entry.method === 0) return entry.data;
    if (entry.method === 8) return zlib.inflateRawSync(entry.data);
    throw new Error(`Unsupported compression method ${entry.method} for ${entry.name}`);
}

/** Replaces an entry's contents (deflate-compressed) and stamps it with the current time. */
export function setEntryContent(entry, content) {
    const now = new Date();
    entry.method = 8;
    entry.data = zlib.deflateRawSync(content);
    entry.crc = crc32(content);
    entry.size = content.length;
    entry.time = (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
    entry.date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
}

/** Creates a new entry from uncompressed content. */
export function makeEntry(name, content) {
    const nameBytes = Buffer.from(name, 'utf8');
    const entry = { name, nameBytes, flags: FLAG_UTF8 };
    setEntryContent(entry, content);
    return entry;
}

/** @param {ZipEntry[]} entries */
export function writeZip(entries) {
    const parts = [];
    const directory = [];
    let offset = 0;

    for (const entry of entries) {
        const flags = entry.flags & FLAG_UTF8; // drop data-descriptor and other bits; sizes are known
        const local = Buffer.alloc(30);
        local.writeUInt32LE(SIG_LOCAL, 0);
        local.writeUInt16LE(20, 4); // version needed: 2.0
        local.writeUInt16LE(flags, 6);
        local.writeUInt16LE(entry.method, 8);
        local.writeUInt16LE(entry.time, 10);
        local.writeUInt16LE(entry.date, 12);
        local.writeUInt32LE(entry.crc, 14);
        local.writeUInt32LE(entry.data.length, 18);
        local.writeUInt32LE(entry.size, 22);
        local.writeUInt16LE(entry.nameBytes.length, 26);
        local.writeUInt16LE(0, 28); // no extra field
        parts.push(local, entry.nameBytes, entry.data);

        const central = Buffer.alloc(46);
        central.writeUInt32LE(SIG_CENTRAL, 0);
        central.writeUInt16LE(20, 4); // version made by
        central.writeUInt16LE(20, 6); // version needed
        central.writeUInt16LE(flags, 8);
        central.writeUInt16LE(entry.method, 10);
        central.writeUInt16LE(entry.time, 12);
        central.writeUInt16LE(entry.date, 14);
        central.writeUInt32LE(entry.crc, 16);
        central.writeUInt32LE(entry.data.length, 20);
        central.writeUInt32LE(entry.size, 24);
        central.writeUInt16LE(entry.nameBytes.length, 28);
        // Host is MS-DOS (version made by = 20), so only the DOS directory bit is meaningful.
        central.writeUInt32LE(entry.name.endsWith('/') ? 0x10 : 0, 38);
        central.writeUInt32LE(offset, 42);
        directory.push(central, entry.nameBytes);

        offset += local.length + entry.nameBytes.length + entry.data.length;
    }

    const directorySize = directory.reduce((sum, part) => sum + part.length, 0);
    const end = Buffer.alloc(22);
    end.writeUInt32LE(SIG_END, 0);
    end.writeUInt16LE(entries.length, 8);
    end.writeUInt16LE(entries.length, 10);
    end.writeUInt32LE(directorySize, 12);
    end.writeUInt32LE(offset, 16);

    return Buffer.concat([...parts, ...directory, end]);
}
