// Generate simple placeholder PNG icons for the add-in
const fs = require("fs");
const path = require("path");

// Minimal 1x1 blue PNG
const bluePixel = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==",
  "base64"
);

const sizes = [16, 32, 80];
const dir = path.join(__dirname, "assets");

if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

// Use SVG to create a simple icon, then encode as base64 data URI
// For Office add-ins, we can actually use SVG or simple colored PNGs

// Let's create a valid minimal PNG for each size using canvas-like approach
// Since we don't have canvas, we'll create a tiny valid PNG

function createPNG(size) {
  // This creates a minimal valid PNG with a blue rectangle
  // For production, replace with real icons
  const IHDR = createIHDR(size, size);
  const IDAT = createIDAT(size, size);
  const IEND = createChunk("IEND", Buffer.alloc(0));

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), // PNG signature
    IHDR,
    IDAT,
    IEND,
  ]);
}

function createIHDR(width, height) {
  const data = Buffer.alloc(13);
  data.writeUInt32BE(width, 0);
  data.writeUInt32BE(height, 4);
  data[8] = 8; // bit depth
  data[9] = 2; // color type RGB
  data[10] = 0; // compression
  data[11] = 0; // filter
  data[12] = 0; // interlace
  return createChunk("IHDR", data);
}

function createIDAT(width, height) {
  const zlib = require("zlib");
  // Create raw image data: blue pixels
  const rawRows = [];
  for (let y = 0; y < height; y++) {
    rawRows.push(0); // filter byte
    for (let x = 0; x < width; x++) {
      rawRows.push(0x42); // R
      rawRows.push(0x85); // G
      rawRows.push(0xf4); // B
    }
  }
  const compressed = zlib.deflateSync(Buffer.from(rawRows));
  return createChunk("IDAT", compressed);
}

function createChunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);

  const crcInput = Buffer.concat([typeBytes, data]);
  const crc = crc32(crcInput);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc >>> 0, 0);

  return Buffer.concat([length, typeBytes, data, crcBuf]);
}

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return crc ^ 0xffffffff;
}

for (const size of sizes) {
  const png = createPNG(size);
  const filePath = path.join(dir, `icon-${size}.png`);
  fs.writeFileSync(filePath, png);
  console.log(`Created ${filePath}`);
}

console.log("Icons generated successfully");
