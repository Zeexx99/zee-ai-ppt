import { mkdirSync, writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const assetsDir = join(__dirname, "..", "public", "assets");

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let j = 0; j < 8; j++) {
      c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0);
    }
  }
  return (c ^ -1) >>> 0;
}

function makeChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, "ascii");
  const crcBuf = Buffer.alloc(4);
  const combined = Buffer.concat([typeBuf, data]);
  crcBuf.writeUInt32BE(crc32(combined), 0);
  return Buffer.concat([len, combined, crcBuf]);
}

function createIconPng(size) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  const radius = Math.max(2, Math.round(size * 0.2));

  for (let y = 0; y < size; y++) {
    const rowStart = y * (size * 4 + 1);
    raw[rowStart] = 0; // Filter type 0 (None)
    for (let x = 0; x < size; x++) {
      const idx = rowStart + 1 + x * 4;

      // Rounded rectangle mask
      let inside = true;
      if (x < radius && y < radius) {
        inside = Math.hypot(x - radius, y - radius) <= radius;
      } else if (x >= size - radius && y < radius) {
        inside = Math.hypot(x - (size - radius - 1), y - radius) <= radius;
      } else if (x < radius && y >= size - radius) {
        inside = Math.hypot(x - radius, y - (size - radius - 1)) <= radius;
      } else if (x >= size - radius && y >= size - radius) {
        inside = Math.hypot(x - (size - radius - 1), y - (size - radius - 1)) <= radius;
      }

      if (!inside) {
        raw[idx] = 0;
        raw[idx + 1] = 0;
        raw[idx + 2] = 0;
        raw[idx + 3] = 0;
        continue;
      }

      // PowerPoint coral-orange gradient (#D24726 -> #A92B12)
      const t = (x + y) / (2 * size);
      let r = Math.round(218 - t * 48);
      let g = Math.round(75 - t * 32);
      let b = Math.round(40 - t * 20);

      // Draw bold white "Z" glyph in the center
      const nx = x / size;
      const ny = y / size;
      const stroke = Math.max(0.11, 1.8 / size);

      const inTopBar = ny >= 0.24 && ny <= 0.24 + stroke && nx >= 0.24 && nx <= 0.76;
      const inBottomBar = ny >= 0.76 - stroke && ny <= 0.76 && nx >= 0.24 && nx <= 0.76;
      const diagY = 0.76 - ((nx - 0.24) / 0.52) * 0.52;
      const inDiag =
        nx >= 0.24 &&
        nx <= 0.76 &&
        ny >= 0.24 &&
        ny <= 0.76 &&
        Math.abs(ny - diagY) <= stroke * 0.72;

      // Small AI sparkle dot at top-right
      const inSparkle = Math.hypot(nx - 0.78, ny - 0.2) <= Math.max(0.055, 1.1 / size);

      if (inTopBar || inBottomBar || inDiag) {
        r = 255;
        g = 255;
        b = 255;
      } else if (inSparkle) {
        r = 255;
        g = 228;
        b = 148;
      }

      raw[idx] = r;
      raw[idx + 1] = g;
      raw[idx + 2] = b;
      raw[idx + 3] = 255;
    }
  }

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const compressed = deflateSync(raw);
  return Buffer.concat([
    signature,
    makeChunk("IHDR", ihdr),
    makeChunk("IDAT", compressed),
    makeChunk("IEND", Buffer.alloc(0)),
  ]);
}

mkdirSync(assetsDir, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  const filePath = join(assetsDir, `icon-${size}.png`);
  writeFileSync(filePath, createIconPng(size));
  console.log(`Generated ${filePath}`);
}
