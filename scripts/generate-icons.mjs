import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

function crc32(buf) {
  let table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  let crc = 0 ^ -1;
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

function makeChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcBuf = Buffer.alloc(4);
  const crcVal = crc32(Buffer.concat([typeBuf, data]));
  crcBuf.writeUInt32BE(crcVal, 0);
  return Buffer.concat([len, typeBuf, data, crcBuf]);
}

function createIconPng(size) {
  // RGBA buffer (size x size)
  const rawData = Buffer.alloc(size * (size * 4 + 1));
  let offset = 0;

  for (let y = 0; y < size; y++) {
    rawData[offset++] = 0; // Filter type 0: None
    const ny = y / (size - 1 || 1);

    for (let x = 0; x < size; x++) {
      const nx = x / (size - 1 || 1);
      const cx = nx - 0.5;
      const cy = ny - 0.5;
      const dist = Math.sqrt(cx * cx + cy * cy);

      // Rounded rectangle mask (radius ~ 22% of size)
      const cornerR = 0.22;
      const qx = Math.max(0, Math.abs(cx) - (0.5 - cornerR));
      const qy = Math.max(0, Math.abs(cy) - (0.5 - cornerR));
      const cornerDist = Math.sqrt(qx * qx + qy * qy);
      const inBox = cornerDist <= cornerR;

      if (!inBox) {
        rawData[offset++] = 0;
        rawData[offset++] = 0;
        rawData[offset++] = 0;
        rawData[offset++] = 0;
        continue;
      }

      // PowerPoint deep rust/coral/amber theme (#C43E1C -> #D24726 -> #E8633A)
      let r = Math.round(196 + ny * 35);
      let g = Math.round(62 + nx * 30);
      let b = Math.round(28 + (nx + ny) * 15);
      let a = 255;

      // Draw stylized "Z" in white/gold
      // Z spans nx: 0.25 to 0.75, ny: 0.25 to 0.75
      const inTopBar = (ny >= 0.25 && ny <= 0.38 && nx >= 0.25 && nx <= 0.75);
      const inBottomBar = (ny >= 0.62 && ny <= 0.75 && nx >= 0.25 && nx <= 0.75);
      // Diagonal: connects (0.75, 0.35) down to (0.25, 0.65)
      // Line equation: nx = 0.75 - (ny - 0.35) * (0.5 / 0.3) = 0.75 - (ny - 0.35) * 1.666
      const diagCenter = 0.75 - (ny - 0.32) * (0.5 / 0.36);
      const inDiag = (ny >= 0.32 && ny <= 0.68 && Math.abs(nx - diagCenter) <= 0.08);

      // Subtle AI dot/sparkle at top right (nx: 0.78, ny: 0.22)
      const dotDist = Math.hypot(nx - 0.78, ny - 0.22);
      const inDot = dotDist < 0.07;

      if (inTopBar || inBottomBar || inDiag || inDot) {
        // Crisp white / soft gold
        r = 255;
        g = inDot ? 230 : 255;
        b = inDot ? 100 : 255;
      }

      rawData[offset++] = r;
      rawData[offset++] = g;
      rawData[offset++] = b;
      rawData[offset++] = a;
    }
  }

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  const ihdrChunk = makeChunk('IHDR', ihdr);
  const idatChunk = makeChunk('IDAT', zlib.deflateSync(rawData));
  const iendChunk = makeChunk('IEND', Buffer.alloc(0));

  return Buffer.concat([signature, ihdrChunk, idatChunk, iendChunk]);
}

const outDir = path.resolve('public/assets');
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

[16, 32, 64, 80].forEach(size => {
  const buf = createIconPng(size);
  const filePath = path.join(outDir, `icon-${size}.png`);
  fs.writeFileSync(filePath, buf);
  console.log(`Generated ${filePath} (${buf.length} bytes)`);
});
