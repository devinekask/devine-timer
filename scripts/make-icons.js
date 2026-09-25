// Generates the add-in icons (a simple stopwatch) as PNGs without any dependencies.
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}

function png(size, pixel) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    for (let x = 0; x < size; x++) {
      // 4x4 supersampling for smooth edges
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < 4; sy++) for (let sx = 0; sx < 4; sx++) {
        const p = pixel((x + (sx + 0.5) / 4) / size, (y + (sy + 0.5) / 4) / size);
        r += p[0] * p[3]; g += p[1] * p[3]; b += p[2] * p[3]; a += p[3];
      }
      const i = y * (size * 4 + 1) + 1 + x * 4;
      raw[i] = a ? r / a : 0; raw[i + 1] = a ? g / a : 0; raw[i + 2] = a ? b / a : 0; raw[i + 3] = (a / 16) * 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const ORANGE = [194, 65, 12, 1], WHITE = [255, 255, 255, 1], NONE = [0, 0, 0, 0];

function stopwatch(u, v) {
  const cx = 0.5, cy = 0.56, R = 0.4;
  const dx = u - cx, dy = v - cy, d = Math.hypot(dx, dy);
  // crown button on top
  if (Math.abs(u - cx) < 0.07 && v > 0.04 && v < 0.16) return ORANGE;
  if (d > R) return NONE;
  if (d > R - 0.07) return ORANGE;       // ring
  // hand pointing up-right
  const t = (dx * 0.6 + -dy * 0.8);      // projection on hand direction
  const perp = Math.abs(dx * 0.8 + dy * 0.6);
  if (t > -0.02 && t < R - 0.13 && perp < 0.045) return ORANGE;
  if (d < 0.06) return ORANGE;           // hub
  return WHITE;
}

const out = path.join(__dirname, "..", "docs", "assets");
fs.mkdirSync(out, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  fs.writeFileSync(path.join(out, `icon-${size}.png`), png(size, stopwatch));
}
console.log("icons written to", out);
