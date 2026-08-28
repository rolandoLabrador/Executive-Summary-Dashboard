import { deflateSync } from 'node:zlib';

export interface PieChartSlice {
  value: number;
  color: string;
}

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const FULL_CIRCLE = Math.PI * 2;

const CRC_TABLE = Array.from({ length: 256 }, (_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit += 1) {
    value = (value & 1) === 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  }
  return value >>> 0;
});

function crc32(data: Buffer): number {
  let checksum = 0xffffffff;
  for (const byte of data) {
    checksum = CRC_TABLE[(checksum ^ byte) & 0xff]! ^ (checksum >>> 8);
  }
  return (checksum ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const typeBuffer = Buffer.from(type, 'ascii');
  const chunk = Buffer.alloc(data.length + 12);
  chunk.writeUInt32BE(data.length, 0);
  typeBuffer.copy(chunk, 4);
  data.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])), data.length + 8);
  return chunk;
}

function rgb(hexColor: string): [number, number, number] {
  const normalized = hexColor.replace(/^#/, '');
  if (!/^[a-f\d]{6}$/i.test(normalized)) {
    throw new Error(`Invalid pie-chart color: ${hexColor}`);
  }
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ];
}

function setPixel(
  pixels: Buffer,
  width: number,
  x: number,
  y: number,
  color: [number, number, number],
): void {
  if (x < 0 || y < 0 || x >= width) return;
  const rowLength = width * 4 + 1;
  const height = pixels.length / rowLength;
  if (y >= height) return;
  const offset = y * rowLength + 1 + x * 4;
  pixels[offset] = color[0];
  pixels[offset + 1] = color[1];
  pixels[offset + 2] = color[2];
  pixels[offset + 3] = 255;
}

function drawSeparator(
  pixels: Buffer,
  width: number,
  centerX: number,
  centerY: number,
  radius: number,
  angle: number,
): void {
  const endX = Math.round(centerX + Math.sin(angle) * radius);
  const endY = Math.round(centerY - Math.cos(angle) * radius);
  const steps = Math.max(Math.abs(endX - centerX), Math.abs(endY - centerY));
  for (let step = 0; step <= steps; step += 1) {
    const progress = steps === 0 ? 0 : step / steps;
    const x = Math.round(centerX + (endX - centerX) * progress);
    const y = Math.round(centerY + (endY - centerY) * progress);
    for (let offsetX = -2; offsetX <= 2; offsetX += 1) {
      for (let offsetY = -2; offsetY <= 2; offsetY += 1) {
        setPixel(pixels, width, x + offsetX, y + offsetY, [255, 255, 255]);
      }
    }
  }
}

/**
 * Produces a dependency-free PNG for embedding in Excel. The image is rendered at
 * high resolution and displayed at half size so the circular edge stays crisp.
 */
export function renderPieChartPng(inputSlices: PieChartSlice[], width = 560, height = 560): Buffer {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 64 || height < 64) {
    throw new Error('Pie-chart dimensions must be integers of at least 64 pixels.');
  }
  const slices = inputSlices.filter((slice) => Number.isFinite(slice.value) && slice.value > 0);
  const total = slices.reduce((sum, slice) => sum + slice.value, 0);
  if (total <= 0) throw new Error('A pie chart requires at least one positive value.');

  const colors = slices.map((slice) => rgb(slice.color));
  const cumulativeAngles: number[] = [];
  let cumulative = 0;
  for (const slice of slices) {
    cumulative += (slice.value / total) * FULL_CIRCLE;
    cumulativeAngles.push(cumulative);
  }

  const rowLength = width * 4 + 1;
  const pixels = Buffer.alloc(rowLength * height);
  const centerX = Math.floor(width / 2);
  const centerY = Math.floor(height / 2);
  const radius = Math.floor(Math.min(width, height) / 2) - 8;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const deltaX = x + 0.5 - centerX;
      const deltaY = y + 0.5 - centerY;
      const distance = Math.hypot(deltaX, deltaY);
      if (distance > radius) continue;
      let angle = Math.atan2(deltaY, deltaX) + Math.PI / 2;
      if (angle < 0) angle += FULL_CIRCLE;
      const sliceIndex = cumulativeAngles.findIndex((boundary) => angle <= boundary);
      setPixel(pixels, width, x, y, colors[sliceIndex === -1 ? colors.length - 1 : sliceIndex]!);
    }
  }

  let separatorAngle = 0;
  for (let index = 0; index < slices.length; index += 1) {
    if (index > 0) drawSeparator(pixels, width, centerX, centerY, radius, separatorAngle);
    separatorAngle += (slices[index]!.value / total) * FULL_CIRCLE;
  }

  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8;
  header[9] = 6;
  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk('IHDR', header),
    pngChunk('IDAT', deflateSync(pixels, { level: 9 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}
