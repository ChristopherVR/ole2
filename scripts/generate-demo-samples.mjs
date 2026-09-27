import { mkdir, readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { buildOle2 } from "../src/ole2-parser-write.ts";
import { buildPptFile } from "../src/legacy-ppt-writer.ts";
import { identifyLegacyOffice } from "../src/legacy-office-detect.ts";
import { inspectLegacyVisio } from "../src/legacy-visio-inspect.ts";
import { inspectLegacyPublisher } from "../src/legacy-publisher-inspect.ts";
import { readOleDocParagraphs } from "../src/ole-document-doc-editor.ts";
import { readOleXlsGrid } from "../src/legacy-excel-biff8.ts";
import { parseOle2 } from "../src/ole2-parser-read.ts";

const args = process.argv.slice(2);
const outputFlag = args.indexOf("--output-dir");
if (outputFlag < 0 || !args[outputFlag + 1]) {
  throw new Error("Usage: bun scripts/generate-demo-samples.mjs --output-dir <path>");
}
const outputDir = path.resolve(args[outputFlag + 1]);
await mkdir(outputDir, { recursive: true });

function bytesOf(buffer) {
  return buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
}

function record(opcode, data) {
  return [opcode & 0xff, opcode >> 8, data.length & 0xff, data.length >> 8, ...data];
}

function bof(type) {
  return record(0x0809, [0x00, 0x06, type & 0xff, type >> 8, ...new Array(12).fill(0)]);
}

const eof = record(0x000a, []);
function sharedStrings(strings) {
  const data = [strings.length, 0, 0, 0, strings.length, 0, 0, 0];
  for (const value of strings) {
    data.push(value.length & 0xff, value.length >> 8, 0x00);
    for (const character of value) data.push(character.charCodeAt(0) & 0xff);
  }
  return record(0x00fc, data);
}

function labelSst(row, col, index) {
  return record(0x00fd, [row & 0xff, row >> 8, col & 0xff, col >> 8, 0, 0, index, 0, 0, 0]);
}

function numberCell(row, col, value) {
  const data = new Uint8Array(8);
  new DataView(data.buffer).setFloat64(0, value, true);
  return record(0x0203, [row & 0xff, row >> 8, col & 0xff, col >> 8, 0, 0, ...data]);
}

function excelSample() {
  const strings = ["Task", "Hours", "Archive index", "Project A", "Project B"];
  const globals = [...bof(0x0005), ...sharedStrings(strings), ...eof];
  const sheet = [
    ...bof(0x0010),
    ...labelSst(0, 0, 0),
    ...labelSst(0, 1, 1),
    ...labelSst(1, 0, 3),
    ...numberCell(1, 1, 6.5),
    ...labelSst(2, 0, 4),
    ...numberCell(2, 1, 3.25),
    ...eof,
  ];
  return bytesOf(buildOle2(new Map([["Workbook", new Uint8Array([...globals, ...sheet])]])));
}

function visioInspectionSample() {
  const stream = new Uint8Array(128);
  const view = new DataView(stream.buffer);
  stream.set(new TextEncoder().encode("Visio (TM) Drawing\r\n"));
  view.setUint16(0x1a, 11, true);
  view.setUint32(0x1c, stream.length, true);
  view.setUint32(0x24, 0x14, true);
  view.setUint32(0x2c, 96, true);
  view.setUint32(0x30, 24, true);
  view.setUint16(0x34, 0x2, true);
  return bytesOf(buildOle2(new Map([["VisioDocument", stream]])));
}

function publisherInspectionSample() {
  return bytesOf(
    buildOle2(new Map([["Contents", new Uint8Array([0xe8, 0xac, 0x22, 0, 0, 0, 0, 0])]])),
  );
}

const samples = [
  [
    "word-97-sample.doc",
    new Uint8Array(
      await readFile(fileURLToPath(new URL("../test/fixtures/ole-word-97.doc", import.meta.url))),
    ),
  ],
  ["excel-biff8-sample.xls", excelSample()],
  [
    "powerpoint-97-sample.ppt",
    bytesOf(
      await buildPptFile({
        widthEmu: 9144000,
        heightEmu: 5143500,
        pictures: [],
        slides: [
          {
            backgroundRgb: "F4F1E9",
            shapes: [
              {
                kind: "shape",
                name: "Sample title",
                spt: 1,
                isConnector: false,
                anchor: { x: 800000, y: 1000000, w: 7600000, h: 1300000 },
                fill: { kind: "solid", rgb: "F4F1E9" },
                line: { kind: "none" },
                text: {
                  textType: 0,
                  paragraphs: [
                    {
                      indentLevel: 0,
                      align: "ctr",
                      runs: [
                        {
                          text: "A tour of OLE2 samples",
                          sizePt: 30,
                          bold: true,
                          colorRgb: "203047",
                        },
                      ],
                    },
                  ],
                },
              },
            ],
          },
        ],
      }),
    ),
  ],
  ["visio-inspection-only.vsd", visioInspectionSample()],
  ["publisher-inspection-only.pub", publisherInspectionSample()],
];

for (const [fileName, bytes] of samples) {
  const identification = identifyLegacyOffice(bytes);
  if (fileName.endsWith(".doc") && !readOleDocParagraphs(bytes)?.length)
    throw new Error("Generated Word sample contains no previewable paragraphs");
  if (fileName.endsWith(".xls") && !readOleXlsGrid(bytes)?.rows.length)
    throw new Error("Generated Excel sample contains no previewable worksheet rows");
  if (
    fileName.endsWith(".ppt") &&
    !parseOle2(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)).getStream(
      "PowerPoint Document",
    )
  )
    throw new Error("Generated PowerPoint sample has no document stream");
  if (fileName.endsWith(".vsd") && !inspectLegacyVisio(bytes))
    throw new Error("Generated Visio inspection sample did not validate");
  if (fileName.endsWith(".pub") && !inspectLegacyPublisher(bytes))
    throw new Error("Generated Publisher inspection sample did not validate");
  await writeFile(path.join(outputDir, fileName), bytes);
  console.log(`${fileName}: ${identification.format} (${bytes.byteLength} bytes)`);
}

console.log(`Generated ${samples.length} sample files in ${outputDir}`);
