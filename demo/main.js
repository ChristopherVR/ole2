import {
  ENTRY_TYPE_STREAM,
  identifyLegacyOffice,
  Ole2ParseError,
  parseOle2,
  readOleDocParagraphs,
  readOleXlsGrid,
} from "./vendor/ole2/index.js";
import { renderMetadata } from "./metadata.js";

const fileInput = document.querySelector("#file-input");
const samplePicker = document.querySelector("#sample-picker");
const loadSampleButton = document.querySelector("#load-sample");
const status = document.querySelector("#status");
const results = document.querySelector("#results");
const streamList = document.querySelector("#streams");
const wordResult = document.querySelector("#word-result");
const wordParagraphs = document.querySelector("#word-paragraphs");
const excelResult = document.querySelector("#excel-result");
const excelGrid = document.querySelector("#excel-grid");
const maximumFileBytes = 100 * 1024 * 1024;
let loadGeneration = 0;

function setStatus(message, kind = "info") {
  status.textContent = message;
  status.dataset.kind = kind;
}

function formatBytes(byteLength) {
  if (byteLength < 1024) return `${byteLength} bytes`;
  if (byteLength < 1024 * 1024) return `${(byteLength / 1024).toFixed(1)} KiB`;
  return `${(byteLength / (1024 * 1024)).toFixed(1)} MiB`;
}

function clearResults() {
  results.hidden = true;
  streamList.replaceChildren();
  wordParagraphs.replaceChildren();
  excelGrid.replaceChildren();
  wordResult.hidden = true;
  excelResult.hidden = true;
  document.querySelector("#metadata-result").hidden = true;
  document.querySelector("#metadata-fields").replaceChildren();
  document.querySelector("#download-metadata").onclick = null;
}

function renderStreams(entries) {
  const rows = entries
    .filter((entry) => entry.type === ENTRY_TYPE_STREAM)
    .sort((first, second) => first.name.localeCompare(second.name));
  for (const entry of rows) {
    const row = document.createElement("tr");
    const name = document.createElement("td");
    const size = document.createElement("td");
    name.textContent = entry.name;
    size.textContent = formatBytes(entry.size);
    row.append(name, size);
    streamList.append(row);
  }
  document.querySelector("#stream-count").textContent = `${rows.length} streams`;
  if (rows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 2;
    cell.textContent = "No streams found.";
    row.append(cell);
    streamList.append(row);
  }
}

function renderWord(bytes) {
  let paragraphs;
  try {
    paragraphs = readOleDocParagraphs(bytes);
  } catch {
    paragraphs = undefined;
  }
  if (!paragraphs) return false;
  wordParagraphs.replaceChildren();
  for (const text of paragraphs) {
    const paragraph = document.createElement("p");
    paragraph.textContent = text || "(empty paragraph)";
    wordParagraphs.append(paragraph);
  }
  wordResult.hidden = false;
  return true;
}

function renderExcel(bytes) {
  let grid;
  try {
    grid = readOleXlsGrid(bytes, 30, 12);
  } catch {
    grid = undefined;
  }
  if (!grid) return false;
  document.querySelector("#sheet-name").textContent = grid.sheetName;
  excelGrid.replaceChildren();
  const table = document.createElement("table");
  table.className = "excel-table";
  const head = document.createElement("thead");
  const headingRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.scope = "col";
  corner.textContent = "#";
  headingRow.append(corner);
  const lastColumn = grid.rows.reduce((max, row) => {
    const used = row.cells.reduce((last, cell, index) => (cell.value ? index : last), -1);
    return Math.max(max, used);
  }, -1);
  for (let col = 0; col <= lastColumn; col++) {
    const heading = document.createElement("th");
    heading.scope = "col";
    heading.textContent = columnName(col);
    headingRow.append(heading);
  }
  head.append(headingRow);
  const body = document.createElement("tbody");
  const lastRow = grid.rows.reduce(
    (last, row, index) => (row.cells.some((cell) => cell.value) ? index : last),
    -1,
  );
  for (let rowIndex = 0; rowIndex <= lastRow; rowIndex++) {
    const row = document.createElement("tr");
    const rowNumber = document.createElement("th");
    rowNumber.scope = "row";
    rowNumber.textContent = String(rowIndex + 1);
    row.append(rowNumber);
    for (let col = 0; col <= lastColumn; col++) {
      const cell = document.createElement("td");
      cell.textContent = grid.rows[rowIndex]?.cells[col]?.value ?? "";
      row.append(cell);
    }
    body.append(row);
  }
  table.append(head, body);
  excelGrid.append(table);
  if (lastRow < 0 || lastColumn < 0) {
    const empty = document.createElement("p");
    empty.className = "empty-grid";
    empty.textContent = "The first worksheet contains no previewable cells.";
    excelGrid.append(empty);
  }
  excelResult.hidden = false;
  return true;
}

function columnName(index) {
  let value = index + 1;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

async function openFile(file, generation) {
  clearResults();
  if (generation !== loadGeneration) return;
  if (file.size > maximumFileBytes) {
    setStatus("For this demo, choose a file smaller than 100 MiB.", "error");
    return;
  }
  setStatus(`Reading ${file.name}…`);
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    if (generation !== loadGeneration) return;
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const compound = parseOle2(view.buffer);
    const identification = identifyLegacyOffice(bytes);
    document.querySelector("#file-heading").textContent = file.name;
    document.querySelector("#file-size").textContent = formatBytes(file.size);
    document.querySelector("#format-heading").textContent =
      `${identification.format.toUpperCase()} · ${identification.support}`;
    document.querySelector("#format-reason").textContent = identification.reason;
    renderStreams(compound.entries);
    renderMetadata(bytes, file.name, setStatus);
    results.hidden = false;
    if (identification.format === "doc" && !renderWord(bytes)) {
      setStatus(
        `${file.name} is a readable CFB file, but its Word main-body text could not be previewed.`,
        "info",
      );
    } else if (identification.format === "xls" && !renderExcel(bytes)) {
      setStatus(
        `${file.name} is a readable CFB file, but its first Excel worksheet could not be previewed.`,
        "info",
      );
    } else {
      setStatus(`${file.name} read successfully. ${identification.reason}`, "success");
    }
  } catch (error) {
    if (generation !== loadGeneration) return;
    const message = error instanceof Error ? error.message : String(error);
    const label =
      error instanceof Ole2ParseError
        ? "This file is not a valid OLE2 compound file."
        : "Could not read this file.";
    setStatus(`${label} ${message}`, "error");
  }
}

async function openSelectedSample() {
  const fileName = samplePicker.value;
  if (!fileName) return;
  const generation = ++loadGeneration;
  loadSampleButton.disabled = true;
  fileInput.value = "";
  clearResults();
  setStatus(`Loading built-in sample ${fileName}…`);
  try {
    const response = await fetch(
      new URL(`./samples/${encodeURIComponent(fileName)}`, import.meta.url),
    );
    if (!response.ok) throw new Error(`Sample file request failed (${response.status}).`);
    const data = await response.arrayBuffer();
    if (generation !== loadGeneration) return;
    const file = new File([data], fileName, {
      type: "application/octet-stream",
    });
    await openFile(file, generation);
  } catch (error) {
    if (generation !== loadGeneration) return;
    clearResults();
    const message = error instanceof Error ? error.message : String(error);
    setStatus(`Could not load built-in sample. ${message}`, "error");
  } finally {
    loadSampleButton.disabled = false;
  }
}

fileInput.addEventListener("change", () => {
  const [file] = fileInput.files ?? [];
  if (file) {
    const generation = ++loadGeneration;
    samplePicker.value = "";
    fileInput.value = "";
    void openFile(file, generation);
  }
});
loadSampleButton.addEventListener("click", () => void openSelectedSample());
void openSelectedSample();
