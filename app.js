const fileInput = document.getElementById("fileInput");
const sheetSelect = document.getElementById("sheetSelect");
const preview = document.getElementById("preview");
const fileName = document.getElementById("fileName");
const sheetCount = document.getElementById("sheetCount");
const status = document.getElementById("status");
const previewHint = document.getElementById("previewHint");
const printPdfButton = document.getElementById("printPdf");
const printDirectButton = document.getElementById("printDirect");
const configInput = document.getElementById("configInput");
const generateTemplateButton = document.getElementById("generateTemplate");

const state = {
  workbook: null,
};

const resetView = () => {
  sheetSelect.innerHTML = "<option value=\"\">请先导入文件</option>";
  sheetSelect.disabled = true;
  preview.innerHTML = "";
  previewHint.textContent = "导入 Excel 文件后显示内容。";
  sheetCount.textContent = "";
  status.textContent = "";
  printPdfButton.disabled = true;
  printDirectButton.disabled = true;
};

const renderTable = (sheetName) => {
  const worksheet = state.workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

  if (!rows.length) {
    preview.innerHTML = "<p class=\"empty\">当前工作表为空。</p>";
    previewHint.textContent = "";
    return;
  }

  const table = document.createElement("table");
  const headerRow = document.createElement("tr");
  const headerCells = rows[0];

  headerCells.forEach((cell) => {
    const th = document.createElement("th");
    th.textContent = cell ?? "";
    headerRow.appendChild(th);
  });

  const thead = document.createElement("thead");
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.slice(1).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell ?? "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  preview.innerHTML = "";
  preview.appendChild(table);
  previewHint.textContent = `当前工作表：${sheetName}`;
};

const renderTemplateTable = (headers) => {
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  const emptyRow = document.createElement("tr");
  headers.forEach(() => {
    const td = document.createElement("td");
    td.textContent = "";
    emptyRow.appendChild(td);
  });
  tbody.appendChild(emptyRow);
  table.appendChild(tbody);

  preview.innerHTML = "";
  preview.appendChild(table);
  previewHint.textContent = "模板预览（来自配置）";
};

const parseTemplateFromConfig = (text) => {
  const rowMatch = text.match(/tableRow2\\.Cells\\.AddRange\\([\\s\\S]*?\\);/);
  if (!rowMatch) {
    return null;
  }

  const cellNames = Array.from(rowMatch[0].matchAll(/this\\.(tableCell\\d+)/g)).map(
    (match) => match[1]
  );

  if (!cellNames.length) {
    return null;
  }

  const headers = cellNames.map((cellName) => {
    const labelMatch = text.match(
      new RegExp(`${cellName}\\.Text = \\"([^\\"]+)\\";`)
    );
    return labelMatch ? labelMatch[1] : cellName;
  });

  return headers;
};

const populateSheets = () => {
  sheetSelect.innerHTML = "";
  state.workbook.SheetNames.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    sheetSelect.appendChild(option);
  });
  sheetSelect.disabled = false;
  renderTable(state.workbook.SheetNames[0]);
};

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;

  if (!file) {
    resetView();
    fileName.textContent = "未选择文件";
    return;
  }

  fileName.textContent = file.name;
  status.textContent = "";

  const reader = new FileReader();
  reader.onload = (loadEvent) => {
    const data = new Uint8Array(loadEvent.target.result);
    state.workbook = XLSX.read(data, { type: "array" });
    sheetCount.textContent = `共 ${state.workbook.SheetNames.length} 个工作表`;
    populateSheets();
    printPdfButton.disabled = false;
    printDirectButton.disabled = false;
  };
  reader.onerror = () => {
    status.textContent = "读取文件失败，请重试。";
    resetView();
  };
  reader.readAsArrayBuffer(file);
});

sheetSelect.addEventListener("change", (event) => {
  const sheetName = event.target.value;
  if (!sheetName) {
    return;
  }
  renderTable(sheetName);
});

const handlePrint = () => {
  if (!state.workbook) {
    const trimmedConfig = configInput.value.trim();
    if (!trimmedConfig) {
      status.textContent = "请先导入 Excel 文件或生成模板。";
      return;
    }
  }
  status.textContent = "";
  window.print();
};

generateTemplateButton.addEventListener("click", () => {
  const trimmedConfig = configInput.value.trim();
  if (!trimmedConfig) {
    status.textContent = "请先粘贴配置内容。";
    return;
  }

  const headers = parseTemplateFromConfig(trimmedConfig);
  if (!headers) {
    status.textContent = "未识别到模板表头，请检查配置内容。";
    return;
  }

  status.textContent = "";
  sheetSelect.innerHTML = "<option value=\"\">模板预览</option>";
  sheetSelect.disabled = true;
  fileName.textContent = "模板预览";
  sheetCount.textContent = `共 ${headers.length} 列`;
  renderTemplateTable(headers);
  printPdfButton.disabled = false;
  printDirectButton.disabled = false;
});

printPdfButton.addEventListener("click", handlePrint);
printDirectButton.addEventListener("click", handlePrint);

resetView();
