const REQUIRED_HEADERS = ["No", "약품명", "불출량", "집계구분"];
const DISPLAY_HEADERS = [
  "약품명",
  "category",
  "병동",
  "환자번호",
  "환자명",
  "불출량",
  "투약량",
  "집계구분",
];

const PARSE_HEADERS = [
  "No",
  "약품명",
  "사용부서",
  "진료의",
  "병동",
  "응급실",
  "환자번호",
  "환자명",
  "불출량",
  "투약량",
  "단위",
  "집계일자",
  "처방일자",
  "집계구분",
  "투약번호",
];

const state = {
  rows: [],
  drugs: [],
  selectedDrugKey: "",
};

const els = {
  dropZone: document.getElementById("dropZone"),
  fileInput: document.getElementById("fileInput"),
  fileName: document.getElementById("fileName"),
  sheetCount: document.getElementById("sheetCount"),
  rowCount: document.getElementById("rowCount"),
  drugCount: document.getElementById("drugCount"),
  visibleDrugCount: document.getElementById("visibleDrugCount"),
  drugButtons: document.getElementById("drugButtons"),
  dataTable: document.getElementById("dataTable"),
  statsTable: document.getElementById("statsTable"),
  tableTitle: document.getElementById("tableTitle"),
  tableScroll: document.getElementById("tableScroll"),
  resetHighlight: document.getElementById("resetHighlight"),
  totalOut: document.getElementById("totalOut"),
  totalReturn: document.getElementById("totalReturn"),
  totalNet: document.getElementById("totalNet"),
  statsScope: document.getElementById("statsScope"),
};

els.fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) readWorkbook(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.add("drag-over");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.remove("drag-over");
  });
});

els.dropZone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (file) readWorkbook(file);
});

els.resetHighlight.addEventListener("click", () => clearDrugSelection());

async function readWorkbook(file) {
  setLoading(file.name);
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: true,
      codepage: 949,
      raw: false,
    });

    const parsed = workbook.SheetNames.flatMap((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
      });
      return parseSheet(matrix, sheetName);
    });

    state.rows = assignDrugGroupOrder(parsed
      .map((row, index) => ({
        ...row,
        id: `row-${index}`,
        originalIndex: index,
        category: classifyDrug(row["약품명"], row["단위"]),
        drugKey: getDrugKey(row["약품명"]),
        drugLabel: getDrugLabel(row["약품명"]),
        quantity: toNumber(row["불출량"]),
      })))
      .sort(compareRowsByDrugGroup);
    state.drugs = summarizeDrugs(state.rows);
    state.selectedDrugKey = "";

    els.fileName.textContent = file.name;
    els.sheetCount.textContent = workbook.SheetNames.length.toLocaleString("ko-KR");
    els.rowCount.textContent = state.rows.length.toLocaleString("ko-KR");
    els.drugCount.textContent = state.drugs.length.toLocaleString("ko-KR");
    els.tableTitle.textContent = `${state.rows.length.toLocaleString("ko-KR")}행 통합`;

    renderTable();
    renderStats();
    renderDrugButtons();
  } catch (error) {
    console.error(error);
    showError("파일을 읽지 못했습니다. 엑셀 형식과 시트 헤더를 확인하세요.");
  }
}

function setLoading(fileName) {
  els.fileName.textContent = fileName;
  els.sheetCount.textContent = "읽는 중";
  els.rowCount.textContent = "읽는 중";
  els.drugCount.textContent = "읽는 중";
  els.tableTitle.textContent = "분석 중";
}

function showError(message) {
  els.tableTitle.textContent = "분석 실패";
  els.sheetCount.textContent = "0";
  els.rowCount.textContent = "0";
  els.drugCount.textContent = "0";
  els.dataTable.tBodies[0].innerHTML = `<tr class="empty-row"><td colspan="8">${escapeHtml(message)}</td></tr>`;
}

function parseSheet(matrix, sheetName) {
  const headerIndex = matrix.findIndex((row) => {
    const values = row.map((value) => normalizeHeader(value));
    return REQUIRED_HEADERS.every((header) => values.includes(header));
  });

  if (headerIndex < 0) return [];

  const headers = matrix[headerIndex].map((value) => normalizeHeader(value));
  const indexByHeader = new Map();
  headers.forEach((header, index) => {
    if (header && !indexByHeader.has(header)) indexByHeader.set(header, index);
  });

  return matrix.slice(headerIndex + 1).flatMap((line, offset) => {
    const drugName = cleanCell(getByHeader(line, indexByHeader, "약품명"));
    if (!drugName) return [];

    const record = {
      sourceSheet: sheetName,
      sourceRow: headerIndex + offset + 2,
    };

    PARSE_HEADERS.forEach((header) => {
      record[header] = cleanCell(getByHeader(line, indexByHeader, header));
    });

    return record;
  });
}

function normalizeHeader(value) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function getByHeader(row, indexByHeader, header) {
  const index = indexByHeader.get(header);
  return index === undefined ? "" : row[index];
}

function cleanCell(value) {
  const text = String(value ?? "").trim();
  return text === "---" ? "" : text;
}

function classifyDrug(name, unit) {
  const normalizedName = String(name || "").toLowerCase();
  const normalizedUnit = String(unit || "").toLowerCase();

  if (/\(p\)/i.test(name)) return "산제";
  if (/(inj|injection|vial|amp|앰플|주사|주\b|ⓥ)/i.test(name)) return "주사";
  if (/(inj|vial|amp|ⓥ|a)/i.test(unit)) return "주사";
  if (/\d+\s*m(l|L)\b/.test(name) || /\/\s*\d+\s*m(l|L)\b/.test(name)) return "주사";
  if (/(tab|tablet|cap|capsule|정\b|t\b)/i.test(name)) return "정제";
  if (/(t|tab)/i.test(unit)) return "정제";
  if (normalizedName.includes("ml") || normalizedUnit.includes("ml")) return "주사";
  return "정제";
}

function getDrugLabel(name) {
  return String(name || "").replace(/^\s*\(p\)\s*/i, "").trim();
}

function getDrugKey(name) {
  return getDrugLabel(name).replace(/\s+/g, " ").toLowerCase();
}

function assignDrugGroupOrder(rows) {
  const orderByKey = new Map();
  return rows.map((row) => {
    if (!orderByKey.has(row.drugKey)) orderByKey.set(row.drugKey, orderByKey.size);
    return {
      ...row,
      groupOrder: orderByKey.get(row.drugKey),
    };
  });
}

function compareRowsByDrugGroup(a, b) {
  return a.groupOrder - b.groupOrder || a.originalIndex - b.originalIndex;
}

function toNumber(value) {
  const text = String(value ?? "").replace(/,/g, "").trim();
  if (!text) return 0;
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function summarizeDrugs(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = row.drugKey;
    if (!map.has(key)) {
      map.set(key, {
        key,
        name: row.drugLabel,
        count: 0,
        net: 0,
        categories: new Set(),
      });
    }
    const item = map.get(key);
    item.count += 1;
    item.net += row.quantity;
    item.categories.add(row.category);
  });

  return [...map.values()]
      .map((drug) => ({
        ...drug,
        category: drug.categories.size === 1 ? [...drug.categories][0] : "혼합",
        shortName: makeShortDrugName(drug.name),
      }))
}

function makeShortDrugName(name) {
  const original = String(name || "");
  const genericMatch = original.match(/\[향정\].*\(([^)]+)\)/i);
  if (genericMatch) return titleCaseDrug(genericMatch[1]);

  return original
    .replace(/^\s*\(p\)\s*/i, "")
    .replace(/^\s*\[[^\]]+\]\s*/g, "")
    .replace(/^\s*\([^)]*가능\)\s*/i, "")
    .replace(/\s*\([^)]*\)\s*/g, " ")
    .replace(/tab\./gi, "tab")
    .replace(/(\d+(?:\.\d+)?\s*(?:mg|mcg|g|%))\s*\/\s*\d+(?:\.\d+)?\s*m?l\b/gi, "$1")
    .replace(/\s+/g, " ")
    .trim();
}

function titleCaseDrug(value) {
  return String(value || "")
    .trim()
    .replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function renderTable() {
  const tbody = els.dataTable.tBodies[0];
  if (!state.rows.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="8">분석할 행이 없습니다.</td></tr>';
    return;
  }

  const fragment = document.createDocumentFragment();
  state.rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.id = row.id;
    tr.dataset.drugKey = row.drugKey;
    tr.dataset.category = row.category;

    DISPLAY_HEADERS.forEach((header) => {
      const td = document.createElement("td");
      if (header === "category") {
        td.innerHTML = `<span class="category-pill cat-${row.category}">${row.category}</span>`;
      } else {
        const value = header === "불출량" || header === "투약량" ? formatNumber(row[header]) : row[header];
        td.textContent = value;
      }

      if (header === "불출량" || header === "투약량") td.classList.add("num");
      if ((header === "불출량" || header === "투약량") && toNumber(row[header]) < 0) td.classList.add("negative");
      fragment.appendChild(tr).appendChild(td);
    });
  });

  tbody.replaceChildren(fragment);
}

function renderStats() {
  const byKind = new Map();
  const categoryTotals = initCategoryBucket();
  const grand = { out: 0, returns: 0, net: 0 };
  const rowsForStats = state.selectedDrugKey
    ? state.rows.filter((row) => row.drugKey === state.selectedDrugKey)
    : state.rows;
  const selectedDrug = state.drugs.find((drug) => drug.key === state.selectedDrugKey);

  els.statsScope.textContent = selectedDrug ? selectedDrug.name : "전체 약품";

  rowsForStats.forEach((row) => {
    const kind = row["집계구분"] || "미지정";
    if (!byKind.has(kind)) byKind.set(kind, initCategoryBucket());
    const amount = row.quantity;
    const bucket = byKind.get(kind)[row.category];

    bucket.net += amount;
    if (amount >= 0) bucket.out += amount;
    else bucket.returns += Math.abs(amount);

    categoryTotals[row.category].net += amount;
    if (amount >= 0) categoryTotals[row.category].out += amount;
    else categoryTotals[row.category].returns += Math.abs(amount);

    grand.net += amount;
    if (amount >= 0) grand.out += amount;
    else grand.returns += Math.abs(amount);
  });

  els.totalOut.textContent = formatNumber(grand.out);
  els.totalReturn.textContent = formatNumber(grand.returns);
  els.totalNet.textContent = formatNumber(grand.net);

  const tbody = els.statsTable.tBodies[0];
  if (!state.rows.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="5">파일을 선택하면 통계가 표시됩니다.</td></tr>';
    return;
  }
  if (!rowsForStats.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="5">선택한 약품의 집계가 없습니다.</td></tr>';
    return;
  }

  const rows = [...byKind.entries()]
    .sort(([a], [b]) => a.localeCompare(b, "ko"))
    .map(([kind, bucket]) => statRow(kind, bucket, "net"));

  rows.push(statRow("불출량", categoryTotals, "out", "summary-row"));
  rows.push(statRow("반납량", categoryTotals, "returns", "summary-row return-row"));
  rows.push(statRow("순사용량", categoryTotals, "net", "summary-row"));
  tbody.innerHTML = rows.join("");
}

function initCategoryBucket() {
  return {
    정제: { out: 0, returns: 0, net: 0 },
    산제: { out: 0, returns: 0, net: 0 },
    주사: { out: 0, returns: 0, net: 0 },
  };
}

function statRow(label, bucket, metric, className = "") {
  const tablet = bucket["정제"][metric];
  const powder = bucket["산제"][metric];
  const injection = bucket["주사"][metric];
  const total = tablet + powder + injection;
  return `
    <tr class="${className}">
      <td>${escapeHtml(label)}</td>
      <td>${formatNumber(tablet)}</td>
      <td>${formatNumber(powder)}</td>
      <td>${formatNumber(injection)}</td>
      <td>${formatNumber(total)}</td>
    </tr>
  `;
}

function renderDrugButtons() {
  els.visibleDrugCount.textContent = `${state.drugs.length.toLocaleString("ko-KR")}개`;

  if (!state.drugs.length) {
    els.drugButtons.innerHTML = "";
    return;
  }

  const fragment = document.createDocumentFragment();
  state.drugs.forEach((drug) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "drug-button";
    if (drug.key === state.selectedDrugKey) button.classList.add("active");
    button.title = drug.name;
    button.innerHTML = `<span>${escapeHtml(drug.shortName)}</span><em>${drug.category} ${formatNumber(drug.net)}</em>`;
    button.addEventListener("click", () => selectDrug(drug.key));
    fragment.appendChild(button);
  });

  els.drugButtons.replaceChildren(fragment);
}

function selectDrug(drugKey) {
  state.selectedDrugKey = drugKey;
  const rows = [...els.dataTable.tBodies[0].querySelectorAll("tr")];
  let firstMatch = null;

  rows.forEach((row) => {
    const isMatch = row.dataset.drugKey === drugKey;
    row.classList.toggle("selected-drug", isMatch);
    row.classList.remove("focus-row");
    if (isMatch && !firstMatch) firstMatch = row;
  });

  if (firstMatch) {
    firstMatch.classList.add("focus-row");
    firstMatch.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
    window.setTimeout(() => firstMatch.classList.remove("focus-row"), 1200);
  }

  renderDrugButtons();
  renderStats();
}

function clearDrugSelection() {
  state.selectedDrugKey = "";
  els.dataTable.tBodies[0].querySelectorAll("tr").forEach((row) => {
    row.classList.remove("selected-drug", "focus-row");
  });
  renderDrugButtons();
  renderStats();
}

function formatNumber(value) {
  const number = toNumber(value);
  if (Math.abs(number % 1) < 0.000001) {
    return number.toLocaleString("ko-KR", { maximumFractionDigits: 0 });
  }
  return number.toLocaleString("ko-KR", { maximumFractionDigits: 4 });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
