const sampleData = {
  g: 9.8,
  rulerU: 0.0005,
  micrometerU: 0.005,
  direct: {
    L: [0.9955, 0.9955, 0.9955, 0.9955, 0.9955, 0.9955],
    littleL: [0.0711, 0.0711, 0.0711, 0.0711, 0.0711, 0.0711],
    D: [1.1860, 1.1860, 1.1860, 1.1860, 1.1860, 1.1860],
    d: [0.302, 0.302, 0.302, 0.302, 0.302, 0.302],
  },
  loads: [
    { mass: 0, add: 0.00, remove: -0.02 },
    { mass: 0.5, add: 1.15, remove: 1.13 },
    { mass: 1.0, add: 2.30, remove: 2.31 },
    { mass: 1.5, add: 3.49, remove: 3.49 },
    { mass: 2.0, add: 4.65, remove: 4.65 },
    { mass: 2.5, add: 5.81, remove: 5.81 },
    { mass: 3.0, add: 7.00, remove: 7.00 },
    { mass: 3.5, add: 8.18, remove: 8.18 },
  ],
};

const directLabels = {
  L: { name: "钢丝长度 L", unit: "m", instrument: "rulerU" },
  littleL: { name: "平面镜到钢丝距离 l", unit: "m", instrument: "rulerU" },
  D: { name: "平面镜到望远镜直尺距离 D", unit: "m", instrument: "rulerU" },
  d: { name: "钢丝直径 d", unit: "mm", instrument: "micrometerU" },
};

const el = {
  gravity: document.getElementById("gravityInput"),
  rulerU: document.getElementById("rulerUInput"),
  micrometerU: document.getElementById("micrometerUInput"),
  directTable: document.getElementById("directTable"),
  loadTableBody: document.querySelector("#loadTable tbody"),
  calculateBtn: document.getElementById("calculateBtn"),
  calculateBottomBtn: document.getElementById("calculateBottomBtn"),
  loadSampleBtn: document.getElementById("loadSampleBtn"),
  clearBtn: document.getElementById("clearBtn"),
  printBtn: document.getElementById("printBtn"),
  backToInputBtn: document.getElementById("backToInputBtn"),
  addDirectColumnBtn: document.getElementById("addDirectColumnBtn"),
  removeDirectColumnBtn: document.getElementById("removeDirectColumnBtn"),
  addLoadRowBtn: document.getElementById("addLoadRowBtn"),
  removeLoadRowBtn: document.getElementById("removeLoadRowBtn"),
  tabButtons: document.querySelectorAll(".tab-button"),
  inputPage: document.getElementById("inputPage"),
  resultPage: document.getElementById("resultPage"),
  status: document.getElementById("statusText"),
  eValue: document.getElementById("eValue"),
  uEValue: document.getElementById("uEValue"),
  relativeValue: document.getElementById("relativeValue"),
  finalValue: document.getElementById("finalValue"),
  canvas: document.getElementById("fitCanvas"),
  averageTable: document.getElementById("averageTable"),
  fitTable: document.getElementById("fitTable"),
  uncertaintyTable: document.getElementById("uncertaintyTable"),
  propagationTable: document.getElementById("propagationTable"),
};

function numberOrNaN(value) {
  if (value === "" || value === null || value === undefined) return NaN;
  return Number(value);
}

function mean(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function sampleStdev(values) {
  if (values.length <= 1) return 0;
  const avg = mean(values);
  const variance = values.reduce((sum, value) => sum + (value - avg) ** 2, 0) / (values.length - 1);
  return Math.sqrt(variance);
}

function directUncertainty(values, instrumentU) {
  const avg = mean(values);
  const sigma = sampleStdev(values);
  const deltaA = 1.05 * sigma;
  const deltaB = instrumentU;
  const u = Math.sqrt(deltaA ** 2 + deltaB ** 2);
  return { mean: avg, sigma, deltaA, deltaB, u };
}

function linearFit(x, y) {
  const n = x.length;
  const xBar = mean(x);
  const yBar = mean(y);
  const sxx = x.reduce((sum, xi) => sum + (xi - xBar) ** 2, 0);
  const syy = y.reduce((sum, yi) => sum + (yi - yBar) ** 2, 0);
  const sxy = x.reduce((sum, xi, index) => sum + (xi - xBar) * (y[index] - yBar), 0);
  const slope = sxy / sxx;
  const intercept = yBar - slope * xBar;
  const fitted = x.map((xi) => intercept + slope * xi);
  const residuals = y.map((yi, index) => yi - fitted[index]);
  const ssr = residuals.reduce((sum, residual) => sum + residual ** 2, 0);
  const residualVariance = ssr / (n - 2);
  const slopeU = Math.sqrt(residualVariance / sxx);
  const interceptU = Math.sqrt(residualVariance * (1 / n + xBar ** 2 / sxx));
  const pearsonR = sxy / Math.sqrt(sxx * syy);
  const r2 = pearsonR ** 2;
  const adjR2 = 1 - (1 - r2) * (n - 1) / (n - 2);

  return {
    n,
    xBar,
    yBar,
    sxx,
    syy,
    sxy,
    slope,
    intercept,
    slopeU,
    interceptU,
    fitted,
    residuals,
    ssr,
    residualVariance,
    pearsonR,
    r2,
    adjR2,
  };
}

function formatNumber(value, digits = 4) {
  if (!Number.isFinite(value)) return "--";
  return Number(value).toFixed(digits);
}

function formatSci(value, digits = 4) {
  if (!Number.isFinite(value)) return "--";
  return Number(value).toExponential(digits);
}

function switchView(viewId) {
  [el.inputPage, el.resultPage].forEach((page) => {
    const isActive = page.id === viewId;
    page.hidden = !isActive;
    page.classList.toggle("active", isActive);
  });

  el.tabButtons.forEach((button) => {
    const isActive = button.dataset.viewTarget === viewId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });
}

function markDataChanged() {
  if (el.status.textContent === "已计算") {
    el.status.textContent = "数据已修改，请重新生成结果";
  }
}

function directRows() {
  return [...el.directTable.querySelectorAll("tbody tr")];
}

function getDirectColumnCount() {
  const headerCells = el.directTable.querySelectorAll("thead th");
  return Math.max(0, headerCells.length - 2);
}

function updateDirectTableWidth() {
  const baseWidth = 262;
  const columnWidth = 86;
  el.directTable.style.minWidth = `${baseWidth + getDirectColumnCount() * columnWidth}px`;
}

function renumberDirectColumns() {
  [...el.directTable.querySelectorAll("thead th")]
    .slice(2)
    .forEach((cell, index) => {
      cell.textContent = index + 1;
    });
  updateDirectTableWidth();
}

function addDirectColumn(values = {}) {
  const headerRow = el.directTable.querySelector("thead tr");
  const th = document.createElement("th");
  th.textContent = getDirectColumnCount() + 1;
  headerRow.appendChild(th);

  directRows().forEach((row) => {
    const key = row.dataset.key;
    const input = document.createElement("input");
    input.type = "number";
    input.step = key === "d" ? "0.001" : "0.0001";
    input.value = values[key] ?? "";

    const cell = document.createElement("td");
    cell.appendChild(input);
    row.appendChild(cell);
  });

  renumberDirectColumns();
}

function removeDirectColumn() {
  if (getDirectColumnCount() <= 1) return;

  el.directTable.querySelector("thead tr").lastElementChild.remove();
  directRows().forEach((row) => row.lastElementChild.remove());
  renumberDirectColumns();
}

function setDirectColumnCount(count) {
  while (getDirectColumnCount() < count) addDirectColumn();
  while (getDirectColumnCount() > count) removeDirectColumn();
  updateDirectTableWidth();
}

function setLoadRows(rows) {
  el.loadTableBody.innerHTML = "";
  rows.forEach((row, index) => addLoadRow(row, index + 1));
}

function addLoadRow(row = { mass: "", add: "", remove: "" }, index = el.loadTableBody.children.length + 1) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>${index}</td>
    <td><input type="number" step="0.1" value="${row.mass ?? ""}"></td>
    <td><input type="number" step="0.01" value="${row.add ?? ""}"></td>
    <td><input type="number" step="0.01" value="${row.remove ?? ""}"></td>
  `;
  el.loadTableBody.appendChild(tr);
}

function renumberLoadRows() {
  [...el.loadTableBody.children].forEach((row, index) => {
    row.cells[0].textContent = index + 1;
  });
}

function loadSample() {
  el.gravity.value = sampleData.g;
  el.rulerU.value = sampleData.rulerU;
  el.micrometerU.value = sampleData.micrometerU;
  setDirectColumnCount(sampleData.direct.L.length);
  Object.entries(sampleData.direct).forEach(([key, values]) => {
    const row = el.directTable.querySelector(`tr[data-key="${key}"]`);
    row.querySelectorAll("input").forEach((input, index) => {
      input.value = values[index];
    });
  });
  setLoadRows(sampleData.loads);
  clearOutputs("示例数据已载入，点击生成结果");
  switchView("inputPage");
}

function clearInputs() {
  document.querySelectorAll("input").forEach((input) => {
    input.value = "";
  });
  setLoadRows(Array.from({ length: 8 }, () => ({ mass: "", add: "", remove: "" })));
  clearOutputs("等待输入");
  switchView("inputPage");
}

function readInputs() {
  const g = numberOrNaN(el.gravity.value);
  const rulerU = numberOrNaN(el.rulerU.value);
  const micrometerU = numberOrNaN(el.micrometerU.value);
  const direct = {};

  Object.keys(directLabels).forEach((key) => {
    const row = el.directTable.querySelector(`tr[data-key="${key}"]`);
    direct[key] = [...row.querySelectorAll("input")].map((input) => numberOrNaN(input.value));
  });

  const loads = [...el.loadTableBody.children].map((row) => {
    const inputs = row.querySelectorAll("input");
    return {
      mass: numberOrNaN(inputs[0].value),
      add: numberOrNaN(inputs[1].value),
      remove: numberOrNaN(inputs[2].value),
    };
  });

  return { g, rulerU, micrometerU, direct, loads };
}

function validateData(data) {
  const problems = [];
  if (!Number.isFinite(data.g) || data.g <= 0) problems.push("重力加速度必须为正数。");
  if (!Number.isFinite(data.rulerU) || data.rulerU < 0) problems.push("米尺 B 类不确定度不能为负。");
  if (!Number.isFinite(data.micrometerU) || data.micrometerU < 0) problems.push("螺旋测微器 B 类不确定度不能为负。");

  Object.entries(data.direct).forEach(([key, values]) => {
    if (values.length === 0) {
      problems.push(`${directLabels[key].name} 至少需要 1 次读数。`);
    }
    if (values.some((value) => !Number.isFinite(value))) {
      problems.push(`${directLabels[key].name} 的 ${values.length} 次读数需要填写完整。`);
    }
  });

  if (data.loads.length < 3) problems.push("最小二乘拟合至少需要 3 组砝码数据。");
  data.loads.forEach((row, index) => {
    if (![row.mass, row.add, row.remove].every(Number.isFinite)) {
      problems.push(`砝码数据第 ${index + 1} 行没有填写完整。`);
    }
  });

  return problems;
}

function analyze(data) {
  const L = directUncertainty(data.direct.L, data.rulerU);
  const littleL = directUncertainty(data.direct.littleL, data.rulerU);
  const D = directUncertainty(data.direct.D, data.rulerU);
  const d = directUncertainty(data.direct.d, data.micrometerU);

  const b = data.loads.map((row) => (row.add + row.remove) / 2);
  const mass = data.loads.map((row) => row.mass);
  const fit = linearFit(b, mass);

  if (!Number.isFinite(fit.slope) || fit.sxx === 0) {
    throw new Error("标尺读数 b 不能全部相同，否则无法做线性拟合。");
  }

  const kKgPerM = fit.slope * 100;
  const ukKgPerM = fit.slopeU * 100;
  const dM = d.mean / 1000;
  const udM = d.u / 1000;

  const E = (8 * data.g / Math.PI) * D.mean * L.mean * kKgPerM / (dM ** 2 * littleL.mean);
  const terms = {
    "uD/D": D.u / D.mean,
    "uL/L": L.u / L.mean,
    "uk/k": ukKgPerM / kKgPerM,
    "2ud/d": 2 * udM / dM,
    "ul/l": littleL.u / littleL.mean,
  };
  const relativeU = Math.sqrt(Object.values(terms).reduce((sum, value) => sum + value ** 2, 0));
  const uE = E * relativeU;

  return {
    raw: data,
    b,
    mass,
    L,
    littleL,
    D,
    d,
    fit,
    kKgPerM,
    ukKgPerM,
    E,
    uE,
    relativeU,
    terms,
  };
}

function renderTable(table, rows) {
  table.innerHTML = "";
  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    row.forEach((value) => {
      const cell = document.createElement(rowIndex === 0 ? "th" : "td");
      cell.textContent = value;
      tr.appendChild(cell);
    });
    table.appendChild(tr);
  });
}

function renderResults(result) {
  el.eValue.textContent = `${formatSci(result.E, 4)} Pa`;
  el.uEValue.textContent = `${formatSci(result.uE, 4)} Pa`;
  el.relativeValue.textContent = `${formatNumber(result.relativeU * 100, 2)}%`;
  el.finalValue.textContent = `(${formatNumber(result.E / 1e11, 2)} ± ${formatNumber(result.uE / 1e11, 2)}) × 10¹¹ Pa`;
  el.status.textContent = "已计算";

  renderAverageTable(result);
  renderFitTable(result);
  renderUncertaintyTable(result);
  renderPropagationTable(result);
  drawChart(result);
}

function renderAverageTable(result) {
  const rows = [["序号", "m/kg", "加砝码/cm", "减砝码/cm", "平均 b/cm"]];
  result.raw.loads.forEach((row, index) => {
    rows.push([
      index + 1,
      formatNumber(row.mass, 2),
      formatNumber(row.add, 2),
      formatNumber(row.remove, 2),
      formatNumber(result.b[index], 3),
    ]);
  });
  renderTable(el.averageTable, rows);
}

function renderFitTable(result) {
  const fit = result.fit;
  const rows = [
    ["项目", "数值"],
    ["拟合方程", `m = ${formatNumber(fit.intercept, 5)} + ${formatNumber(fit.slope, 5)} b`],
    ["截距 a/kg", `${formatNumber(fit.intercept, 5)} ± ${formatNumber(fit.interceptU, 5)}`],
    ["斜率 k/(kg/cm)", `${formatNumber(fit.slope, 5)} ± ${formatNumber(fit.slopeU, 5)}`],
    ["斜率 k/(kg/m)", `${formatNumber(result.kKgPerM, 4)} ± ${formatNumber(result.ukKgPerM, 4)}`],
    ["残差平方和", formatNumber(fit.ssr, 7)],
    ["Pearson's r", formatNumber(fit.pearsonR, 6)],
    ["R²(COD)", formatNumber(fit.r2, 6)],
    ["调整后 R²", formatNumber(fit.adjR2, 6)],
  ];
  renderTable(el.fitTable, rows);
}

function renderUncertaintyTable(result) {
  const items = [
    ["L", "m", result.L],
    ["l", "m", result.littleL],
    ["D", "m", result.D],
    ["d", "mm", result.d],
  ];
  const rows = [["量", "单位", "平均值", "σ", "ΔA", "ΔB", "合成 u"]];
  items.forEach(([name, unit, value]) => {
    rows.push([
      name,
      unit,
      formatNumber(value.mean, unit === "mm" ? 3 : 4),
      formatNumber(value.sigma, 6),
      formatNumber(value.deltaA, 6),
      formatNumber(value.deltaB, unit === "mm" ? 3 : 4),
      formatNumber(value.u, unit === "mm" ? 3 : 4),
    ]);
  });
  renderTable(el.uncertaintyTable, rows);
}

function renderPropagationTable(result) {
  const rows = [["分量", "相对量", "平方贡献"]];
  Object.entries(result.terms).forEach(([name, value]) => {
    rows.push([name, formatNumber(value, 6), formatNumber(value ** 2, 8)]);
  });
  rows.push(["合成 uE/E", formatNumber(result.relativeU, 6), formatNumber(result.relativeU ** 2, 8)]);
  rows.push(["uE/Pa", formatSci(result.uE, 4), ""]);
  renderTable(el.propagationTable, rows);
}

function drawChart(result) {
  const canvas = el.canvas;
  const ctx = canvas.getContext("2d");
  const width = canvas.width;
  const height = canvas.height;

  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);

  const margin = { left: 92, right: 44, top: 72, bottom: 82 };
  const plot = {
    left: margin.left,
    right: width - margin.right,
    top: margin.top,
    bottom: height - margin.bottom,
  };
  plot.width = plot.right - plot.left;
  plot.height = plot.bottom - plot.top;

  const xMin = Math.min(-0.5, Math.min(...result.b) - 0.5);
  const xMax = Math.max(9, Math.max(...result.b) + 0.5);
  const yMin = Math.min(-0.2, Math.min(...result.mass) - 0.2);
  const yMax = Math.max(4, Math.max(...result.mass) + 0.3);

  const px = (x) => plot.left + ((x - xMin) / (xMax - xMin)) * plot.width;
  const py = (y) => plot.bottom - ((y - yMin) / (yMax - yMin)) * plot.height;

  ctx.strokeStyle = "#d9dee7";
  ctx.lineWidth = 1;
  ctx.font = "20px Microsoft YaHei, sans-serif";
  ctx.fillStyle = "#162033";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  const xTicks = [0, 2, 4, 6, 8];
  const yTicks = [0, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4];
  xTicks.forEach((tick) => {
    const x = px(tick);
    ctx.beginPath();
    ctx.moveTo(x, plot.top);
    ctx.lineTo(x, plot.bottom);
    ctx.stroke();
    ctx.fillText(String(tick), x, plot.bottom + 28);
  });
  yTicks.forEach((tick) => {
    const y = py(tick);
    ctx.beginPath();
    ctx.moveTo(plot.left, y);
    ctx.lineTo(plot.right, y);
    ctx.stroke();
    ctx.textAlign = "right";
    ctx.fillText(Number.isInteger(tick) ? String(tick) : tick.toFixed(1), plot.left - 16, y);
  });

  ctx.strokeStyle = "#111827";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(plot.left, plot.bottom);
  ctx.lineTo(plot.right, plot.bottom);
  ctx.moveTo(plot.left, plot.top);
  ctx.lineTo(plot.left, plot.bottom);
  ctx.stroke();

  ctx.textAlign = "center";
  ctx.font = "bold 28px Microsoft YaHei, sans-serif";
  ctx.fillText("砝码质量-标尺读数关系图", width / 2, 36);
  ctx.font = "22px Microsoft YaHei, sans-serif";
  ctx.fillText("标尺读数 b/cm", width / 2, height - 30);

  ctx.save();
  ctx.translate(30, height / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText("砝码质量 m/kg", 0, 0);
  ctx.restore();

  const fit = result.fit;
  const x1 = xMin;
  const x2 = xMax;
  const y1 = fit.intercept + fit.slope * x1;
  const y2 = fit.intercept + fit.slope * x2;
  ctx.strokeStyle = "#bf5146";
  ctx.lineWidth = 3;
  ctx.beginPath();
  ctx.moveTo(px(x1), py(y1));
  ctx.lineTo(px(x2), py(y2));
  ctx.stroke();

  ctx.fillStyle = "#333333";
  result.b.forEach((b, index) => {
    const x = px(b);
    const y = py(result.mass[index]);
    ctx.fillRect(x - 6, y - 6, 12, 12);
  });

  drawLegend(ctx, plot);
  drawStatsTable(ctx, result, plot);
}

function drawLegend(ctx, plot) {
  const x = plot.left + 24;
  const y = plot.top + 22;
  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#9ca3af";
  ctx.lineWidth = 1.5;
  ctx.fillRect(x, y, 360, 86);
  ctx.strokeRect(x, y, 360, 86);

  ctx.font = "18px Microsoft YaHei, sans-serif";
  ctx.fillStyle = "#333333";
  ctx.fillRect(x + 18, y + 20, 12, 12);
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillText("砝码质量", x + 46, y + 26);
  ctx.strokeStyle = "#bf5146";
  ctx.lineWidth = 3;
  ctx.beginPath();
  ctx.moveTo(x + 18, y + 62);
  ctx.lineTo(x + 88, y + 62);
  ctx.stroke();
  ctx.fillStyle = "#162033";
  ctx.fillText("砝码质量-标尺读数的线性拟合", x + 104, y + 62);
}

function drawStatsTable(ctx, result, plot) {
  const rows = [
    ["方程", "m = a + k b"],
    ["权重", "不加权"],
    ["截距", `${formatNumber(result.fit.intercept, 5)} ± ${formatNumber(result.fit.interceptU, 5)} kg`],
    ["斜率", `${formatNumber(result.fit.slope, 5)} ± ${formatNumber(result.fit.slopeU, 5)} kg/cm`],
    ["残差平方和", formatNumber(result.fit.ssr, 7)],
    ["Pearson's r", formatNumber(result.fit.pearsonR, 6)],
    ["R²(COD)", formatNumber(result.fit.r2, 6)],
    ["调整后R²", formatNumber(result.fit.adjR2, 6)],
  ];
  const width = 430;
  const rowHeight = 32;
  const x = plot.right - width - 34;
  const y = plot.bottom - rows.length * rowHeight - 32;

  ctx.font = "17px Microsoft YaHei, sans-serif";
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillStyle = "#f7f7f7";
  ctx.strokeStyle = "#8a8a8a";
  ctx.lineWidth = 1;
  ctx.fillRect(x, y, width, rows.length * rowHeight);
  ctx.strokeRect(x, y, width, rows.length * rowHeight);
  rows.forEach((row, index) => {
    const rowY = y + index * rowHeight;
    ctx.beginPath();
    ctx.moveTo(x, rowY);
    ctx.lineTo(x + width, rowY);
    ctx.moveTo(x + 150, rowY);
    ctx.lineTo(x + 150, rowY + rowHeight);
    ctx.stroke();
    ctx.fillStyle = "#162033";
    ctx.fillText(row[0], x + 12, rowY + rowHeight / 2);
    ctx.fillText(row[1], x + 165, rowY + rowHeight / 2);
  });
}

function clearOutputs(status = "等待计算") {
  el.status.textContent = status;
  el.eValue.textContent = "--";
  el.uEValue.textContent = "--";
  el.relativeValue.textContent = "--";
  el.finalValue.textContent = "--";
  [el.averageTable, el.fitTable, el.uncertaintyTable, el.propagationTable].forEach((table) => {
    table.innerHTML = "";
  });
  const ctx = el.canvas.getContext("2d");
  ctx.clearRect(0, 0, el.canvas.width, el.canvas.height);
}

function showError(message) {
  el.status.textContent = "数据有误";
  const rows = [["提示"], [message]];
  renderTable(el.averageTable, rows);
  el.averageTable.classList.add("error");
}

function calculate() {
  el.averageTable.classList.remove("error");
  const data = readInputs();
  const problems = validateData(data);
  if (problems.length) {
    clearOutputs("数据有误");
    showError(problems[0]);
    switchView("resultPage");
    return false;
  }

  try {
    const result = analyze(data);
    renderResults(result);
    switchView("resultPage");
    return true;
  } catch (error) {
    clearOutputs("数据有误");
    showError(error.message);
    switchView("resultPage");
    return false;
  }
}

el.calculateBtn.addEventListener("click", calculate);
el.calculateBottomBtn.addEventListener("click", calculate);
el.loadSampleBtn.addEventListener("click", loadSample);
el.clearBtn.addEventListener("click", clearInputs);
el.printBtn.addEventListener("click", () => {
  const resultReady = !el.resultPage.hidden || calculate();
  if (resultReady) window.print();
});
el.backToInputBtn.addEventListener("click", () => switchView("inputPage"));
el.addDirectColumnBtn.addEventListener("click", () => {
  addDirectColumn();
  markDataChanged();
});
el.removeDirectColumnBtn.addEventListener("click", () => {
  removeDirectColumn();
  markDataChanged();
});
el.addLoadRowBtn.addEventListener("click", () => {
  addLoadRow();
  markDataChanged();
});
el.removeLoadRowBtn.addEventListener("click", () => {
  if (el.loadTableBody.children.length > 3) {
    el.loadTableBody.lastElementChild.remove();
    renumberLoadRows();
    markDataChanged();
  }
});

document.addEventListener("input", (event) => {
  if (event.target.matches("input")) markDataChanged();
});

el.tabButtons.forEach((button) => {
  button.addEventListener("click", () => switchView(button.dataset.viewTarget));
});

loadSample();
