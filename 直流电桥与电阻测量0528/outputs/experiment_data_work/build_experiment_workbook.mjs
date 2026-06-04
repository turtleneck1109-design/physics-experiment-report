import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL("../final/", import.meta.url);
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const readme = workbook.worksheets.add("README");
const summary = workbook.worksheets.add("结果汇总");
const wheat = workbook.worksheets.add("惠斯通电桥");
const raw4 = workbook.worksheets.add("四端法原始数据");
const fit4 = workbook.worksheets.add("四端法拟合");

const bridgeRaw = [
  ["RA", "109.9 kΩ", 10000, 100, 1100.3, 1100.4, 1099.9, 9.366, 0.1, 0.09],
  ["RB", "912.4 Ω", 1000, 1000, 911.9, 911.9, 912.0, 5.451, 0.1, 0.143],
  ["RC", "51.15 Ω", 100, 1000, 509.2, 509.3, 509.3, 17.79, 0.1, 0.11],
];

const fourTerminal = {
  "S-1": {
    d: 0.506,
    L: 14.20,
    I: [0, 10.03, 20.09, 29.99, 39.98, 50.09, 60.02, 70.01, 80.00, 89.99, 100.03],
    U: [0, 0.46, 0.92, 1.38, 1.85, 2.28, 2.76, 3.32, 3.69, 4.16, 4.62],
  },
  "S-2": {
    d: 0.510,
    L: 14.20,
    I: [0, 10.18, 19.99, 30.03, 39.94, 49.97, 59.98, 70.00, 79.95, 90.03, 99.93],
    U: [0, 0.30, 0.58, 0.89, 1.18, 1.48, 1.77, 2.10, 2.36, 2.66, 2.95],
  },
  "S-3": {
    d: 0.452,
    L: 14.20,
    I: [0, 10.01, 20.04, 29.99, 40.04, 49.99, 60.00, 70.01, 80.01, 90.02, 99.98],
    U: [0, 1.01, 2.02, 3.01, 4.02, 5.02, 6.04, 7.04, 8.04, 9.05, 10.05],
  },
};

function setValues(sheet, range, values) {
  sheet.getRange(range).values = values;
}

function setFormulas(sheet, range, formulas) {
  sheet.getRange(range).formulas = formulas;
}

function dBFormula(cell) {
  return `=INT(${cell}/10000)*10000*0.1%+INT(MOD(${cell},10000)/1000)*1000*0.1%+INT(MOD(${cell},1000)/100)*100*0.1%+INT(MOD(${cell},100)/10)*10*0.1%+INT(MOD(${cell},10)/1)*1*0.5%+MOD(${cell},1)*2%+0.020`;
}

for (const sheet of [readme, summary, wheat, raw4, fit4]) {
  sheet.showGridLines = false;
}

setValues(readme, "A1:B12", [
  ["直流电桥与电阻测量实验数据处理", ""],
  ["数据来源", "当前文件夹内手写实验记录照片；RB 的 ΔU 读作 0.143 mV，若原记录不是该值可在“惠斯通电桥”中直接修改。"],
  ["惠斯通公式", "Rx = (R1/R2) × R0"],
  ["A 类不确定度", "ΔA = 2.48 × σ，n = 3"],
  ["B 类不确定度", "按电阻箱各十进位误差累加，并加残余电阻 0.020 Ω；取三次测量 ΔB 最大值。"],
  ["R0 合成不确定度", "u_R0 = SQRT(ΔA^2 + ΔBmax^2)"],
  ["Rx 相对不确定度", "u_rRx = SQRT(0.05%^2 + 0.05%^2 + u_rR0^2)"],
  ["四端法拟合", "用最小二乘法拟合 U(mV) = R(Ω) × I(mA) + b，斜率即电阻 R。"],
  ["电阻率", "ρ = R × π(d/2)^2 / L，d 和 L 在公式中转换为 m。"],
  ["有效数字", "最终结果按 PPT 要求保留：平均值有效数字充足，不确定度取 1 到 2 位，相对不确定度取 1 到 2 位。"],
  ["输出说明", "所有主要派生量均保留为公式，便于检查和修改原始读数。"],
  ["", ""],
]);

setValues(summary, "A1:G1", [["项目", "对象", "最终结果", "相对不确定度", "拟合电阻 R/Ω", "电阻率 ρ/(Ω·m)", "备注"]]);
setValues(summary, "A2:G4", [
  ["惠斯通电桥", "RA", "", "", "", "", "R1/R2 = 100"],
  ["惠斯通电桥", "RB", "", "", "", "", "R1/R2 = 1"],
  ["惠斯通电桥", "RC", "", "", "", "", "R1/R2 = 0.1"],
]);
setFormulas(summary, "C2:D4", [
  ["='惠斯通电桥'!O14", "='惠斯通电桥'!M14/100"],
  ["='惠斯通电桥'!O15", "='惠斯通电桥'!M15/100"],
  ["='惠斯通电桥'!O16", "='惠斯通电桥'!M16/100"],
]);
setValues(summary, "A6:G8", [
  ["四端法", "S-1", "", "", "", "", "d=0.506 mm, L=14.20 mm"],
  ["四端法", "S-2", "", "", "", "", "d=0.510 mm, L=14.20 mm"],
  ["四端法", "S-3", "", "", "", "", "d=0.452 mm, L=14.20 mm"],
]);
setFormulas(summary, "E6:F8", [
  ["='四端法拟合'!B3", "='四端法拟合'!F3"],
  ["='四端法拟合'!B4", "='四端法拟合'!F4"],
  ["='四端法拟合'!B5", "='四端法拟合'!F5"],
]);

setValues(wheat, "A1:J1", [["Rx", "万用表粗测", "R1 (Ω)", "R2 (Ω)", "R0-1 (Ω)", "R0-2 (Ω)", "R0-3 (Ω)", "I (mA)", "ΔR0 (Ω)", "ΔU (mV)"]]);
setValues(wheat, "A2:J4", bridgeRaw);
setValues(wheat, "A7:O7", [["Rx", "R0平均 (Ω)", "σ (Ω)", "ΔA (Ω)", "ΔB1 (Ω)", "ΔB2 (Ω)", "ΔB3 (Ω)", "ΔBmax (Ω)", "u_R0 (Ω)", "u_rR0 (%)", "R1/R2", "Rx平均 (Ω)", "u_rRx (%)", "u_Rx (Ω)", "最终表示"]]);

const processedRows = [14, 15, 16];
for (let i = 0; i < bridgeRaw.length; i++) {
  const sourceRow = i + 2;
  const row = i + 14;
  const finalFormula = i === 0
    ? `=TEXT(L${row}/1000,"0.00")&" ± "&TEXT(N${row}/1000,"0.00")&" kΩ"`
    : i === 1
      ? `=TEXT(L${row},"0.0")&" ± "&TEXT(N${row},"0.0")&" Ω"`
      : `=TEXT(L${row},"0.00")&" ± "&TEXT(N${row},"0.00")&" Ω"`;
  setValues(wheat, `A${row}:A${row}`, [[bridgeRaw[i][0]]]);
  setFormulas(wheat, `B${row}:O${row}`, [[
    `=AVERAGE(E${sourceRow}:G${sourceRow})`,
    `=STDEV.S(E${sourceRow}:G${sourceRow})`,
    `=2.48*C${row}`,
    dBFormula(`E${sourceRow}`),
    dBFormula(`F${sourceRow}`),
    dBFormula(`G${sourceRow}`),
    `=MAX(E${row}:G${row})`,
    `=SQRT(D${row}^2+H${row}^2)`,
    `=I${row}/B${row}*100`,
    `=C${sourceRow}/D${sourceRow}`,
    `=B${row}*K${row}`,
    `=SQRT(0.05%^2+0.05%^2+(J${row}/100)^2)*100`,
    `=L${row}*(M${row}/100)`,
    finalFormula,
  ]]);
}

setValues(wheat, "A20:D24", [
  ["灵敏度计算", "", "", ""],
  ["Rx", "ΔU/(ΔR0/R0)", "ΔU (mV)", "ΔR0/R0"],
  ["RA", "", "", ""],
  ["RB", "", "", ""],
  ["RC", "", "", ""],
]);
setFormulas(wheat, "B22:D24", [
  ["=J2/(I2/B14)", "=J2", "=I2/B14"],
  ["=J3/(I3/B15)", "=J3", "=I3/B15"],
  ["=J4/(I4/B16)", "=J4", "=I4/B16"],
]);

const rawRows = [["样品", "点号", "I (mA)", "U (mV)", "直径 d (mm)", "长度 L (mm)"]];
for (const [sample, v] of Object.entries(fourTerminal)) {
  for (let i = 0; i < v.I.length; i++) {
    rawRows.push([sample, i, v.I[i], v.U[i], v.d, v.L]);
  }
}
setValues(raw4, `A1:F${rawRows.length}`, rawRows);

setValues(fit4, "A1:G1", [["样品", "拟合电阻 R/Ω", "截距 b/mV", "R²", "横截面积 S/m²", "电阻率 ρ/(Ω·m)", "拟合方程"]]);
setValues(fit4, "A3:A5", [["S-1"], ["S-2"], ["S-3"]]);
setFormulas(fit4, "B3:G5", [
  ["=SLOPE('四端法原始数据'!D2:D12,'四端法原始数据'!C2:C12)", "=INTERCEPT('四端法原始数据'!D2:D12,'四端法原始数据'!C2:C12)", "=RSQ('四端法原始数据'!D2:D12,'四端法原始数据'!C2:C12)", "=PI()*('四端法原始数据'!E2/1000)^2/4", "=B3*E3/('四端法原始数据'!F2/1000)", `="U = "&TEXT(B3,"0.00000")&" I "&IF(C3>=0,"+ ","- ")&TEXT(ABS(C3),"0.00000")`],
  ["=SLOPE('四端法原始数据'!D13:D23,'四端法原始数据'!C13:C23)", "=INTERCEPT('四端法原始数据'!D13:D23,'四端法原始数据'!C13:C23)", "=RSQ('四端法原始数据'!D13:D23,'四端法原始数据'!C13:C23)", "=PI()*('四端法原始数据'!E13/1000)^2/4", "=B4*E4/('四端法原始数据'!F13/1000)", `="U = "&TEXT(B4,"0.00000")&" I "&IF(C4>=0,"+ ","- ")&TEXT(ABS(C4),"0.00000")`],
  ["=SLOPE('四端法原始数据'!D24:D34,'四端法原始数据'!C24:C34)", "=INTERCEPT('四端法原始数据'!D24:D34,'四端法原始数据'!C24:C34)", "=RSQ('四端法原始数据'!D24:D34,'四端法原始数据'!C24:C34)", "=PI()*('四端法原始数据'!E24/1000)^2/4", "=B5*E5/('四端法原始数据'!F24/1000)", `="U = "&TEXT(B5,"0.00000")&" I "&IF(C5>=0,"+ ","- ")&TEXT(ABS(C5),"0.00000")`],
]);

setValues(fit4, "A8:G8", [["I (mA)", "S-1 实测", "S-1 拟合", "S-2 实测", "S-2 拟合", "S-3 实测", "S-3 拟合"]]);
for (let i = 0; i < 11; i++) {
  const row = i + 9;
  setFormulas(fit4, `A${row}:G${row}`, [[
    `='四端法原始数据'!C${i + 2}`,
    `='四端法原始数据'!D${i + 2}`,
    `=$B$3*A${row}+$C$3`,
    `='四端法原始数据'!D${i + 13}`,
    `=$B$4*A${row}+$C$4`,
    `='四端法原始数据'!D${i + 24}`,
    `=$B$5*A${row}+$C$5`,
  ]]);
}

const chart = fit4.charts.add("line", fit4.getRange("A8:G19"));
chart.title = "四端法伏安特性曲线与最小二乘拟合";
chart.hasLegend = true;
chart.xAxis = { axisType: "textAxis" };
chart.yAxis = { numberFormatCode: "0.00" };
chart.setPosition("I2", "Q20");

for (const [sheet, widths] of [
  [readme, [22, 110]],
  [summary, [16, 12, 24, 16, 16, 18, 36]],
  [wheat, [10, 14, 10, 10, 12, 12, 12, 10, 10, 10, 12, 12, 12, 14, 20]],
  [raw4, [10, 8, 12, 12, 14, 14]],
  [fit4, [12, 16, 14, 12, 18, 18, 30, 2, 16, 16, 16, 16, 16, 16, 16, 16, 16]],
]) {
  for (let i = 0; i < widths.length; i++) {
    sheet.getRange(`${String.fromCharCode(65 + i)}:${String.fromCharCode(65 + i)}`).format.columnWidth = widths[i];
  }
}

for (const sheet of [readme, summary, wheat, raw4, fit4]) {
  sheet.getRange("A1:Q1").format.font.bold = true;
  sheet.getRange("A1:Q1").format.fill.color = "#D9EAF7";
  sheet.freezePanes.freezeRows(1);
}

summary.getRange("D2:D4").format.numberFormat = "0.000%";
summary.getRange("E6:E8").format.numberFormat = "0.00000";
summary.getRange("F6:F8").format.numberFormat = "0.00E+00";
wheat.getRange("B14:N16").format.numberFormat = "0.000";
wheat.getRange("J14:M16").format.numberFormat = "0.0000";
wheat.getRange("B22:D24").format.numberFormat = "0.000";
wheat.getRange("D22:D24").format.numberFormat = "0.000000";
raw4.getRange("C2:F34").format.numberFormat = "0.00";
fit4.getRange("B3:C5").format.numberFormat = "0.00000";
fit4.getRange("D3:D5").format.numberFormat = "0.000000";
fit4.getRange("E3:F5").format.numberFormat = "0.00E+00";
fit4.getRange("A9:G19").format.numberFormat = "0.00";

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(new URL("直流电桥与电阻测量_数据处理.xlsx", outputDir));

for (const sheetName of ["结果汇总", "惠斯通电桥", "四端法拟合"]) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(new URL(`${sheetName}.png`, outputDir), new Uint8Array(await preview.arrayBuffer()));
}

const inspection = await workbook.inspect({
  kind: "table",
  range: "结果汇总!A1:G8",
  include: "values,formulas",
});
console.log(JSON.stringify(inspection, null, 2));
