import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __filename = fileURLToPath(import.meta.url);
const ROOT = path.dirname(__filename);
const FIG_DIR = path.join(ROOT, "output_figures");
const OUTPUT_DIR = path.join(ROOT, "outputs", "figure_legends_excel");
const OUTPUT_XLSX = path.join(OUTPUT_DIR, "简谐振动实验图片与详细图例.xlsx");
const PREVIEW_PNG = path.join(OUTPUT_DIR, "figure_legends_preview.png");

const FIGURES = [
  {
    filename: "00_results_summary.png",
    title: "图0  简谐振动实验数据处理结果汇总",
    caption:
      "图像内容：汇总本次简谐振动实验的主要计算结果，包括焦利秤法、质量-周期法、斜面倾角影响、振幅-周期关系和能量关系。\n" +
      "图例说明：表中 K 表示两弹簧等效劲度系数，m0 表示系统等效附加质量；T²-M 图用于由斜率反推 K，A-T 图用于检验周期是否随振幅变化，Vmax²-A² 图用于检验能量关系。\n" +
      "关键结果：焦利秤法得到 k1=2.2206 N/m、k2=2.5255 N/m、k1+k2=4.7461 N/m；周期法得到的 K 约为 4.86-4.91 N/m，与焦利秤法接近。\n" +
      "结论：实验总体支持简谐振动模型；斜面倾角主要改变平衡位置，对周期影响不明显；周期基本与振幅无关。",
  },
  {
    filename: "01_static_spring1_fit.png",
    title: "图1  焦利秤法测量弹簧1劲度系数",
    caption:
      "坐标含义：横轴为弹簧1伸长量 ΔL，单位 m；纵轴为外加拉力 F，单位 N。\n" +
      "图例说明：蓝色散点为实验测量数据，红色直线为不加权最小二乘线性拟合结果。\n" +
      "拟合模型：F = a + kΔL，其中斜率 k 即弹簧1劲度系数。\n" +
      "拟合结果：a = -0.001207 ± 0.000833 N，k1 = 2.220580 ± 0.011086 N/m，Pearson r = 0.999938，R² = 0.999875。\n" +
      "物理含义：数据点几乎落在同一直线上，说明弹簧1在实验范围内满足胡克定律，线性关系可靠。",
  },
  {
    filename: "02_static_spring2_fit.png",
    title: "图2  焦利秤法测量弹簧2劲度系数",
    caption:
      "坐标含义：横轴为弹簧2伸长量 ΔL，单位 m；纵轴为外加拉力 F，单位 N。\n" +
      "图例说明：蓝色散点为实验测量数据，红色直线为不加权最小二乘线性拟合结果。\n" +
      "拟合模型：F = a + kΔL，其中斜率 k 即弹簧2劲度系数。\n" +
      "拟合结果：a = -0.000442 ± 0.000379 N，k2 = 2.525540 ± 0.005755 N/m，Pearson r = 0.999987，R² = 0.999974。\n" +
      "物理含义：弹簧2的线性拟合优度很高，得到的劲度系数大于弹簧1，说明弹簧2相对更硬。",
  },
  {
    filename: "03_mass_period_fit.png",
    title: "图3  水平条件下 T²-M 关系",
    caption:
      "坐标含义：横轴为振子质量 M，单位 kg；纵轴为周期平方 T²，单位 s²。\n" +
      "图例说明：蓝色散点为水平导轨条件下的实验数据，红色直线为 T² 对 M 的线性拟合。\n" +
      "理论关系：T² = 4π²/(k1+k2) · M + 4π²/(k1+k2) · m0，斜率可反推出两弹簧等效劲度系数 K=k1+k2。\n" +
      "拟合结果：T² = 0.062375 + 8.119147 M，K = 4.8624 N/m，m0 = 7.683 g，R² = 0.999971。\n" +
      "结论：T² 与 M 呈高度线性关系，说明质量-周期法适用于本组水平实验数据。",
  },
  {
    filename: "04_incline1_mass_period_fit.png",
    title: "图4  斜面1条件下 T²-M 关系",
    caption:
      "坐标含义：横轴为振子质量 M，单位 kg；纵轴为周期平方 T²，单位 s²。\n" +
      "图例说明：蓝色散点为斜面1条件下的实验数据，红色直线为线性拟合结果。\n" +
      "实验条件：斜面高度 h = 1.240 cm，对应倾角 θ = 0.822°。\n" +
      "拟合结果：截距 a = 0.064953 ± 0.003447 s²，斜率 = 8.04251 ± 0.02553 s²/kg；由斜率得到 K = 4.9087 N/m，m0 = 8.076 g，R² = 0.999950。\n" +
      "结论：斜面1中 T²-M 仍保持良好线性，K 与水平条件结果接近，倾角对周期影响较小。",
  },
  {
    filename: "05_incline2_mass_period_fit.png",
    title: "图5  斜面2条件下 T²-M 关系",
    caption:
      "坐标含义：横轴为振子质量 M，单位 kg；纵轴为周期平方 T²，单位 s²。\n" +
      "图例说明：蓝色散点为斜面2条件下的实验数据，红色直线为线性拟合结果。\n" +
      "实验条件：斜面高度 h = 2.510 cm，对应倾角 θ = 1.665°。\n" +
      "拟合结果：截距 a = 0.063836 ± 0.003427 s²，斜率 = 8.05108 ± 0.02539 s²/kg；由斜率得到 K = 4.9035 N/m，m0 = 7.929 g，R² = 0.999950。\n" +
      "结论：斜面2与斜面1、水平条件所得 K 值非常接近，说明小倾角主要影响平衡位置，而不是振动周期。",
  },
  {
    filename: "06_amplitude_period_fit.png",
    title: "图6  振幅 A 与周期 T 的关系",
    caption:
      "坐标含义：横轴为振幅 A，单位 cm；纵轴为平均周期 T，单位 ms。\n" +
      "图例说明：蓝色散点为不同振幅下的周期测量结果，红色直线为线性拟合趋势线。\n" +
      "拟合结果：平均周期 Tmean = 1020.973 ms，周期标准差 sT = 0.245 ms；线性斜率 = -0.003657 ms/cm，Pearson r = -0.139628，R² = 0.019496。\n" +
      "物理含义：拟合斜率接近 0，R² 很小，说明振幅变化不能有效解释周期变化。\n" +
      "结论：在实验振幅范围内，周期基本与振幅无关，符合简谐振动等时性。",
  },
  {
    filename: "07_energy_vmax2_a2_fit.png",
    title: "图7  Vmax² 与 A² 的能量关系",
    caption:
      "坐标含义：横轴为振幅平方 A²，纵轴为最大速度平方 Vmax²；其中 Vmax 由光电门挡光宽度 Δx=0.995 cm 与通过时间 t 计算得到。\n" +
      "图例说明：蓝色散点为实验计算点，红色直线为 Vmax² 对 A² 的线性拟合。\n" +
      "理论关系：Vmax² = (k1+k2)/(M+m0) · A²，因此拟合斜率可用于估算等效劲度系数。\n" +
      "拟合结果：斜率 = 31.4350 s^-2，截距 = 0.034430，R² = 0.998860；使用水平 T²-M 得到的 m0=7.683 g 时，K = 4.0319 N/m。\n" +
      "结论：Vmax² 与 A² 呈显著线性关系，支持机械能转化关系；该方法得到的 K 小于焦利秤法和周期法，可能与测速、振幅读数或阻尼损耗有关。",
  },
];

function columnName(indexOneBased) {
  let n = indexOneBased;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

async function pngDataUrlAndSize(filePath) {
  const bytes = await fs.readFile(filePath);
  if (
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47
  ) {
    throw new Error(`${filePath} is not a PNG file`);
  }
  return {
    dataUrl: `data:image/png;base64,${bytes.toString("base64")}`,
    width: bytes.readUInt32BE(16),
    height: bytes.readUInt32BE(20),
  };
}

function setRangeStyle(range, style) {
  range.format = {
    ...range.format,
    ...style,
  };
}

async function buildWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("图表与图例");
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(2);

  for (let col = 1; col <= 15; col += 1) {
    const letter = columnName(col);
    const widthPx =
      col <= 8 ? 84 : col === 9 ? 18 : col === 15 ? 52 : 104;
    sheet.getRange(`${letter}1:${letter}190`).format.columnWidthPx = widthPx;
  }

  for (let row = 1; row <= 190; row += 1) {
    sheet.getRangeByIndexes(row - 1, 0, 1, 15).format.rowHeightPx = 24;
  }

  sheet.getRange("A1:O1").merge();
  sheet.getRange("A1").values = [["简谐振动实验图片与详细图例"]];
  setRangeStyle(sheet.getRange("A1:O1"), {
    fill: "#17324D",
    font: { bold: true, color: "#FFFFFF", size: 18, name: "Microsoft YaHei" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  });
  sheet.getRangeByIndexes(0, 0, 1, 15).format.rowHeightPx = 38;

  sheet.getRange("A2:O2").merge();
  sheet.getRange("A2").values = [[
    "左侧为实验处理图，右侧为详细图例说明；图中散点均为实验数据，拟合线均为不加权最小二乘线性拟合。",
  ]];
  setRangeStyle(sheet.getRange("A2:O2"), {
    fill: "#EAF2FA",
    font: { color: "#1F2937", size: 11, name: "Microsoft YaHei" },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
  });
  sheet.getRangeByIndexes(1, 0, 1, 15).format.rowHeightPx = 30;

  let startRow = 4;
  const imageWidthPx = 650;
  const blockRows = 19;

  for (const [index, figure] of FIGURES.entries()) {
    const filePath = path.join(FIG_DIR, figure.filename);
    const image = await pngDataUrlAndSize(filePath);
    const imageHeightPx = Math.round(imageWidthPx * image.height / image.width);
    const titleRow = startRow;
    const contentTopRow = startRow + 1;
    const contentBottomRow = startRow + blockRows - 1;

    sheet.getRange(`A${titleRow}:O${titleRow}`).merge();
    sheet.getRange(`A${titleRow}`).values = [[figure.title]];
    setRangeStyle(sheet.getRange(`A${titleRow}:O${titleRow}`), {
      fill: index % 2 === 0 ? "#DDEAF6" : "#E7F0E8",
      font: { bold: true, color: "#102A43", size: 13, name: "Microsoft YaHei" },
      horizontalAlignment: "left",
      verticalAlignment: "center",
      borders: { preset: "outside", style: "thin", color: "#9CA3AF" },
    });
    sheet.getRangeByIndexes(titleRow - 1, 0, 1, 15).format.rowHeightPx = 28;

    sheet.images.add({
      dataUrl: image.dataUrl,
      anchor: {
        from: {
          row: contentTopRow - 1,
          col: 0,
          rowOffsetPx: 8,
          colOffsetPx: 8,
        },
        extent: {
          widthPx: imageWidthPx,
          heightPx: imageHeightPx,
        },
      },
    });

    sheet.getRange(`J${contentTopRow}:N${contentBottomRow}`).merge();
    sheet.getRange(`J${contentTopRow}`).values = [[figure.caption]];
    setRangeStyle(sheet.getRange(`J${contentTopRow}:N${contentBottomRow}`), {
      fill: "#F8FAFC",
      font: { color: "#111827", size: 10, name: "Microsoft YaHei" },
      horizontalAlignment: "left",
      verticalAlignment: "top",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#CBD5E1" },
    });

    sheet.getRange(`O${contentTopRow}:O${contentBottomRow}`).merge();
    sheet.getRange(`O${contentTopRow}`).values = [[figure.filename]];
    setRangeStyle(sheet.getRange(`O${contentTopRow}:O${contentBottomRow}`), {
      fill: "#EEF2F7",
      font: { color: "#475569", size: 9, name: "Microsoft YaHei" },
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
      borders: { preset: "all", style: "thin", color: "#CBD5E1" },
    });

    sheet.getRange(`A${contentTopRow}:H${contentBottomRow}`).format.borders = {
      preset: "outside",
      style: "thin",
      color: "#CBD5E1",
    };

    startRow += blockRows + 2;
  }

  const finalRow = startRow - 2;
  sheet.getRange(`A1:O${finalRow}`).format.font = {
    name: "Microsoft YaHei",
  };

  const preview = await workbook.render({
    sheetName: "图表与图例",
    range: `A1:O${finalRow}`,
    scale: 0.6,
    format: "png",
  });
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  await fs.writeFile(PREVIEW_PNG, new Uint8Array(await preview.arrayBuffer()));

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 50 },
    summary: "formula error scan",
    maxChars: 1000,
  });

  const tableCheck = await workbook.inspect({
    kind: "table",
    sheetId: "图表与图例",
    range: "A1:O12",
    include: "values",
    tableMaxRows: 12,
    tableMaxCols: 15,
    tableMaxCellChars: 60,
    maxChars: 2500,
  });

  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(OUTPUT_XLSX);

  return {
    output: OUTPUT_XLSX,
    preview: PREVIEW_PNG,
    finalRow,
    checks: {
      formulaErrorScan: errors.ndjson,
      tablePreviewRows: 12,
      plannedImages: FIGURES.length,
      tableCheckLength: tableCheck.ndjson.length,
    },
  };
}

const result = await buildWorkbook();
console.log(JSON.stringify(result, null, 2));
process.exitCode = 0;
