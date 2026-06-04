"""生成适合打印的 PDF 报告。

PDF 中的数值不重新写一套算法，而是复用 make_final_tables.py 中的
读取、计算和绘图函数，保证 PDF 与 Excel 的结果一致。
"""

from __future__ import annotations

from pathlib import Path

import matplotlib

# PDF 生成只需要保存文件，使用 Agg 后端可避免依赖图形窗口。
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

from make_final_tables import OUTPUT_DIR, PLOT_PATH, analyze, load_data, make_plot


ROOT = Path(__file__).resolve().parent
PDF_PATH = ROOT / "数据处理_打印版报告.pdf"
COMPACT_PDF_PATH = ROOT / "数据处理_打印版报告_紧凑版.pdf"
TIGHT_PDF_PATH = ROOT / "数据处理_打印版报告_紧凑优化版.pdf"


def set_chinese_font() -> None:
    """设置中文字体，并让 PDF 尽量嵌入可复制的文字。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["pdf.fonttype"] = 42
    plt.rcParams["ps.fonttype"] = 42


def new_page(title: str, page_no: int):
    """创建一张 A4 纵向页面，并加入统一页眉和页码。"""
    fig = plt.figure(figsize=(8.27, 11.69), dpi=180)
    fig.patch.set_facecolor("white")
    fig.text(0.5, 0.982, title, ha="center", va="top", fontsize=18, fontweight="bold")
    fig.text(0.5, 0.028, f"第 {page_no} 页", ha="center", va="bottom", fontsize=9, color="#666666")
    return fig


def add_table(
    fig,
    rect: list[float],
    rows: list[list],
    font_size: float = 9,
    header_color: str = "#1f77b4",
    col_widths: list[float] | None = None,
    y_scale: float = 1.25,
):
    """在 PDF 页面指定位置绘制表格。

    rect 是 matplotlib 的相对坐标：[left, bottom, width, height]。
    bbox=[0,0,1,1] 让表格填满该区域，避免标题和表格之间出现大空白。
    """
    ax = fig.add_axes(rect)
    ax.axis("off")
    table = ax.table(
        cellText=rows[1:],
        colLabels=rows[0],
        cellLoc="center",
        bbox=[0, 0, 1, 1],
        colWidths=col_widths,
    )
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (r, _c), cell in table.get_celld().items():
        cell.set_edgecolor("#9a9a9a")
        cell.set_linewidth(0.55)
        if r == 0:
            cell.set_facecolor(header_color)
            cell.get_text().set_color("white")
            cell.get_text().set_fontweight("bold")
        else:
            cell.set_facecolor("#ffffff" if r % 2 else "#f7fbff")
    return table


def add_section(fig, y: float, text: str) -> None:
    """绘制蓝底小节标题。"""
    fig.text(
        0.08,
        y,
        text,
        ha="left",
        va="center",
        fontsize=13,
        fontweight="bold",
        color="#1f4e79",
        bbox={"facecolor": "#d9eaf7", "edgecolor": "none", "pad": 4},
    )


def page_summary(pdf: PdfPages, data: dict, result: dict) -> None:
    """第 1 页：最终结果、主要公式和直接测量量摘要。"""
    fig = new_page("拉伸法测金属丝杨氏模量数据处理报告", 1)
    fig.text(0.5, 0.940, "按 PPT 要求：平均值、不确定度、最小二乘拟合与不确定度传递", ha="center", fontsize=11, color="#555555")

    add_section(fig, 0.900, "一、最终结果")
    summary_rows = [
        ["项目", "结果", "说明"],
        ["拟合方程", f"m = {result['fit']['intercept']:.5f} + {result['fit']['slope']:.5f} b", "b 单位 cm，m 单位 kg"],
        ["斜率 k", f"{result['fit']['slope']:.5f} ± {result['fit']['slope_u']:.5f} kg/cm", "最小二乘拟合"],
        ["斜率 k(SI)", f"{result['k_kg_per_m']:.4f} ± {result['uk_kg_per_m']:.4f} kg/m", "计算 E 时使用"],
        ["Pearson's r", f"{result['fit']['pearson_r']:.6f}", "相关系数"],
        ["R²(COD)", f"{result['fit']['r2']:.6f}", "决定系数"],
        ["E", f"{result['E']:.4e} Pa", "E = 8gDLk/(πd²l)"],
        ["uE", f"{result['u_E']:.4e} Pa", "绝对不确定度"],
        ["相对不确定度", f"{result['rel_u_E'] * 100:.2f}%", "uE/E"],
        ["最终结果", f"({result['E'] / 1e11:.2f} ± {result['u_E'] / 1e11:.2f}) × 10^11 Pa", "报告写法"],
    ]
    add_table(fig, [0.08, 0.545, 0.84, 0.33], summary_rows, font_size=9.3, col_widths=[0.24, 0.38, 0.38])

    add_section(fig, 0.505, "二、计算公式")
    formula_rows = [
        ["内容", "公式或说明"],
        ["杨氏模量", "E = 8gDLk / (πd²l)"],
        ["不确定度传递", "uE/E = √[(uD/D)² + (uL/L)² + (uk/k)² + (2ud/d)² + (ul/l)²]"],
        ["直接测量合成不确定度", "u = √(ΔA² + ΔB²)，其中 ΔA = 1.05σ"],
    ]
    add_table(fig, [0.08, 0.330, 0.84, 0.145], formula_rows, font_size=9.1, col_widths=[0.30, 0.70])

    add_section(fig, 0.280, "三、实验原始量")
    basic_rows = [
        ["量", "平均值", "合成不确定度"],
        ["钢丝长度 L/m", f"{result['L']['mean']:.4f}", f"{result['L']['u']:.4f}"],
        ["平面镜到钢丝距离 l/m", f"{result['l']['mean']:.4f}", f"{result['l']['u']:.4f}"],
        ["平面镜到望远镜直尺距离 D/m", f"{result['D']['mean']:.4f}", f"{result['D']['u']:.4f}"],
        ["钢丝直径 d/mm", f"{result['d']['mean']:.3f}", f"{result['d']['u']:.3f}"],
    ]
    add_table(fig, [0.08, 0.105, 0.84, 0.135], basic_rows, font_size=9.2, col_widths=[0.42, 0.29, 0.29])

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def page_raw_and_uncertainty(pdf: PdfPages, data: dict, result: dict) -> None:
    """第 2 页：标尺读数平均值和直接测量不确定度。"""
    fig = new_page("原始数据、平均值与直接测量不确定度", 2)

    add_section(fig, 0.925, "一、标尺读数平均值 b")
    raw_rows = [["m/kg", "加砝码读数/cm", "减砝码读数/cm", "平均 b/cm"]]
    for m, add, remove, b in zip(data["mass"], data["add"], data["remove"], data["b"]):
        raw_rows.append([f"{m:.2f}", f"{add:.2f}", f"{remove:.2f}", f"{b:.3f}"])
    add_table(fig, [0.08, 0.605, 0.84, 0.285], raw_rows, font_size=9.5, col_widths=[0.20, 0.27, 0.27, 0.26])

    add_section(fig, 0.550, "二、直接测量量的不确定度")
    direct_rows = [
        ["量", "单位", "平均值", "σ", "ΔA", "ΔB", "合成 u", "表示结果"],
        ["L", "m", f"{result['L']['mean']:.6f}", f"{result['L']['sigma']:.6f}", f"{result['L']['delta_a']:.6f}", f"{result['L']['delta_b']:.6f}", f"{result['L']['u']:.6f}", f"{result['L']['mean']:.4f} ± {result['L']['u']:.4f} m"],
        ["l", "m", f"{result['l']['mean']:.6f}", f"{result['l']['sigma']:.6f}", f"{result['l']['delta_a']:.6f}", f"{result['l']['delta_b']:.6f}", f"{result['l']['u']:.6f}", f"{result['l']['mean']:.4f} ± {result['l']['u']:.4f} m"],
        ["D", "m", f"{result['D']['mean']:.6f}", f"{result['D']['sigma']:.6f}", f"{result['D']['delta_a']:.6f}", f"{result['D']['delta_b']:.6f}", f"{result['D']['u']:.6f}", f"{result['D']['mean']:.4f} ± {result['D']['u']:.4f} m"],
        ["d", "mm", f"{result['d']['mean']:.6f}", f"{result['d']['sigma']:.6f}", f"{result['d']['delta_a']:.6f}", f"{result['d']['delta_b']:.6f}", f"{result['d']['u']:.6f}", f"{result['d']['mean']:.3f} ± {result['d']['u']:.3f} mm"],
    ]
    add_table(
        fig,
        [0.055, 0.355, 0.89, 0.165],
        direct_rows,
        font_size=7.4,
        col_widths=[0.08, 0.08, 0.14, 0.13, 0.13, 0.13, 0.13, 0.18],
        y_scale=1.35,
    )

    add_section(fig, 0.295, "三、直接测量原始读数")
    direct_raw_rows = [
        ["量", "第1次", "第2次", "第3次", "第4次", "第5次", "第6次"],
        ["L/m", *[f"{v:.4f}" for v in data["L_values"]]],
        ["l/m", *[f"{v:.4f}" for v in data["l_values"]]],
        ["D/m", *[f"{v:.4f}" for v in data["D_values"]]],
        ["d/mm", *[f"{v:.3f}" for v in data["d_values"]]],
    ]
    add_table(fig, [0.055, 0.105, 0.89, 0.155], direct_raw_rows, font_size=8.4, col_widths=[0.16, 0.14, 0.14, 0.14, 0.14, 0.14, 0.14])

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def page_fit(pdf: PdfPages, data: dict, result: dict) -> None:
    """第 3 页：最小二乘拟合辅助量和拟合参数。"""
    fig = new_page("最小二乘线性拟合计算过程", 3)
    fit = result["fit"]

    add_section(fig, 0.925, "一、拟合辅助量")
    rows = [["序号", "b/cm", "m/kg", "拟合m/kg", "b-b均", "m-m均", "残差", "残差²"]]
    for idx, (b, m, fitted, residual) in enumerate(zip(data["b"], data["mass"], fit["fitted"], fit["residuals"]), start=1):
        rows.append(
            [
                idx,
                f"{b:.3f}",
                f"{m:.2f}",
                f"{fitted:.5f}",
                f"{b - fit['x_bar']:.5f}",
                f"{m - fit['y_bar']:.5f}",
                f"{residual:.5f}",
                f"{residual**2:.7f}",
            ]
        )
    add_table(fig, [0.055, 0.600, 0.89, 0.29], rows, font_size=7.9, col_widths=[0.08, 0.12, 0.11, 0.14, 0.14, 0.14, 0.13, 0.14])

    add_section(fig, 0.545, "二、拟合参数")
    summary_rows = [
        ["项目", "数值", "说明"],
        ["n", f"{fit['n']}", "数据点数"],
        ["b均/cm", f"{fit['x_bar']:.6f}", "横坐标平均值"],
        ["m均/kg", f"{fit['y_bar']:.6f}", "纵坐标平均值"],
        ["Sxx", f"{fit['sxx']:.6f}", "Σ(bi-b均)²"],
        ["Syy", f"{fit['syy']:.6f}", "Σ(mi-m均)²"],
        ["Sxy", f"{fit['sxy']:.6f}", "Σ(bi-b均)(mi-m均)"],
        ["截距 a/kg", f"{fit['intercept']:.5f} ± {fit['intercept_u']:.5f}", "m=a+kb"],
        ["斜率 k/(kg/cm)", f"{fit['slope']:.5f} ± {fit['slope_u']:.5f}", "Sxy/Sxx"],
        ["k/(kg/m)", f"{result['k_kg_per_m']:.4f} ± {result['uk_kg_per_m']:.4f}", "SI 单位"],
        ["残差平方和", f"{fit['ssr']:.7f}", "RSS"],
        ["Pearson's r", f"{fit['pearson_r']:.6f}", "相关系数"],
        ["R² / 调整后R²", f"{fit['r2']:.6f} / {fit['adj_r2']:.6f}", "拟合优度"],
    ]
    add_table(fig, [0.08, 0.095, 0.84, 0.405], summary_rows, font_size=8.4, col_widths=[0.30, 0.35, 0.35], y_scale=1.18)

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def page_plot_and_propagation(pdf: PdfPages, result: dict) -> None:
    """第 4 页：拟合图和 E 的不确定度传递分量。"""
    fig = new_page("拟合图与不确定度传递", 4)

    add_section(fig, 0.925, "一、Python matplotlib 线性拟合图")
    image = plt.imread(PLOT_PATH)
    ax = fig.add_axes([0.08, 0.445, 0.84, 0.445])
    ax.imshow(image)
    ax.axis("off")

    add_section(fig, 0.390, "二、不确定度传递分量")
    prop_rows = [["分量", "相对量", "平方贡献"]]
    for label, value in result["terms"].items():
        prop_rows.append([label, f"{value:.6f}", f"{value**2:.8f}"])
    prop_rows.extend(
        [
            ["合成 uE/E", f"{result['rel_u_E']:.6f}", f"{result['rel_u_E'] ** 2:.8f}"],
            ["uE/Pa", f"{result['u_E']:.4e}", ""],
            ["最终结果", f"({result['E'] / 1e11:.2f} ± {result['u_E'] / 1e11:.2f}) × 10^11 Pa", ""],
        ]
    )
    add_table(fig, [0.11, 0.105, 0.78, 0.24], prop_rows, font_size=8.9, col_widths=[0.32, 0.34, 0.34], y_scale=1.18)

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def build_pdf(path: Path, data: dict, result: dict) -> None:
    """按固定页序写出完整 PDF。"""
    with PdfPages(path) as pdf:
        page_summary(pdf, data, result)
        page_raw_and_uncertainty(pdf, data, result)
        page_fit(pdf, data, result)
        page_plot_and_propagation(pdf, result)


def main() -> None:
    """脚本入口：读取数据、计算、刷新拟合图并导出 PDF。"""
    set_chinese_font()
    OUTPUT_DIR.mkdir(exist_ok=True)
    data = load_data()
    result = analyze(data)
    make_plot(data, result)
    try:
        build_pdf(PDF_PATH, data, result)
        print(f"pdf={PDF_PATH}")
    except PermissionError:
        try:
            build_pdf(COMPACT_PDF_PATH, data, result)
            print(f"pdf={COMPACT_PDF_PATH}")
            print(f"warning=原 PDF 可能正在打开，已另存为紧凑版。")
        except PermissionError:
            build_pdf(TIGHT_PDF_PATH, data, result)
            print(f"pdf={TIGHT_PDF_PATH}")
            print(f"warning=前两个 PDF 可能正在打开，已另存为紧凑优化版。")


if __name__ == "__main__":
    main()
