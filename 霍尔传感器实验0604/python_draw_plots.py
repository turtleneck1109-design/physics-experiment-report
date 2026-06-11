from pathlib import Path
import math

import numpy as np
import openpyxl

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "霍尔传感器实验数据记录表_已填写.xlsx"
OUT = ROOT / "outputs" / "python_plots"
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 160


def as_float(value):
    if value is None or value == "":
        return None
    return float(str(value).strip())


def row_values(ws, row, start_col, end_col):
    values = []
    for col in range(start_col, end_col + 1):
        value = as_float(ws.cell(row=row, column=col).value)
        if value is not None:
            values.append(value)
    return np.array(values, dtype=float)


def finite_solenoid_b(x_m, current_a, turns=3000, length_m=0.26, diameter_m=0.025):
    mu0 = 4 * math.pi * 1e-7
    term1 = (length_m + 2 * x_m) / (2 * np.sqrt(diameter_m**2 + (length_m + 2 * x_m) ** 2))
    term2 = (length_m - 2 * x_m) / (2 * np.sqrt(diameter_m**2 + (length_m - 2 * x_m) ** 2))
    return mu0 * turns / length_m * current_a * (term1 + term2)


def least_squares_line(x, y):
    x_mean = np.mean(x)
    y_mean = np.mean(y)
    slope = np.sum((x - x_mean) * (y - y_mean)) / np.sum((x - x_mean) ** 2)
    intercept = y_mean - slope * x_mean
    y_fit = slope * x + intercept
    ss_res = np.sum((y - y_fit) ** 2)
    ss_tot = np.sum((y - y_mean) ** 2)
    r2 = 1 - ss_res / ss_tot
    pearson_r = np.corrcoef(x, y)[0, 1]
    n = len(x)
    sxx = np.sum((x - x_mean) ** 2)
    sigma2 = ss_res / (n - 2)
    slope_se = np.sqrt(sigma2 / sxx)
    intercept_se = np.sqrt(sigma2 * (1 / n + x_mean**2 / sxx))
    adj_r2 = 1 - (1 - r2) * (n - 1) / (n - 2)
    return {
        "slope": slope,
        "intercept": intercept,
        "y_fit": y_fit,
        "r2": r2,
        "pearson_r": pearson_r,
        "sse": ss_res,
        "adj_r2": adj_r2,
        "slope_se": slope_se,
        "intercept_se": intercept_se,
    }


def style_axes(ax):
    ax.grid(True, linestyle="--", alpha=0.35)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


wb = openpyxl.load_workbook(INPUT, data_only=True)
ws = wb["填写模板"]

us = row_values(ws, 5, 2, 16)
u0 = row_values(ws, 6, 2, 16)
u = row_values(ws, 7, 2, 16)

im = row_values(ws, 13, 2, 11)
u_prime2_mv = row_values(ws, 14, 2, 11)

x_raw_cm = np.concatenate([
    row_values(ws, 20, 2, 18),
    row_values(ws, 27, 2, 18),
    row_values(ws, 33, 2, 10),
])
u_prime3_mv = np.concatenate([
    row_values(ws, 21, 2, 18),
    row_values(ws, 28, 2, 18),
    row_values(ws, 34, 2, 10),
])

length_m = 0.26
diameter_m = 0.025
delta_x_cm = 2.3

b0_025_t = finite_solenoid_b(0, 0.25, length_m=length_m, diameter_m=diameter_m)
b0_025_mt = b0_025_t * 1000

k_vs_us = (u - u0) / b0_025_t
k_over_us = k_vs_us / us
k_us_fit = least_squares_line(us, k_vs_us)

b2_t = finite_solenoid_b(0, im, length_m=length_m, diameter_m=diameter_m)
b2_mt = b2_t * 1000
fit = least_squares_line(b2_mt, u_prime2_mv)
k_fit_v_per_t = fit["slope"]
intercept_mv = fit["intercept"]
u_fit_mv = fit["y_fit"]
r2 = fit["r2"]
reference_k = 31.25
relative_error = (k_fit_v_per_t - reference_k) / reference_k * 100

x_center_cm = x_raw_cm - (delta_x_cm + length_m * 100 / 2)
b_exp_mt = u_prime3_mv / k_fit_v_per_t
b_theory_mt = finite_solenoid_b(x_center_cm / 100, 0.25, length_m=length_m, diameter_m=diameter_m) * 1000

fig, ax = plt.subplots(figsize=(7.2, 4.6))
ax.scatter(us, k_vs_us, s=36, color="#1f77b4", label="实验数据")
us_fit_grid = np.linspace(np.min(us), np.max(us), 200)
ax.plot(
    us_fit_grid,
    k_us_fit["slope"] * us_fit_grid + k_us_fit["intercept"],
    linewidth=1.8,
    color="#d62728",
    label="最小二乘拟合",
)
fit_rows = [
    ["方程", "K = a + bUs"],
    ["权重", "不加权"],
    ["截距 a", f"{k_us_fit['intercept']:.4f} ± {k_us_fit['intercept_se']:.4f}"],
    ["斜率 b", f"{k_us_fit['slope']:.4f} ± {k_us_fit['slope_se']:.4f}"],
    ["Pearson's r", f"{k_us_fit['pearson_r']:.6f}"],
    ["R²(COD)", f"{k_us_fit['r2']:.6f}"],
    ["调整后R²", f"{k_us_fit['adj_r2']:.6f}"],
]
fit_table = ax.table(
    cellText=fit_rows,
    cellLoc="left",
    colWidths=[0.26, 0.52],
    bbox=[0.04, 0.54, 0.45, 0.38],
)
fit_table.auto_set_font_size(False)
fit_table.set_fontsize(8.5)
for cell in fit_table.get_celld().values():
    cell.set_edgecolor("#9e9e9e")
    cell.set_linewidth(0.5)
    cell.set_facecolor("#f5f5f5")
    cell.get_text().set_fontfamily("Microsoft YaHei")
ax.annotate(
    "低电压点偏离明显",
    xy=(2.5, k_vs_us[0]),
    xytext=(3.25, 8),
    arrowprops=dict(arrowstyle="->", color="#555555", lw=0.9),
    fontsize=9,
)
ax.set_title("实验一：灵敏度 K 随工作电压 Us 的变化")
ax.set_xlabel("Us / V")
ax.set_ylabel("K / (V/T)")
ax.legend(frameon=False, loc="lower right")
style_axes(ax)
fig.tight_layout()
fig.savefig(OUT / "实验一_K随Us变化.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(7.2, 4.6))
ax.plot(us, k_over_us, marker="o", linewidth=1.8, color="#2ca02c", label="实验数据")
ax.set_title("实验一：K/Us 随工作电压 Us 的变化")
ax.set_xlabel("Us / V")
ax.set_ylabel("K/Us / (V/(T·V))")
ax.legend(frameon=False)
style_axes(ax)
fig.tight_layout()
fig.savefig(OUT / "实验一_K除以Us随Us变化.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(7.2, 4.8))
order = np.argsort(b2_mt)
ax.scatter(b2_mt, u_prime2_mv, s=42, color="#1f77b4", label="实验数据")
ax.plot(b2_mt[order], u_fit_mv[order], color="#d62728", linewidth=1.8, label="最小二乘拟合")
ax.set_title("实验二：U' 与磁场强度 B 的线性关系")
ax.set_xlabel("B / mT")
ax.set_ylabel("U' = U - U0 / mV")
fit_rows = [
    ["方程", "U' = a + kB"],
    ["绘图", "实验数据"],
    ["权重", "不加权"],
    ["截距 a", f"{intercept_mv:.5f} ± {fit['intercept_se']:.5f} mV"],
    ["斜率 k", f"{k_fit_v_per_t:.5f} ± {fit['slope_se']:.5f} V/T"],
    ["残差平方和", f"{fit['sse']:.6f}"],
    ["Pearson's r", f"{fit['pearson_r']:.6f}"],
    ["R²(COD)", f"{fit['r2']:.6f}"],
    ["调整后R²", f"{fit['adj_r2']:.6f}"],
]
fit_table = ax.table(
    cellText=fit_rows,
    cellLoc="left",
    colWidths=[0.28, 0.55],
    bbox=[0.04, 0.58, 0.43, 0.36],
)
fit_table.auto_set_font_size(False)
fit_table.set_fontsize(9.5)
for cell in fit_table.get_celld().values():
    cell.set_edgecolor("#9e9e9e")
    cell.set_linewidth(0.5)
    cell.set_facecolor("#f5f5f5")
    cell.get_text().set_fontfamily("Microsoft YaHei")
ax.legend(frameon=False, loc="lower right")
style_axes(ax)
fig.tight_layout()
fig.savefig(OUT / "实验二_Uprime_B最小二乘拟合.png")
plt.close(fig)

fig, ax = plt.subplots(figsize=(8.2, 4.8))
order = np.argsort(x_center_cm)
ax.plot(
    x_center_cm[order],
    b_exp_mt[order],
    marker="o",
    markersize=4.5,
    linewidth=1.8,
    color="black",
    linestyle="-",
    label="实验值：实线+圆点",
)
ax.plot(
    x_center_cm[order],
    b_theory_mt[order],
    marker="s",
    markersize=4,
    linewidth=1.6,
    color="black",
    linestyle="--",
    label="理论值：虚线+方块",
)
ax.set_title("实验三：螺线管内磁场 B(x) 分布")
ax.set_xlabel("以螺线管中心为原点的位置 x / cm")
ax.set_ylabel("B / mT")
ax.annotate(
    "实验值",
    xy=(-4, b_exp_mt[order][np.argmin(np.abs(x_center_cm[order] + 4))]),
    xytext=(-8.2, 3.85),
    arrowprops=dict(arrowstyle="->", color="black", lw=0.9),
    fontsize=10,
)
ax.annotate(
    "理论值",
    xy=(6, b_theory_mt[order][np.argmin(np.abs(x_center_cm[order] - 6))]),
    xytext=(7.8, 3.25),
    arrowprops=dict(arrowstyle="->", color="black", lw=0.9),
    fontsize=10,
)
ax.legend(frameon=False, loc="lower center", ncol=2)
style_axes(ax)
fig.tight_layout()
fig.savefig(OUT / "实验三_Bx实验值与理论值对比.png")
plt.close(fig)

summary = OUT / "拟合结果.txt"
summary.write_text(
    "\n".join([
        f"B0(Im=0.25 A) = {b0_025_mt:.4f} mT",
        f"最小二乘拟合方程: U' = {k_fit_v_per_t:.4f} B + {intercept_mv:.4f}",
        "说明: U' 单位为 mV, B 单位为 mT, 因此斜率数值等于 V/T。",
        f"K = {k_fit_v_per_t:.4f} V/T",
        f"R^2 = {r2:.6f}",
        f"相对误差(相对 31.25 V/T) = {relative_error:.2f}%",
    ]),
    encoding="utf-8",
)

print(f"输出目录: {OUT}")
print(summary.read_text(encoding="utf-8"))
