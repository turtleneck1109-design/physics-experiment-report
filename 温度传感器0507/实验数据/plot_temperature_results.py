from pathlib import Path
import csv
import math

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "processed_results"


def read_rows(path):
    rows = []
    with open(path, encoding="gbk", errors="ignore") as f:
        next(f, None)
        for line in f:
            parts = line.split()
            if len(parts) < 6:
                continue
            try:
                rows.append(tuple(float(x) for x in parts[:6]))
            except ValueError:
                pass
    return rows


def heating_branch(rows, skip_count):
    usable = rows[skip_count:]
    max_idx = max(range(len(usable)), key=lambda i: usable[i][0])
    return usable[: max_idx + 1]


def pick_nearest(rows, targets):
    return [min(rows, key=lambda r: abs(r[0] - target)) for target in targets]


def linear_fit(xs, ys):
    n = len(xs)
    mx = sum(xs) / n
    my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    k = sxy / sxx
    b = my - k * mx
    yh = [k * x + b for x in xs]
    ss_res = sum((y - p) ** 2 for y, p in zip(ys, yh))
    ss_tot = sum((y - my) ** 2 for y in ys)
    r2 = 1 - ss_res / ss_tot if ss_tot else 1
    return k, b, r2


def find_font(size):
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simsun.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for name in candidates:
        if Path(name).exists():
            return ImageFont.truetype(name, size)
    return ImageFont.load_default()


def nice_ticks(vmin, vmax, count=6):
    if vmin == vmax:
        return [vmin]
    step = (vmax - vmin) / (count - 1)
    return [vmin + i * step for i in range(count)]


def text_size(draw, text, font):
    box = draw.textbbox((0, 0), text, font=font)
    return box[2] - box[0], box[3] - box[1]


def draw_text_box(draw, xy, lines, font, padding=12, line_gap=5, border=3):
    x, y = xy
    widths = [text_size(draw, line, font)[0] for line in lines]
    heights = [text_size(draw, line, font)[1] for line in lines]
    w = max(widths) + padding * 2
    h = sum(heights) + line_gap * (len(lines) - 1) + padding * 2
    draw.rectangle((x + 6, y + 6, x + w + 6, y + h + 6), fill=(20, 20, 20))
    draw.rectangle((x, y, x + w, y + h), fill="white", outline=(20, 20, 20), width=border)
    cy = y + padding
    for line, height in zip(lines, heights):
        draw.text((x + padding, cy), line, fill=(20, 20, 20), font=font)
        cy += height + line_gap
    return w, h


def draw_fit_table(draw, xy, rows, font, col_widths=(130, 205), row_h=34):
    x, y = xy
    w = sum(col_widths)
    h = row_h * len(rows)
    draw.rectangle((x, y, x + w, y + h), fill="white", outline=(20, 20, 20), width=2)
    draw.line((x + col_widths[0], y, x + col_widths[0], y + h), fill=(20, 20, 20), width=1)
    for i in range(1, len(rows)):
        yy = y + i * row_h
        draw.line((x, yy, x + w, yy), fill=(20, 20, 20), width=1)
    for i, (left, right) in enumerate(rows):
        cy = y + i * row_h + row_h / 2
        draw.text((x + 8, cy), left, fill=(20, 20, 20), font=font, anchor="lm")
        draw.text((x + col_widths[0] + col_widths[1] / 2, cy), right, fill=(20, 20, 20), font=font, anchor="mm")


def draw_rotated_label(img, text, center, font, fill):
    temp = Image.new("RGBA", (360, 70), (255, 255, 255, 0))
    td = ImageDraw.Draw(temp)
    box = td.textbbox((0, 0), text, font=font)
    tw, th = box[2] - box[0], box[3] - box[1]
    td.text(((360 - tw) / 2, (70 - th) / 2), text, fill=fill + (255,), font=font)
    rotated = temp.rotate(90, expand=True)
    x = int(center[0] - rotated.width / 2)
    y = int(center[1] - rotated.height / 2)
    img.paste(rotated, (x, y), rotated)


def draw_title_ribbon(draw, title, font, center_x):
    tw, th = text_size(draw, title, font)
    w = max(520, tw + 130)
    h = th + 26
    x = center_x - w / 2
    y = 30
    draw.rectangle((x + 9, y + 9, x + w + 9, y + h + 9), fill=(20, 20, 20))
    draw.rectangle((x, y, x + w, y + h), fill="white", outline=(20, 20, 20), width=3)
    draw.text((center_x, y + h / 2), title, fill=(20, 20, 20), font=font, anchor="mm")


def draw_legend(draw, xy, series_label, fit_label, font):
    x, y = xy
    w, h = 250, 82
    draw.rectangle((x, y, x + w, y + h), fill="white", outline=(20, 20, 20), width=2)
    draw.rectangle((x + 24, y + 20, x + 36, y + 32), fill=(20, 20, 20))
    draw.text((x + 52, y + 27), series_label, fill=(20, 20, 20), font=font, anchor="lm")
    draw.line((x + 20, y + 58, x + 42, y + 58), fill=(20, 20, 20), width=3)
    draw.text((x + 52, y + 58), fit_label, fill=(20, 20, 20), font=font, anchor="lm")


def draw_plot(path, xs, ys, k, b, title, xlabel, ylabel, series_label, fit_label, table_rows, xfmt, yfmt):
    w, h = 1400, 950
    ml, mr, mt, mb = 150, 480, 145, 130
    plot_w, plot_h = w - ml - mr, h - mt - mb

    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    # Leave a little more breathing room than a tight auto-scale. The combined
    # data joins two source files, so a wider scale keeps that join from looking
    # visually harsher than the measured values justify.
    xpad = (xmax - xmin) * 0.09 or 1
    ypad = (ymax - ymin) * 0.16 or 1
    xmin, xmax = xmin - xpad, xmax + xpad
    ymin, ymax = ymin - ypad, ymax + ypad

    def sx(x):
        return ml + (x - xmin) / (xmax - xmin) * plot_w

    def sy(y):
        return mt + (ymax - y) / (ymax - ymin) * plot_h

    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    font_title = find_font(34)
    font = find_font(25)
    font_small = find_font(21)
    font_table = find_font(17)

    axis = (30, 30, 30)
    grid = (238, 238, 238)
    point = (20, 20, 20)
    line = (20, 20, 20)

    draw_title_ribbon(d, title, font_title, w / 2)

    for xt in nice_ticks(xmin, xmax, count=5):
        x = sx(xt)
        d.line((x, mt, x, mt + plot_h), fill=grid, width=1)
        d.text((x, mt + plot_h + 18), xfmt(xt), fill=axis, font=font_small, anchor="ma")

    for yt in nice_ticks(ymin, ymax, count=5):
        y = sy(yt)
        d.line((ml, y, ml + plot_w, y), fill=grid, width=1)
        d.text((ml - 16, y), yfmt(yt), fill=axis, font=font_small, anchor="rm")

    d.line((ml, mt, ml, mt + plot_h), fill=axis, width=4)
    d.line((ml, mt + plot_h, ml + plot_w, mt + plot_h), fill=axis, width=4)

    x1, x2 = min(xs), max(xs)
    d.line((sx(x1), sy(k * x1 + b), sx(x2), sy(k * x2 + b)), fill=line, width=4)

    for x, y in zip(xs, ys):
        px, py = sx(x), sy(y)
        d.rectangle((px - 6, py - 6, px + 6, py + 6), fill=point)

    d.text((ml + plot_w / 2, h - 55), xlabel, fill=axis, font=font, anchor="ma")
    draw_rotated_label(img, ylabel, (48, mt + plot_h / 2), font, axis)
    side_x = ml + plot_w + 55
    draw_legend(d, (side_x, mt + 35), series_label, fit_label, font_small)
    draw_fit_table(d, (side_x, mt + 145), table_rows, font_table)
    draw_text_box(
        d,
        (side_x, mt + 465),
        ["姓名：陈卓远", "学号：525031910036", "班级：计科2533", "日期：2026.5.7"],
        font_small,
        padding=10,
        line_gap=4,
        border=3,
    )

    img.save(path)


def write_selected_csv(path, selected):
    t_c = [r[0] for r in selected]
    r_ohm = [r[3] for r in selected]
    u_pn = [r[4] for r in selected]
    inv_t = [1 / (t + 273.15) for t in t_c]
    ln_r = [math.log(r) for r in r_ohm]

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["t / degC", "R / ohm (CH1)", "ln(R/ohm)", "1/T / K^-1", "U / V (CH2)"])
        for row in zip(t_c, r_ohm, ln_r, inv_t, u_pn):
            writer.writerow([f"{row[0]:.3f}", f"{row[1]:.6f}", f"{row[2]:.8f}", f"{row[3]:.10f}", f"{row[4]:.8f}"])

    return t_c, r_ohm, ln_r, inv_t, u_pn


def process_one(label, path, skip_count, targets):
    rows = read_rows(path)
    heating = heating_branch(rows, skip_count)
    selected = pick_nearest(heating, targets)

    csv_path = OUT_DIR / f"{label}_selected_temperature_data.csv"
    t_c, _r_ohm, ln_r, inv_t, u_pn = write_selected_csv(csv_path, selected)
    t_k = [t + 273.15 for t in t_c]

    b_const, ln_a, r2_therm = linear_fit(inv_t, ln_r)
    k_k, ug_k, r2_pn_k = linear_fit(t_k, u_pn)

    thermistor_path = OUT_DIR / f"{label}_thermistor_lnR_invT.png"
    pn_path = OUT_DIR / f"{label}_pn_T_U.png"

    draw_plot(
        thermistor_path,
        inv_t,
        ln_r,
        b_const,
        ln_a,
        "热敏电阻lnR与T^-1曲线",
        "T^-1/K^-1",
        "lnR/lnΩ",
        "lnR/lnΩ",
        "fitted lnR",
        [
            ("方程", "y = a + b*x"),
            ("绘图", "lnR/lnΩ"),
            ("截距", f"{ln_a:.5f}"),
            ("斜率", f"{b_const:.5f}"),
            ("残差平方和", "-"),
            ("Pearson's r", f"{math.sqrt(max(r2_therm, 0)):.5f}"),
            ("R平方(COD)", f"{r2_therm:.5f}"),
            ("调整后R平方", f"{r2_therm:.5f}"),
        ],
        lambda x: f"{x:.5f}",
        lambda y: f"{y:.2f}",
    )

    draw_plot(
        pn_path,
        t_k,
        u_pn,
        k_k,
        ug_k,
        "PN结电压与T曲线",
        "T/K",
        "PN结电压U/V",
        "PN结电压U/V",
        "fitted U",
        [
            ("方程", "y = a + b*x"),
            ("绘图", "PN结电压U/V"),
            ("截距", f"{ug_k:.5f}"),
            ("斜率", f"{k_k:.8f}"),
            ("残差平方和", "-"),
            ("Pearson's r", f"{-math.sqrt(max(r2_pn_k, 0)):.5f}"),
            ("R平方(COD)", f"{r2_pn_k:.5f}"),
            ("调整后R平方", f"{r2_pn_k:.5f}"),
        ],
        lambda x: f"{x:.0f}",
        lambda y: f"{y:.3f}",
    )

    return {
        "label": label,
        "source": path.name,
        "selected_count": len(selected),
        "temperature_range": (min(t_c), max(t_c)),
        "csv": csv_path,
        "thermistor_plot": thermistor_path,
        "pn_plot": pn_path,
        "thermistor_fit": (b_const, ln_a, r2_therm),
        "pn_fit": (k_k, ug_k, r2_pn_k),
        "selected": selected,
    }


def process_combined(label, selected, source_names):
    selected = sorted(selected, key=lambda row: row[0])

    csv_path = OUT_DIR / f"{label}_selected_temperature_data.csv"
    t_c, _r_ohm, ln_r, inv_t, u_pn = write_selected_csv(csv_path, selected)
    t_k = [t + 273.15 for t in t_c]

    b_const, ln_a, r2_therm = linear_fit(inv_t, ln_r)
    k_k, ug_k, r2_pn_k = linear_fit(t_k, u_pn)

    thermistor_path = OUT_DIR / f"{label}_thermistor_lnR_invT.png"
    pn_path = OUT_DIR / f"{label}_pn_T_U.png"

    draw_plot(
        thermistor_path,
        inv_t,
        ln_r,
        b_const,
        ln_a,
        "热敏电阻lnR与T^-1曲线",
        "T^-1/K^-1",
        "lnR/lnΩ",
        "lnR/lnΩ",
        "fitted lnR",
        [
            ("方程", "y = a + b*x"),
            ("绘图", "lnR/lnΩ"),
            ("截距", f"{ln_a:.5f}"),
            ("斜率", f"{b_const:.5f}"),
            ("残差平方和", "-"),
            ("Pearson's r", f"{math.sqrt(max(r2_therm, 0)):.5f}"),
            ("R平方(COD)", f"{r2_therm:.5f}"),
            ("调整后R平方", f"{r2_therm:.5f}"),
        ],
        lambda x: f"{x:.5f}",
        lambda y: f"{y:.2f}",
    )

    draw_plot(
        pn_path,
        t_k,
        u_pn,
        k_k,
        ug_k,
        "PN结电压与T曲线",
        "T/K",
        "PN结电压U/V",
        "PN结电压U/V",
        "fitted U",
        [
            ("方程", "y = a + b*x"),
            ("绘图", "PN结电压U/V"),
            ("截距", f"{ug_k:.5f}"),
            ("斜率", f"{k_k:.8f}"),
            ("残差平方和", "-"),
            ("Pearson's r", f"{-math.sqrt(max(r2_pn_k, 0)):.5f}"),
            ("R平方(COD)", f"{r2_pn_k:.5f}"),
            ("调整后R平方", f"{r2_pn_k:.5f}"),
        ],
        lambda x: f"{x:.0f}",
        lambda y: f"{y:.3f}",
    )

    return {
        "label": label,
        "source": " + ".join(source_names),
        "selected_count": len(selected),
        "temperature_range": (min(t_c), max(t_c)),
        "csv": csv_path,
        "thermistor_plot": thermistor_path,
        "pn_plot": pn_path,
        "thermistor_fit": (b_const, ln_a, r2_therm),
        "pn_fit": (k_k, ug_k, r2_pn_k),
        "selected": selected,
    }


def main():
    OUT_DIR.mkdir(exist_ok=True)
    txt_files = sorted(ROOT.glob("*.txt*"), key=lambda p: p.stat().st_size)
    room_path = txt_files[0]
    main_path = txt_files[-1]

    results = [
        process_one("room_to_30C", room_path, 15, list(range(23, 35))),
        process_one("30C_to_90C", main_path, 22, list(range(31, 91, 2))),
    ]
    results.append(
        process_combined(
            "combined_25C_to_90C",
            pick_nearest(heating_branch(read_rows(room_path), 15), [25, 27, 29])
            + pick_nearest(heating_branch(read_rows(main_path), 22), list(range(31, 90, 2))),
            [room_path.name, main_path.name],
        )
    )

    for result in results:
        print(f"[{result['label']}] source: {result['source']}")
        print(f"  selected points: {result['selected_count']}")
        print(f"  temperature range: {result['temperature_range'][0]:.3f} to {result['temperature_range'][1]:.3f} degC")
        print(f"  csv: {result['csv']}")
        print(f"  thermistor plot: {result['thermistor_plot']}")
        print(f"  PN plot: {result['pn_plot']}")
        b_const, ln_a, r2_therm = result["thermistor_fit"]
        k_k, ug_k, r2_pn = result["pn_fit"]
        print(f"  thermistor fit: lnR = {b_const:.6f}*(1/T) + {ln_a:.6f}, R^2 = {r2_therm:.6f}")
        print(f"  PN fit: U = {k_k:.9f}*T + {ug_k:.9f}, R^2 = {r2_pn:.6f}")
    print("Processing note: selected points are copied from raw TXT rows; only ln(R) and 1/T are calculated.")


if __name__ == "__main__":
    main()
