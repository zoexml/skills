from __future__ import annotations

from datetime import date

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# 根目录与中文字体沿用 generate_time_heatmap.py 里的那一份定义：四个脚本成组分发，
# 各自再抄一份的话，换机器或改数据目录就要改多处，迟早漏掉一个。
from generate_time_heatmap import FONT_PATH, PROJECT_ROOT, require_columns, require_source, require_time_buckets


SOURCE_XLSX = (
    PROJECT_ROOT
    / "data/丽达时段销售/报表_超市分析_实时销售分析_时段销售_去海天仙侠岭.xlsx"
)
OUTPUT_PNG = (
    PROJECT_ROOT / "outputs/时段销售_日均笔数热力图_75天口径_3以下置0.png"
)

# 画布几何：与 generate_time_heatmap.py 保持同一套 LEFT/TOP/CELL_HEIGHT，
# apply_store_type_legend.py 才能对两张热力图共用一套叠加坐标。
CANVAS_WIDTH = 1600
CANVAS_HEIGHT = 900
LEFT_MARGIN = 185
RIGHT_MARGIN = 150
TOP_MARGIN = 185
CELL_WIDTH = 84
CELL_HEIGHT = 58

START_DATE = date(2026, 3, 1)
END_DATE = date(2026, 5, 14)
DATE_DAYS = (END_DATE - START_DATE).days + 1

# 口径：平均每天交易笔数 = 交易笔数 ÷ 75 天；低于该阈值的格子按 0 展示
MIN_DISPLAY = 3.0

TIME_ORDER = [
    "8点前",
    "8点~9点",
    "9点~10点",
    "10点~11点",
    "11点~12点",
    "12点~13点",
    "13点~14点",
    "14点~15点",
    "15点~16点",
    "16点~17点",
    "17点~18点",
    "18点~19点",
    "19点~20点",
    "20点~21点",
    "21点后",
]

LOW_COLOR = (208, 232, 236)  # 浅青：低
MID_COLOR = (240, 173, 78)  # 黄橙：中
HIGH_COLOR = (192, 32, 32)  # 红：高


def font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH, size=size)


def mix_color(low: tuple[int, int, int], high: tuple[int, int, int], ratio: float) -> tuple[int, int, int]:
    ratio = max(0.0, min(1.0, ratio))
    return tuple(int(a + (b - a) * ratio) for a, b in zip(low, high))


def heat_color(ratio: float) -> tuple[int, int, int]:
    """浅青 → 黄橙 → 红 三段渐变，ratio 取 0~1。"""
    if ratio <= 0.5:
        return mix_color(LOW_COLOR, MID_COLOR, ratio / 0.5)
    return mix_color(MID_COLOR, HIGH_COLOR, (ratio - 0.5) / 0.5)


def text_size(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.FreeTypeFont) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=text_font)
    return box[2] - box[0], box[3] - box[1]


def clean_store_name(store: str) -> str:
    if "]" in store:
        code, name = store.split("]", 1)
        return f"{code}] {name}"
    return store


def load_daily_data() -> tuple[list[str], list[str], list[list[int]], list[float]]:
    df = pd.read_excel(require_source(SOURCE_XLSX), header=1)
    require_columns(df, ("门店", "时段", "交易笔数"), SOURCE_XLSX)
    require_time_buckets(df, TIME_ORDER, SOURCE_XLSX)
    df = df[df["门店"].notna() & df["时段"].notna()].copy()
    df = df[df["时段"] != "小计"].copy()

    records: list[tuple[str, float, float, dict[str, int]]] = []
    for store_name, group in df.groupby("门店"):
        total_transactions = float(group["交易笔数"].sum())
        if total_transactions <= 0:
            continue

        daily_by_time = (group.set_index("时段")["交易笔数"] / DATE_DAYS).to_dict()
        display = {
            time_label: (0 if daily_by_time.get(time_label, 0.0) < MIN_DISPLAY else round(daily_by_time.get(time_label, 0.0)))
            for time_label in TIME_ORDER
        }
        records.append(
            (
                clean_store_name(store_name),
                total_transactions / DATE_DAYS,
                float(group["销售收入"].sum()),
                display,
            )
        )

    # 行序同原报告第 8 页：按销售收入降序
    records.sort(key=lambda item: item[2], reverse=True)

    stores = [item[0] for item in records]
    daily_total = [item[1] for item in records]
    matrix = [[item[3][time_label] for time_label in TIME_ORDER] for item in records]
    return stores, TIME_ORDER, matrix, daily_total


def draw_heatmap() -> None:
    stores, times, matrix, daily_total = load_daily_data()

    image_width = CANVAS_WIDTH
    image_height = CANVAS_HEIGHT
    left_margin = LEFT_MARGIN
    right_margin = RIGHT_MARGIN
    top_margin = TOP_MARGIN
    # 15 个时段 + 左右边距须落在画布内：185 + 15×84 + 150 = 1595
    cell_width = CELL_WIDTH
    cell_height = CELL_HEIGHT
    grid_width = cell_width * len(times)

    img = Image.new("RGB", (image_width, image_height), "#f7f7f5")
    draw = ImageDraw.Draw(img)

    title_font = font(26)
    subtitle_font = font(14)
    header_font = font(16)
    axis_font = font(13)
    label_font = font(12)
    value_font = font(17)
    total_font = font(18)

    title = "各门店时段日均交易笔数热力图"
    subtitle = f"数据周期：{START_DATE.year}/{START_DATE.month}/{START_DATE.day} - {END_DATE.year}/{END_DATE.month}/{END_DATE.day}（共 {DATE_DAYS} 天）"
    note = (
        f"口径：平均每天交易笔数 = 交易笔数 ÷ {DATE_DAYS} 天  ·  "
        f"{MIN_DISPLAY:.0f} 笔/日以下按 0 展示  ·  门店按销售收入降序"
    )

    title_w, _ = text_size(draw, title, title_font)
    subtitle_w, _ = text_size(draw, subtitle, subtitle_font)
    note_w, _ = text_size(draw, note, subtitle_font)

    draw.text(((image_width - title_w) / 2, 26), title, fill="#17396b", font=title_font)
    draw.text(((image_width - subtitle_w) / 2, 68), subtitle, fill="#7d8794", font=subtitle_font)
    draw.text(((image_width - note_w) / 2, 92), note, fill="#9aa3af", font=subtitle_font)

    legend_left = 690
    legend_top = 124
    legend_width = 220
    legend_height = 16
    for x in range(legend_width):
        draw.line(
            [(legend_left + x, legend_top), (legend_left + x, legend_top + legend_height)],
            fill=heat_color(x / max(legend_width - 1, 1)),
            width=1,
        )
    draw.text((legend_left - 22, legend_top - 1), "低", fill="#8a8f98", font=label_font)
    draw.text((legend_left + legend_width + 8, legend_top - 1), "高", fill="#8a8f98", font=label_font)
    draw.text(
        (legend_left + legend_width + 40, legend_top - 1),
        "单位：笔/日",
        fill="#8a8f98",
        font=label_font,
    )

    header_top = top_margin - 32
    draw.rounded_rectangle(
        [(0, header_top), (image_width, header_top + 34)],
        radius=0,
        fill="#2f5b9e",
    )

    store_header = "门店"
    total_header = "日均\n交易量"
    store_header_w, _ = text_size(draw, store_header, header_font)
    draw.text((left_margin / 2 - store_header_w / 2, header_top + 7), store_header, fill="white", font=header_font)

    for index, time_label in enumerate(times):
        cell_left = left_margin + index * cell_width
        label_w, _ = text_size(draw, time_label, axis_font)
        draw.text((cell_left + (cell_width - label_w) / 2, header_top + 8), time_label, fill="white", font=axis_font)

    total_header_lines = total_header.splitlines()
    for line_index, line_text in enumerate(total_header_lines):
        line_w, _ = text_size(draw, line_text, axis_font)
        draw.text(
            (left_margin + grid_width + (right_margin - line_w) / 2, header_top + 2 + line_index * 15),
            line_text,
            fill="white",
            font=axis_font,
        )

    max_value = max(max(row) for row in matrix)
    max_value = max_value if max_value > 0 else 1

    for row_index, store_name in enumerate(stores):
        row_top = top_margin + row_index * cell_height
        row_mid_y = row_top + cell_height / 2

        if row_index % 2 == 1:
            draw.rectangle([(0, row_top), (image_width, row_top + cell_height)], fill="#fbfbfa")

        store_w, store_h = text_size(draw, store_name, axis_font)
        draw.text((left_margin - 14 - store_w, row_mid_y - store_h / 2), store_name, fill="#304c7f", font=axis_font)

        total_text = f"{daily_total[row_index]:,.0f}"
        total_w, total_h = text_size(draw, total_text, total_font)
        draw.text(
            (left_margin + grid_width + (right_margin - total_w) / 2, row_mid_y - total_h / 2 - 2),
            total_text,
            fill="#3b72d5",
            font=total_font,
        )

        for col_index, value in enumerate(matrix[row_index]):
            cell_left = left_margin + col_index * cell_width
            ratio = value / max_value
            draw.rectangle(
                [(cell_left, row_top), (cell_left + cell_width - 3, row_top + cell_height - 3)],
                fill=heat_color(ratio),
            )

            value_text = f"{value}"
            value_w, value_h = text_size(draw, value_text, value_font)
            text_fill = "white" if ratio > 0.55 else "#5c6470"
            draw.text(
                (cell_left + (cell_width - value_w) / 2, row_top + (cell_height - value_h) / 2 - 2),
                value_text,
                fill=text_fill,
                font=value_font,
            )

    draw.rectangle(
        [(left_margin, top_margin), (left_margin + grid_width, top_margin + len(stores) * cell_height)],
        outline="#b4c5e2",
        width=2,
    )
    draw.line(
        [(left_margin + grid_width, top_margin), (left_margin + grid_width, top_margin + len(stores) * cell_height)],
        fill="#d9dde4",
        width=2,
    )

    OUTPUT_PNG.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUTPUT_PNG)
    print(OUTPUT_PNG)


if __name__ == "__main__":
    draw_heatmap()
