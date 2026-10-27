"""后处理：为时段热力图叠加门店类型图标与说明图例。

变体（位置参数，默认 ratio）：
  ratio  输入 outputs/时段销售_时段占比热力图_75天口径.png
  daily  输入 outputs/时段销售_日均笔数热力图_75天口径_3以下置0.png
输出均为同目录下 `<原文件名>_门店类型图例版.png`。

门店类型在 STORE_TYPES 中按"店名 -> 类型列表"配置；
门店顺序由对应生成脚本动态给出（占比版=日均交易量降序，日均版=销售收入降序），
不依赖硬编码行序，便于后续源数据变动时图标自动跟随。

几何参数从 generate_time_heatmap.py import，避免两处手抄不一致。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from generate_time_daily_heatmap import OUTPUT_PNG as DAILY_PNG
from generate_time_daily_heatmap import load_daily_data
from generate_time_heatmap import CELL_HEIGHT, FONT_PATH, LEFT_MARGIN, TOP_MARGIN
from generate_time_heatmap import OUTPUT_PNG as RATIO_PNG
from generate_time_heatmap import load_heatmap_data

# 变体 -> (底图, 取门店行序的函数)
VARIANTS = {
    "ratio": (RATIO_PNG, load_heatmap_data),
    "daily": (DAILY_PNG, load_daily_data),
}

# 图标窄栏宽度。店名文字约占 x73~171（left_margin=185 减去 14 间距），
# 所以图标只占最左侧 x<ICON_STRIP_RIGHT，不得侵入店名区域。
ICON_STRIP_RIGHT = 68

# 店名 -> 类型列表。键使用生成脚本 clean_store_name() 后的字符串。
STORE_TYPES: dict[str, list[str]] = {
    "[1028] 崂山超市":    ["flagship"],
    "[1101] 凯德店":       ["flow"],
    "[1002] 胶州店":       ["community"],
    "[1106] 辽阳路店":     ["flow"],
    "[1107] 览秀城店":     ["flow"],
    "[1003] 城阳店":       ["community"],
    "[1015] 绿城店":       ["flow"],
    "[1010] 市北店":       ["community"],
    "[1105] 万象城店":     ["morning"],
    "[1011] 海阳店":       ["flow", "afternoon"],
}

NAVY = "#17396b"
GREEN = "#1F9D55"
BURGUNDY = "#8A1538"
PURPLE = "#7A4CC2"
YELLOW = "#F5B700"
BORDER = "#D8E0EE"
ROW_ALT = "#fbfbfa"


def font(size):
    return ImageFont.truetype(FONT_PATH, size=size)


def draw_flow(draw, x, y, color, scale=1.0):
    w = max(2, int(4 * scale))
    pts = [
        (x - 14 * scale, y + 10 * scale),
        (x - 6 * scale, y - 3 * scale),
        (x + 2 * scale, y + 4 * scale),
        (x + 12 * scale, y - 11 * scale),
    ]
    draw.line(pts, fill=color, width=w, joint="curve")
    for px, py in pts:
        r = 4 * scale
        draw.ellipse([px - r, py - r, px + r, py + r], fill="white", outline=color, width=w)


def draw_building(draw, x, y, color, scale=1.0):
    w = max(2, int(4 * scale))
    draw.rounded_rectangle(
        [x - 14 * scale, y - 15 * scale, x + 14 * scale, y + 15 * scale],
        radius=max(1, int(3 * scale)),
        outline=color,
        width=w,
        fill="white",
    )
    for dx in (-7, 2):
        for dy in (-8, 1):
            draw.rectangle(
                [x + dx * scale, y + dy * scale, x + (dx + 4) * scale, y + (dy + 4) * scale],
                fill=color,
            )
    draw.line([(x - 14 * scale, y + 15 * scale), (x + 14 * scale, y + 15 * scale)], fill=color, width=w)


def draw_flagship(draw, x, y, color, scale=1.0):
    w = max(2, int(4 * scale))
    draw.line([(x - 11 * scale, y - 15 * scale), (x - 11 * scale, y + 15 * scale)], fill=color, width=w)
    flag = [
        (x - 9 * scale, y - 14 * scale),
        (x + 15 * scale, y - 10 * scale),
        (x + 7 * scale, y - 1 * scale),
        (x - 9 * scale, y + 1 * scale),
    ]
    draw.polygon(flag, outline=color, fill="white")
    draw.line(flag + [flag[0]], fill=color, width=w, joint="curve")
    draw.line([(x - 18 * scale, y + 15 * scale), (x + 3 * scale, y + 15 * scale)], fill=color, width=w)


def draw_morning(draw, x, y, color, scale=1.0):
    w = max(2, int(4 * scale))
    draw.arc([x - 13 * scale, y - 2 * scale, x + 13 * scale, y + 24 * scale], 180, 360, fill=color, width=w)
    draw.line([(x - 18 * scale, y + 11 * scale), (x + 18 * scale, y + 11 * scale)], fill=color, width=w)
    for dx, dy in [(-15, -6), (-7, -13), (0, -16), (7, -13), (15, -6)]:
        draw.line([(x, y - 2 * scale), (x + dx * scale, y + dy * scale)], fill=color, width=max(2, int(3 * scale)))


def draw_afternoon(draw, x, y, color, accent, scale=1.0):
    w = max(2, int(4 * scale))
    draw.arc([x - 15 * scale, y - 17 * scale, x + 9 * scale, y + 7 * scale], 210, 520, fill=color, width=w)
    wave = [
        (x - 17 * scale, y + 9 * scale),
        (x - 7 * scale, y + 2 * scale),
        (x + 3 * scale, y + 9 * scale),
        (x + 14 * scale, y + 2 * scale),
    ]
    draw.line(wave, fill=accent, width=w, joint="curve")
    draw.line([(x - 18 * scale, y + 16 * scale), (x + 17 * scale, y + 16 * scale)], fill=accent, width=w)


def draw_kind(draw, kind, x, y, scale=1.0):
    if kind == "flow":
        draw_flow(draw, x, y, GREEN, scale)
    elif kind == "community":
        draw_building(draw, x, y, BURGUNDY, scale)
    elif kind == "flagship":
        draw_flagship(draw, x, y, PURPLE, scale)
    elif kind == "morning":
        draw_morning(draw, x, y - 1 * scale, YELLOW, scale)
    elif kind == "afternoon":
        draw_afternoon(draw, x, y - 1 * scale, YELLOW, GREEN, scale)


def overlay_icon_column(image, stores):
    draw = ImageDraw.Draw(image)
    missing = [s for s in stores if not STORE_TYPES.get(s)]
    if missing:
        print(
            f"⚠️  {len(missing)}/{len(stores)} 家门店在 STORE_TYPES 里没有配置，这些行不会有类型图标：\n"
            f"    {missing}\n"
            f"    换一家企业时，需要在 STORE_TYPES 里补上该企业的门店名（键用 clean_store_name() 之后的字符串）。",
            file=sys.stderr,
        )
    for idx in range(len(stores)):
        row_top = TOP_MARGIN + idx * CELL_HEIGHT
        row_bottom = row_top + CELL_HEIGHT
        fill = ROW_ALT if idx % 2 == 1 else "#f7f7f5"
        # 只清最左侧窄栏（店名从 x≈73 起，不受影响），保证重复运行不叠影
        draw.rectangle([(0, row_top), (ICON_STRIP_RIGHT, row_bottom)], fill=fill)
        kinds = STORE_TYPES.get(stores[idx], [])
        cy = row_top + CELL_HEIGHT / 2
        if len(kinds) == 1:
            draw_kind(draw, kinds[0], 32, cy, 1.0)
        elif len(kinds) >= 2:
            draw_kind(draw, kinds[0], 17, cy, 0.78)
            draw_kind(draw, kinds[1], 48, cy, 0.78)


def overlay_legend(image):
    # 顶部留白被标题/渐变色图例/note 占满，图例放到底部 800~852 留白区
    draw = ImageDraw.Draw(image)
    title_font = font(17)
    text_font = font(14)
    box_left, box_top = 220, 800
    box_right, box_bottom = 1380, 852
    draw.rounded_rectangle(
        [(box_left, box_top), (box_right, box_bottom)],
        radius=8,
        fill=(255, 255, 255, 244),
        outline=BORDER,
        width=1,
    )
    draw.text((box_left + 16, box_top + 15), "门店类型说明", fill=NAVY, font=title_font)
    # 起点从 box_left+138 右移到 +150，与标题文字留出间距（原先第一个图标贴着标题）
    items = [
        ("flow",      "商圈流量型",     box_left + 150),
        ("flagship",   "标杆旗舰型",     box_left + 360),
        ("community", "社区民生型",     box_left + 570),
        ("morning",   "商圈上午流量型", box_left + 780),
        ("afternoon", "商圈下午流量型", box_left + 1015),
    ]
    y = box_top + 28
    for kind, label, x in items:
        draw_kind(draw, kind, x, y, 0.58)
        draw.text((x + 22, y - 8), label, fill=NAVY, font=text_font)


def main():
    parser = argparse.ArgumentParser(description="为时段热力图叠加门店类型图标与图例")
    parser.add_argument(
        "variant",
        nargs="?",
        default="ratio",
        choices=sorted(VARIANTS),
        help="叠加哪个变体（默认 ratio=占比版，daily=日均笔数版）",
    )
    variant = parser.parse_args().variant

    input_png, load_stores = VARIANTS[variant]
    stores, _times, _matrix, _avg = load_stores()
    image = Image.open(input_png).convert("RGBA")
    overlay_icon_column(image, stores)
    overlay_legend(image)
    output_png = input_png.with_name(f"{input_png.stem}_门店类型图例版.png")
    output_png.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(output_png, quality=95)
    print(output_png)


if __name__ == "__main__":
    main()
