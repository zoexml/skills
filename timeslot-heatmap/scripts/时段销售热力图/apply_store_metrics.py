"""后处理：为时段热力图右侧追加门店三件套指标（SKU 数 / 销售额 / 客单价）。

数据源：`data/丽达时段销售/丽达客单价等数据_0301-0514.xlsx` —— **门店粒度**，不是门店×时段粒度。
取三列：`近三月有销售SKU` / `近三月有销售额` / `近三月客单价`。
该表窗口 0301-0514 与热力图的 75 天口径（2026/03/01-2026/05/14）一致。

为什么必须用第二张表：
  时段销售表 21 列里**没有 SKU 数**；而且它的客单价口径也不同
  （销售额 ÷ 交易笔数 = 55.2，本表客单价 = 86）。
  本表客单价由上游按未暴露的第三量（大概是去重小单数）算好，只能整列引用，不要自己复算。

做法：**在底图右侧拼接一块指标面板**，不改动底图任何像素，
所以 generate_*_heatmap.py 的几何和已逐格验证过的数值都不受影响。

门店行序由对应生成脚本动态给出（占比版=日均交易量降序，日均版=销售收入降序），
与底图行序一致；join 键是门店编码（底图 `[1028] 崂山超市` ↔ 本表 `1028`）。

用法：
  python apply_store_metrics.py                                   # 占比版
  python apply_store_metrics.py daily                             # 日均笔数版
  python apply_store_metrics.py ratio --input outputs/xxx_门店类型图例版.png
                                                                  # 与图标图例版串联
输出：`<底图名>_门店指标版.png`
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

from generate_time_daily_heatmap import OUTPUT_PNG as DAILY_PNG
from generate_time_daily_heatmap import load_daily_data
from generate_time_heatmap import CELL_HEIGHT, FONT_PATH, PROJECT_ROOT, TOP_MARGIN, require_columns, require_source
from generate_time_heatmap import OUTPUT_PNG as RATIO_PNG
from generate_time_heatmap import load_heatmap_data


SOURCE_XLSX = PROJECT_ROOT / "data/丽达时段销售/丽达客单价等数据_0301-0514.xlsx"

# 变体 -> (底图, 取门店行序的函数)
VARIANTS = {
    "ratio": (RATIO_PNG, load_heatmap_data),
    "daily": (DAILY_PNG, load_daily_data),
}

# 面板宽度：够放「销售额：30,718万」这一行（4 个汉字 + 冒号 + 6 位数字 + 汉字）
PANEL_WIDTH = 260
PANEL_PAD = 16

# 与 generate_*_heatmap.py 里表头/隔行底色保持一致（那两处是字面量，改色需两处同步）
HEADER_FILL = "#2f5b9e"
ROW_ALT = "#fbfbfa"
ROW_FILL = "#f7f7f5"
DIVIDER = "#d9dde4"
NAVY = "#17396b"
MUTED = "#5c6470"

CODE_RE = r"\[(\d+)\]"


def font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH, size=size)


def load_store_metrics() -> dict[str, tuple[int, int, int]]:
    """门店编码 -> (有销售SKU, 销售额去尾到万, 客单价)。"""
    df = pd.read_excel(require_source(SOURCE_XLSX))
    require_columns(
        df, ("门店编码", "近三月有销售SKU", "近三月有销售额", "近三月客单价"), SOURCE_XLSX
    )
    metrics: dict[str, tuple[int, int, int]] = {}
    for _, row in df.iterrows():
        code = str(int(row["门店编码"]))
        metrics[code] = (
            int(row["近三月有销售SKU"]),
            # 原报告第 8 页附近那版图上是「去尾」不是四舍五入：30,718,497 → 3071 万
            int(row["近三月有销售额"]) // 10000,
            int(row["近三月客单价"]),
        )
    return metrics


def overlay_metrics_panel(image: Image.Image, stores: list[str], metrics: dict[str, tuple[int, int, int]]) -> Image.Image:
    width, height = image.size
    canvas = Image.new("RGB", (width + PANEL_WIDTH, height), ROW_FILL)
    canvas.paste(image, (0, 0))
    draw = ImageDraw.Draw(canvas)

    panel_left = width
    # 表头横条延续到面板，视觉上连成一条
    header_top = TOP_MARGIN - 32
    draw.rectangle([(panel_left, header_top), (width + PANEL_WIDTH, header_top + 34)], fill=HEADER_FILL)
    header_font = font(13)
    header = "SKU / 销售额 / 客单"
    header_w, header_h = draw.textbbox((0, 0), header, font=header_font)[2:4]
    draw.text(
        (panel_left + (PANEL_WIDTH - header_w) / 2, header_top + 17 - header_h / 2 - 1),
        header,
        fill="white",
        font=header_font,
    )

    label_font = font(12)
    value_font = font(12)
    misses: list[str] = []
    for idx, store in enumerate(stores):
        row_top = TOP_MARGIN + idx * CELL_HEIGHT
        if idx % 2 == 1:
            draw.rectangle([(panel_left, row_top), (width + PANEL_WIDTH, row_top + CELL_HEIGHT)], fill=ROW_ALT)

        code_match = re.search(CODE_RE, store)
        values = metrics.get(code_match.group(1)) if code_match else None
        if values is None:
            lines = [("—", "不在指标表中")]
            misses.append(store)
        else:
            sku, revenue_wan, ticket = values
            lines = [
                ("SKU", f"{sku:,}"),
                ("销售额", f"{revenue_wan:,} 万"),
                ("客单", f"¥{ticket:,}"),
            ]

        for line_index, (label, value) in enumerate(lines):
            line_y = row_top + 6 + line_index * 17
            draw.text((panel_left + PANEL_PAD, line_y), label, fill=MUTED, font=label_font)
            draw.text((panel_left + PANEL_PAD + 52, line_y), value, fill=NAVY, font=value_font)

    if misses:
        print(
            f"⚠️  {len(misses)}/{len(stores)} 家门店在三件套表里没匹配到，面板显示「不在指标表中」：\n"
            f"    {misses}\n"
            f"    门店编码按正则 {CODE_RE} 从热力图门店名里提取。换一家企业时，"
            f"若其门店名不是「[编码]名称」这种写法，需要同步改 CODE_RE。",
            file=sys.stderr,
        )

    draw.line([(panel_left, header_top), (panel_left, height)], fill=DIVIDER, width=2)
    return canvas


def main() -> None:
    parser = argparse.ArgumentParser(description="为时段热力图追加门店三件套指标面板")
    parser.add_argument(
        "variant",
        nargs="?",
        default="ratio",
        choices=sorted(VARIANTS),
        help="基于哪个变体（默认 ratio=占比版，daily=日均笔数版）",
    )
    parser.add_argument("--input", type=Path, default=None, help="覆盖底图路径，用于与 apply_store_type_legend.py 串联")
    args = parser.parse_args()

    default_png, load_stores = VARIANTS[args.variant]
    input_png = args.input or default_png
    stores = load_stores()[0]

    image = Image.open(input_png).convert("RGB")
    canvas = overlay_metrics_panel(image, stores, load_store_metrics())

    output_png = input_png.with_name(f"{input_png.stem}_门店指标版.png")
    output_png.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_png)
    print(output_png)


if __name__ == "__main__":
    main()
