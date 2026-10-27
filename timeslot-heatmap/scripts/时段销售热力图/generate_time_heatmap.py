from __future__ import annotations

import os
from datetime import date
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


# 根目录：默认由本文件位置上溯两级（<root>/scripts/<子目录>/xxx.py）。
# 可用环境变量 HEATMAP_ROOT 覆盖 —— 数据放在别处、或把脚本装进 Skill 目录时用得上。
# 四个脚本（两个生成 + 两个后处理）共用这一个定义。
PROJECT_ROOT = Path(os.environ.get("HEATMAP_ROOT") or Path(__file__).resolve().parents[2])

SOURCE_XLSX = (
    PROJECT_ROOT
    / "data/丽达时段销售/报表_超市分析_实时销售分析_时段销售_去海天仙侠岭.xlsx"
)
OUTPUT_PNG = PROJECT_ROOT / "outputs/时段销售_时段占比热力图_75天口径.png"

# 中文字体候选，按平台顺序探测。四个脚本共用这一张表：换机器只需在这里补一个本机字体路径。
# 必须含中文字形，否则中文会渲染成方框。
FONT_CANDIDATES = (
    "/System/Library/Fonts/Hiragino Sans GB.ttc",             # macOS 冬青黑体
    "/System/Library/Fonts/PingFang.ttc",                     # macOS 苹方
    "C:/Windows/Fonts/msyh.ttc",                              # Windows 微软雅黑
    "C:/Windows/Fonts/simhei.ttf",                            # Windows 黑体
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",  # Linux Noto CJK
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",          # Linux 文泉驿微米黑
)


def resolve_font_path() -> str:
    for candidate in FONT_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    raise SystemExit(
        "找不到可用的中文字体，请在 FONT_CANDIDATES 中补充本机字体绝对路径。已尝试：\n  "
        + "\n  ".join(FONT_CANDIDATES)
    )


FONT_PATH = resolve_font_path()


def require_columns(df: pd.DataFrame, columns: tuple[str, ...], source: Path) -> None:
    """确认源表有需要的列。

    报表导出格式一变（少列、改名、表头行数不同）就会在这里停下，
    否则后面 df["门店"] 之类的取值会 KeyError，或者更糟——静默取到错列。
    """
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise SystemExit(
            f"源表缺少必需列：{missing}\n"
            f"  文件：{source}\n"
            f"  实际列：{list(df.columns)[:12]}{' …' if len(df.columns) > 12 else ''}\n"
            f"  提示：本脚本按「标题行 + 表头行」两行结构读取（header=1）。"
            f"若该企业的导出没有标题行，把 header 改成 0。"
        )


def require_time_buckets(df: pd.DataFrame, time_order: tuple[str, ...] | list[str], source: Path) -> None:
    """确认源表的时段分档与 TIME_ORDER 完全一致。

    这条检查是为「换一家企业」准备的：不同企业的报表可能用不同的时段桶
    （比如两小时一档）。load_*() 是按 TIME_ORDER 逐档取值的，取不到会兜成 0，
    于是会安静地产出一张全零热力图——看着正常，实际没有任何数据。
    所以这里宁可停下，并把差异摊开。
    """
    actual = {str(t).strip() for t in df["时段"].dropna().unique()} - {"小计", "合计"}
    expected = set(time_order)
    if actual == expected:
        return
    raise SystemExit(
        "源表的时段分档与 TIME_ORDER 不一致，继续跑会得到一张全零的图。\n"
        f"  文件：{source}\n"
        f"  报表里有而 TIME_ORDER 没有（{len(actual - expected)} 个）：{sorted(actual - expected)}\n"
        f"  TIME_ORDER 有而报表里没有（{len(expected - actual)} 个）：{sorted(expected - actual)}\n"
        "  改法：把 generate_time_heatmap.py 的 TIME_ORDER 改成这份报表实际的时段标签，"
        "顺序按真实时间先后排。"
    )


def require_source(path: Path) -> Path:
    """读源表前先确认存在。

    直接抛 FileNotFoundError 会让人以为是代码坏了；实际上九成是数据没放对位置，
    或者 HEATMAP_ROOT 指向的目录没有镜像出 `data/丽达时段销售/` 这一层。
    所以这里给出可操作的两种改法，而不是让调用方去看栈。
    """
    if path.exists():
        return path
    relative = path.parent.relative_to(PROJECT_ROOT) if path.is_relative_to(PROJECT_ROOT) else path.parent
    raise SystemExit(
        f"找不到源表：{path}\n"
        f"\n"
        f"两种改法：\n"
        f"  1) 把报表放到上面那个路径下；或\n"
        f"  2) 设 HEATMAP_ROOT 指向一个包含 {relative}/ 子目录的根目录\n"
        f"     （它必须镜像包内结构：<root>/data/丽达时段销售/ 与 <root>/outputs/）\n"
        f"\n"
        f"当前 PROJECT_ROOT = {PROJECT_ROOT}"
    )

# 画布几何：apply_store_type_legend.py 会 import 这几个常量叠加图标列与图例，
# 改这里两边同时生效。须满足 LEFT_MARGIN + CELL_WIDTH×15 + RIGHT_MARGIN ≤ CANVAS_WIDTH，
# 否则右侧「20点-21点 / 21点后」两列会被裁掉（原 cell_width=104 时为 1855 > 1600，已修）。
CANVAS_WIDTH = 1600
CANVAS_HEIGHT = 900
LEFT_MARGIN = 185
RIGHT_MARGIN = 110
TOP_MARGIN = 185
CELL_WIDTH = 85
CELL_HEIGHT = 58

START_DATE = date(2026, 3, 1)
END_DATE = date(2026, 5, 14)
DATE_DAYS = (END_DATE - START_DATE).days + 1

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


def font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH, size=size)


def mix_color(low: tuple[int, int, int], high: tuple[int, int, int], ratio: float) -> tuple[int, int, int]:
    ratio = max(0.0, min(1.0, ratio))
    return tuple(int(a + (b - a) * ratio) for a, b in zip(low, high))


def text_size(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.FreeTypeFont) -> tuple[int, int]:
    box = draw.textbbox((0, 0), text, font=text_font)
    return box[2] - box[0], box[3] - box[1]


def clean_store_name(store: str) -> str:
    if "]" in store:
        code, name = store.split("]", 1)
        return f"{code}] {name}"
    return store


def load_heatmap_data() -> tuple[list[str], list[str], list[list[float]], list[float]]:
    df = pd.read_excel(require_source(SOURCE_XLSX), header=1)
    require_columns(df, ("门店", "时段", "交易笔数"), SOURCE_XLSX)
    require_time_buckets(df, TIME_ORDER, SOURCE_XLSX)
    df = df[df["门店"].notna() & df["时段"].notna()].copy()
    df = df[df["时段"] != "小计"].copy()

    records: list[tuple[str, float, dict[str, float]]] = []
    for store_name, group in df.groupby("门店"):
        total_transactions = float(group["交易笔数"].sum())
        if total_transactions <= 0:
            continue

        ratio_by_time = (group.set_index("时段")["交易笔数"] / total_transactions * 100).to_dict()
        ordered_ratio = {time_label: round(float(ratio_by_time.get(time_label, 0.0)), 1) for time_label in TIME_ORDER}
        avg_daily_transactions = total_transactions / DATE_DAYS
        records.append((clean_store_name(store_name), avg_daily_transactions, ordered_ratio))

    records.sort(key=lambda item: item[1], reverse=True)

    stores = [item[0] for item in records]
    avg_daily = [item[1] for item in records]
    matrix = [[item[2][time_label] for time_label in TIME_ORDER] for item in records]
    return stores, TIME_ORDER, matrix, avg_daily


def draw_heatmap() -> None:
    stores, times, matrix, avg_daily = load_heatmap_data()

    image_width = CANVAS_WIDTH
    image_height = CANVAS_HEIGHT
    left_margin = LEFT_MARGIN
    right_margin = RIGHT_MARGIN
    top_margin = TOP_MARGIN
    bottom_margin = 55
    cell_width = CELL_WIDTH
    cell_height = CELL_HEIGHT
    grid_width = cell_width * len(times)
    grid_height = cell_height * len(stores)

    img = Image.new("RGB", (image_width, image_height), "#f7f7f5")
    draw = ImageDraw.Draw(img)

    title_font = font(26)
    subtitle_font = font(14)
    header_font = font(16)
    axis_font = font(13)
    label_font = font(12)
    value_font = font(15)
    avg_font = font(20)

    title = "各门店时段交易占比热力图（日均）"
    subtitle = f"数据周期：{START_DATE.year}/{START_DATE.month}/{START_DATE.day} - {END_DATE.year}/{END_DATE.month}/{END_DATE.day}（共 {DATE_DAYS} 天）"
    note = "门店按日均交易量降序  ·  单元格=该时段交易占比  ·  右侧=日均交易量"

    title_w, _ = text_size(draw, title, title_font)
    subtitle_w, _ = text_size(draw, subtitle, subtitle_font)
    note_w, _ = text_size(draw, note, subtitle_font)

    draw.text(((image_width - title_w) / 2, 28), title, fill="#17396b", font=title_font)
    draw.text(((image_width - subtitle_w) / 2, 72), subtitle, fill="#7d8794", font=subtitle_font)
    draw.text(((image_width - note_w) / 2, 96), note, fill="#9aa3af", font=subtitle_font)

    legend_left = 145
    legend_top = 118
    legend_width = 220
    legend_height = 18
    low_color = (247, 239, 206)
    high_color = (126, 15, 19)
    for x in range(legend_width):
        ratio = x / max(legend_width - 1, 1)
        draw.line(
            [(legend_left + x, legend_top), (legend_left + x, legend_top + legend_height)],
            fill=mix_color(low_color, high_color, ratio),
            width=1,
        )
    draw.text((legend_left - 22, legend_top - 1), "低", fill="#8a8f98", font=label_font)
    draw.text((legend_left + legend_width + 8, legend_top - 1), "高", fill="#8a8f98", font=label_font)

    header_top = top_margin - 32
    draw.rounded_rectangle(
        [(0, header_top), (image_width, header_top + 34)],
        radius=0,
        fill="#2f5b9e",
    )

    store_header = "门店"
    avg_header = "日均\n交易量"
    store_header_w, _ = text_size(draw, store_header, header_font)
    draw.text((left_margin / 2 - store_header_w / 2, header_top + 7), store_header, fill="white", font=header_font)

    for index, time_label in enumerate(times):
        cell_left = left_margin + index * cell_width
        label_w, _ = text_size(draw, time_label, axis_font)
        draw.text((cell_left + (cell_width - label_w) / 2, header_top + 8), time_label, fill="white", font=axis_font)

    avg_header_lines = avg_header.splitlines()
    for line_index, line_text in enumerate(avg_header_lines):
        line_w, _ = text_size(draw, line_text, axis_font)
        draw.text(
            (left_margin + grid_width + (right_margin - line_w) / 2, header_top + 2 + line_index * 15),
            line_text,
            fill="white",
            font=axis_font,
        )

    min_value = min(min(row) for row in matrix)
    max_value = max(max(row) for row in matrix)

    for row_index, store_name in enumerate(stores):
        row_top = top_margin + row_index * cell_height
        row_mid_y = row_top + cell_height / 2

        if row_index % 2 == 1:
            draw.rectangle([(0, row_top), (image_width, row_top + cell_height)], fill="#fbfbfa")

        store_w, store_h = text_size(draw, store_name, axis_font)
        draw.text((left_margin - 14 - store_w, row_mid_y - store_h / 2), store_name, fill="#304c7f", font=axis_font)

        avg_text = f"{avg_daily[row_index] / 1000:.1f}k"
        avg_w, avg_h = text_size(draw, avg_text, avg_font)
        draw.text(
            (left_margin + grid_width + (right_margin - avg_w) / 2, row_mid_y - avg_h / 2 - 2),
            avg_text,
            fill="#3b72d5",
            font=avg_font,
        )

        for col_index, value in enumerate(matrix[row_index]):
            cell_left = left_margin + col_index * cell_width
            cell_top = row_top
            ratio = (value - min_value) / (max_value - min_value) if max_value > min_value else 0.0
            fill = mix_color(low_color, high_color, ratio)
            draw.rectangle(
                [(cell_left, cell_top), (cell_left + cell_width - 3, cell_top + cell_height - 3)],
                fill=fill,
            )

            value_text = f"{value:.1f}%"
            value_w, value_h = text_size(draw, value_text, value_font)
            text_fill = "white" if ratio > 0.45 else "#8d939c"
            draw.text(
                (cell_left + (cell_width - value_w) / 2, cell_top + (cell_height - value_h) / 2 - 2),
                value_text,
                fill=text_fill,
                font=value_font,
            )

    draw.rectangle(
        [(left_margin, top_margin), (left_margin + grid_width, top_margin + grid_height)],
        outline="#b4c5e2",
        width=2,
    )
    draw.line(
        [(left_margin + grid_width, top_margin), (left_margin + grid_width, top_margin + grid_height)],
        fill="#d9dde4",
        width=2,
    )

    OUTPUT_PNG.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUTPUT_PNG, quality=95)


if __name__ == "__main__":
    draw_heatmap()
