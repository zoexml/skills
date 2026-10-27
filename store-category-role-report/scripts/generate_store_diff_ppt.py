from pathlib import Path
from collections import Counter

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


BASE_DIR = Path("/Users/joet/Downloads/大数据畅销品推荐")
INPUT_XLSX = BASE_DIR / "各门店代表性差异品类_企业15门店12.xlsx"
OUTPUT_PPTX = BASE_DIR / "各门店代表性差异品类分析报告_企业15门店12.pptx"
STORE_ROLE_XLSX = Path("/Users/joet/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/E28C68EA2571E42BDD1407AA361583ED/Caches/Files/2026-05/a808316c5629a684a30b8d3d697057f8/门店下的品类角色-企业15%门店12%.xlsx")
ENTERPRISE_ROLE_XLSX = Path("/Users/joet/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/E28C68EA2571E42BDD1407AA361583ED/Caches/Files/2026-05/170d005971ff83788c0db8724513e187/企业下的品类角色-企业15%门店12%.xlsx")
LEVEL_LABELS = ["3级", "4级", "5级"]
LEVEL_SHEETS = {"3级": "3级品类", "4级": "4级品类", "5级": "5级品类"}
DEPT_MAP = {"生鲜部": "生鲜", "食品部": "食品", "非食品部": "非食"}
STORE_NAME_MAP = {
    "1002": "胶州店",
    "1003": "城阳店",
    "1010": "市北店",
    "1011": "海阳店",
    "1015": "绿城店",
    "1028": "崂山超市",
    "1101": "凯德店",
    "1105": "万象城店",
    "1106": "辽阳路店",
    "1107": "览秀城店",
}
THEME = {
    "navy": (24, 43, 73),
    "blue": (46, 91, 255),
    "teal": (0, 163, 137),
    "gold": (243, 156, 18),
    "green": (39, 174, 96),
    "ink": (49, 62, 82),
    "muted": (112, 124, 145),
    "light": (246, 248, 252),
    "line": (224, 229, 236),
}
DISADV_MIN_TARGET_COVERAGE = 4


def load_data():
    summary = pd.read_excel(INPUT_XLSX, sheet_name="门店层级汇总")
    top = pd.read_excel(INPUT_XLSX, sheet_name="每店Top5代表差异品类")
    detail = {
        level: pd.read_excel(INPUT_XLSX, sheet_name=sheet)
        for level, sheet in LEVEL_SHEETS.items()
    }

    for frame in [summary, top, *detail.values()]:
        if "门店编码" in frame.columns:
            frame["门店编码"] = frame["门店编码"].astype(str).str.strip()
        if "层级" in frame.columns:
            frame["层级"] = frame["层级"].astype(str).str.strip()
        if "部门" in frame.columns:
            frame["部门"] = frame["部门"].astype(str).str.strip()
        if "品类名称" in frame.columns:
            frame["品类名称"] = frame["品类名称"].astype(str).str.strip()
    raw_store = {}
    raw_ent = {}
    for level, sheet in LEVEL_SHEETS.items():
        s = pd.read_excel(STORE_ROLE_XLSX, sheet_name=sheet, dtype=str)
        e = pd.read_excel(ENTERPRISE_ROLE_XLSX, sheet_name=sheet, dtype=str)
        s = s[s.iloc[:, 0] != s.columns[0]].copy()
        e = e[e.iloc[:, 0] != e.columns[0]].copy()
        for df in [s, e]:
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype(str).str.strip()
            for col in ["c_score", "c_sale", "c_sale_index"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
        raw_store[level] = s
        raw_ent[level] = e
    disadvantage = {}
    disadvantage_meta = {}
    for level in LEVEL_LABELS:
        s = raw_store[level]
        e = raw_ent[level]
        ent_target = e[e["c_class_role_id"] == "目标性品类"].copy()
        coverage_rows = []
        for code, group in ent_target.groupby("c_ccode"):
            sr = s[s["c_ccode"].astype(str) == str(code)].copy()
            coverage_rows.append({
                "品类编码": code,
                "品类名称": group["c_cname"].iloc[0],
                "目标性覆盖门店数": int(sr[sr["c_class_role_id"] == "目标性品类"]["c_store_id"].nunique()),
                "是否作废": "作废" in str(group["c_cname"].iloc[0]),
            })
        coverage_df = pd.DataFrame(coverage_rows)
        valid_codes = set(
            coverage_df[
                (coverage_df["目标性覆盖门店数"] >= DISADV_MIN_TARGET_COVERAGE)
                & (~coverage_df["是否作废"])
            ]["品类编码"].astype(str)
        )
        ent_target = ent_target[ent_target["c_ccode"].astype(str).isin(valid_codes)].copy()
        disadvantage_meta[level] = coverage_df
        rows = []
        for store_id in sorted([v for v in s["c_store_id"].unique().tolist() if str(v).isdigit()]):
            st_target_codes = set(s[(s["c_store_id"] == store_id) & (s["c_class_role_id"] == "目标性品类")]["c_ccode"].astype(str))
            miss = ent_target[~ent_target["c_ccode"].astype(str).isin(st_target_codes)].copy()
            if len(miss):
                miss["门店编码"] = str(store_id)
                miss["层级"] = level
                miss["品类编码"] = miss["c_ccode"]
                miss["品类名称"] = miss["c_cname"]
                miss["部门"] = miss["c_dept_name"]
                store_rows = s[(s["c_store_id"] == store_id) & (s["c_ccode"].astype(str).isin(miss["c_ccode"].astype(str)))][
                    ["c_ccode", "c_class_role_id", "c_score", "c_sale", "c_sale_index"]
                ].drop_duplicates("c_ccode")
                store_rows = store_rows.rename(
                    columns={
                        "c_ccode": "品类编码",
                        "c_class_role_id": "门店角色",
                        "c_score": "门店综合得分",
                        "c_sale": "门店销售额",
                        "c_sale_index": "门店卖力指数",
                    }
                )
                miss["企业综合得分"] = miss["c_score"]
                miss["企业销售额"] = miss["c_sale"]
                miss["企业卖力指数"] = miss["c_sale_index"]
                miss["品类编码"] = miss["c_ccode"]
                miss = miss.merge(store_rows, on="品类编码", how="left")
                rows.append(
                    miss[
                        [
                            "门店编码",
                            "层级",
                            "品类编码",
                            "品类名称",
                            "部门",
                            "企业综合得分",
                            "企业销售额",
                            "企业卖力指数",
                            "门店角色",
                            "门店综合得分",
                            "门店销售额",
                            "门店卖力指数",
                        ]
                    ]
                )
        disadvantage[level] = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(columns=["门店编码", "层级", "品类编码", "品类名称", "部门", "企业综合得分", "企业销售额", "企业卖力指数"])
    return summary, top, detail, disadvantage, disadvantage_meta


def store_label(store_id):
    store_id = str(store_id).strip()
    return STORE_NAME_MAP.get(store_id, store_id)


def compute_metrics(summary, detail, disadvantage, disadvantage_meta):
    counts_3 = int(summary[summary["层级"] == "3级"]["差异目标性品类数"].sum())
    counts_4 = int(summary[summary["层级"] == "4级"]["差异目标性品类数"].sum())
    counts_5 = int(summary[summary["层级"] == "5级"]["差异目标性品类数"].sum())
    stores = sorted(summary["门店编码"].unique().tolist())

    store_5 = summary[summary["层级"] == "5级"][["门店编码", "差异目标性品类数", "平均综合得分"]].copy()
    top_count_store = store_5.sort_values("差异目标性品类数", ascending=False).iloc[0]
    top_score_store = store_5.sort_values("平均综合得分", ascending=False).iloc[0]

    dept_counts = Counter()
    for frame in detail.values():
        dept_counts.update(frame["部门"].tolist())

    protein_keywords = ["蛋", "鸡", "鸭", "猪", "牛", "羊", "虾", "鱼", "水产", "海鲜"]
    level5 = detail["5级"].copy()
    protein_mask = level5["品类名称"].astype(str).apply(lambda x: any(k in x for k in protein_keywords))
    protein_top5 = (
        level5[protein_mask]
        .groupby(["品类编码", "品类名称", "部门"])
        .agg(
            覆盖门店数=("门店编码", "nunique"),
            销售额汇总=("销售额", "sum"),
            平均综合得分=("综合得分", "mean"),
            平均卖力指数=("卖力指数", "mean"),
        )
        .reset_index()
        .sort_values(["覆盖门店数", "平均综合得分", "销售额汇总"], ascending=[False, False, False])
        .head(5)
    )

    disadvantage_summary = []
    for level in LEVEL_LABELS:
        frame = disadvantage[level]
        if len(frame):
            per_store = frame.groupby("门店编码").size().reset_index(name="劣势品类数")
            per_store["层级"] = level
            disadvantage_summary.append(per_store)
    disadvantage_summary = pd.concat(disadvantage_summary, ignore_index=True) if disadvantage_summary else pd.DataFrame(columns=["门店编码", "劣势品类数", "层级"])

    level_ratio_rows = []
    ent_target_counts = {
        level: int(len(disadvantage[level]["品类编码"].drop_duplicates())) for level in LEVEL_LABELS
    }
    ent_target_raw_counts = {
        level: int(len(disadvantage_meta[level][["品类编码"]].drop_duplicates())) for level in LEVEL_LABELS
    }
    store_target_counts = {}
    for level, sheet in LEVEL_SHEETS.items():
        store_raw = pd.read_excel(STORE_ROLE_XLSX, sheet_name=sheet, dtype=str)
        store_raw = store_raw[store_raw.iloc[:, 0] != store_raw.columns[0]].copy()
        store_raw["c_store_id"] = store_raw["c_store_id"].astype(str).str.strip()
        store_raw["c_class_role_id"] = store_raw["c_class_role_id"].astype(str).str.strip()
        store_raw = store_raw[store_raw["c_store_id"].str.isdigit()]
        store_target_counts[level] = int(len(store_raw[store_raw["c_class_role_id"] == "目标性品类"]))
    adv_counts = {"3级": counts_3, "4级": counts_4, "5级": counts_5}
    disadv_counts = {
        level: int(disadvantage_summary.loc[disadvantage_summary["层级"] == level, "劣势品类数"].sum()) if len(disadvantage_summary) else 0
        for level in LEVEL_LABELS
    }
    for level in LEVEL_LABELS:
        level_ratio_rows.append({
            "层级": level,
            "企业原始目标性品类数": ent_target_raw_counts[level],
            "企业目标性品类数": ent_target_counts[level],
            "门店目标性品类数": store_target_counts[level],
            "优势品类数": adv_counts[level],
            "优势占门店目标性品类比例": adv_counts[level] / store_target_counts[level] if store_target_counts[level] else 0,
            "劣势品类数": disadv_counts[level],
            "劣势占企业目标性品类比例": disadv_counts[level] / (ent_target_counts[level] * len(stores)) if ent_target_counts[level] else 0,
        })
    level_ratio_df = pd.DataFrame(level_ratio_rows)

    return {
        "stores": stores,
        "counts_3": counts_3,
        "counts_4": counts_4,
        "counts_5": counts_5,
        "top_count_store": top_count_store,
        "top_score_store": top_score_store,
        "dept_counts": dept_counts,
        "protein_top5": protein_top5,
        "disadvantage_summary": disadvantage_summary,
        "level_ratio_df": level_ratio_df,
        "disadvantage_meta": disadvantage_meta,
    }


def guess_tags(store_top):
    category_rules = [
        (
            "水果细分经营",
            ["果", "莓", "瓜", "柑", "桔", "梨", "苹果", "桃", "提", "蜜柚", "椰青", "火龙果", "牛油果", "青枣", "菠萝蜜", "猕猴桃", "桑葚", "草莓", "西瓜"],
        ),
        (
            "鲜食主食经营",
            ["米制品", "馄饨", "馒头", "蒸", "卷", "热炒", "盒饭", "煮制", "面包", "蛋糕", "果切", "半成品", "谷物", "凉拌素菜", "凉拌菜", "包装熟食", "冷熟食", "冷加工熟食"],
        ),
        (
            "熟食现制经营",
            ["卤", "烤", "煎", "炸", "肠类", "火腿", "热炸", "酱卤", "熟食"],
        ),
        (
            "乳品饮品经营",
            ["酸奶", "鲜奶", "纯奶", "保久奶", "乳酸菌", "果蔬汁", "饮料", "麦片", "冲调"],
        ),
        (
            "鲜蛋白类经营",
            ["鲜蛋", "鸡", "鸭", "猪", "牛", "羊", "虾", "水产", "肉类", "畜类", "禽类"],
        ),
        (
            "非食补充经营",
            ["纸", "卫生巾", "洗衣液", "洗涤", "图书", "音像", "香烟"],
        ),
        (
            "调味零食经营",
            ["厨房调味", "饼干", "零食", "挂面", "巧克力", "糖果", "膨化", "调味"],
        ),
    ]

    scores = {label: 0.0 for label, _ in category_rules}
    for idx, row in store_top.reset_index(drop=True).iterrows():
        name = str(row["品类名称"])
        weight = float(row.get("代表性指数", 0)) if pd.notna(row.get("代表性指数", 0)) else 0
        weight = weight if weight > 0 else max(10, 100 - idx * 5)
        for label, keywords in category_rules:
            if any(k in name for k in keywords):
                scores[label] += weight

    ranked = [label for label, score in sorted(scores.items(), key=lambda item: item[1], reverse=True) if score > 0]
    if not ranked:
        return "区域化长尾机会"

    # Avoid over-tagging weak secondary buckets.
    primary = ranked[0]
    tags = [primary]
    if len(ranked) > 1 and scores[ranked[1]] >= scores[primary] * 0.45:
        tags.append(ranked[1])
    if "水果细分经营" in tags and "鲜食主食经营" in tags and scores["水果细分经营"] >= scores["鲜食主食经营"] * 0.7:
        tags = ["水果细分经营", "鲜食主食经营"]
    return " / ".join(tags)


def dominant_dept(store_frames):
    counts = Counter()
    for frame in store_frames:
        counts.update(frame["部门"].tolist())
    if not counts:
        return "无明显主导"
    dept, count = counts.most_common(1)[0]
    total = sum(counts.values())
    return f"{dept}主导（占比{count / total:.0%}）"


def summarize_focus_items(store_top):
    if len(store_top) == 0:
        return "无明显集中品类"
    ranked = (
        store_top.sort_values(["代表性指数", "综合得分"], ascending=[False, False])["品类名称"]
        .astype(str)
        .drop_duplicates()
        .tolist()
    )
    return "、".join(ranked[:3])


def build_store_profiles(summary, top, detail, disadvantage):
    theme_overrides = {
        "1011": "禽肉水产更突出，熟食鲜食同步发力",
        "1105": "水果细分经营 / 鲜食主食经营",
    }
    profiles = []
    for store in sorted(summary["门店编码"].unique().tolist()):
        top5 = top[top["门店编码"] == store].copy()
        detail_frames = [frame[frame["门店编码"] == store].copy() for frame in detail.values()]
        store_summary = summary[summary["门店编码"] == store].copy()
        peak_row = store_summary.sort_values("差异目标性品类数", ascending=False).iloc[0]
        disadv5 = disadvantage["5级"][disadvantage["5级"]["门店编码"].astype(str) == str(store)].copy()
        disadv_top = (
            disadv5.sort_values(["企业综合得分", "企业销售额"], ascending=[False, False])["品类名称"]
            .astype(str)
            .drop_duplicates()
            .tolist()
        )
        profiles.append(
            {
                "store": store,
                "theme": theme_overrides.get(str(store), guess_tags(top5.head(8))),
                "focus_items": summarize_focus_items(top5),
                "dominant_dept": dominant_dept(detail_frames),
                "peak_level": peak_row["层级"],
                "peak_count": int(peak_row["差异目标性品类数"]),
                "note": "",
                "disadv_count_5": int(len(disadv5)),
                "disadv_items": "、".join(disadv_top[:3]) if disadv_top else "无明显劣势品类",
            }
        )
    return profiles


def compute_score_cutoffs():
    cutoffs = {"enterprise": {}, "store": {}}
    for level, sheet in LEVEL_SHEETS.items():
        e = pd.read_excel(ENTERPRISE_ROLE_XLSX, sheet_name=sheet, dtype=str)
        e = e[e.iloc[:, 0] != e.columns[0]].copy()
        e["c_dept_name"] = e["c_dept_name"].astype(str).str.strip()
        e["c_class_role_id"] = e["c_class_role_id"].astype(str).str.strip()
        e["c_score"] = pd.to_numeric(e["c_score"], errors="coerce")
        cutoffs["enterprise"][level] = {}
        for dept, grp in e[e["c_class_role_id"] == "目标性品类"].groupby("c_dept_name"):
            cutoffs["enterprise"][level][dept] = float(grp["c_score"].min()) if len(grp) else None

        s = pd.read_excel(STORE_ROLE_XLSX, sheet_name=sheet, dtype=str)
        s = s[s.iloc[:, 0] != s.columns[0]].copy()
        s["c_store_id"] = s["c_store_id"].astype(str).str.strip()
        s["c_dept_name"] = s["c_dept_name"].astype(str).str.strip()
        s["c_class_role_id"] = s["c_class_role_id"].astype(str).str.strip()
        s["c_score"] = pd.to_numeric(s["c_score"], errors="coerce")
        for store_id, grp in s[s["c_store_id"].str.isdigit()].groupby("c_store_id"):
            cutoffs["store"].setdefault(str(store_id), {})[level] = {}
            for dept, dept_grp in grp[grp["c_class_role_id"] == "目标性品类"].groupby("c_dept_name"):
                cutoffs["store"][str(store_id)][level][dept] = float(dept_grp["c_score"].min()) if len(dept_grp) else None
    return cutoffs


def get_threshold(cutoffs, level, dept, store_id=None):
    dept = str(dept).strip()
    if store_id is None:
        return cutoffs["enterprise"].get(level, {}).get(dept)
    return cutoffs["store"].get(str(store_id), {}).get(level, {}).get(dept)


def format_threshold_summary(cutoffs, store_id):
    lines = []
    for level in LEVEL_LABELS:
        parts = []
        for dept in ["生鲜部", "食品部", "非食品部"]:
            ent_cut = get_threshold(cutoffs, level, dept)
            store_cut = get_threshold(cutoffs, level, dept, store_id=store_id)
            if ent_cut is not None or store_cut is not None:
                dept_label = DEPT_MAP.get(dept, dept)
                ent_text = f"{ent_cut:.0f}" if ent_cut is not None else "-"
                store_text = f"{store_cut:.0f}" if store_cut is not None else "-"
                parts.append(f"{dept_label}企{ent_text}/店{store_text}")
        if parts:
            lines.append(f"{level} " + " ".join(parts))
    return "；".join(lines)


def add_column_chart(slide, left, top, width, height, categories, series_dict, chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_dict.items():
        chart_data.add_series(name, values)
    chart = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data).chart
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.value_axis.has_major_gridlines = True
    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.value_axis.tick_labels.font.size = Pt(11)
    chart.chart_title.has_text_frame = False
    return chart


def add_line_chart(slide, left, top, width, height, categories, series_dict):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_dict.items():
        chart_data.add_series(name, values)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, left, top, width, height, chart_data).chart
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.value_axis.has_major_gridlines = True
    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.value_axis.tick_labels.font.size = Pt(11)
    return chart


def add_full_bg(slide, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*color)
    shape.line.fill.background()
    slide.shapes._spTree.remove(shape._element)
    slide.shapes._spTree.insert(2, shape._element)


def add_textbox(slide, left, top, width, height, text, font_size=20, bold=False, color=(0, 0, 0), align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    run = p.runs[0]
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    run.font.name = "PingFang SC"
    return box


def add_bullets(slide, left, top, width, height, bullets, font_size=19, color=(40, 40, 40)):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.margin_left = Pt(3)
    tf.margin_right = Pt(3)
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    for idx, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.font.size = Pt(font_size)
        p.font.color.rgb = RGBColor(*color)
        p.font.name = "PingFang SC"
        p.space_after = Pt(4)
        p.line_spacing = 1.1
    return box


def add_table(slide, left, top, width, height, data, font_size=13, header_fill=(46, 91, 255), col_widths=None, compact=False):
    rows = len(data)
    cols = len(data[0]) if rows else 0
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table
    if col_widths:
        for idx, col_width in enumerate(col_widths[:cols]):
            table.columns[idx].width = col_width
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = str(data[r][c])
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
            cell.margin_left = Pt(1 if compact else 4)
            cell.margin_right = Pt(1 if compact else 4)
            cell.margin_top = Pt(0 if compact else 2)
            cell.margin_bottom = Pt(0 if compact else 2)
            cell.text_frame.word_wrap = True
            cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER if r == 0 else PP_ALIGN.LEFT
                paragraph.font.size = Pt(font_size)
                paragraph.font.name = "PingFang SC"
                paragraph.space_after = Pt(0)
                paragraph.line_spacing = 1.0
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(*header_fill)
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.bold = True
                    paragraph.font.color.rgb = RGBColor(255, 255, 255)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252) if r % 2 else RGBColor(255, 255, 255)
    return table


def add_card(slide, left, top, width, height, fill=(255, 255, 255), line=(224, 229, 236), radius_shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    shape = slide.shapes.add_shape(radius_shape, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill)
    shape.line.color.rgb = RGBColor(*line)
    shape.line.width = Pt(1)
    return shape


def add_metric_card(slide, left, top, width, height, title, value, accent):
    add_card(slide, left, top, width, height, fill=(255, 255, 255), line=THEME["line"])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height)
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(*accent)
    bar.line.fill.background()
    add_textbox(slide, left + Inches(0.22), top + Inches(0.18), width - Inches(0.3), Inches(0.22), title, font_size=11, color=THEME["muted"])
    add_textbox(slide, left + Inches(0.22), top + Inches(0.45), width - Inches(0.3), Inches(0.35), value, font_size=21, bold=True, color=THEME["navy"])


def format_level_items(level_top, cutoffs=None, store_id=None, level=None):
    if len(level_top) == 0:
        return ["无代表性独有目标性品类"]
    items = []
    for _, row in level_top.head(2).iterrows():
        ratio_text = ""
        dept_cutoff = None
        if cutoffs is not None and store_id is not None and level is not None:
            dept_cutoff = get_threshold(cutoffs, level, row["部门"], store_id=store_id)
        if dept_cutoff and dept_cutoff > 0:
            over_gate = row["综合得分"] / dept_cutoff - 1
            ratio_text = f"｜高于阈值{over_gate:.0%}"
        items.append(f"{row['品类名称']}｜{row['部门']}｜得分{row['综合得分']:.0f}{ratio_text}｜卖力{row['卖力指数']:.1f}")
    return items


def calc_ratio(score, cutoff):
    if pd.isna(score) or cutoff in (None, 0) or pd.isna(cutoff):
        return None
    return float(score) / float(cutoff)


def format_ratio_delta(ratio):
    if ratio is None:
        return "阈值缺失"
    delta = ratio - 1
    if delta >= 0:
        return f"+{delta:.0%}"
    return f"-{abs(delta):.0%}"


def short_name(text, limit=10):
    text = str(text)
    return text if len(text) <= limit else text[:limit] + "…"


def build_advantage_dept_lists(detail_df, cutoffs, store_id):
    result = {}
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = detail_df[detail_df["部门"] == dept].copy()
        df["优势比例"] = df.apply(
            lambda r: calc_ratio(r["综合得分"], get_threshold(cutoffs, r["层级"], dept, store_id=store_id)),
            axis=1,
        )
        df = df[df["优势比例"].notna()].sort_values(["优势比例", "综合得分"], ascending=[False, False])
        items = []
        for _, row in df.head(3).iterrows():
            items.append(
                f"{short_name(row['品类名称'])}｜{row['层级']}｜优势{row['优势比例']:.2f}｜得分{row['综合得分']:.0f}"
            )
        result[dept] = items or ["无"]
    return result


def build_advantage_ratio_top(detail_df, cutoffs, store_id):
    df = detail_df.copy()
    if len(df) == 0:
        return ["无"]
    df["优势比例"] = df.apply(
        lambda r: calc_ratio(r["综合得分"], get_threshold(cutoffs, r["层级"], r["部门"], store_id=store_id)),
        axis=1,
    )
    df = df[df["优势比例"].notna()].sort_values(["优势比例", "综合得分"], ascending=[False, False])
    items = []
    for _, row in df.head(5).iterrows():
        items.append(
            f"{short_name(row['品类名称'])}｜{DEPT_MAP.get(row['部门'], row['部门'])}｜{row['层级']}｜优势{row['优势比例']:.2f}"
        )
    return items or ["无"]


def build_disadvantage_dept_lists(disadv_df, cutoffs, store_id):
    result = {}
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = disadv_df[disadv_df["部门"] == dept].copy()
        df["劣势比例"] = df.apply(
            lambda r: calc_ratio(r["门店综合得分"], get_threshold(cutoffs, r["层级"], dept, store_id=store_id)),
            axis=1,
        )
        df = df.sort_values(["劣势比例", "企业综合得分"], ascending=[True, False], na_position="last")
        items = []
        for _, row in df.head(3).iterrows():
            ratio = row.get("劣势比例")
            score_text = f"{row['门店综合得分']:.0f}" if pd.notna(row.get("门店综合得分")) else "-"
            items.append(
                f"{short_name(row['品类名称'])}｜{row['层级']}｜劣势{ratio:.2f}｜店得分{score_text}" if ratio is not None else
                f"{short_name(row['品类名称'])}｜{row['层级']}｜店得分{score_text}｜阈值缺失"
            )
        result[dept] = items or ["无"]
    return result


def build_disadvantage_ratio_top(disadv_df, cutoffs, store_id):
    df = disadv_df.copy()
    if len(df) == 0:
        return ["无"]
    df["劣势比例"] = df.apply(
        lambda r: calc_ratio(r["门店综合得分"], get_threshold(cutoffs, r["层级"], r["部门"], store_id=store_id)),
        axis=1,
    )
    df = df[df["劣势比例"].notna()].sort_values(["劣势比例", "企业综合得分"], ascending=[True, False])
    items = []
    for _, row in df.head(5).iterrows():
        score_text = f"{row['门店综合得分']:.0f}" if pd.notna(row.get("门店综合得分")) else "-"
        items.append(
            f"{short_name(row['品类名称'])}｜{DEPT_MAP.get(row['部门'], row['部门'])}｜{row['层级']}｜劣势{row['劣势比例']:.2f}"
        )
    return items or ["无"]


def build_advantage_gap_dept_lists(detail_df, cutoffs, store_id):
    result = {}
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = detail_df[detail_df["部门"] == dept].copy()
        df["阈值"] = df.apply(
            lambda r: get_threshold(cutoffs, r["层级"], dept, store_id=store_id),
            axis=1,
        )
        df = df[df["阈值"].notna()].copy()
        df["差距"] = df["综合得分"] - df["阈值"]
        df = df.sort_values(["差距", "综合得分"], ascending=[False, False])
        items = []
        for _, row in df.head(3).iterrows():
            items.append(
                f"{short_name(row['品类名称'])}｜{row['层级']}｜高阈值{row['差距']:.0f}分"
            )
        result[dept] = items or ["无"]
    return result


def build_disadvantage_gap_dept_lists(disadv_df, cutoffs, store_id):
    result = {}
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = disadv_df[disadv_df["部门"] == dept].copy()
        df["阈值"] = df.apply(
            lambda r: get_threshold(cutoffs, r["层级"], dept, store_id=store_id),
            axis=1,
        )
        df = df[df["阈值"].notna() & df["门店综合得分"].notna()].copy()
        df["差距"] = df["阈值"] - df["门店综合得分"]
        df = df.sort_values(["差距", "企业综合得分"], ascending=[False, False])
        items = []
        for _, row in df.head(3).iterrows():
            items.append(
                f"{short_name(row['品类名称'])}｜{row['层级']}｜低阈值{row['差距']:.0f}分"
            )
        result[dept] = items or ["无"]
    return result


def build_advantage_level_table(detail_df, cutoffs, store_id):
    rows = [["部门", "层级", "优势比例Top3", "得分差距Top3"]]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        for level in LEVEL_LABELS:
            df = detail_df[(detail_df["部门"] == dept) & (detail_df["层级"] == level)].copy()
            if len(df):
                cutoff = get_threshold(cutoffs, level, dept, store_id=store_id)
                df["优势比例"] = df["综合得分"].apply(lambda v: calc_ratio(v, cutoff))
                ratio_df = df[df["优势比例"].notna()].sort_values(["优势比例", "综合得分"], ascending=[False, False])
                gap_df = df[df["优势比例"].notna()].copy()
                gap_df["差距"] = gap_df["综合得分"] - cutoff if cutoff else pd.NA
                gap_df = gap_df[gap_df["差距"].notna()].sort_values(["差距", "综合得分"], ascending=[False, False])
                ratio_items = [
                    f"{short_name(r['品类名称'], 8)} {r['优势比例']:.2f}"
                    for _, r in ratio_df.head(3).iterrows()
                ]
                gap_items = [
                    f"{short_name(r['品类名称'], 8)} +{r['差距']:.0f}"
                    for _, r in gap_df.head(3).iterrows()
                ]
            else:
                ratio_items = []
                gap_items = []
            rows.append([
                DEPT_MAP.get(dept, dept),
                level,
                "\n".join(ratio_items) if ratio_items else "-",
                "\n".join(gap_items) if gap_items else "-",
            ])
    return rows


def build_disadvantage_level_table(disadv_df, cutoffs, store_id):
    rows = [["部门", "层级", "劣势指数Top3", "得分差距Top3"]]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        for level in LEVEL_LABELS:
            df = disadv_df[(disadv_df["部门"] == dept) & (disadv_df["层级"] == level)].copy()
            if len(df):
                cutoff = get_threshold(cutoffs, level, dept, store_id=store_id)
                df["劣势指数"] = df["门店综合得分"].apply(lambda v: calc_ratio(v, cutoff))
                ratio_df = df[df["劣势指数"].notna()].sort_values(["劣势指数", "企业综合得分"], ascending=[True, False])
                gap_df = df[df["劣势指数"].notna()].copy()
                gap_df["差距"] = cutoff - gap_df["门店综合得分"] if cutoff else pd.NA
                gap_df = gap_df[gap_df["差距"].notna()].sort_values(["差距", "企业综合得分"], ascending=[False, False])
                ratio_items = [
                    f"{short_name(r['品类名称'], 8)} {r['劣势指数']:.2f}"
                    for _, r in ratio_df.head(3).iterrows()
                ]
                gap_items = [
                    f"{short_name(r['品类名称'], 8)} -{r['差距']:.0f}"
                    for _, r in gap_df.head(3).iterrows()
                ]
            else:
                ratio_items = []
                gap_items = []
            rows.append([
                DEPT_MAP.get(dept, dept),
                level,
                "\n".join(ratio_items) if ratio_items else "-",
                "\n".join(gap_items) if gap_items else "-",
            ])
    return rows


def format_score_gap(value):
    if pd.isna(value):
        return "差距-"
    if value >= 0:
        return f"+{value:.0f}分"
    return f"-{abs(value):.0f}分"


def format_disadvantage_gap(value):
    if pd.isna(value):
        return "差距-"
    if value >= 0:
        return f"-{value:.0f}分"
    return f"+{abs(value):.0f}分"


def build_advantage_metric_matrix(detail_df, cutoffs, store_id, metric):
    headers = ["部门", *LEVEL_LABELS]
    rows = [headers]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        row_values = [DEPT_MAP.get(dept, dept)]
        for level in LEVEL_LABELS:
            df = detail_df[(detail_df["部门"] == dept) & (detail_df["层级"] == level)].copy()
            if len(df) == 0:
                row_values.append("-")
                continue
            cutoff = get_threshold(cutoffs, level, dept, store_id=store_id)
            df["优势比例"] = df["综合得分"].apply(lambda v: calc_ratio(v, cutoff))
            df = df[df["优势比例"].notna()].copy()
            if len(df) == 0:
                row_values.append("-")
                continue
            df["得分差距"] = df["综合得分"] - cutoff
            if metric == "ratio":
                df = df.sort_values(["优势比例", "综合得分"], ascending=[False, False])
                items = [
                    f"{short_name(r['品类名称'], 6)}｜{DEPT_MAP.get(dept, dept)}｜得分{r['综合得分']:.0f}｜{format_ratio_delta(r['优势比例'])}｜卖力{r['卖力指数']:.1f}"
                    for _, r in df.head(3).iterrows()
                ]
            else:
                df = df.sort_values(["得分差距", "综合得分"], ascending=[False, False])
                items = [
                    f"{short_name(r['品类名称'], 6)}｜{DEPT_MAP.get(dept, dept)}｜{format_score_gap(r['得分差距'])}｜得分{r['综合得分']:.0f}｜卖力{r['卖力指数']:.1f}"
                    for _, r in df.head(3).iterrows()
                ]
            row_values.append("\n".join(items))
        rows.append(row_values)
    return rows


def build_disadvantage_metric_matrix(disadv_df, cutoffs, store_id, metric):
    headers = ["部门", *LEVEL_LABELS]
    rows = [headers]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        row_values = [DEPT_MAP.get(dept, dept)]
        for level in LEVEL_LABELS:
            df = disadv_df[(disadv_df["部门"] == dept) & (disadv_df["层级"] == level)].copy()
            if len(df) == 0:
                row_values.append("-")
                continue
            cutoff = get_threshold(cutoffs, level, dept, store_id=store_id)
            df["劣势指数"] = df["门店综合得分"].apply(lambda v: calc_ratio(v, cutoff))
            df = df[df["劣势指数"].notna()].copy()
            if len(df) == 0:
                row_values.append("-")
                continue
            df["得分差距"] = cutoff - df["门店综合得分"]
            if metric == "ratio":
                df = df.sort_values(["劣势指数", "企业综合得分"], ascending=[True, False])
                items = [
                    f"{short_name(r['品类名称'], 6)}｜{DEPT_MAP.get(dept, dept)}｜店得分{r['门店综合得分']:.0f}｜{format_ratio_delta(r['劣势指数'])}｜卖力{r['门店卖力指数']:.1f}"
                    for _, r in df.head(3).iterrows()
                ]
            else:
                df = df.sort_values(["得分差距", "企业综合得分"], ascending=[False, False])
                items = [
                    f"{short_name(r['品类名称'], 6)}｜{DEPT_MAP.get(dept, dept)}｜{format_disadvantage_gap(r['得分差距'])}｜店得分{r['门店综合得分']:.0f}｜卖力{r['门店卖力指数']:.1f}"
                    for _, r in df.head(3).iterrows()
                ]
            row_values.append("\n".join(items))
        rows.append(row_values)
    return rows


def build_advantage_gap_cross_level_table(detail_df, cutoffs, store_id):
    rows = [["部门", "Top1", "Top2", "Top3"]]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = detail_df[detail_df["部门"] == dept].copy()
        if len(df):
            df["阈值"] = df.apply(lambda r: get_threshold(cutoffs, r["层级"], dept, store_id=store_id), axis=1)
            df = df[df["阈值"].notna()].copy()
            df["得分差距"] = df["综合得分"] - df["阈值"]
            df = df.sort_values(["得分差距", "综合得分"], ascending=[False, False])
            items = [
                f"{short_name(r['品类名称'], 7)}｜{r['层级']}｜{DEPT_MAP.get(dept, dept)}｜{format_score_gap(r['得分差距'])}｜得分{r['综合得分']:.0f}｜卖力{r['卖力指数']:.1f}"
                for _, r in df.head(3).iterrows()
            ]
        else:
            items = []
        while len(items) < 3:
            items.append("-")
        rows.append([DEPT_MAP.get(dept, dept), *items[:3]])
    return rows


def build_disadvantage_gap_cross_level_table(disadv_df, cutoffs, store_id):
    rows = [["部门", "Top1", "Top2", "Top3"]]
    for dept in ["生鲜部", "食品部", "非食品部"]:
        df = disadv_df[disadv_df["部门"] == dept].copy()
        if len(df):
            df["阈值"] = df.apply(lambda r: get_threshold(cutoffs, r["层级"], dept, store_id=store_id), axis=1)
            df = df[df["阈值"].notna() & df["门店综合得分"].notna()].copy()
            df["得分差距"] = df["阈值"] - df["门店综合得分"]
            df = df.sort_values(["得分差距", "企业综合得分"], ascending=[False, False])
            items = [
                f"{short_name(r['品类名称'], 7)}｜{r['层级']}｜{DEPT_MAP.get(dept, dept)}｜{format_disadvantage_gap(r['得分差距'])}｜店得分{r['门店综合得分']:.0f}｜卖力{r['门店卖力指数']:.1f}"
                for _, r in df.head(3).iterrows()
            ]
        else:
            items = []
        while len(items) < 3:
            items.append("-")
        rows.append([DEPT_MAP.get(dept, dept), *items[:3]])
    return rows


def add_store_matrix_slide(prs, title_layout, title, subtitle, threshold_text, table_data, header_fill, note):
    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (255, 255, 255))
    add_section_title(slide, title, subtitle)
    add_card(slide, Inches(0.72), Inches(1.18), Inches(11.58), Inches(0.64), fill=(249, 251, 255), line=THEME["line"])
    add_textbox(slide, Inches(0.95), Inches(1.31), Inches(11.0), Inches(0.28), "目标性品类得分阈值（按部门：企业/门店）：" + threshold_text, font_size=9, color=THEME["muted"])
    add_table(
        slide,
        Inches(0.45),
        Inches(2.02),
        Inches(12.45),
        Inches(4.02),
        table_data,
        font_size=8,
        header_fill=header_fill,
        col_widths=[Inches(0.7), Inches(3.9), Inches(3.9), Inches(3.9)],
        compact=True,
    )
    add_card(slide, Inches(0.72), Inches(6.54), Inches(11.58), Inches(0.28), fill=(249, 251, 255), line=THEME["line"])
    add_textbox(slide, Inches(0.95), Inches(6.61), Inches(11.0), Inches(0.12), note, font_size=7, color=THEME["muted"])
    return slide


def add_section_title(slide, title, subtitle=None):
    add_textbox(slide, Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.55), title, font_size=26, bold=True, color=THEME["navy"])
    if subtitle:
        add_textbox(slide, Inches(0.62), Inches(0.83), Inches(12.0), Inches(0.35), subtitle, font_size=10, color=THEME["muted"])


def build_ppt(summary, top, detail, disadvantage, metrics, profiles, cutoffs):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    title_layout = prs.slide_layouts[6]

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, THEME["light"])
    accent = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.7), Inches(2.2), Inches(0.4))
    accent.fill.solid()
    accent.fill.fore_color.rgb = RGBColor(*THEME["blue"])
    accent.line.fill.background()
    add_textbox(slide, Inches(0.85), Inches(0.74), Inches(2.6), Inches(0.28), "门店双向机会分析", font_size=15, bold=True, color=(255, 255, 255))
    add_textbox(slide, Inches(0.8), Inches(1.45), Inches(7.3), Inches(1.15), "各门店目标性品类双向分析\n报告", font_size=27, bold=True, color=THEME["navy"])
    add_textbox(slide, Inches(0.82), Inches(2.95), Inches(6.2), Inches(1.0), "基于门店目标性品类与企业目标性品类对比\n识别门店在3级、4级、5级的本地化经营机会", font_size=18, color=(66, 77, 99))
    add_textbox(slide, Inches(0.82), Inches(6.45), Inches(5), Inches(0.3), "生成日期：2026-05-26", font_size=11, color=(100, 110, 130))
    hero = add_card(slide, Inches(8.1), Inches(1.2), Inches(4.45), Inches(4.8), fill=(255, 255, 255), line=THEME["line"])
    add_textbox(slide, Inches(8.45), Inches(1.65), Inches(3.8), Inches(0.35), "关键结论", font_size=14, bold=True, color=THEME["blue"])
    add_bullets(slide, Inches(8.35), Inches(2.05), Inches(3.9), Inches(3.2), [
        "报告同时区分优势品类和劣势品类两个方向。",
        "优势品类主要集中在5级小类，生鲜、熟食和水果是主要来源。",
        "劣势品类则用于识别企业重点但门店未重点经营的小类。"
    ], font_size=16, color=THEME["ink"])

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (255, 255, 255))
    add_section_title(slide, "分析口径与报告结构", "对比对象：门店目标性品类 vs 企业目标性品类；按优势/劣势两个方向拆解")
    bullets = [
        "分析范围：10家门店，覆盖3级、4级、5级品类。",
        "优势品类：门店为目标性，但企业同层级不是目标性。",
        f"劣势品类：企业为目标性、且目标性覆盖门店数≥{DISADV_MIN_TARGET_COVERAGE}，但门店同层级不是目标性。",
        "优势品类排序：综合得分为主，叠加卖力指数与销售额，形成“代表性指数”。",
        "报告结构：总体洞察、门店分型、各门店单页、经营建议。"
    ]
    add_card(slide, Inches(0.7), Inches(1.3), Inches(5.7), Inches(4.4), fill=(249, 251, 255), line=THEME["line"])
    add_bullets(slide, Inches(0.88), Inches(1.55), Inches(5.25), Inches(3.8), bullets, font_size=16, color=THEME["ink"])
    ratio_df = metrics["level_ratio_df"]
    data = [["层级", "企业目标性\n(3个月原始)", "门店目标性\n平均数", "优势占门店目标性比例", "劣势占企业共性目标性比例"]]
    for _, row in ratio_df.iterrows():
        data.append([
            row["层级"],
            int(row["企业原始目标性品类数"]),
            f"{int(row['门店目标性品类数'])} / 10 = {row['门店目标性品类数'] / 10:.1f}",
            f"{row['优势品类数']} / {int(row['门店目标性品类数'])} = {row['优势占门店目标性品类比例']:.1%}",
            f"{row['劣势品类数']} / ({int(row['企业目标性品类数'])}×10) = {row['劣势占企业目标性品类比例']:.1%}",
        ])
    add_table(
        slide,
        Inches(6.35),
        Inches(1.45),
        Inches(6.0),
        Inches(3.1),
        data,
        font_size=9,
        col_widths=[Inches(0.7), Inches(1.2), Inches(1.2), Inches(1.45), Inches(1.45)],
    )
    add_card(slide, Inches(7.15), Inches(4.3), Inches(5.0), Inches(1.35), fill=(245, 248, 252), line=THEME["line"])
    add_textbox(slide, Inches(6.95), Inches(4.5), Inches(5.2), Inches(1.05), "核心判断：企业目标性数量按3个月原始口径展示；劣势占比则按过滤后的企业共性目标性口径计算，避免少数门店集中型品类误伤其他门店。", font_size=13, color=(35, 45, 68))

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (248, 250, 255))
    add_section_title(slide, "总体结论", "先看优势品类，再看劣势品类，避免把双向问题混成一个口径")
    bullets = [
        f"优势品类主要集中在5级：共{metrics['counts_5']}个，说明门店超配机会更多发生在精细化小类层面。",
        f"优势品类数量最多的门店是{store_label(metrics['top_count_store']['门店编码'])}，5级有{int(metrics['top_count_store']['差异目标性品类数'])}个。",
        f"5级平均得分最高的门店是{store_label(metrics['top_score_store']['门店编码'])}，平均综合得分为{metrics['top_score_store']['平均综合得分']:.2f}。",
        f"5级劣势品类总量为{int(metrics['disadvantage_summary'][metrics['disadvantage_summary']['层级'] == '5级']['劣势品类数'].sum()) if len(metrics['disadvantage_summary']) else 0}个；已剔除覆盖门店数过低和作废品类。",
        "高代表性优势机会多与熟食现制、水果细分、禽肉水产、乳品饮品经营有关。"
    ]
    add_card(slide, Inches(0.72), Inches(1.25), Inches(6.2), Inches(4.95), fill=(255, 255, 255), line=THEME["line"])
    add_bullets(slide, Inches(0.92), Inches(1.5), Inches(5.75), Inches(4.4), bullets, font_size=16, color=THEME["ink"])
    dept_total = sum(metrics["dept_counts"].values())
    dept_table = [["部门", "独有目标性品类数", "占比"]]
    for dept, count in metrics["dept_counts"].most_common():
        dept_table.append([dept, count, f"{count / dept_total:.1%}"])
    add_table(slide, Inches(7.05), Inches(1.55), Inches(5.2), Inches(2.3), dept_table, font_size=13, header_fill=(0, 163, 137))
    add_card(slide, Inches(7.0), Inches(4.15), Inches(5.2), Inches(1.55), fill=(240, 249, 244), line=(210, 233, 221))
    add_textbox(slide, Inches(7.22), Inches(4.45), Inches(4.8), Inches(0.9), "管理启示：企业级清单做经营底盘，门店独有目标性品类更适合做区域化动销与引流突破。", font_size=15, color=(32, 58, 46))

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (255, 255, 255))
    add_section_title(slide, "5级优势/劣势品类分布", "同一张图同时看门店超配机会与企业重点缺失")
    pivot_counts = summary.pivot(index="门店编码", columns="层级", values="差异目标性品类数").fillna(0)[LEVEL_LABELS]
    disadv5 = metrics["disadvantage_summary"]
    disadv5 = disadv5[disadv5["层级"] == "5级"].set_index("门店编码")["劣势品类数"] if len(disadv5) else pd.Series(dtype=float)
    add_column_chart(
        slide,
        Inches(0.75),
        Inches(1.35),
        Inches(11.6),
        Inches(4.8),
        [store_label(v) for v in pivot_counts.index.tolist()],
        {
            "5级优势品类": pivot_counts["5级"].tolist(),
            "5级劣势品类": [int(disadv5.get(v, 0)) for v in pivot_counts.index.tolist()],
        },
    )
    add_textbox(slide, Inches(0.85), Inches(6.45), Inches(11.5), Inches(0.32), "同一家门店可能同时存在优势品类和劣势品类，后续动作要区分“继续加码”与“补齐短板”两类策略。", font_size=14, color=(95, 108, 126))

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (255, 255, 255))
    add_section_title(slide, "独有目标性品类质量与部门结构", "平均综合得分反映门店独有目标性品类的质量；部门结构反映优势来源")
    pivot_score = summary.pivot(index="门店编码", columns="层级", values="平均综合得分").fillna(0)[LEVEL_LABELS]
    add_line_chart(
        slide,
        Inches(0.55),
        Inches(1.35),
        Inches(6.0),
        Inches(4.7),
        [store_label(v) for v in pivot_score.index.tolist()],
        {level: [round(v, 1) for v in pivot_score[level].tolist()] for level in LEVEL_LABELS},
    )
    dept_rows = []
    for level, frame in detail.items():
        dept_counts = frame["部门"].value_counts()
        for dept, count in dept_counts.items():
            dept_rows.append({"层级": level, "部门": DEPT_MAP.get(dept, dept), "数量": count})
    dept_df = pd.DataFrame(dept_rows)
    pivot_dept = dept_df.pivot(index="层级", columns="部门", values="数量").fillna(0)
    add_column_chart(
        slide,
        Inches(6.9),
        Inches(1.35),
        Inches(5.7),
        Inches(4.7),
        pivot_dept.index.tolist(),
        {dept: pivot_dept[dept].tolist() for dept in pivot_dept.columns.tolist()},
        chart_type=XL_CHART_TYPE.COLUMN_STACKED,
    )
    add_textbox(slide, Inches(0.8), Inches(6.45), Inches(11.4), Inches(0.3), "5级平均分最高的门店集中在崂山超市、凯德店、览秀城店，说明这些门店的小类优势机会既多又强。", font_size=14, color=(95, 108, 126))

    profile_chunks = [profiles[:5], profiles[5:]]
    for idx, chunk in enumerate(profile_chunks, start=1):
        slide = prs.slides.add_slide(title_layout)
        add_full_bg(slide, (247, 249, 253))
        add_section_title(slide, f"门店优势画像总览（{idx}/2）", "根据各店代表性独有目标性品类的经营方向、主要强化品类与独有目标性品类主导部门级比例生成门店画像")
        rows = [["门店", "优势方向总结", "主要强化品类", "独有目标性品类主导部门级比例", "优势峰值层级", "峰值数量"]]
        for item in chunk:
            rows.append([
                store_label(item["store"]),
                item["theme"],
                item["focus_items"],
                item["dominant_dept"],
                item["peak_level"],
                item["peak_count"],
            ])
        add_table(
            slide,
            Inches(0.45),
            Inches(1.45),
            Inches(12.45),
            Inches(4.95),
            rows,
            font_size=10,
            col_widths=[Inches(1.0), Inches(2.6), Inches(2.9), Inches(3.0), Inches(1.05), Inches(0.85)],
        )
        add_textbox(slide, Inches(0.75), Inches(6.45), Inches(11.5), Inches(0.28), "这些画像适合直接用于门店分层经营讨论，尤其适合生鲜、熟食和节令品类策略。", font_size=14, color=(95, 108, 126))

    for profile in profiles:
        store = profile["store"]
        store_name = store_label(store)
        store_summary = summary[summary["门店编码"] == store].copy().sort_values("层级")
        summary_map = {row["层级"]: row for _, row in store_summary.iterrows()}
        threshold_text = format_threshold_summary(cutoffs, str(store))
        store_detail = pd.concat(
            [frame[frame["门店编码"].astype(str) == str(store)].copy() for frame in detail.values()],
            ignore_index=True,
        )
        store_disadv = pd.concat(
            [frame[frame["门店编码"].astype(str) == str(store)].copy() for frame in disadvantage.values()],
            ignore_index=True,
        )
        add_store_matrix_slide(
            prs,
            title_layout,
            f"{store_name} 优势比例Top3",
            f"优势主题：{profile['theme']} | 按部门与3/4/5级分别排序",
            threshold_text,
            build_advantage_metric_matrix(store_detail, cutoffs, str(store), "ratio"),
            THEME["blue"],
            "说明：优势比例=门店得分/门店阈值，按部门和层级分别取最大的Top3；每条保留部门、得分、阈值比例和卖力。",
        )
        add_store_matrix_slide(
            prs,
            title_layout,
            f"{store_name} 优势得分差距Top3",
            f"优势主题：{profile['theme']} | 按部门跨3/4/5级排序",
            threshold_text,
            build_advantage_gap_cross_level_table(store_detail, cutoffs, str(store)),
            THEME["teal"],
            "说明：优势得分差距=门店得分-门店阈值；本页按部门跨3/4/5级取差距最大的Top3，避免和优势比例页重复。",
        )
        add_store_matrix_slide(
            prs,
            title_layout,
            f"{store_name} 劣势指数Top3",
            f"企业目标性但该店不是目标性 | 5级劣势数：{profile['disadv_count_5']}",
            threshold_text,
            build_disadvantage_metric_matrix(store_disadv, cutoffs, str(store), "ratio"),
            THEME["gold"],
            "说明：劣势指数=门店得分/门店阈值，越小越弱；按部门和层级分别取最小的Top3。",
        )
        add_store_matrix_slide(
            prs,
            title_layout,
            f"{store_name} 劣势得分差距Top3",
            f"企业目标性但该店不是目标性 | 按部门跨3/4/5级排序",
            threshold_text,
            build_disadvantage_gap_cross_level_table(store_disadv, cutoffs, str(store)),
            THEME["gold"],
            "说明：劣势得分差距=门店阈值-门店得分；本页按部门跨3/4/5级取差距最大的Top3，避免和劣势指数页重复。",
        )

    protein_slide = prs.slides.add_slide(title_layout)
    add_full_bg(protein_slide, (255, 255, 255))
    add_section_title(
        protein_slide,
        "鲜蛋白相关Top5小类",
        "全门店汇总口径：在所有5级独有目标性品类中，筛选禽肉、水产、蛋类等鲜蛋白相关小类后汇总排序",
    )
    add_card(protein_slide, Inches(0.72), Inches(1.28), Inches(11.58), Inches(1.0), fill=(245, 248, 252), line=THEME["line"])
    add_textbox(
        protein_slide,
        Inches(0.95),
        Inches(1.58),
        Inches(11.0),
        Inches(0.35),
        "这里的“鲜蛋白”用更直白的话说，就是禽肉、水产、蛋类等新鲜蛋白商品。下表展示所有门店汇总后，最值得关注的Top5独有目标性小类。",
        font_size=14,
        color=THEME["ink"],
    )
    protein_table = [["5级小类", "部门", "覆盖门店数", "平均综合得分", "平均卖力指数", "销售额汇总"]]
    for _, row in metrics["protein_top5"].iterrows():
        protein_table.append([
            row["品类名称"],
            row["部门"],
            int(row["覆盖门店数"]),
            f"{row['平均综合得分']:.0f}",
            f"{row['平均卖力指数']:.1f}",
            f"{row['销售额汇总'] / 10000:.1f}万",
        ])
    add_table(protein_slide, Inches(0.72), Inches(2.55), Inches(11.58), Inches(2.8), protein_table, font_size=13, header_fill=THEME["blue"])
    add_card(protein_slide, Inches(0.72), Inches(5.7), Inches(11.58), Inches(0.75), fill=(249, 251, 255), line=THEME["line"])
    add_textbox(
        protein_slide,
        Inches(0.96),
        Inches(5.94),
        Inches(11.0),
        Inches(0.22),
        "业务解读：这些小类在部分门店被提升为目标性品类，但企业整体未统一列为目标性，更适合作为区域化或门店化优势机会来观察。",
        font_size=12,
        color=THEME["muted"],
    )

    disadv_rows = [["门店", "5级劣势品类数", "优先关注Top1", "优先关注Top2", "优先关注Top3"]]
    disadv5 = disadvantage["5级"].copy()
    store_order = (
        metrics["disadvantage_summary"][metrics["disadvantage_summary"]["层级"] == "5级"]
        .sort_values(["劣势品类数", "门店编码"], ascending=[False, True])["门店编码"]
        .astype(str)
        .tolist()
    )
    for store in store_order:
        store_df = disadv5[disadv5["门店编码"].astype(str) == str(store)].copy()
        store_df = store_df.sort_values(["企业综合得分", "企业销售额"], ascending=[False, False])
        names = store_df["品类名称"].astype(str).drop_duplicates().tolist()[:3]
        while len(names) < 3:
            names.append("-")
        disadv_rows.append([
            store_label(store),
            int(len(store_df)),
            names[0],
            names[1],
            names[2],
        ])
    disadvantage_chunks = [disadv_rows[:6], [disadv_rows[0], *disadv_rows[6:]]]
    for idx, chunk in enumerate(disadvantage_chunks, start=1):
        disadvantage_slide = prs.slides.add_slide(title_layout)
        add_full_bg(disadvantage_slide, (255, 255, 255))
        add_section_title(
            disadvantage_slide,
            f"各门店劣势品类Top清单（{idx}/2）",
            f"口径：仅保留企业目标性覆盖门店数≥{DISADV_MIN_TARGET_COVERAGE}、且非作废品类；按5级劣势品类数量和企业综合得分筛选每店最值得关注的短板",
        )
        add_table(
            disadvantage_slide,
            Inches(0.55),
            Inches(1.45),
            Inches(12.2),
            Inches(4.9),
            chunk,
            font_size=11,
            header_fill=THEME["gold"],
            col_widths=[Inches(1.2), Inches(1.25), Inches(3.25), Inches(3.25), Inches(3.25)],
        )
        add_card(disadvantage_slide, Inches(0.7), Inches(6.42), Inches(11.7), Inches(0.45), fill=(249, 251, 255), line=THEME["line"])
        add_textbox(
            disadvantage_slide,
            Inches(0.92),
            Inches(6.54),
            Inches(11.2),
            Inches(0.18),
            "使用建议：这页适合直接给商品、营运和采购团队做补短板复盘；已排除少数门店集中拉高的企业目标性品类，结果更接近真实短板。",
            font_size=11,
            color=THEME["muted"],
        )

    slide = prs.slides.add_slide(title_layout)
    add_full_bg(slide, (245, 248, 252))
    add_section_title(slide, "经营建议", "从企业底盘到门店优势机会，建议形成双层经营机制")
    bullets = [
        "企业层：保留统一目标性品类清单，作为基础经营盘和统一资源投放依据。",
        "门店层：将门店独有目标性品类纳入季度复盘，形成“门店优势机会池”。",
        "生鲜与熟食：对崂山超市、海阳店、凯德店等差异强店重点做熟食加工、水果强化、鲜品蛋白的专项复盘。",
        "商品策略：对只在1-2家门店成立的5级独有目标性品类，优先作为区域化、节令化或客群化策略验证样本。",
        "组织动作：建议建立“企业统一 + 门店补充”的目标性品类双清单管理机制。 "
    ]
    add_card(slide, Inches(0.75), Inches(1.3), Inches(11.8), Inches(4.95), fill=(255, 255, 255), line=THEME["line"])
    add_bullets(slide, Inches(0.95), Inches(1.55), Inches(11.1), Inches(4.35), bullets, font_size=17, color=THEME["ink"])
    add_textbox(slide, Inches(0.9), Inches(6.4), Inches(10.8), Inches(0.3), "附：本报告为可编辑PPT，可继续追加门店照片、销售趋势和经营动作复盘页。", font_size=13, color=(92, 104, 121))

    prs.save(OUTPUT_PPTX)


def main():
    summary, top, detail, disadvantage, disadvantage_meta = load_data()
    metrics = compute_metrics(summary, detail, disadvantage, disadvantage_meta)
    profiles = build_store_profiles(summary, top, detail, disadvantage)
    cutoffs = compute_score_cutoffs()
    build_ppt(summary, top, detail, disadvantage, metrics, profiles, cutoffs)
    print(OUTPUT_PPTX)


if __name__ == "__main__":
    main()
