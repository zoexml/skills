from pathlib import Path

import pandas as pd


BASE_DIR = Path("/Users/joet/Downloads/大数据畅销品推荐")
STORE_ROLE_XLSX = Path("/Users/joet/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/E28C68EA2571E42BDD1407AA361583ED/Caches/Files/2026-05/a808316c5629a684a30b8d3d697057f8/门店下的品类角色-企业15%门店12%.xlsx")
ENTERPRISE_XLSX = Path("/Users/joet/Library/Containers/com.tencent.WeWorkMac/Data/Documents/Profiles/E28C68EA2571E42BDD1407AA361583ED/Caches/Files/2026-05/170d005971ff83788c0db8724513e187/企业下的品类角色-企业15%门店12%.xlsx")
OUTPUT_XLSX = BASE_DIR / "各门店代表性差异品类_企业15门店12.xlsx"

LEVEL_SHEETS = {"3级": "3级品类", "4级": "4级品类", "5级": "5级品类"}


def clean_store_sheet(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df = df[df.iloc[:, 0] != df.columns[0]].copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    num_cols = ["c_pi", "c_fi", "c_sale", "c_skus", "c_sale_index", "c_sale_rate", "c_growth_store", "c_score", "c_score_avg"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def clean_enterprise_sheet(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df = df[df.iloc[:, 0] != df.columns[0]].copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    num_cols = ["c_pi", "c_fi", "c_sale", "c_skus", "c_sale_index", "c_sale_rate", "c_growth_store", "c_score", "c_score_avg"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def main():
    writer = pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl")
    summary_rows = []
    top_rows = []

    for level, sheet in LEVEL_SHEETS.items():
        store_df = clean_store_sheet(STORE_ROLE_XLSX, sheet)
        ent_df = clean_enterprise_sheet(ENTERPRISE_XLSX, sheet)

        ent_target_codes = set(ent_df.loc[ent_df["c_class_role_id"] == "目标性品类", "c_ccode"].astype(str))
        store_target = store_df[store_df["c_class_role_id"] == "目标性品类"].copy()
        diff = store_target[~store_target["c_ccode"].astype(str).isin(ent_target_codes)].copy()

        diff["层级"] = level
        diff["代表性指数"] = (
            diff["c_score"].rank(pct=True, method="average") * 0.6
            + diff["c_sale_index"].rank(pct=True, method="average") * 0.25
            + diff["c_sale"].rank(pct=True, method="average") * 0.15
        )
        diff["代表性指数"] = (diff["代表性指数"] * 100).round(2)

        out = diff[
            [
                "c_store_id",
                "层级",
                "c_ccode",
                "c_cname",
                "c_dept_name",
                "c_score",
                "c_sale_index",
                "c_sale",
                "c_sale_rate",
                "c_growth_store",
                "c_pi",
                "c_fi",
                "代表性指数",
            ]
        ].copy()
        out.columns = [
            "门店编码",
            "层级",
            "品类编码",
            "品类名称",
            "部门",
            "综合得分",
            "卖力指数",
            "销售额",
            "销售占比",
            "超市增长率",
            "千人购买率",
            "千次购买率",
            "代表性指数",
        ]
        out = out.sort_values(["门店编码", "代表性指数", "综合得分"], ascending=[True, False, False])
        out.to_excel(writer, sheet_name=sheet, index=False)

        for store_id, group in out.groupby("门店编码"):
            summary_rows.append(
                {
                    "门店编码": str(store_id),
                    "层级": level,
                    "差异目标性品类数": int(len(group)),
                    "平均综合得分": round(group["综合得分"].mean(), 2),
                    "最高综合得分": round(group["综合得分"].max(), 2),
                    "平均卖力指数": round(group["卖力指数"].mean(), 2),
                }
            )
            for _, row in group.head(5).iterrows():
                top_rows.append(row.to_dict())

    summary_df = pd.DataFrame(summary_rows).sort_values(["门店编码", "层级"])
    top_df = pd.DataFrame(top_rows).sort_values(["门店编码", "层级", "代表性指数"], ascending=[True, True, False])
    summary_df.to_excel(writer, sheet_name="门店层级汇总", index=False)
    top_df.to_excel(writer, sheet_name="每店Top5代表差异品类", index=False)
    writer.close()
    print(OUTPUT_XLSX)


if __name__ == "__main__":
    main()
