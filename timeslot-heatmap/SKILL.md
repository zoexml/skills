---
name: timeslot-heatmap
description: 生成门店「时段销售」热力图 —— 从门店×时段销售报表算出各时段交易占比或日均交易笔数，渲染成热力图 PNG，并可叠加门店类型图标/图例与门店三件套指标（SKU 数、销售额、客单价）。当用户提到时段销售、时段热力图、分时经营、门店时段交易占比、日均交易笔数热力图、高峰时段或排班分析，或要复现、更新丽达那套「各门店时段交易占比热力图」（含门店类型图例版、75 天口径、3 笔/日以下置 0）时，务必使用本 skill —— 即使他们没提「热力图」三个字，只说「那个分时的图再出一版」也要用。
---

# 门店时段销售热力图

把一张**门店 × 时段**的销售报表渲染成分时热力图，回答「哪家店、哪个时段是高峰、哪个是尾流」。

## 先确定要哪一种口径

同一个 75 天窗口下有两种图，**先按用户已有的产物或原话判断要哪张**，不确定就问：

| 口径 | 单元格含义 | 行序 | 脚本 |
|---|---|---|---|
| **占比版** | 该时段交易笔数 ÷ 该店全天交易笔数 | 门店日均交易量降序 | `generate_time_heatmap.py` |
| **日均笔数版** | 该时段交易笔数 ÷ 75 天，**低于 3 笔/日照 0 展示** | 门店销售收入降序 | `generate_time_daily_heatmap.py` |

两版都跳过交易笔数为 0 的门店（丽达这一期是万象汇店，75 天全零），所以图上只有 10 行而不是 11 行。
**这不是 bug，不要"修"成 11 行。**

日均笔数版是外部报告第 8 页的本地复刻（口径已逐格校验 150/150），要跟那份 pptx 对齐就用它；
占比版对应丽达 2026-05-28 那批 PNG。

## 运行

脚本在 `<root>/scripts/时段销售热力图/`，从那里跑：

```bash
python generate_time_heatmap.py            # 占比版底图
python generate_time_daily_heatmap.py      # 日均笔数版底图

# 两个后处理：都不改底图任何像素（图例是叠画，指标是在右侧拼一块面板），可单独用也可串联
python apply_store_type_legend.py ratio    # 叠门店类型图标 + 图例
python apply_store_type_legend.py daily
python apply_store_metrics.py ratio        # 右侧拼门店三件套面板（SKU/销售额/客单）
python apply_store_metrics.py daily

# 串联出「图标 + 图例 + 三件套」完整形态（= 2026-05-28 那版的复现）
python apply_store_type_legend.py ratio
python apply_store_metrics.py ratio \
  --input ../../outputs/时段销售_时段占比热力图_75天口径_门店类型图例版.png
```

产物一律落在 `<root>/outputs/`。各脚本 `--help` 看参数。

## 换新一期数据

窗口是**硬编码常量**，不是从报表里推出来的，所以换期要显式改：

1. 新报表放进 `data/丽达时段销售/`，改两个生成脚本里的 `SOURCE_XLSX`
2. 改 `START_DATE` / `END_DATE` —— `DATE_DAYS = (END - START).days + 1` 就是分母，
   「75 天口径」的 75 就是这么来的；改了它，日均笔数整列跟着变
3. 有新一期的三件套表就改 `apply_store_metrics.py` 的 `SOURCE_XLSX`
4. 门店有增减，同步 `apply_store_type_legend.py` 里的 `STORE_TYPES`

## 目录与路径约定（别动层级）

```
<root>/
├── scripts/时段销售热力图/*.py    ← 必须保持两级嵌套，见下
├── data/丽达时段销售/*.xlsx
├── outputs/
└── references/
```

脚本用 `Path(__file__).resolve().parents[2]` 定位根目录，所以 `scripts/` 下面**必须再套一层**。

要把数据放到别处，设环境变量，别挪脚本 —— 注意 `HEATMAP_ROOT` 指的是**一个镜像了包内结构的根目录**，
不是「你的 xlsx 放在哪」：

```bash
# ✅ 目录下要有 data/丽达时段销售/ 和 outputs/
HEATMAP_ROOT=/data/heatmap-root python generate_time_heatmap.py
```

指错了会得到一条明确的提示（而不是裸的 FileNotFoundError），告诉你要放哪个路径、或者该指向哪一层。

## 数据源：两张表，用途不同

| 表 | 粒度 | 提供什么 |
|---|---|---|
| `报表_超市分析_实时销售分析_时段销售_去海天仙侠岭.xlsx` | 门店 × 时段（165 行） | 热力图**单元格**：交易笔数、销售收入、毛利、会员销售… |
| `丽达客单价等数据_0301-0514.xlsx` | 门店（10 行） | 右侧三件套：SKU 数、销售额、客单价 |

时段表实际有 **177 行 = 165 有效 + 11 行「小计」+ 1 行「合计」**，用之前两个都要剔：
按 `门店.notna()` 排掉合计行、按 `时段 != "小计"` 排掉小计行。漏剔会让所有合计**翻倍**
（曾经把 268.5 万笔报成 537 万笔）。

改口径、换数据、或想知道某个数字的来历，读 `references/口径说明.md`。

## 已经踩过的坑

- **客单价不要自己算**。整列引用三件套表即可 —— `销售额 ÷ 交易次数` 对不上（85.6 vs 70），
  因为它用的是上游另一个未暴露的量（大概是去重小单数）。
- **销售额是向下取整到万位**，不是四舍五入：30,718,497 → 3071 万。
- **`STORE_TYPES` 按店名硬编码**。新增门店必须补条目，否则那行**没有图标且不报错**（静默跳过）。
- 改布局守住 `LEFT_MARGIN + CELL_WIDTH×15 + RIGHT_MARGIN ≤ CANVAS_WIDTH`，
  超了右侧时段列会被**静默裁掉**（发生过一次）。
- 图标栏只占最左 `x<68`。店名画在 `x73~171`，行底色刷进这个区间会把店名整条盖掉。

## 依赖

`pandas` / `openpyxl`（读 xlsx）/ `pillow`（画图），见 `requirements.txt`。

中文字体按平台自动探测（macOS 冬青黑体 → Windows 微软雅黑 → Linux Noto），
找不到会直接报错并列出候选路径。把本机字体绝对路径补进
`generate_time_heatmap.py` 的 `FONT_CANDIDATES` 即可 —— 四个脚本共用那一张表。
