# skills

个人维护的 agent skills 集合，Claude Code / Codex / WorkBuddy 共用。

## 技能

| 技能 | 用途 | 装载 agent |
|---|---|---|
| `frontend-standards` | 前端编码规范：SCSS 嵌套（`&` 变体、≤3 层）、箭头函数、改代码同步文档 | claude, codex, workbuddy |
| `timeslot-heatmap` | 门店×时段销售热力图（丽达口径，含门店类型图例） | claude, codex, workbuddy |
| `category-role-calculation` | 品类角色计算：3/4/5 级、月度/四季维度，兼容旧版算法 | claude, codex, workbuddy |
| `store-category-role-report` | 企业-门店品类角色差异分析：明细工作簿 + 可编辑 PPT | codex |

## 结构

```
skills/
├── <skill-name>/
│   ├── SKILL.md          # frontmatter: name / description / version
│   ├── scripts/          # 可选：可执行脚本
│   ├── references/       # 可选：模板、数据字典
│   └── requirements.txt  # 可选：Python 依赖
└── README.md
```

每个技能一个顶层目录，`SKILL.md` 直接放在目录下 —— 不要再套 `skills/` 或隐藏目录，magpie 解析源路径时会多一层。

## 维护约定

1. **改内容必须升 `version`**（patch/minor/major 按语义），方便和 magpie 里装的副本对账。
2. **产物不进仓库**：`outputs/`、`__pycache__`、`.DS_Store` 已在 `.gitignore`。
3. **改完要同步到 magpie**。`magpie library sync` 只检查 agent 有没有装，**不会**把 folder 源的内容变更推出去，必须手动覆盖：

   ```bash
   for d in ~/.config/magpie/library/skills/<name> ~/.claude/skills/<name>; do
     [ -d "$d" ] && cp -R "<name>/." "$d/"
   done
   ```

   （codex / workbuddy 是指向 library 副本的软链，覆盖 library 即可。）
4. **提交信息**用 `feat(<skill>): …` / `fix(<skill>): …`，一个技能一次改动一条。

## 来源

这些技能为自维护，不依赖上游仓库。若将来迁到 GitHub 源（`zoexml/skills`），每个技能按 `repo + path` 单独安装，magpie 会记录各自的 `commit` 用于更新检查。
