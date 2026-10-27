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
3. **改完要 push 再让 magpie 重新拉**。技能在 magpie 里以 GitHub 源安装（`zoexml/skills` + path），本地改动不会自动同步：

   ```bash
   git add -A && git commit -m "feat(<name>): …" && git push

   # 安装/更新只能走 web 模式（没有 CLI 子命令）
   HTTPS_PROXY=socks5://127.0.0.1:7890 HTTP_PROXY=socks5://127.0.0.1:7890 \
     MAGPIE_WEB_KEY=<key> /Applications/magpie.app/Contents/MacOS/magpie web \
     --no-open --addr 127.0.0.1:3499 &
   magpie library skill update <name> --replace-edits
   ```

   下载走 `codeload.github.com`，该域名在本机被代理 fake-ip 劫持，不挂 socks5 会报 EOF；GUI 里直接点更新同样会失败。
4. **提交信息**用 `feat(<skill>): …` / `fix(<skill>): …`，一个技能一次改动一条。

## 来源

全部自维护，以 GitHub 源安装：`zoexml/skills::<name>`，magpie 为每个技能记录内容 `hash` 用于更新检查。技能之间互不依赖，可单独安装到任意 agent。
