---
name: finance-news-warmap
description: 金融/行业「新闻作战地图」构建 skill。把真实搜索新闻按经纬度落点，生成 Palantir 风格、可离线双击打开的单文件 HTML（ECharts 地理散点 + 中国/世界切换 + 可筛选可悬停 + 右侧数据驱动「建议区」）。配套 search→compile→fill 流水线，把每个细分领域自动补满 N 条真实新闻并去重，绝不编造数据。触发：做新闻情报地图、作战地图、创新动态地图、搜新闻自动补满、给高管出趋势与建议面板、要离线自包含 HTML 交付物。**调用前须由 agent 执行 WebSearch 联网能力探测（不可用则中止，不编造/不占位；此为 agent 侧流程约束，非代码强制断言）。所有库/字体/地图数据均随技能本地分发并在构建时内联，运行时无任何 CDN/外部请求。**
agent_created: true
version: 1.1.0
---

# 金融创新作战地图（Finance News War Map）

一个「**真实数据驱动 + 单一离线 HTML**」的新闻情报地图工作流。它把 WebSearch 搜来的真实国内新闻按城市经纬度落点，用 ECharts 地理散点呈现，并带一个**数据驱动的右侧「建议区」**（自动总结趋势 + 给银行/保险/金融科技高管建议）。

核心立场（来自用户强约束）：**只用真实搜索新闻，绝不编造；每个细分领域搜满目标条数；长字段写足；产出必须离线可双击打开。**

---

## 〇、启动前需向用户确认的问题（Question Flow）

运行本 skill 前，按顺序确认以下问题；标注「默认不询问」的按默认值直接执行，标注「必问」的必须拿到答案再跑。

| # | 问题 | 默认 | 处置 |
|---|---|---|---|
| 1 | 关注哪些领域 / 分类法？ | 银行(10)+金融科技(5)+保险(5)=20 项 | 默认不询问；用户要增删才改 |
| 2 | 每个子方向填满几条？ | 5 条 | 默认不询问 |
| 3 | 时间窗口？ | **最近一周（freshness=`d7`）** | **默认不询问**；除非用户要近 1 月/指定日期 |
| 4 | 地理范围？ | 中国国内为主 | 默认不询问 |
| 5 | 新闻来源 / 语言口径？ | 国内真实新闻，官媒+主流财经媒体 | 默认不询问 |
| 6 | **你是什么角色 / 视角？** | ——（无安全默认） | **必问**：如"浙江省分行行长""零售条线总监""合规官"。答案经 `build --role "角色名"` 固化进建议区，**不在浏览器里做输入框/切换器** |
| 7 | 视觉风格？ | Palantir 深色 + 三栏 | 默认不询问 |
| 8 | 长字段强制要求？ | event/why/innovation/learn 各 ≥100 字 | 默认不询问 |
| 9 | 输出位置 / 文件名？ | `金融创新全球作战地图.html` + `news.xlsx` | 默认不询问 |
| 10 | 要不要定时自动刷新？ | 手动 build | 默认不询问 |

> 角色（#6）是**唯一必须主动问**的一项；其余一律默认，除非用户另行指定。角色决定建议区"高管建议"那一段的视角，命中种子场景用种子文案，否则按关键词从 `DIRECTION_INSIGHTS` 动态生成并固化。

---

## 〇·五、环境前置检查（调用时强制第一步 · Gate）

本 skill 的整个数据流依赖**真实联网搜索**——没有可工作的 `WebSearch`，后续 search→compile→fill 全部失去意义，且会直接触碰"只用真实新闻、绝不编造"的硬底线。因此在做**任何**搜索 / 编译 / 构建之前，**必须先确认环境具备联网搜索能力**；确认通过才进入 Step 1，不通过则立即中止。

**检查步骤（agent 在调用 skill 时执行）：**
1. 先按 Question Flow 向用户问清必要项（其中 #6 角色必问；其余默认）。
2. **探测**：执行一次轻量 `WebSearch`（如查询 `金融科技 创新 新闻 2026` 或任何与任务相关的真实关键词，freshness 用 `d7`）。
3. **判定**：
   - ✅ 成功返回 ≥1 条真实结果（有标题 / 链接 / 摘要）→ 环境可用，继续 Step 1。
   - ❌ 工具不可用 / 调用被拒 / 沙箱网络被阻断 / 返回空或报错 / 只返回"无法访问"之类占位 → **立即中止**。
4. **中止处置**：明确告知用户——"本环境当前不支持联网搜索（WebSearch 不可用 / 网络受限）。按硬规则『只用真实新闻、绝不编造』，无法执行后续流程。"然后停止。**不要**尝试用训练记忆编造新闻，也**不要**改用离线占位数据蒙混。可建议用户：
   - 开启联网 / 搜索工具后重试；或
   - 改为仅基于用户已提供的本地新闻文件（`news.xlsx` 或传入的 JSON）做 `build`（此时需用户**明确认可**"只用本地既有数据、不再补新搜"，且仍遵守长字段≥100字 / 去重等规则）。

**注意**：`WebSearch` 是 agent 侧的服务器工具；脚本内的 `pipeline.py search` 只打印查询供人工 / agent 去搜，本身不联网。所以"是否支持联网搜索"只能在 agent 调用 skill 时由上述探测判定，不能靠脚本断言。

---

## 〇·六、数据安全与对外传输声明（透明性 · 回应供应链/外传审查）

本 skill 是**纯本地生成器**，不存在隐藏或自动化的对外网络调用。所有对外传输都需用户显式触发并自带凭据：

1. **联网搜索（WebSearch）由调用方 agent 执行，不是 skill 代码联网**：`build.py` / `pipeline.py` / `template.html` 在运行时**不发起任何 WebSearch 或互联网请求**——`pipeline.py search` 只把查询打印出来供人工/agent 去搜。
2. **离线单文件是强制硬指标（无 CDN / 无外部字体）**：`template.html` 不再引用任何外部 `https://` 的 `<script>` 或 `<link>`。ECharts 库、world/china GeoJSON、**xlsx 导出库**全部随技能分发于 `assets/` 并在构建时**内联**进 HTML；字体仅用系统字体栈回退（Microsoft YaHei / PingFang SC 等）。生成的 HTML 双击离线可用，不经过任何第三方服务器。
3. **远程同步（Supabase pull/push）完全可选 + 显式 + 用户凭据**：
   - 仅当用户主动运行 `build.py pull` / `push` / `build --source supabase` 时才会发生；
   - 需要用户在本地 `config.toml` 显式填入**自己的** Supabase url/key/table，缺文件即报错中止，绝不静默连接；
   - 默认 `build`（无 `--source`）只读写本地 `news.xlsx`，**完全不触碰任何远程**；
   - 不存在自动上传、环境变量 harvesting、文件系统枚举用于外传等行为。
4. **不编造 / 不做离线占位**：WebSearch 不可用且用户未提供本地数据时，按规则立即中止，绝不用记忆编造或离线占位数据蒙混（见「〇·五」）。

> 说明：本 skill 的「只用真实新闻、绝不编造」「调用即探测 WebSearch」等是**对调用 agent 的流程约束（指令）**，由 agent 在调用时落实，而非由 skill 代码在运行时做断言式强制。agent 应如实向用户说明当前环境是否具备联网搜索能力。

---

## 一、何时用这个 skill

- 用户想做 Palantir 风格的「新闻作战地图 / 情报地图 / 创新动态地图」：散点落点 + 可筛选 + 可悬停 + 中国/世界切换。
- 用户想做「**搜新闻 → 编译 JSON → 去重并入 → 自动补满每个细分领域 N 条**」的流水线（杜绝虚拟/占位数据）。
- 用户想要一个「**建议区 / 洞察面板**」，从实时新闻数据自动总结趋势与高管可执行建议（不是写死的模板话术）。
- 用户要「**双击可打开、可邮件分享**」的单一离线 HTML（库与地图数据全部内联，不依赖 CDN/网络）。

---

## 二、硬规则（默认开启，除非用户明确授权跳过）

1. **联网搜索能力前置 Gate（最先执行）**：调用本 skill 时**第一步**先探测当前环境是否支持真实联网搜索（`WebSearch` 可用且能返回真实结果）。**不支持则立即中止整个流程**，绝不退化为"用记忆编造新闻"或"离线占位数据蒙混"。详见下一节「环境前置检查」。
2. **只用真实新闻**：全部来自 WebSearch（默认 freshness=`d7` 最近一周），绝不编造、绝不虚构事件/链接。
3. **每细分领域搜满目标条数**（默认 TARGET=5）：不足时只报缺口，**不编造占位补齐**。
4. **时间窗口默认最近一周（freshness=`d7`）**：直接默认，**不询问用户**；如需其他窗口（近 1 月/指定日期）才单独确认。
5. **长字段写足**：每条新闻的 `event / why / innovation / learn` 四个字段各 ≥ 100 字。注意去重保留首次出现，所以**必须在 fill 前补全**，否则合并后无法覆盖。
6. **离线单文件**：产出 HTML 必须把 echarts 库与 world/china GeoJSON 全部内联，双击即可打开、可邮件分享。
7. **去重不重复**：已存在新闻按 norm_key 去重，合并时 existing 在前，绝不重复存储。

---

## 三、架构（单一离线 HTML + 占位符模板）

```
template.html  ──(占位符替换)──►  金融创新全球作战地图.html   ← 双击离线打开
build.py      读 news.xlsx（可编辑数据源）注入 __NEWS__ / __SUGGESTIONS__
pipeline.py   分类法真理源 + 搜索/状态/补满 流水线
assets/       echarts.min.js + world.json + china.json（34 省）← 构建时内联
```

`template.html` 用 6 个占位符，由 `build.py` 替换：

| 占位符 | 内容 |
|---|---|
| `__ECHARTS_LIB__` | echarts.min.js 全文 |
| `__WORLD_JSON__` | world.json（ECharts 世界地图 GeoJSON） |
| `__CHINA_JSON__` | china.json（34 省级边界 GeoJSON） |
| `__NEWS__` | 新闻数组 JSON（13 字段） |
| `__SAMPLE__` | `true`/`false`（是否示例水印） |
| `__SUGGESTIONS__` | 建议区结构化 JSON（build 时计算） |

`references/` 目录里是可直接复用的脚手架（**自包含，无需联网下载额外文件**）：

```
references/
├── template.html          # Palantir 风三栏地图模板（含占位符，离线单文件目标）
├── build.py               # 生成器：init/build/merge/replace + 建议区计算 + 去重
├── pipeline.py            # 20 子领域 taxonomy 驱动的 search/status/fill/rebuild 补满流水线
├── news_sample.json       # 内置示例数据源（11 条，init/build 缺 xlsx 时回退用）
├── db.py                  # Supabase 读写（可选，pull/push 用；凭据走本地 config.toml）
├── config.example.toml    # Supabase 配置样例
└── assets/               # 离线单文件必需，已随技能分发
    ├── echarts.min.js     # ~1MB，内联进 HTML
    ├── world.json         # ~1MB，世界地图边界
    └── china.json         # 中国 34 省边界
```

### 三·5、环境准备（依赖与目录约定）

1. **装依赖**：`pip install -r requirements.txt`（核心只要 `openpyxl`；`supabase` 仅 pull/push 时需要）。
2. **数据/产物落在「当前工作目录（CWD），不污染技能安装目录**——这是本 skill 的可分发约定：
   - 只读资产（`template.html` / `assets/` / `news_sample.json`）随技能分发，安装后不应改动；
   - `build.py` 把 `news.xlsx`（数据源）与 `金融创新全球作战地图.html`（产物）写到你**运行命令时所在的目录**；
   - 所以在你自己的项目目录里运行 `build.py` / `pipeline.py` 即可，不要在技能安装目录内直接写。
3. **快速开始**（在你自己的项目目录执行）：
   ```bash
   python build.py init          # 用内置示例生成 news.xlsx（可编辑数据源）
   python build.py build         # 读 news.xlsx -> 生成离线单文件 HTML 到当前目录
   # 或跳过 init，直接 build（缺 news.xlsx 时自动回退示例）
   python build.py build --role "浙江省分行行长"   # 指定建议区视角后固化进 HTML
   ```
4. **联网能力前置（Gate）**：本 skill 的「新闻」来自真实 WebSearch，**必须**先按「〇·五、环境前置检查」探测 WebSearch 可用；不可用时立即停止，绝不编造/占位。

---

## 四、数据 schema（13 列，news.xlsx 与 JSON 通用）

`time / category / direction / company / city / lng / lat / title / link / event / why / innovation / learn`

- JSON 用 `coord:[经度,纬度]`；xlsx 用 `lng/lat` 两列（`build.news_to_xlsx` / `xlsx_to_news` 自动互转）。
- `event/why/innovation/learn` 各 ≥ 100 字——这是用户的硬指标。
- `category` 取值示例：`银行 / 保险 / 金融科技`（对应配色 cyan / violet / amber）。

---

## 五、工作流（端到端）

0. **环境前置检查（Gate，必须最先执行）**：按「〇·五、环境前置检查」做 WebSearch 探测。不可用 → 立即中止并告知用户，**不进入后续任何步骤**。可用 → 继续。
1. **定义分类法 TAXONOMY**（唯一真理源）：`(大类, 方向, 搜索查询)` 三元组列表——图例、筛选、补满标准都以它为准。示例 20 项：银行(存款理财/公募私募/银保/信托/个人信用贷款/信用卡分期/汽融/个人房贷/个人房地产抵押贷款/小微企业主贷款) 10 项 + 金融科技(产品创新/模式创新/技术创新/需求走向/权益升级) 5 项 + 保险(寿险/财险/健康险/养老险/普惠保险) 5 项。
2. **搜索**：对每个子方向 `WebSearch`（默认 freshness `d7` 最近一周，不询问），收集真实国内新闻候选。
3. **编译 JSON**：每条含 13 字段；`coord=[经度,纬度]`；`link` 用真实 URL；四长字段写足 ≥100 字。
4. **去重并入**：`pipeline.py fill <json>`（或 `build.py merge <json>`），existing 在前 → 不重复 → 重写 news.xlsx + 重建 HTML。
5. **查缺口**：`pipeline.py status` 列出每类 当前 vs 目标，打印缺口方向。
6. **补满**：对缺口方向再搜、再 fill，循环直到每类达标。
7. **确定视角/角色（问用户，不写死）**：在 build 前问用户「你是什么角色/视角？」（如"浙江省分行行长""零售条线总监""合规官"）。该角色通过 `build.py build --role "角色名"` 固化进 HTML 的建议区——**不在浏览器里做角色输入框或切换器**。
8. **重建**：每次改完 news.xlsx 后 `build.py build [--role ...]` 重生成单文件 HTML。
9. **校验**：占位符无残留、JSON 合法、每子方向计数、分类值在 {银行,保险,金融科技} 内。

---

## 六、关键代码模式（直接复用）

### 去重主键（保证已存在不重复存储）
```python
import re
def norm_key(n):
    link = (n.get("link") or "").strip().lower()
    if link:
        return re.sub(r"[?#].*$", "", link)        # 链接去 ?# 参数后小写
    return re.sub(r"\s+", "", (n.get("title") or ""))  # 否则标题去空格兜底

def dedup(news, keep_first=True):
    seen, out = set(), []
    for n in news:
        k = norm_key(n)
        if not k or k in seen:
            continue
        seen.add(k); out.append(n)
    return out
```

### 数据驱动「建议区」（build 时实时算，非写死）
`build_suggestions(news, role=None)` 返回结构（`role` 由 `--role` 在 build 时传入，默认 `金融高管视角`）：
- `total / span[起,止] / catCounts`（大类分布条）
- `trends`：取出现频次 ≥2 的子方向 Top6，每条带 `trend`(趋势判断)+`advice`(高管动作)+`evidence`(2 条真实标题)
- `execAdvice`：**单一角色块**（不再是浏览器里的多受众切换）。命中 `ROLE_ADVICE` 种子（金融高管视角/银行高管/保险高管/金融科技负责人/分行行长等）用种子文案；否则按 `ROLE_KEYWORDS` 关键词匹配最相关业务方向，从 `DIRECTION_INSIGHTS` 动态生成 4 条建议并固化。**角色不在浏览器交互选择**——只在 build 时由 `--role` 决定，默认 `金融高管视角`。

趋势/建议文案放两张常量表，由数据**驱动选取**，避免模板话术：
- `DIRECTION_INSIGHTS[方向]` = `{trend, advice}`（每个子方向一条实战洞察）
- `ROLE_ADVICE[角色]` = `{color, icon, points:[...]}`（种子场景，可增删改；任意未命中角色由关键词动态生成，绝不写死受众集）

> 文案要「真判断、敢拆穿伪创新」，不要包装式空话。例如银保写「从趸交理财转向期交保障+养老，银行渠道从通道变客户经营伙伴」；AI 写「从 PoC 到生产，先建数据底座与 AI 中台，重定义可计算的非结构化资产」。

### 占位符替换 + 断言无残留
```python
out = (tpl
  .replace("__ECHARTS_LIB__", echarts).replace("__WORLD_JSON__", world.strip())
  .replace("__CHINA_JSON__", china.strip()).replace("__NEWS__", news_js)
  .replace("__SAMPLE__", "true" if is_sample else "false")
  .replace("__SUGGESTIONS__", sugg_js))
for tok in ("__ECHARTS_LIB__","__WORLD_JSON__","__CHINA_JSON__","__NEWS__","__SAMPLE__","__SUGGESTIONS__"):
    assert tok not in out, f"占位符未替换: {tok}"
```

### 光点按时间「挨个闪烁」（非齐闪）
`animationDelay: function(idx){ return order.indexOf(idx)*160; }`，其中 `order` 是按 `time` 排序后的原始下标序列。

---

## 七、性能与体验要点

- **大体积 GeoJSON 解析后置**：首屏先画界面框架，`window.onload` 后再 `initMap()`，避免阻塞首屏。
- **扫描/动画从简**：去掉开机动画、扫描线；用 `effectScatter` 涟漪即可，保持清爽（用户明确嫌扫描特效丑、嫌慢）。
- **字体**：仅用系统字体栈回退（Microsoft YaHei / PingFang SC / Hiragino Sans GB 等），**不再引用任何外部字体 CDN**，离线打开即生效，体感无明显塌方即可。

---

## 八、地图合规

- 用官方/标准 `china.json`（含 34 省级行政区边界），**不要自绘、不要缺省省份**。
- 中国为默认视图，`center:[105,36]`、`zoom:1.5`；世界视图 `center:[68,22]`、`zoom:1.25`。

---

## 九、常见坑（已踩过）

- **Windows schannel 证书**：git/网络若报 `CRYPT_E_NO_REVOCATION_CHECK`，curl 加 `--ssl-no-revoke`、git `http.sslVerify false`（仅本仓库）。
- **WebSearch 无 URL（标 Untitled）**：部分长文/聚合页无有效链接 → 编译 JSON 时改用同领域其他带 URL 结果，避免 norm_key 碰撞或无效链接。
- **短字段**：dedup 保首次，短字段必须在 fill 前补全（可用 `expand.py` 批量把 <100 字字段补到 ≥100 字，基于该条的 company/city/title/direction 写具体不泛化的内容）。
- **分类错配**：深圳建行「经营贷」应归「个人房地产抵押贷款」而非「个人房贷」；分类必须严格对齐 TAXONOMY，否则计数错位、建议区错配。
- **NEWS[0].time 空**：`stats()` 里 `NEWS[0]?NEWS[0].time.slice(0,10):...` 要兜底，避免首屏报错。

---

## 十、交付物清单

- `金融创新全球作战地图.html` —— 离线单文件（库+地图内联），双击即开、可邮件分享。
- `finance-map/news.xlsx` —— 可编辑数据源，用户可直接改表、再 `build` 重建。
- `references/` —— 可复用脚手架：`template.html` / `build.py` / `pipeline.py`。

> 注：`build.py` 依赖 `assets/echarts.min.js`、`assets/world.json`、`assets/china.json`、`assets/xlsx.full.min.js` 四个文件，均已随技能分发、置于脚本同级的 `assets/` 目录，构建时全部内联进 HTML。**不依赖任何运行时 CDN**——离线双击即可打开。xlsx 导出库用于浏览器端"导出当前数据为 .xlsx"功能。
