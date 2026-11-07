# 模拟富豪 · tycoon-sim

> 一个给豆包用的 Skill：查真实在售的游艇、豪宅、油气区块、私人飞机、岛屿、艺术品，用虚构的高端卡包虚拟下单（**绝不真实付款**），用真实行情虚拟炒股，再由一位"私人资产秘书"跟你口语汇报家底、本周盈亏、卡账和关注清单。

它是一个可记账的富豪模拟器，也是一套"想买什么先帮你查清楚"的资产检索工具。所有交易都是假的，所有价格都是真的。

---

## 能做什么

- **找真实货**：游艇（YachtWorld / Boats.com）、豪宅（Sotheby's / Zillow / 链家）、油气区块（BLM / 矿业权市场）、私人飞机、岛屿、艺术品（苏富比 / 佳士得）
- **看图找货**：丢一张微信截图或手机照片，识别品牌型号 → 找同款真实在售价 → 问"要下单吗、刷哪张卡"
- **虚拟下单**：五张虚构高端卡（百夫长黑金、Visa 无限、白金卡……），信用卡进本期账单、借记卡扣余额，额度超了会被拦
- **虚拟炒股**：按真实行情建仓、加减仓、刷价，算本周已实现 + 浮动盈亏
- **记账**：账本存飞书云端表格（手机也能看能记），也可退回本地 JSON + Excel
- **秘书汇报**：固定称呼、口语化讲人话，不贴表格；主动提醒卡账到期、现金趴太多、关注项有新货

## 安装

### 方式一：放进技能目录（推荐）

把 `tycoon-sim/` 整个文件夹放进你的工作环境目录：

```
<你的环境>/workspace/.user_skills/tycoon-sim/
```

新开一个对话，说一句「打开富豪秘书」即被激活。

### 方式二：豆包客户端上传

豆包 PC 端 →「插件 · 技能 · 伙伴」→ 技能页 → 右上角「+ 添加」→ 上传技能，把 `tycoon-sim/` 拖进去。

## 首次使用

技能激活后第一步永远是 `setup`（幂等，重复执行只补缺、不覆盖）：

```bash
# 全新用户：在飞书里建一张云端账本并绑定
python3 <skill>/scripts/ledger.py newsheet --name 富豪账本

# 已有云端表：直接绑定
python3 <skill>/scripts/ledger.py setup --sheet <飞书表格URL>

# 只想本地记账（不连云端）
python3 <skill>/scripts/ledger.py setup --cash 300000000
```

执行后会在 `~/.tycoon/` 生成账本、Excel 台账、`config.json`、`秘书备忘.md`；指定了云端表时，账本存在表格里。

## 常用命令

| 想干的事 | 命令 |
| --- | --- |
| 秘书简报（口语） | `ledger.py brief --chat` |
| 买资产 / 购物 | `ledger.py buy --name "Predator 60 EVO" --price 9944139 --site YachtWorld --url <链接> --card 1005 --type asset` |
| 建仓 / 加减仓 | `ledger.py stock-buy --code 600519 --name 贵州茅台 --shares 200 --price 1453` |
| 刷最新价 | `ledger.py mark 600519=1453 NVDA=1246.05` |
| 本周账单 | `ledger.py report --period week --detail` |
| 关注清单 | `ledger.py watch --action add --category 游艇 --keyword "60英尺 Predator"` |
| 卡包与还款 | `ledger.py card --action list` / `card --action pay --id card-centurion --amount 9944139` |
| 换称呼 | `ledger.py profile --address 李总` |
| 导出 Excel | `ledger.py export` |

完整说明见 [tycoon-sim/SKILL.md](tycoon-sim/SKILL.md)。

## 目录结构

```
tycoon-sim/
├── SKILL.md            # 技能主文档：触发条件、工作流、秘书人格、云端账本
├── scripts/
│   └── ledger.py       # 账本引擎（纯标准库；openpyxl 可选，用于导出 Excel）
└── references/
    └── sources.md      # 真实货源清单：站点、检索话术、常见翻车点
```

## 数据与隐私

- **不产生任何真实交易**：订单只写进你自己的账本，不调用支付接口。
- **卡包全为虚构**：卡名、尾号、额度都是示例；技能从不索取真实卡号、CVV、验证码。
- **账本归你**：云端表在你自己的飞书云空间，只有你（和被你授权的人）能看；不连云端时数据只在你本机。
- **仓库里没有数据**：`.gitignore` 已排除账本、台账、配置与备忘。

## 免责声明

行情价与挂牌价来自公开站点，仅作参考，**不构成任何投资建议**。本项目为个人娱乐与技能示例，与字节跳动 / 豆包官方无关，未获其背书。

## License

MIT，见 [LICENSE](LICENSE)。
