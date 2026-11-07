# 真实货源清单（tycoon-sim）

找货的第一原则：**只报站上真实存在的挂牌，报不出就换站或降级为"搜索结果参考"，绝不编价。**

访问方式优先级：浏览器（`browser-use-automation`，淘宝系可先用 `taobao-native`）→ `web.fetch` 抓挂牌页 → `general_search` 摘录。抓不到具体挂牌时，只给"该平台上这类货的价格区间 + 检索入口"，不要虚构具体标价。

## 一、游艇

| 站点 | 用途 | 备注 |
|---|---|---|
| yachtworld.com | 全球最大游艇/艇经纪挂牌库 | 按 Length、Price、Location、Type（Motor/Sail）筛；每条有经纪商与年份 |
| boats.com | 同集团，新艇与二手艇 | 适合查新艇定价与配置 |
| boattrader.com | 北美中小型艇 | 价格更贴地气 |

检索词：`Sunseeker 131`、`Benetti 50m`、`Azimut Grande`、`Ferretti 720`、`yacht for sale Hong Kong / Singapore / Monaco`。
量级：20–30 米二手艇约 300 万–2000 万美元；50 米以上新造 3000 万美元起。年年有维护费（按船价 5%–10%/年计），可以当游戏里的"持续支出"讲给用户。

## 二、豪宅 / 别墅

| 站点 | 用途 |
|---|---|
| sothebysrealty.com | 全球豪宅，按国家/城市筛，挂牌含实景图 |
| listsothebysrealtyhk.com | 香港豪宅（山顶、南区） |
| zillow.com | 美国房源，价格/估值为美元 |
| ke.com（贝壳）、lianjia.com | 国内二手房/别墅，人民币挂牌价 |

检索词：`Beverly Hills estate for sale`、`Malibu oceanfront`、`Hong Kong Peak house`、`上海 佘山 独栋别墅`、`杭州 西湖 别墅`。
量级：国内一线城市独栋 3000 万–3 亿；比佛利山庄园 2000 万–2 亿美元；香港山顶豪宅 1 亿–10 亿港币。

## 三、油田 / 油气区块（这个最硬核，也最容易编错）

| 来源 | 是什么 | 怎么用 |
|---|---|---|
| energynet.com | 油气资产交易平台，含 BLM 等政府租约在线拍卖 | 站内有 active listings / government lease sales，按州、面积、拍卖截止时间看 |
| blm.gov | 美国联邦土地油气租约拍卖公告（PDF） | 公告里写明拍卖时间、区块、最低报价（bonus bid）与 royalty rate |
| ky.mnr.gov.cn（自然资源部矿业权市场） | 中国油气探矿权挂牌出让公告 | 公告含区块名称、面积、起始价、出让收益率；新疆/黑龙江/渤海湾等区块都在这里 |

口径提醒：油气区块买到的是**探矿权/租约**，不是"油田所有权"。回复用户时要说清"拿下的是 XX 区块探矿权，起始价/中标价 X"，别混成"买了个油田"。量级：国内区块起始价常见几百万到几亿元人民币；美国单个租约 parcel 的 bonus bid 从每英亩几十美元到上千美元不等。

## 四、私人飞机（备选品类）

| 站点 | 用途 |
|---|---|
| controller.com | 飞机/航空器交易挂牌（访问前先核实可用性） |
| avbuyer.com | 公务机挂牌与市场报告 |

检索词：`Gulfstream G650 for sale`、`Bombardier Global 7500 pre-owned`。量级：G650 二手约 3000 万–4500 万美元，Global 7500 新机约 7500 万美元。

## 五、岛屿 / 艺术品（备选品类）

- 岛屿：`privateislandsinc.com`、`vladi-private-islands.de`（访问前核实）；检索 `private island for sale Fiji / Caribbean`。
- 艺术品：`sothebys.com`、`christies.com` 的 auction results / upcoming lots，按 lot 号报价，落槌价与含佣金价要分开说。

## 六、购物网站（虚拟下单只记账）

| 平台 | 说明 |
|---|---|
| 淘宝 / 天猫 taobao.com、tmall.com | 用 `taobao-native` 或浏览器搜索；记 SKU 标题 + 价格 + 链接 |
| 京东 jd.com | 自营价更稳定，适合买家电、数码 |
| 拼多多 yangkeduo.com | 低价日用品，价格波动大 |
| 亚马逊 amazon.com | 海外购，美元计价，注意运费与关税 |
| Farfetch / Net-a-Porter | 奢侈品服饰，欧元/美元标价 |

下单动作永远是：搜索 → 选中具体 SKU → 读真实价格 → `buy --type order --site ... --url ...`。**绝不点"提交订单/支付"**。

## 七、汇率与计价口径

账本一律人民币计价。外币挂牌用 `seed_finance_lookup` 查当日中间价换算，回复里写："$12,000,000 ≈ ¥8,640 万（按 7.20）"。汇率当天查、不要用记忆值。

## 八、检索话术模板

- 游艇：`site:yachtworld.com <品牌> <长度> for sale`、`<城市> yacht brokerage listing`
- 豪宅：`site:sothebysrealty.com <城市> estate`、`<城市> 独栋别墅 挂牌价`
- 油气：`site:energynet.com oil gas properties for sale <州>`、`site:ky.mnr.gov.cn 石油天然气 探矿权 挂牌出让`
- 购物：`<商品名> 价格 site:jd.com`、`<商品名> 淘宝`

## 九、常见翻车点

1. 把资讯文章里的"估值/市场价"当成在售挂牌价 → 只在挂牌页取价。
2. 美元当人民币直接记账 → 必须换算并标注。
3. 油气区块说成"买下油田" → 说清是探矿权/租约。
4. 报价过期 → 挂牌页若显示已售/下架，换一条再报，别硬用。
