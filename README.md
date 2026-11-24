# 股票日线增量表

用 **Baostock**（免费、官方、无需 token）每天抓一次 A 股日线，把「当日最低 / 当日最高 / 收盘 / 涨跌额 / 涨跌幅」增量追加进一个 Excel 表，跑在 **GitHub Actions** 上，可选往**企业微信群**或**飞书群**发一条摘要。

- 零成本：不用服务器、不用数据库、不用大模型
- 增量不重复：按 `(日期, 股票代码)` 去重，跑多少次都只保留一份
- 持久化靠 Git：每次跑完把 `excel/daily_stock.xlsx` 提交回本仓库，第二天接着往下加
- 涨红 / 跌绿 / 平灰，表头冻结；「股票」列是 `代码 名称` 合成的一列（如 `688981 中芯国际`），筛选只开在「日期 / 股票」两列（数值列不放下拉箭头，选一只股票整行跟着筛）

## 表格长什么样

| 日期 | 股票代码 | 股票名称 | 当日最低 | 当日最高 | 收盘 | 涨跌额 | 涨跌幅(%) | 涨/跌 |
|---|---|---|---|---|---|---|---|---|
| 2026-11-20 | 688981 | 中芯国际 | 119.33 | 125.68 | 125.46 | +3.21 | +2.63% | 涨 |
| 2026-11-20 | 688825 | 长鑫科技 | 35.16 | 36.96 | 35.16 | -2.34 | -6.24% | 跌 |

按日期倒序（最新在最上面）。同时会写一份 `excel/daily_stock.csv`，在 GitHub 网页上点开就能看表格。

## 目录

```
stock-daily-excel/
├─ scripts/baostock_daily_excel.py   抓数 + 写表 + 推送（单文件，可本地跑）
├─ .github/workflows/daily.yml      按 RUN_TIME（默认北京 19:00）自动跑 + 手动触发
├─ excel/daily_stock.xlsx           数据表（由 Action 自动提交）
├─ excel/daily_stock.csv            同内容 CSV，方便网页查看
├─ requirements.txt
├─ .env.example                     本地运行用的配置样例
└─ run.ps1                          Windows 一键本地运行
```

## 一、本地跑（先验证数据对不对）

```powershell
cd F:\Plugin\stock-daily-excel
python -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt

Copy-Item .env.example .env      # 填 STOCK_LIST，想要微信就填 WECHAT_WEBHOOK_URL
.\run.ps1                        # 等价于 python scripts\baostock_daily_excel.py --days 10
```

直接指定股票也可以，不必写 .env：

```powershell
.\.venv\Scripts\python.exe scripts\baostock_daily_excel.py --stocks '300750,688256,688836' --days 30
```

> ⚠️ PowerShell 里**逗号必须加引号**（`'600519,000001'`）。不引号会被当成数组拆成两个参数，
> 报 `unrecognized arguments`；另外 `000001` 这种前导 0 的代码不引号还会被吃成数字 `1`。

参数：

| 参数 | 默认 | 说明 |
|---|---|---|
| `--stocks` | 环境变量 `STOCK_LIST` | 逗号分隔的 6 位 A 股代码 |
| `--days` | 10 | 回溯天数；增量会去重，多取几天没坏处 |
| `--excel` | `excel/daily_stock.xlsx` | 输出路径，相对仓库根目录 |
| `--no-wechat` | 关 | 只写表，不推送企业微信 |
| `--no-feishu` | 关 | 只写表，不推送飞书 |

## 二、部署成 GitHub Action（真正常用的方式）

1. 在 GitHub 新建一个**空仓库**（public 更省额度；不要勾选自动生成 README），然后把本目录推上去
   （本目录已经 `git init` 好了，`.venv` 和 `.env` 已在 `.gitignore` 里，不会被提交）：

   ```powershell
   cd F:\Plugin\stock-daily-excel
   git add .
   git commit -m "init: baostock daily excel"
   git remote add origin https://github.com/<你的用户名>/stock-daily-excel.git
   git push -u origin main
   ```

2. 仓库 **Settings → Secrets and variables → Actions**：

   | 类型 | 名称 | 值 |
   |---|---|---|
   | Variables | `STOCK_LIST` | `300750,688256,688836`（逗号分隔，无空格） |
   | Variables（可选） | `RUN_TIME` | `19:00`（**北京时间** HH:MM，不填默认 19:00） |
   | Secrets（可选） | `WECHAT_WEBHOOK_URL` | 企业微信群机器人 Webhook |
   | Secrets（可选） | `FEISHU_WEBHOOK_URL` | 飞书群自定义机器人 Webhook |

   两个推送渠道互相独立，填哪个发哪个，都填就都发。

3. **Actions** 页 → 左侧 `股票日线` → **Run workflow** 跑一次。
   跑完仓库里会多一个 `chore(excel): 更新股票日线 ...` 的提交，说明增量落盘成功。

4. 之后每天按 `RUN_TIME`（北京时间）自动跑。

**定时是怎么做到"可配置"的：** GitHub 的 `cron` 只接受写死的字符串，
`${{ vars.X }}` 塞不进去；而且 GitHub 的定时唤醒**本身就不准时**——实测晚 30~40 分钟很常见，
偶尔整次丢掉。所以 workflow 不指望"准点醒"，而是 **每 20 分钟醒一次**（`*/20 * * * 1-5`，UTC），
由第一步「时间判断」决定要不要真跑：

1. 今天已经跑过（`state/last_run.txt` 记着今天的日期）→ 跳过
2. 现在还没到 `RUN_TIME` → 跳过
3. 其余情况 → 执行。**哪怕晚了两三个小时才醒，只要当天还没跑就会补上；跑完写标记，当天不再重复。**

- 想改成 `18:30`？把 `RUN_TIME` 填 `18:30` 就行，**不用改代码、不用提交**
- 误差范围：正常情况下在 `RUN_TIME` 之后 0~20 分钟内执行（每 20 分钟醒一次 + GitHub 自身延迟）
- 慢一点也没关系：窗口不是"1 小时"，是"当天剩余时间"，所以晚 2 小时也能跑上
- **同一天只会真正执行一次**：跳过的那些唤醒 10 秒内退出，日志里会写 `今天 ... 已经跑过了` / `还没到 ...`，**这是正常的**
- `RUN_TIME` 写错格式（如 `8:00`、`20:70`）会直接报 `变量 RUN_TIME=... 格式不对`，不会静默跑错
- 手动点 **Run workflow** 不受 `RUN_TIME` 和"今天已跑过"限制，随时跑

> 持久化的原理：Action 的工作区每次都是全新的，所以脚本跑完会 `git commit && git push` 把
> Excel 写回仓库；下次 checkout 下来就是同一个文件，继续追加。**不需要服务器。**

## 三、飞书群怎么配

1. 打开飞书群 → **设置**（右上角 `…`）→ **群机器人** → **添加机器人** → **自定义机器人** → 复制 Webhook
2. 形如 `https://open.feishu.cn/open-apis/bot/v2/hook/65c464a6-xxxx-xxxx`
3. 填进 Secrets 的 `FEISHU_WEBHOOK_URL`，下次运行群里会收到一张卡片：

```
📈 每日行情速览 · 2026-11-20
16 只 ｜ 涨 8 · 跌 7 · 平 1
表内累计 336 行
─────────────────────
🔴 688981 中芯国际  涨跌 +2.63%（3.21）  高 125.68 · 低 119.33  收盘 125.46
🟢 688825 长鑫科技  涨跌 -6.24%（-2.34）  高 36.96 · 低 35.16  收盘 35.16
```

字段顺序：股票代码、名称、涨跌幅（涨跌额）、当日最高、当日最低、收盘，
按涨跌幅从高到低排，涨红跌绿。

**注意事项：**

- 自定义机器人（webhook）只支持 `text` / `post` / `image` / `interactive`（卡片），**发不了文件**。
  想发 xlsx 本体必须建企业自建应用（要 `im:resource` 权限），而且个人版飞书（`my.feishu.cn`）
  目前调用开放平台 API 会被拒（`21001 Access denied due to platform policy`）——所以这里只发卡片，
  表格本体留在仓库里。
- 机器人「安全设置」如果用**签名校验**，本脚本不支持（只支持无校验或关键词模式）；用关键词模式时，
  确保关键词出现在消息里（默认抬头是「每日行情」，可作关键词）。
- 飞书限流 100 次/分钟、5 次/秒，本脚本一次只发 1 条，不会触发。

## 四、企业微信怎么配

1. 企业微信里建一个群（可以只有自己）→ 群右上角 `…` → **群机器人** → 添加机器人 → 复制 Webhook
2. 形如 `https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxxxxxx`
3. 填进 Secrets 的 `WECHAT_WEBHOOK_URL`，下次运行就会收到：

```
## 股票日线更新 2026-11-20
> 新增 273 行 ｜ 表内共 336 行
> 涨 114 ｜ 跌 155 ｜ 平 4
**涨幅前三**
- 300223 君正股份 +18.37%（+17.38）
```

只发文字摘要，是因为文字最稳（企业微信 markdown 单条约 4096 字节，脚本会自动截断）。
如果以后想把 xlsx 文件本身发进群也可以：先 `POST webhook/upload_media?type=file` 拿
`media_id`，再发 `msgtype=file`（文件 ≤ 20MB），见
[群机器人文档](https://developer.work.weixin.qq.com/document/path/99110)。

## 五、数据源说明

- 来源：`baostock.query_history_k_data_plus(..., frequency="d", adjustflag="3")`
- `adjustflag="3"` = **不复权**，所以「当日最低/最高/收盘」就是当天真实成交价
- 涨跌幅直接用官方字段 `pctChg`（除权口径由交易所日线数据保证），缺失时才用 `close/preclose` 反算
- 涨跌额 = `close - preclose`
- 停牌日 Baostock 返回空串，脚本已跳过

## 六、已知限制

| 限制 | 说明 |
|---|---|
| 只有 A 股 | 港股/美股会被跳过；Baostock 也不提供北交所（`4`/`8` 开头） |
| 当日数据 17:30 后才更新 | 默认 `RUN_TIME=19:00`；下午手动跑会得到「今日没有新数据」 |
| 首次要跑久一点 | `--days 30` 补历史，之后每天只加 1 天 |
| 定时不会精确到分 | GitHub 的 scheduled 唤醒本身就会延迟（实测 30~40 分钟常见，偶尔整次丢），所以实际执行时间是 `RUN_TIME` 之后 0~20 分钟（还没醒就得等下一次唤醒，最坏几小时，但当天一定会跑） |
| 跳过也会计费 | 每 20 分钟醒一次，跳过的那次也按 1 分钟计费（约 72 分钟/天）；公开仓库免费 |
| 企业微信单条 4096 字节 | 脚本按字节截断，不会发送失败 |
| 别用表头下拉里的「排序」 | 筛选区只含 A:B，用它排序只会重排这两列、把数值列甩开；表本身已按日期倒序，要排序请告诉我改脚本 |
