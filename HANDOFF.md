# 客户爬取任务交接文档 (HANDOFF.md)

> **使用指引**：在新的会话窗口或更换新账号时，直接将本文档内容与**新账号的 cURL 请求信息**一同发送给 AI 助手，即可无缝继续执行爬取。

---

## 一、项目基础信息

- **项目本地路径**：`d:\我的APP\客户爬取`
- **GitHub 仓库**：[https://github.com/haojie10/customer-crawl](https://github.com/haojie10/customer-crawl)
- **环境依赖**：Python 3.10+，核心依赖：`httpx`, `openpyxl`, `requests` (`pip install -r requirements.txt`)

---

## 二、当前进度与下一批次

| 批次范围 | 规则标准 | 状态 | 成果文件 |
| :--- | :--- | :---: | :--- |
| **第 001 ~ 200 页** | Purchase 职位 + Valid 验证邮箱 | 已完成 | 分卷 1~4 卷已落盘 |
| **第 351 ~ 400 页** | Leader+Purchase 双标签 + 独立验证列 + 单邮箱拆行 | 已完成 | `欧盟采购商_(351-400页)_单邮箱导入版.xlsx` |
| **第 401 ~ 450 页** | Leader+Purchase 双标签 + 独立验证列 + 单邮箱拆行 | 已完成 | `欧盟采购商_(401-450页)_单邮箱导入版.xlsx` |
| **第 451 ~ 500 页** | Leader+Purchase 双标签 + 独立验证列 + 单邮箱拆行 | 已完成 | `欧盟采购商_(451-500页)_单邮箱导入版.xlsx` |
| **第 501 ~ 550 页** | Leader+Purchase 双标签 + 独立验证列 + 单邮箱拆行 | 已完成 | `欧盟采购商_近1年高频采购_有效采购决策人_(501-550页_500家公司)_单邮箱导入版.xlsx` |
| **第 551 ~ 600 页** | Leader+Purchase 双标签 + 独立验证列 + 单邮箱拆行 | 已完成 | `欧盟采购商_(551-600页)_单邮箱导入版.xlsx` |
| **第 601 ~ 650 页** | **当前待执行批次**（共 50 页 / 500 家企业） | **待启动** | 待新账号提供 Token 后执行 |

---

## 三、新账号需提供的信息

换用新账号后，只需在网页端任意点开一次搜索/联系人请求，复制 cURL 或提供以下两项：
1. **`authorization` / `topease_access_token`**（32 位 Hex Token）
2. **`historySearchId`**（检索请求 payload 中的搜索 ID）

---

## 四、最新业务筛选与数据规范

1. **企业筛选条件**：
   - 欧盟 27 国、近 1 年交易频次 > 20 次、含中国采购经验、排除微型企业/货代/银行。
2. **联系人提取规范**：
   - 接口：`getCompanyContacts`
   - 参数：`statusList="1,5"`（同时包含 Leader 企业负责人/高管 + Purchase 采购决策人）、`status=0`（筛选全部状态 ALL）。
   - 角色映射：
     - `IsBoss and IsPurchase` $\rightarrow$ `采购 + 企业负责人 (Leader)`
     - `IsBoss` $\rightarrow$ `企业负责人/高管 (Leader)`
     - `IsPurchase` $\rightarrow$ `采购负责人 (Purchase)`
     - 其它 $\rightarrow$ `关键决策人 (Decision Maker)`
3. **邮箱验证状态独立列**：
   - `Valid (有效)`：`valid == True` 或 `grade == 'A'`
   - `Risky (可能有效)`：`valid == False` 或 `grade in ['B', 'C', 'D', 'F']`
   - `未验证`：其余未校验情况
4. **单邮箱单行规范（CRM 导入友好）**：
   - 每行仅保留一个电子邮箱及其对应的一个验证状态。
   - 若原数据同一人有多个邮箱，拆分成多行，连续重编序号 `#`，完整保留企业画像。

---

## 五、执行指令与核心脚本

### 1. 爬取执行（以 551~600 页为例）
基于已有模版 `crawl_eu_buyers_501_550.py` 另存为 `crawl_eu_buyers_551_600.py`，更新：
- `START_PAGE = 551`, `END_PAGE = 600`
- `AUTH_TOKEN = "<新Token>"`
- `historySearchId = "<新SearchId>"`
- `OUTPUT_EXCEL = "欧盟采购商_近1年高频采购_有效采购决策人_(551-600页_500家公司).xlsx"`
- `PROGRESS_FILE = "crawl_progress_eu_551_600.json"`
- `load_progress` 中默认 `last_completed_page = 550`

启动命令（后台防风控运行）：
```bash
python -u crawl_eu_buyers_551_600.py
```

### 2. 多邮箱拆行为单行（导入版生成）
爬取完成后，运行项目中已就绪的通用拆解工具：
```bash
# 保留无邮箱行（标准导入版）：
python split_emails_to_rows.py "欧盟采购商_近1年高频采购_有效采购决策人_(551-600页_500家公司).xlsx"

# 若导入系统要求每行必须有邮箱，加上 --filter-empty：
python split_emails_to_rows.py "欧盟采购商_近1年高频采购_有效采购决策人_(551-600页_500家公司).xlsx" --filter-empty
```
