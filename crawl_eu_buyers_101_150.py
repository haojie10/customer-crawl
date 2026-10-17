"""
特易资讯 (TopEase) - 欧盟活跃采购商 (近1年交易>20次) 全景画像与有效采购决策人采集工具
【高风控保护温和采集版 - 仅抓取前50页 / 500家企业】

筛选条件：
1. 区域/国家：欧盟 27 国 (EU)
2. 交易时间：近 1 年 (tradeDate: 1)
3. 交易次数：> 20 次 (tradeCount: 20)
4. 企业筛选：含中国采购经验，排除物流、银行、小微企业
5. 联系方式：必须含邮箱
6. 深度联系人：仅抓取【Purchase 采购职位 (statusList=5)】且【邮箱验证状态为 Valid (status=2)】

防封控策略：
- 请求之间添加 1.0 ~ 2.0 秒随机延迟，拟人化访问节奏
- 每 5 页自动增量落盘保存一次 Excel 与进度
- 遇到异常或会话失效自动安全退出
"""

import os
import sys
import time
import json
import random
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# API 端点
SEARCH_COMPANY_URL = "https://et.topease.net/api/gt8/search/search"
COMPANY_PROFILE_URL = "https://et.topease.net/api/gt8/search/getCompanyProfile"
SEARCH_KEY_URL = "https://et.topease.net/api/sou/search/searchCompanys"
CONTACTS_URL = "https://et.topease.net/api/sou/search/getCompanyContacts"

# 本次任务配置
START_PAGE = 101
END_PAGE = 150         # 第 101 ~ 150 页，共 500 家企业
PAGE_SIZE = 10         # 每页 10 家公司，总计 500 家
SAVE_INTERVAL = 5      # 每 5 页保存一次

OUTPUT_EXCEL = "欧盟采购商_近1年高频采购_有效采购决策人_(101-150页_500家公司).xlsx"
PROGRESS_FILE = "crawl_progress_eu_101_150.json"

# 认证与请求头
AUTH_TOKEN = "f67d75ce85a6506ba28033490fda3b44"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "authorization": AUTH_TOKEN,
    "cookie": f"topease_access_token={AUTH_TOKEN}",
    "language": "en",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=1&wlf=CompanyIndex",
    "unionkey": "gt8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36",
    "website": "https://et.topease.net"
}

# 欧盟 27 国 + 近1年 + 交易>20次 + 含中国采购经验
SEARCH_PAYLOAD_TEMPLATE = {
    "autoTranslate": True,
    "fullAddressMatchs": [],
    "productDescriptions": [],
    "hsCodes": [],
    "industryCategorys": [],
    "industryType": "",
    "companyNames": [],
    "companyNumbers": [],
    "companyWebsites": [],
    "buyerCompanys": [],
    "supplierCompanys": [],
    "countryIds": [76, 194, 51, 68, 80, 143, 53, 167, 12, 112, 166, 18, 54, 127, 181, 99, 64, 170, 69, 118, 30, 186, 119, 92, 182, 101, 56],
    "tradeDate": 1,
    "tradeCount": 20,
    "buyFromArea": 43,
    "dataSourceCountryId": 43,
    "excludeLogistics": True,
    "excludeBank": True,
    "excludeMicroAgent": True,
    "showContactIcon": True,
    "productTypes": [],
    "surnames": [],
    "tradeShowNames": [],
    "tradeShowNameAndSessions": [],
    "linkTypes": [],
    "industryTyCodes": [],
    "containContact": 3,
    "sortType": "desc",
    "sort": "intelligent",
    "metric": "",
    "type": 1,
    "sortChange": True,
    "paixunt": True,
    "countryId": 0,
    "historySearchId": "72826896"
}

COLUMNS = [
    ("#", 8),
    ("公司名称 (Company)", 35),
    ("国家/地区 (Country)", 16),
    ("注册号 (Tax Code)", 18),
    ("企业官网 (Website)", 28),
    ("最近交易日期", 15),
    ("总交易单数", 12),
    ("近3年主营/采购商品 (Top3)", 40),
    ("近3年采购HS编码 (Top3)", 30),
    ("近3年主要采购来源国", 25),
    ("近3年核心合作供应商", 40),
    ("采购决策人姓名", 20),
    ("职位头衔", 30),
    ("有效电子邮箱 (Valid)", 30),
    ("社媒主页 (LinkedIn等)", 35),
    ("是否采购决策人", 16),
    ("数据更新时间", 14),
]

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def init_excel_workbook(filepath):
    if os.path.exists(filepath):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "欧盟采购决策人"
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    headers = [col[0] for col in COLUMNS]
    ws.append(headers)
    
    for col_idx, (_, col_width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = col_width
    
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    wb.save(filepath)
    return wb, ws

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "last_completed_page": 100,
        "total_companies": 0,
        "total_contacts": 0,
        "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
    }

def save_progress(page: int, companies: int, contacts: int):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "last_completed_page": page,
            "total_companies": companies,
            "total_contacts": contacts,
            "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
        }, f, indent=2, ensure_ascii=False)

def fetch_company_profile(client: httpx.Client, company_id: str, retries=2):
    if not company_id:
        return {}
    headers_gt8 = HEADERS.copy()
    headers_gt8["unionkey"] = "gt8"
    for attempt in range(retries):
        try:
            resp = client.post(COMPANY_PROFILE_URL, headers=headers_gt8, json={"id": company_id}, timeout=15.0)
            if resp.status_code != 200:
                continue
            data = resp.json().get("data", {})
            cards = data.get("cards", [])
            profile = {"products": "", "hs_codes": "", "origin_countries": "", "suppliers": ""}
            for card in cards:
                ctype = card.get("type")
                items = card.get("items", [])
                if ctype == "purchaseProduct":
                    profile["products"] = ", ".join([clean_text(it.get("name")) for it in items if it.get("name")])
                elif ctype == "purchaseHsCode":
                    profile["hs_codes"] = ", ".join([clean_text(it.get("name") or it.get("code")) for it in items if it.get("name") or it.get("code")])
                elif ctype == "purchaseOriginCountry":
                    profile["origin_countries"] = ", ".join([clean_text(it.get("name")) for it in items if it.get("name")])
                elif ctype == "supplier":
                    profile["suppliers"] = ", ".join([clean_text(it.get("name")) for it in items if it.get("name")])
            return profile
        except Exception:
            time.sleep(1.0)
    return {}

def fetch_valid_purchase_contacts(client: httpx.Client, website: str, retries=2):
    """
    精准抓取 Purchase 职位 (statusList=5) 且 邮箱为 Valid 状态 (status=2) 的联系人
    """
    if not website:
        return []
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    for attempt in range(retries):
        try:
            # 步骤 1: 用官网换取 searchKey
            resp_sc = client.get(
                SEARCH_KEY_URL,
                params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": website, "type": 2, "sourceType": 1},
                headers=headers_yisou,
                timeout=15.0
            )
            if resp_sc.status_code != 200:
                continue
            sc_data = resp_sc.json()
            if sc_data.get("code") == 1004:
                print("\n[提示] 易搜 AI 搜索量已达今日账号限额，将仅保留贸易画像！")
                return []
            result_list = sc_data.get("data", {}).get("result", [])
            if not result_list:
                return []
            search_key = result_list[0].get("searchKey")
            if not search_key:
                return []
            
            # 拟人化小微停顿
            time.sleep(random.uniform(0.3, 0.6))
            
            # 步骤 2: 精准请求 status=2 (Valid) & statusList=5 (Purchase职位)
            contact_params = {
                "status": 2,          # Valid 邮箱验证
                "pageIndex": 1,
                "pageSize": 50,
                "keyword": "",
                "fuzzyMatch": 2,
                "countryIds": "",
                "searchKey": search_key,
                "sourceType": 1,
                "statusList": 5,      # Purchase 职位
                "excludeEmail": ""
            }
            resp_c = client.get(CONTACTS_URL, params=contact_params, headers=headers_yisou, timeout=15.0)
            if resp_c.status_code == 200:
                c_data = resp_c.json()
                return c_data.get("data", {}).get("result", [])
        except Exception:
            time.sleep(0.8)
    return []

def run():
    prog = load_progress()
    last_page = prog.get("last_completed_page", 0)
    current_start = last_page + 1
    
    if current_start > END_PAGE:
        print(f"[提示] 前 {END_PAGE} 页已全部采集完毕！")
        return
        
    wb, ws = init_excel_workbook(OUTPUT_EXCEL)
    data_font = Font(name="Microsoft YaHei", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    row_count = ws.max_row - 1
    company_count = prog.get("total_companies", 0)
    contact_count = prog.get("total_contacts", 0)
    
    print("=" * 80)
    print(" [启动] 欧盟高频采购商（近1年>20次交易）深度采集")
    print(f" 采集范围: 第 {current_start} 页 ~ 第 {END_PAGE} 页 (严格限制 50 页 / 500 家企业)")
    print(" 筛选规则: 欧盟27国 + 近1年交易>20次 + 含中国采购经验 + Purchase采购职位 + Valid有效邮箱")
    print(f" 防封保护: 拟人化随机延迟 (1.0~2.0s/企业) | 输出文件: {OUTPUT_EXCEL}")
    print("=" * 80)
    
    start_time = time.time()
    
    with httpx.Client() as client:
        for p in range(current_start, END_PAGE + 1):
            payload = SEARCH_PAYLOAD_TEMPLATE.copy()
            payload["pageIndex"] = p
            payload["pageSize"] = PAGE_SIZE
            
            headers_gt8 = HEADERS.copy()
            headers_gt8["unionkey"] = "gt8"
            
            companies = []
            for attempt in range(3):
                try:
                    resp = client.post(SEARCH_COMPANY_URL, headers=headers_gt8, json=payload, timeout=25.0)
                    data = resp.json()
                    code = data.get("code")
                    msg = str(data.get("msg", ""))
                    if code not in [200, 0]:
                        print(f"\n[警告] 第 {p} 页接口返回代码 {code}: {msg}")
                        if code == 401 or any(k in msg.lower() for k in ["token", "login", "expire", "session"]) or "过期" in msg:
                            print("\n[停止] 登录态已过期，已为您安全落盘！请更新 Token 后继续。")
                            wb.save(OUTPUT_EXCEL)
                            save_progress(p - 1, company_count, contact_count)
                            return
                    companies = data.get("data", {}).get("result", [])
                    break
                except Exception as e:
                    print(f"\n[重试] 第 {p} 页网络请求异常 (尝试 {attempt+1}/3): {e}")
                    time.sleep(2.0)
                    
            if not companies:
                print(f"[提示] 第 {p} 页无数据或请求未返回结果。")
                continue
                
            for comp in companies:
                company_count += 1
                cid = comp.get("id")
                cname = clean_text(comp.get("companyName"))
                ccountry = clean_text(comp.get("country"))
                ctax = clean_text(comp.get("companyTaxCode"))
                cweb = clean_text(comp.get("website"))
                clast_date = clean_text(comp.get("lastTradeDate"))
                ctrade_total = clean_text(comp.get("companyTradeTotal"))
                
                # 1. 抓取贸易画像 (Top3商品/HS/货源国/供应商)
                profile = fetch_company_profile(client, cid)
                p_products = profile.get("products", "")
                p_hscodes = profile.get("hs_codes", "")
                p_origins = profile.get("origin_countries", "")
                p_suppliers = profile.get("suppliers", "")
                
                # 2. 拟人化延迟
                time.sleep(random.uniform(0.5, 1.0))
                
                # 3. 抓取 Purchase 职位 & Valid 邮箱联系人
                contacts = fetch_valid_purchase_contacts(client, cweb) if cweb else []
                
                if not contacts:
                    # 即使无采购职位联系人，也完整保留公司画像行
                    row_count += 1
                    row_data = [
                        row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                        p_products, p_hscodes, p_origins, p_suppliers,
                        "", "", "", "", "", ""
                    ]
                    ws.append(row_data)
                    for col_idx in range(1, len(COLUMNS) + 1):
                        cell = ws.cell(row=row_count + 1, column=col_idx)
                        cell.font = data_font
                        cell.border = thin_border
                        cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 16, 17] else left_align
                else:
                    contact_count += len(contacts)
                    for c in contacts:
                        row_count += 1
                        fname = clean_text(c.get("fullName"))
                        jtitle = clean_text(c.get("jobTitle"))
                        emails = ", ".join([clean_text(e.get("value")) for e in c.get("emails", []) if e.get("value")])
                        links = ", ".join([clean_text(l.get("value")) for l in c.get("links", []) if l.get("value")])
                        is_purchase = "是"  # 已精准筛选 Purchase 职位
                        update_d = clean_text(c.get("updateDate"))
                        
                        row_data = [
                            row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                            p_products, p_hscodes, p_origins, p_suppliers,
                            fname, jtitle, emails, links, is_purchase, update_d
                        ]
                        ws.append(row_data)
                        for col_idx in range(1, len(COLUMNS) + 1):
                            cell = ws.cell(row=row_count + 1, column=col_idx)
                            cell.font = data_font
                            cell.border = thin_border
                            cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 16, 17] else left_align
                
                # 拟人化随机节奏 (0.8 ~ 1.5 秒/公司)，全面降低风控风险
                time.sleep(random.uniform(0.8, 1.5))
            
            # 单页完成进度打印
            pct = (p / END_PAGE) * 100
            elapsed = time.time() - start_time
            print(f"[{pct:5.1f}%] 已完成第 {p:2d}/{END_PAGE} 页 | 累计企业: {company_count} 家 | 有效采购决策人: {contact_count} 位", flush=True)
            
            # 定期安全落盘
            if p % SAVE_INTERVAL == 0 or p == END_PAGE:
                wb.save(OUTPUT_EXCEL)
                save_progress(p, company_count, contact_count)
            
            # 每页之间停顿 1.5 ~ 2.5 秒
            time.sleep(random.uniform(1.5, 2.5))
            
    wb.save(OUTPUT_EXCEL)
    save_progress(END_PAGE, company_count, contact_count)
    total_mins = (time.time() - start_time) / 60
    print("\n" + "=" * 80)
    print(f" [完成] 前 {END_PAGE} 页 (共 {company_count} 家企业) 深度调研全部完成！")
    print(f" 累计获取有效采购决策人: {contact_count} 位")
    print(f" 耗时: {total_mins:.2f} 分钟 | 输出文件: {os.path.abspath(OUTPUT_EXCEL)}")
    print("=" * 80)

if __name__ == "__main__":
    run()
