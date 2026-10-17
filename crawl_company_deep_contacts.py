"""
特易资讯 (TopEase) - 采购商全景数据与深度联系人批量采集工具 (前50页/500家公司)
特性：
1. 包含企业基础、贸易画像主营产品（Top商品、Top HS编码、Top来源国、Top供应商）、AI深度联系人（姓名、职位、全部邮箱、社媒、决策人标识）
2. 每 5 页（50 家公司）自动落盘保存一次，防止数据丢失
3. 断点续传与异常重试
4. 全面兼容 Windows 控制台编码
"""

import os
import sys
import time
import json
import re
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

SEARCH_COMPANY_URL = "https://et.topease.net/api/gt8/search/search"
COMPANY_PROFILE_URL = "https://et.topease.net/api/gt8/search/getCompanyProfile"
SEARCH_KEY_URL = "https://et.topease.net/api/sou/search/searchCompanys"
CONTACTS_URL = "https://et.topease.net/api/sou/search/getCompanyContacts"

OUTPUT_EXCEL = "采购商_全景画像与联系人_前50页(500家公司).xlsx"
PROGRESS_FILE = "crawl_progress_company_contacts.json"

START_PAGE = 1
END_PAGE = 50  # 50 页 = 500 家公司
SAVE_INTERVAL = 5  # 每 5 页自动存盘一次

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "authorization": "ce65ab8a0db896ba6a5af961dc7ad69e",
    "cookie": "topease_access_token=ce65ab8a0db896ba6a5af961dc7ad69e",
    "language": "zh",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=1&wlf=CompanyIndex",
    "unionkey": "gt8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36",
    "website": "https://et.topease.net"
}

BASE_SEARCH_PAYLOAD = {
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
    "countryIds": [0],
    "tradeDate": 3,
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
    "historySearchId": "72300240"
}

COLUMNS = [
    ("#", 8),
    ("公司名称 (Company)", 35),
    ("国家/地区 (Country)", 16),
    ("注册号 (Tax Code)", 18),
    ("企业官网 (Website)", 28),
    ("最近交易日期", 15),
    ("总交易单数", 12),
    ("近3年主营/采购商品", 35),
    ("近3年采购HS编码", 28),
    ("近3年采购来源国", 20),
    ("近3年核心供应商", 35),
    ("联系人姓名 (Full Name)", 22),
    ("职位头衔 (Job Title)", 35),
    ("电子邮箱 (Email)", 35),
    ("社媒主页 (Social Media)", 35),
    ("联系电话 (Phone)", 20),
    ("是否采购/决策人", 16),
    ("数据更新时间", 14),
]

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    s = ILLEGAL_CHARACTERS_RE.sub("", s)
    return s

def init_excel_workbook(filepath):
    if os.path.exists(filepath):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "采购商全景联系人"
    
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
    return {"last_completed_page": 0, "total_companies": 0, "total_contacts": 0}

def save_progress(page: int, companies: int, contacts: int):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "last_completed_page": page,
            "total_companies": companies,
            "total_contacts": contacts,
            "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
        }, f, indent=2, ensure_ascii=False)

def fetch_company_profile(client: httpx.Client, company_id: str):
    if not company_id:
        return {}
    headers_gt8 = HEADERS.copy()
    headers_gt8["unionkey"] = "gt8"
    try:
        resp = client.post(COMPANY_PROFILE_URL, headers=headers_gt8, json={"id": company_id}, timeout=15.0)
        if resp.status_code != 200:
            return {}
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
        return {}

def fetch_company_contacts(client: httpx.Client, website: str):
    if not website:
        return []
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    try:
        resp_sc = client.get(
            SEARCH_KEY_URL,
            params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": website, "type": 2, "sourceType": 1},
            headers=headers_yisou,
            timeout=20.0
        )
        if resp_sc.status_code != 200:
            return []
        sc_data = resp_sc.json().get("data", {}).get("result", [])
        if not sc_data:
            return []
        search_key = sc_data[0].get("searchKey")
        if not search_key:
            return []
            
        resp_c = client.get(
            CONTACTS_URL,
            params={"status": 0, "pageIndex": 1, "pageSize": 100, "keyword": "", "fuzzyMatch": 2, "countryIds": "", "searchKey": search_key, "sourceType": 1, "statusList": -1, "excludeEmail": ""},
            headers=headers_yisou,
            timeout=20.0
        )
        if resp_c.status_code == 200:
            return resp_c.json().get("data", {}).get("result", [])
    except Exception:
        pass
    return []

def run_crawler():
    prog = load_progress()
    last_page = prog.get("last_completed_page", 0)
    current_start = max(START_PAGE, last_page + 1)
    
    if current_start > END_PAGE:
        print(f"[提示] 前 {END_PAGE} 页已全部采集完成！")
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
    
    print("=" * 70)
    print(f" 采购商全景画像与深度联系人批量采集启动 (前 50 页 / 500 家企业)")
    print(f" 起始页码: 第 {current_start} 页 ~ 第 {END_PAGE} 页")
    print(f" 输出文件: {OUTPUT_EXCEL}")
    print("=" * 70)
    
    start_time = time.time()
    
    with httpx.Client() as client:
        for p in range(current_start, END_PAGE + 1):
            payload = BASE_SEARCH_PAYLOAD.copy()
            payload["pageIndex"] = p
            payload["pageSize"] = 10
            
            headers_gt8 = HEADERS.copy()
            headers_gt8["unionkey"] = "gt8"
            
            try:
                resp = client.post(SEARCH_COMPANY_URL, headers=headers_gt8, json=payload, timeout=25.0)
                data = resp.json()
                if data.get("code") not in [200, 0]:
                    print(f"\n[警告] 第 {p} 页接口返回代码 {data.get('code')}: {data.get('msg')}")
                    if any(k in str(data.get("msg", "")).lower() for k in ["token", "login", "expire"]):
                        print("[停止] 登录态已失效，安全保存退出！")
                        break
                companies = data.get("data", {}).get("result", [])
            except Exception as e:
                print(f"\n[错误] 第 {p} 页公司列表请求失败: {e}")
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
                
                # 获取贸易画像
                profile = fetch_company_profile(client, cid)
                p_products = profile.get("products", "")
                p_hscodes = profile.get("hs_codes", "")
                p_origins = profile.get("origin_countries", "")
                p_suppliers = profile.get("suppliers", "")
                
                # 获取深度联系人
                contacts = fetch_company_contacts(client, cweb) if cweb else []
                
                if not contacts:
                    row_count += 1
                    row_data = [
                        row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                        p_products, p_hscodes, p_origins, p_suppliers,
                        "", "", "", "", "", "", ""
                    ]
                    ws.append(row_data)
                    for col_idx in range(1, len(COLUMNS) + 1):
                        cell = ws.cell(row=row_count + 1, column=col_idx)
                        cell.font = data_font
                        cell.border = thin_border
                        cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 17, 18] else left_align
                else:
                    contact_count += len(contacts)
                    for c in contacts:
                        row_count += 1
                        fname = clean_text(c.get("fullName"))
                        jtitle = clean_text(c.get("jobTitle"))
                        emails = ", ".join([clean_text(e.get("value")) for e in c.get("emails", []) if e.get("value")])
                        links = ", ".join([clean_text(l.get("value")) for l in c.get("links", []) if l.get("value")])
                        phones = ", ".join([clean_text(ph.get("value")) for ph in c.get("phones", []) if ph.get("value")])
                        is_purchase = "是" if (c.get("IsPurchase") or c.get("IsBoss")) else "否"
                        update_d = clean_text(c.get("updateDate"))
                        
                        row_data = [
                            row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                            p_products, p_hscodes, p_origins, p_suppliers,
                            fname, jtitle, emails, links, phones, is_purchase, update_d
                        ]
                        ws.append(row_data)
                        for col_idx in range(1, len(COLUMNS) + 1):
                            cell = ws.cell(row=row_count + 1, column=col_idx)
                            cell.font = data_font
                            cell.border = thin_border
                            cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 17, 18] else left_align
                
                time.sleep(0.2)
            
            # 每页进度
            elapsed = time.time() - start_time
            pages_done = p - current_start + 1
            speed = pages_done / max(1, elapsed)
            remaining_pages = END_PAGE - p
            eta_m = (remaining_pages / max(0.01, speed)) / 60
            percent = (p / END_PAGE) * 100
            print(f"[{percent:5.1f}%] 已完成第 {p:2d}/{END_PAGE} 页 | 累计企业: {company_count:3d} 家 | 累计联系人: {contact_count:5d} 条 | 预计剩余: {eta_m:.1f}分钟", flush=True)
            
            # 定期存盘
            if p % SAVE_INTERVAL == 0 or p == END_PAGE:
                wb.save(OUTPUT_EXCEL)
                save_progress(p, company_count, contact_count)
                
    wb.save(OUTPUT_EXCEL)
    save_progress(p, company_count, contact_count)
    total_time = (time.time() - start_time) / 60
    print("\n" + "=" * 70)
    print(f"[完成] 前 {END_PAGE} 页全部采集完毕！")
    print(f"总计处理企业: {company_count} 家 | 获取深度联系人: {contact_count} 条")
    print(f"总耗时: {total_time:.2f} 分钟")
    print(f"文件保存至: {os.path.abspath(OUTPUT_EXCEL)}")
    print("=" * 70)

if __name__ == "__main__":
    run_crawler()
