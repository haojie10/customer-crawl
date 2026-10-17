"""
特易资讯 (TopEase) - 采购商全景贸易画像与主营产品大批量自动分卷挂机采集工具
功能：
1. 自动分卷：每 100 页（1,000 家公司）自动保存并封存为一个独立 Excel 文件（如 第03卷_(151-250页).xlsx）
2. 全景贸易画像：
   - 基础档案：公司名、国家、税号/注册号、企业官网、最近交易日、总交易单数
   - 近3年主营/采购商品 Top3
   - 近3年采购 HS 编码 Top3
   - 近3年采购商品主要来源国 Top3
   - 近3年核心合作供应商 Top3
   - 深度联系人：姓名、职位、完整邮箱、社媒主页、决策人标识
3. 断点续传与持久化：每 10 页自动存盘一次，支持随时中断并无缝恢复
4. 安全保障：Token 失效自停保护、异常自动重试
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

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# 接口地址
SEARCH_COMPANY_URL = "https://et.topease.net/api/gt8/search/search"
COMPANY_PROFILE_URL = "https://et.topease.net/api/gt8/search/getCompanyProfile"
SEARCH_KEY_URL = "https://et.topease.net/api/sou/search/searchCompanys"
CONTACTS_URL = "https://et.topease.net/api/sou/search/getCompanyContacts"

# 挂机配置
TOTAL_MAX_PAGES = 3000       # 最大挂机总页数（可按需调大）
PAGES_PER_VOLUME = 100       # 每卷页数（100页 = 1,000家企业）
SAVE_INTERVAL = 10           # 每 10 页自动保存一次 Excel 与进度
GLOBAL_START_PAGE = 151      # 本轮挂机起始页（前 150 页已在第 1、2 批完工）

PROGRESS_FILE = "crawl_progress_daemon.json"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "authorization": "fc9cae7feaee7780b5b662c6beb16cd2",
    "cookie": "topease_access_token=fc9cae7feaee7780b5b662c6beb16cd2",
    "language": "zh",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=1&wlf=CompanyIndex",
    "unionkey": "gt8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36",
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
    "historySearchId": "72372538"
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
    ("联系人姓名", 20),
    ("职位头衔", 30),
    ("电子邮箱", 30),
    ("社媒主页", 35),
    ("是否采购/决策人", 16),
    ("数据更新时间", 14),
]

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def get_volume_info(page_num: int):
    """
    根据页码计算对应的卷号及该卷的起止页码
    例如:
    - 第 1 ~ 100 页 -> 第 01 卷 (1-100页)
    - 第 101 ~ 200 页 -> 第 02 卷 (101-200页)
    - 第 151 ~ 250 页 -> 第 03 卷 (151-250页)  (当从 151 开始 100 页每卷时)
    """
    # 按照 1-100, 101-200, 201-300 或 151-250 规则对齐
    # 这里按用户明确的 151-250, 251-350, 351-450 顺推
    if page_num < 151:
        vol_idx = 1 if page_num <= 50 else 2
        v_start = 1 if vol_idx == 1 else 51
        v_end = 50 if vol_idx == 1 else 150
    else:
        offset = page_num - 151
        vol_idx = 3 + (offset // PAGES_PER_VOLUME)
        v_start = 151 + (vol_idx - 3) * PAGES_PER_VOLUME
        v_end = v_start + PAGES_PER_VOLUME - 1

    filename = f"采购商_全景贸易画像与主营产品_第{vol_idx:02d}卷_({v_start}-{v_end}页).xlsx"
    return vol_idx, v_start, v_end, filename

def init_excel_workbook(filepath):
    if os.path.exists(filepath):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "采购商贸易画像与主营产品"
    
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
    # 默认从 150 之后开始
    return {
        "last_completed_page": 150,
        "current_volume": 3,
        "volume_companies": 0,
        "volume_contacts": 0,
        "total_all_companies": 1500,
        "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
    }

def save_progress(page: int, vol_idx: int, vol_companies: int, vol_contacts: int, total_comps: int):
    data = {
        "last_completed_page": page,
        "current_volume": vol_idx,
        "volume_companies": vol_companies,
        "volume_contacts": vol_contacts,
        "total_all_companies": total_comps,
        "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def fetch_company_profile(client: httpx.Client, company_id: str, retries=3):
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
            time.sleep(0.5)
    return {}

def fetch_company_contacts(client: httpx.Client, website: str, retries=2):
    if not website:
        return []
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    for attempt in range(retries):
        try:
            resp_sc = client.get(
                SEARCH_KEY_URL,
                params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": website, "type": 2, "sourceType": 1},
                headers=headers_yisou,
                timeout=15.0
            )
            if resp_sc.status_code != 200:
                continue
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
                timeout=15.0
            )
            if resp_c.status_code == 200:
                return resp_c.json().get("data", {}).get("result", [])
        except Exception:
            time.sleep(0.3)
    return []

def run_daemon():
    prog = load_progress()
    last_page = prog.get("last_completed_page", 150)
    current_page = last_page + 1
    total_comps = prog.get("total_all_companies", 1500)
    
    print("=" * 80)
    print(" [启动] 特易资讯 - 采购商全景贸易画像【大批量自动分卷挂机模式】已启动")
    print(f" 起始页码: 第 {current_page} 页 | 单卷规格: {PAGES_PER_VOLUME} 页 / 1,000 家企业")
    print(f" 挂机目标: 连续处理至第 {TOTAL_MAX_PAGES} 页")
    print("=" * 80)
    
    data_font = Font(name="Microsoft YaHei", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    current_vol_idx = None
    wb = None
    ws = None
    current_vol_file = None
    vol_companies = 0
    vol_contacts = 0
    
    with httpx.Client() as client:
        while current_page <= TOTAL_MAX_PAGES:
            vol_idx, v_start, v_end, vol_file = get_volume_info(current_page)
            
            # 如果进入新分卷，初始化对应的 Excel
            if current_vol_idx != vol_idx:
                if wb is not None:
                    # 保存并封存上一卷
                    wb.save(current_vol_file)
                    print(f"\n[分卷完成] 第 {current_vol_idx:02d} 卷已全部封存保存至: {current_vol_file}")
                
                current_vol_idx = vol_idx
                current_vol_file = vol_file
                wb, ws = init_excel_workbook(current_vol_file)
                vol_companies = 0
                vol_contacts = 0
                print(f"\n[开启新卷] 进入第 {current_vol_idx:02d} 卷 (页码范围: 第 {v_start} ~ {v_end} 页) -> {current_vol_file}")
            
            row_count = ws.max_row - 1
            
            # 请求当前页数据
            payload = BASE_SEARCH_PAYLOAD.copy()
            payload["pageIndex"] = current_page
            payload["pageSize"] = 10
            
            headers_gt8 = HEADERS.copy()
            headers_gt8["unionkey"] = "gt8"
            
            page_success = False
            for attempt in range(3):
                try:
                    resp = client.post(SEARCH_COMPANY_URL, headers=headers_gt8, json=payload, timeout=25.0)
                    data = resp.json()
                    if data.get("code") not in [200, 0]:
                        msg = str(data.get("msg", ""))
                        code = data.get("code")
                        print(f"\n[警告] 第 {current_page} 页接口返回代码 {code}: {msg}")
                        if code == 401 or any(k in msg.lower() for k in ["token", "login", "expire", "session"]) or "过期" in msg:
                            print("\n[登录态失效] 当前会话 Token 已过期，已为您安全落盘！请更新 Token 后继续。")
                            if wb:
                                wb.save(current_vol_file)
                                save_progress(current_page - 1, vol_idx, vol_companies, vol_contacts, total_comps)
                            return
                    companies = data.get("data", {}).get("result", [])
                    page_success = True
                    break
                except Exception as e:
                    print(f"\n[重试] 第 {current_page} 页网络请求异常 (尝试 {attempt+1}/3): {e}")
                    time.sleep(2)
            
            if not page_success:
                print(f"[跳过] 第 {current_page} 页多次请求失败，继续下一页...")
                current_page += 1
                continue
                
            if not companies:
                print(f"\n[提醒] 第 {current_page} 页已无更多企业数据，全量采集结束。")
                break
                
            for comp in companies:
                vol_companies += 1
                total_comps += 1
                cid = comp.get("id")
                cname = clean_text(comp.get("companyName"))
                ccountry = clean_text(comp.get("country"))
                ctax = clean_text(comp.get("companyTaxCode"))
                cweb = clean_text(comp.get("website"))
                clast_date = clean_text(comp.get("lastTradeDate"))
                ctrade_total = clean_text(comp.get("companyTradeTotal"))
                
                # 获取贸易画像（主营商品、HS编码、来源国、供应商）
                profile = fetch_company_profile(client, cid)
                p_products = profile.get("products", "")
                p_hscodes = profile.get("hs_codes", "")
                p_origins = profile.get("origin_countries", "")
                p_suppliers = profile.get("suppliers", "")
                
                # 尝试获取联系人（若有已缓存）
                contacts = fetch_company_contacts(client, cweb) if cweb else []
                
                if not contacts:
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
                    vol_contacts += len(contacts)
                    for c in contacts:
                        row_count += 1
                        fname = clean_text(c.get("fullName"))
                        jtitle = clean_text(c.get("jobTitle"))
                        emails = ", ".join([clean_text(e.get("value")) for e in c.get("emails", []) if e.get("value")])
                        links = ", ".join([clean_text(l.get("value")) for l in c.get("links", []) if l.get("value")])
                        is_purchase = "是" if (c.get("IsPurchase") or c.get("IsBoss")) else "否"
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
                
                time.sleep(0.08)
                
            # 计算当前卷内进度
            vol_progress = ((current_page - v_start + 1) / PAGES_PER_VOLUME) * 100
            print(f"[第{vol_idx:02d}卷 | {vol_progress:5.1f}%] 已完成第 {current_page} 页 (本卷: {current_page - v_start + 1}/{PAGES_PER_VOLUME}页) | 本卷企业: {vol_companies} 家 | 全局累计: {total_comps} 家", flush=True)
            
            # 定期保存
            if current_page % SAVE_INTERVAL == 0 or current_page == v_end:
                wb.save(current_vol_file)
                save_progress(current_page, vol_idx, vol_companies, vol_contacts, total_comps)
            
            current_page += 1
            
    if wb:
        wb.save(current_vol_file)
        save_progress(current_page - 1, current_vol_idx, vol_companies, vol_contacts, total_comps)
    print("\n" + "=" * 80)
    print(" [完成] 挂机采集任务圆满完成！")
    print("=" * 80)

if __name__ == "__main__":
    run_daemon()
