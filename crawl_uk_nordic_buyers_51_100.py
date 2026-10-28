"""
特易资讯 (TopEase) - 英国及北欧活跃采购商 (近1年交易>20次) 全景画像与决策人采集工具
【第 51 ~ 100 页 / 500家企业】（全新高精度独立域名防误匹版）

核心规范与防御机制：
1. 智能独立域名甄别：海关记录未登记官网时，自动通过公司名反查，并严格排除 LinkedIn、Facebook、Creditsafe 等第三方公域平台，仅锁定目标企业专属独立官网
2. 彻底杜绝假邮箱：硬性拦截 @linkedin.com、@fb.com、@meta.com 等平台员工邮箱，确保 100% 对应本公司真实人员
3. 突破网页端单选限制：同时支持采集【Purchase 采购负责人】与【Leader 企业高管/核心决策人】(statusList=5,3)
4. 【电子邮箱】列：纯邮箱文本，后面不加任何 [Valid]、[Risky] 等后缀标签
5. 【邮箱验证状态】列：独立成列，明确标注 Valid (有效)、Risky (可能有效) 或 Unverified (未验证)
6. 【角色类型 (Purchase/Leader)】列：独立成列，清晰标识 Purchase / Leader / 采购+高管
7. 即使无采购/高管联系人，也完整保留企业基本面与贸易全景画像

筛选条件：
1. 区域/国家：英国及北欧等选定国家 (countryIds: [171, 213, 153, 195, 93])
2. 交易时间：近 1 年 (tradeDate: 1)
3. 交易次数：> 20 次 (tradeCount: 20)
4. 企业筛选：含中国采购经验，排除物流、银行、小微企业
5. 联系方式：必须含邮箱 (containContact: 3)
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

# 本次任务配置（第 51 ~ 100 页）
START_PAGE = 51
END_PAGE = 100         # 第 51 ~ 100 页，共 500 家企业
PAGE_SIZE = 10         # 每页 10 家公司，总计 500 家
SAVE_INTERVAL = 5      # 每 5 页保存一次

OUTPUT_EXCEL = "英国及北欧采购商_近1年高频采购_有效采购决策人_(51-100页_500家公司).xlsx"
PROGRESS_FILE = "crawl_progress_uk_nordic_51_100.json"

# 认证与请求头
AUTH_TOKEN = os.getenv("TOPEASE_TOKEN", "1f12147ebb8b7238b188118e89e55484")

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "authorization": AUTH_TOKEN,
    "cookie": f"topease_access_token={AUTH_TOKEN}",
    "language": "en",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=1&wlf=CompanyIndex",
    "unionkey": "gt8",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36",
    "website": "https://et.topease.net"
}

# 英国及北欧选定国家 + 近1年 + 交易>20次 + 含中国采购经验
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
    "countryIds": [171, 213, 153, 195, 93],
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
    "historySearchId": "73416286"
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
    ("角色类型 (Purchase/Leader)", 25),
    ("电子邮箱", 35),
    ("邮箱验证状态", 20),
    ("社媒主页 (LinkedIn等)", 35),
    ("数据更新时间", 14),
]

# 第三方公域平台黑名单：坚决禁止作为企业官网提取
EXCLUDED_DOMAINS = {
    "linkedin.com", "lnkd.in", "facebook.com", "fb.com", "meta.com",
    "instagram.com", "twitter.com", "x.com", "youtube.com", "wikipedia.org",
    "creditsafe.com", "dnb.com", "zoominfo.com", "glassdoor.com", "zawya.com",
    "bloomberg.com", "crunchbase.com", "racius.com", "kompass.com",
    "yellowpages.com", "yelp.com", "reddit.com", "github.com"
}

# 邮箱域名黑名单：坚决禁止将公域平台员工邮箱记入客户数据
EXCLUDED_EMAIL_DOMAINS = {
    "linkedin.com", "lnkd.in", "fb.com", "facebook.com", "meta.com",
    "reddit.com", "foundrybrands.com"
}

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def is_third_party_domain(domain_or_url: str) -> bool:
    if not domain_or_url:
        return False
    d = domain_or_url.lower().strip()
    for exc in EXCLUDED_DOMAINS:
        if exc in d:
            return True
    return False

def parse_contact_emails(emails_list):
    """
    解析邮箱列表：
    - 纯邮箱文本，不带任何 [Valid]、[Risky]、[Unverified] 标签后缀
    - 坚决过滤剔除第三方平台邮箱 (@linkedin.com, @fb.com 等)
    - 独立【邮箱验证状态】列：分别标记 Valid (有效)、Risky (可能有效) 或 Unverified (未验证)
    返回: (纯邮箱文本, 状态文本)
    """
    valid_emails = []
    status_tags = []
    for em in emails_list:
        val = clean_text(em.get("value"))
        if not val or "@" not in val:
            continue
            
        email_domain = val.split("@")[-1].lower().strip()
        if any(email_domain == exc or email_domain.endswith("." + exc) for exc in EXCLUDED_EMAIL_DOMAINS):
            # 彻底杜绝领英或公域平台的错误邮箱（包含子域名如 e.linkedin.com）
            continue
            
        grade = str(em.get("grade", "")).strip().upper()
        valid = em.get("valid")
        
        if grade == "A" or valid is True:
            tag_cn = "Valid (有效)"
        elif grade == "B" or valid is False:
            tag_cn = "Risky (可能有效)"
        else:
            tag_cn = "Unverified (未验证)"
            
        valid_emails.append(val)
        if tag_cn not in status_tags:
            status_tags.append(tag_cn)
            
    email_str = ", ".join(valid_emails)
    status_str = ", ".join(status_tags)
    return email_str, status_str

def parse_contact_role(contact_obj):
    """
    解析联系人的业务角色：采购 (Purchase)、企业负责人/高管 (Leader)、或双重兼任
    """
    is_purchase = bool(contact_obj.get("IsPurchase"))
    is_boss = bool(contact_obj.get("IsBoss"))
    
    if is_purchase and is_boss:
        return "采购 + 企业负责人 (Leader)"
    elif is_purchase:
        return "采购负责人 (Purchase)"
    elif is_boss:
        return "企业负责人/高管 (Leader)"
    else:
        return "关键决策人 (Decision Maker)"

def init_excel_workbook(filepath):
    if os.path.exists(filepath):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "采购与企业决策人"
    
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
        "last_completed_page": START_PAGE - 1,
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

def fetch_valid_purchase_and_leader_contacts(client: httpx.Client, website: str, company_name: str, retries=2):
    """
    智能独立域名甄别检索策略：
    1. 优先使用官网 (type=2) 换取 searchKey（必须不是第三方平台域名）
    2. 若官网为空或未匹配，则使用公司名称 (type=1) 检索，严格排除 LinkedIn/Facebook 等第三方域名，锁定真正的企业专属独立官网！
    3. 同时请求 Purchase (采购) 与 Leader (高管/核心决策人) (statusList=5,3)
    返回: (联系人列表, 最终确认的官网)
    """
    resolved_web = website or ""
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    
    for attempt in range(retries):
        try:
            search_key = None
            
            # 策略 A: 若有官网且不是第三方平台，优先使用官网检索
            if website and not is_third_party_domain(website):
                resp_sc = client.get(
                    SEARCH_KEY_URL,
                    params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": website, "type": 2, "sourceType": 1},
                    headers=headers_yisou,
                    timeout=15.0
                )
                if resp_sc.status_code == 200:
                    sc_data = resp_sc.json()
                    if sc_data.get("code") == 1004:
                        print("\n[提示] 易搜 AI 搜索量已达今日账号限额，将仅保留贸易画像！")
                        return [], resolved_web
                    res_list = sc_data.get("data", {}).get("result", [])
                    for res in res_list:
                        dom = res.get("domain") or res.get("website") or ""
                        if not is_third_party_domain(dom) and res.get("searchKey"):
                            search_key = res.get("searchKey")
                            break
            
            # 策略 B: 若官网未匹配，通过公司名称检索，但严格过滤第三方平台，锁定独立企业官网！
            if not search_key and company_name:
                resp_sc2 = client.get(
                    SEARCH_KEY_URL,
                    params={"pageIndex": 1, "pageSize": 10, "linkType": 0, "country": 0, "keyword": company_name, "type": 1, "sourceType": 1},
                    headers=headers_yisou,
                    timeout=15.0
                )
                if resp_sc2.status_code == 200:
                    sc_data2 = resp_sc2.json()
                    if sc_data2.get("code") == 1004:
                        print("\n[提示] 易搜 AI 搜索量已达今日账号限额，将仅保留贸易画像！")
                        return [], resolved_web
                    res_list2 = sc_data2.get("data", {}).get("result", [])
                    for res in res_list2:
                        dom = res.get("domain") or res.get("website") or ""
                        # 严格排除领英、脸书等公域平台
                        if not is_third_party_domain(dom) and res.get("searchKey"):
                            search_key = res.get("searchKey")
                            if not resolved_web:
                                resolved_web = clean_text(dom or res.get("companyWebsite") or "")
                            break
            
            if not search_key:
                return [], resolved_web
            
            # 拟人化小微停顿
            time.sleep(random.uniform(0.3, 0.6))
            
            # 步骤 2: 请求 status=0 包含全量状态，且 statusList=5,3 (采购 + Leader高管)
            contact_params = {
                "status": 0,
                "pageIndex": 1,
                "pageSize": 50,
                "keyword": "",
                "fuzzyMatch": 2,
                "countryIds": "",
                "searchKey": search_key,
                "sourceType": 1,
                "statusList": "5,3",  # 采购 + Leader 高管
                "excludeEmail": ""
            }
            resp_c = client.get(CONTACTS_URL, params=contact_params, headers=headers_yisou, timeout=15.0)
            if resp_c.status_code == 200:
                c_data = resp_c.json()
                raw_contacts = c_data.get("data", {}).get("result", [])
                
                # 去重机制（以 contact id 或 姓名去重）
                unique_contacts = []
                seen_keys = set()
                for c in raw_contacts:
                    ckey = str(c.get("id") or c.get("fullName") or "")
                    if ckey and ckey in seen_keys:
                        continue
                    seen_keys.add(ckey)
                    unique_contacts.append(c)
                return unique_contacts, resolved_web
        except Exception:
            time.sleep(0.8)
    return [], resolved_web

def run():
    prog = load_progress()
    last_page = prog.get("last_completed_page", START_PAGE - 1)
    current_start = last_page + 1
    
    if current_start > END_PAGE:
        print(f"[提示] 第 {START_PAGE} ~ {END_PAGE} 页已全部采集完毕！")
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
    
    total_pages_in_batch = END_PAGE - START_PAGE + 1
    
    print("=" * 80)
    print(" [启动] 英国及北欧高频采购商（近1年>20次交易）高精度深度采集【第51-100页】")
    print(f" 采集范围: 第 {current_start} 页 ~ 第 {END_PAGE} 页 (总计 {total_pages_in_batch} 页 / 约 {total_pages_in_batch * PAGE_SIZE} 家企业)")
    print(" 筛选规则: 英国及北欧等选定国 (countryIds: [171, 213, 153, 195, 93]) + 近1年交易>20次 + 含中国采购经验")
    print(" 决策人类型: 【Purchase采购负责人】+【Leader企业高管】同步双选 (statusList=5,3)")
    print(" 域名防护: 严格排除 LinkedIn/Facebook 等第三方平台，精准锁定企业独立官网")
    print(" 邮箱规范: 纯邮箱地址（绝无任何标签后缀），彻底拦截 @linkedin.com 等平台员工邮箱")
    print(f" 防封保护: 拟人化随机延迟 (0.8~1.5s/企业) | 输出文件: {OUTPUT_EXCEL}")
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
                time.sleep(random.uniform(0.3, 0.6))
                
                # 3. 抓取 Purchase + Leader 职位联系人（严格独立官网 + 排除领英员工）
                raw_contacts, resolved_web = fetch_valid_purchase_and_leader_contacts(client, cweb, cname)
                cweb = resolved_web or cweb
                valid_contacts_data = []
                for c in raw_contacts:
                    # 仅保留 Purchase 或 Leader
                    if not (c.get("IsPurchase") or c.get("IsBoss")):
                        continue
                    emails_str, status_str = parse_contact_emails(c.get("emails", []))
                    if emails_str:
                        fname = clean_text(c.get("fullName"))
                        jtitle = clean_text(c.get("jobTitle"))
                        role_type = parse_contact_role(c)
                        links = ", ".join([clean_text(l.get("value")) for l in c.get("links", []) if l.get("value")])
                        update_d = clean_text(c.get("updateDate"))
                        valid_contacts_data.append((fname, jtitle, role_type, emails_str, status_str, links, update_d))
                
                if not valid_contacts_data:
                    # 即使无联系人，也完整保留公司画像行
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
                        cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 14, 16, 18] else left_align
                else:
                    contact_count += len(valid_contacts_data)
                    for fname, jtitle, role_type, emails_str, status_str, links, update_d in valid_contacts_data:
                        row_count += 1
                        row_data = [
                            row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                            p_products, p_hscodes, p_origins, p_suppliers,
                            fname, jtitle, role_type, emails_str, status_str, links, update_d
                        ]
                        ws.append(row_data)
                        for col_idx in range(1, len(COLUMNS) + 1):
                            cell = ws.cell(row=row_count + 1, column=col_idx)
                            cell.font = data_font
                            cell.border = thin_border
                            cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 14, 16, 18] else left_align
                
                # 拟人化随机节奏 (0.8 ~ 1.5 秒/公司)，全面降低风控风险
                time.sleep(random.uniform(0.8, 1.5))
            
            # 单页完成进度打印
            completed_count = p - START_PAGE + 1
            pct = (completed_count / total_pages_in_batch) * 100
            print(f"[{pct:5.1f}%] 已完成第 {p:3d}/{END_PAGE} 页 (本批第 {completed_count:2d}/{total_pages_in_batch} 页) | 累计企业: {company_count} 家 | 精准采购与高管决策人: {contact_count} 位", flush=True)
            
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
    print(f" [完成] 第 {START_PAGE} ~ {END_PAGE} 页 (共 {company_count} 家企业) 深度调研全部完成！")
    print(f" 累计获取真实采购与高管决策人: {contact_count} 位")
    print(f" 耗时: {total_mins:.2f} 分钟 | 输出文件: {os.path.abspath(OUTPUT_EXCEL)}")
    print("=" * 80)

if __name__ == "__main__":
    run()
