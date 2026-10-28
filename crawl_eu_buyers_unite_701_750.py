"""
特易资讯 (TopEase) - 欧盟活跃采购商 (近1年交易>20次) 全景画像与决策人采集工具
【优耐特账号专属脚本 - 防领英公域污染 & 双轨智能穿透版】
【第 701 ~ 750 页 / 500家企业】

核心防护升级：
1. 【公域与黄页域名黑名单】：内置 EXCLUDED_DOMAINS（LinkedIn、Facebook、Twitter、ZoomInfo 等），严禁将其作为官网！
2. 【智能级联域名反查】：
   - 优先使用海关数据自带的官网（若在黑名单内则舍弃）；
   - 若无官网或官网属于黑名单，以企业名称在易搜中遍历检索，跳过所有黑名单，只提取第一个独立企业官网及其 searchKey；
   - 若未匹配到独立官网，绝不拿领英等公域换取联系人，直接终止联系人请求，防止领英数千员工垃圾邮箱污染！
3. 【邮箱二次防御过滤】：联系人解析时，二次过滤 @linkedin.com、@facebook.com 等黑名单邮箱后缀；
4. 【双选与格式规范】：
   - 角色：【Purchase 采购负责人】+【Leader 企业高管】(statusList="5,3")
   - 邮箱：涵盖 Valid、Risky 及 Unverified
   - O 列【电子邮箱】：纯净邮箱，严禁拼接标签后缀
   - P 列【邮箱验证状态】：独立显示验证状态 (Valid (有效), Risky (可能有效), 未验证)
5. 【兜底与熔断保护】：即使个别企业无联系人，也完整保留企业贸易画像，保证 500 家企业 100% 收录；遇 1004 安全落盘。

防封控策略：
- 请求之间添加 0.8 ~ 1.5 秒随机延迟，拟人化访问节奏
- 每 5 页自动增量落盘保存一次 Excel 与进度
"""

import os
import sys
import time
import json
import random
from urllib.parse import urlparse
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

# 本次任务配置（优耐特账号 701-750页 修正版）
START_PAGE = 701
END_PAGE = 750         # 第 701 ~ 750 页，共 500 家企业
PAGE_SIZE = 10         # 每页 10 家公司，总计 500 家
SAVE_INTERVAL = 5      # 每 5 页保存一次

OUTPUT_EXCEL = "欧盟采购商_近1年高频采购_有效采购决策人_(701-750页_500家公司_优耐特账号).xlsx"
PROGRESS_FILE = "crawl_progress_eu_unite_701_750.json"

# 优耐特账号 认证与请求头
AUTH_TOKEN = os.getenv("TOPEASE_TOKEN", "32009feb5b4ed13d169ce30a6c92089d")

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
    "historySearchId": "73253680"
}

# 1. 公共公域/社交/黄页黑名单域名（严禁将其视为企业独立官网，防止串线拉取平台员工）
EXCLUDED_DOMAINS = {
    "linkedin.com", "lnkd.in", "facebook.com", "fb.com", "meta.com",
    "instagram.com", "twitter.com", "x.com", "youtube.com", "wikipedia.org",
    "creditsafe.com", "dnb.com", "zoominfo.com", "glassdoor.com", "zawya.com",
    "kompass.com", "bloomberg.com", "reuters.com", "crunchbase.com",
    "google.com", "baidu.com", "yahoo.com", "bing.com", "duckduckgo.com",
    "apple.com", "microsoft.com", "amazon.com", "tiktok.com"
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

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def clean_domain(url_or_domain: str) -> str:
    """提取纯净根主机名，去除 http://、路径及 www 前缀"""
    if not url_or_domain:
        return ""
    val = url_or_domain.strip().lower()
    if not val.startswith("http://") and not val.startswith("https://"):
        val = "http://" + val
    try:
        parsed = urlparse(val)
        host = parsed.netloc or parsed.path.split('/')[0]
        host = host.split(':')[0]
        if host.startswith("www."):
            host = host[4:]
        return host
    except Exception:
        return ""

def is_excluded_domain(domain_or_url: str) -> bool:
    """检查是否属于公域黑名单"""
    dom = clean_domain(domain_or_url)
    if not dom:
        return True
    return any(dom == exc or dom.endswith("." + exc) for exc in EXCLUDED_DOMAINS)

def parse_contact_emails(emails_list):
    """
    解析联系人的所有邮箱（含 Valid、Risky、Unverified），
    严密过滤 LinkedIn 等公域黑名单邮箱后缀。
    O列输出纯邮箱，P列输出对应的验证状态标签。
    返回: (纯邮箱文本, 状态文本)
    """
    clean_emails = []
    status_tags = []
    for em in emails_list:
        val = clean_text(em.get("value"))
        if not val or "@" not in val:
            continue
        e_lower = val.lower()
        # 邮箱后缀防伪过滤：严禁 linkedin.com 等公域邮箱
        e_domain = e_lower.split("@")[-1].strip()
        if is_excluded_domain(e_domain):
            continue
            
        grade = str(em.get("grade", "")).strip().upper()
        valid = em.get("valid")
        
        if grade == "A" or valid is True:
            tag_cn = "Valid (有效)"
        elif grade == "B" or valid is False:
            tag_cn = "Risky (可能有效)"
        else:
            tag_cn = "未验证"
            
        clean_emails.append(val)
        if tag_cn not in status_tags:
            status_tags.append(tag_cn)
            
    email_str = ", ".join(clean_emails)
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
    ws.title = "欧盟采购与企业决策人"
    
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

def resolve_independent_domain_and_key(client: httpx.Client, headers_yisou: dict, company_name: str, raw_website: str):
    """
    智能解析目标企业的独立官网，彻底跳过 LinkedIn 等公域黑名单。
    返回: (resolved_website, search_key, is_quota_exceeded)
    """
    resolved_web = ""
    search_key = ""
    
    # 策略 A: 优先使用海关数据自带的官网（必须非黑名单）
    dom_candidate = clean_domain(raw_website)
    if dom_candidate and not is_excluded_domain(dom_candidate):
        try:
            resp_sc = client.get(
                SEARCH_KEY_URL,
                params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": raw_website, "type": 2, "sourceType": 1},
                headers=headers_yisou,
                timeout=15.0
            )
            if resp_sc.status_code == 200:
                sc_data = resp_sc.json()
                if sc_data.get("code") == 1004:
                    return "", "", True
                res_list = sc_data.get("data", {}).get("result", [])
                for r in res_list:
                    cand = r.get("domain") or clean_domain(r.get("url")) or ""
                    if cand and not is_excluded_domain(cand):
                        resolved_web = r.get("url") or raw_website
                        search_key = r.get("searchKey")
                        if search_key:
                            break
        except Exception:
            pass

    # 策略 B: 若无官网或官网为公域/未匹配成功，通过易搜搜索公司名称反查独立官网
    if not search_key and company_name:
        for search_type in [1, 2]:  # 先查精确(type=1)，再查模糊(type=2)
            try:
                resp_sc = client.get(
                    SEARCH_KEY_URL,
                    params={"pageIndex": 1, "pageSize": 10, "linkType": 0, "country": 0, "keyword": company_name, "type": search_type, "sourceType": 1},
                    headers=headers_yisou,
                    timeout=20.0
                )
                if resp_sc.status_code == 200:
                    sc_data = resp_sc.json()
                    if sc_data.get("code") == 1004:
                        return "", "", True
                    results = sc_data.get("data", {}).get("result", [])
                    for res in results:
                        candidate_domain = res.get("domain") or clean_domain(res.get("url")) or ""
                        # 核心防线：必须不是 LinkedIn 等三方公域！
                        if candidate_domain and not is_excluded_domain(candidate_domain):
                            resolved_web = res.get("url") or candidate_domain
                            search_key = res.get("searchKey")
                            if search_key:
                                break
                if search_key:
                    break
            except Exception:
                pass

    return resolved_web, search_key, False

def fetch_clean_contacts(client: httpx.Client, company_name: str, raw_website: str, retries=2):
    """
    智能获取采购与高管决策人：
    1. 严格过滤 LinkedIn 等公域域名，只以独立官网 searchKey 查询；
    2. 严格做邮箱黑名单二次防伪过滤。
    返回: (final_website, contacts_list, is_quota_exceeded)
    """
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    
    resolved_web, search_key, is_quota = resolve_independent_domain_and_key(client, headers_yisou, company_name, raw_website)
    if is_quota:
        return raw_website, [], True
    if not search_key:
        return resolved_web or raw_website, [], False
        
    time.sleep(random.uniform(0.3, 0.6))
    
    # 策略 C: 使用独立官网 searchKey 提取决策人，并进行双选与二次邮箱防伪
    for attempt in range(retries):
        try:
            contact_params = {
                "status": 0,          # 0=获取全部邮箱(Valid+Risky+Unverified)
                "pageIndex": 1,
                "pageSize": 50,
                "keyword": "",
                "fuzzyMatch": 2,
                "countryIds": "",
                "searchKey": search_key,
                "sourceType": 1,
                "statusList": "5,3",  # 5=Purchase采购负责人, 3=Leader企业高管
                "excludeEmail": ""
            }
            resp_c = client.get(CONTACTS_URL, params=contact_params, headers=headers_yisou, timeout=20.0)
            if resp_c.status_code == 200:
                c_data = resp_c.json()
                if c_data.get("code") == 1004:
                    return resolved_web or raw_website, [], True
                items = c_data.get("data", {}).get("result", [])
                
                # 去重与二次邮箱防伪
                unique_contacts = []
                seen_keys = set()
                for c in items:
                    is_purchase = bool(c.get("IsPurchase"))
                    is_boss = bool(c.get("IsBoss"))
                    if not (is_purchase or is_boss):
                        continue
                        
                    emails_str, status_str = parse_contact_emails(c.get("emails", []))
                    
                    ckey = str(c.get("id") or c.get("fullName") or "")
                    if ckey and ckey in seen_keys:
                        continue
                    seen_keys.add(ckey)
                    
                    c["_clean_emails_str"] = emails_str
                    c["_clean_status_str"] = status_str
                    unique_contacts.append(c)
                return resolved_web or raw_website, unique_contacts, False
        except Exception:
            time.sleep(0.8)
            
    return resolved_web or raw_website, [], False

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
    print(" 【优耐特账号脚本启动 - 防领英污染 & 智能级联反查版】 欧盟高频采购商修正采集")
    print(f" 采集范围: 第 {current_start} 页 ~ 第 {END_PAGE} 页 (总计 {total_pages_in_batch} 页 / 约 {total_pages_in_batch * PAGE_SIZE} 家企业)")
    print(" 核心防御: 彻底剔除 LinkedIn/FB/Twitter 等公域垃圾域名与邮箱，仅锁定企业真实独立官网！")
    print(" 级联机制: 列表官网优先 -> 企业名精准反查 -> 企业名模糊反查 (自动跳过公域黑名单)")
    print(" 决策人角色: 【Purchase 采购负责人】+【Leader 企业高管】(statusList=5,3)")
    print(" 格式规范: O列纯邮箱，P列单独显示验证状态")
    print(f" 防封保护: 拟人化随机延迟 (0.8~1.5s/企业) | 输出文件: {OUTPUT_EXCEL}")
    print("=" * 80)
    
    start_time = time.time()
    quota_exceeded = False
    
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
                raw_cweb = clean_text(comp.get("website"))
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
                
                # 3. 智能级联反查独立官网 + 采购高管联系人（黑名单严密防御）
                final_web = raw_cweb
                raw_contacts = []
                if not quota_exceeded:
                    final_web, raw_contacts, is_quota = fetch_clean_contacts(client, cname, raw_cweb)
                    if is_quota:
                        quota_exceeded = True
                        print("\n[提示] 优耐特账号 易搜 AI 搜索量已达今日账号限额，将仅保留海关贸易画像！")
                
                if not raw_contacts:
                    # 无联系人，完整保留公司画像行
                    row_count += 1
                    row_data = [
                        row_count, cname, ccountry, ctax, final_web, clast_date, ctrade_total,
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
                    contact_count += len(raw_contacts)
                    for c in raw_contacts:
                        row_count += 1
                        fname = clean_text(c.get("fullName"))
                        jtitle = clean_text(c.get("jobTitle"))
                        role_type = parse_contact_role(c)
                        
                        emails_str = c.get("_clean_emails_str", "")
                        status_str = c.get("_clean_status_str", "")
                        
                        links = ", ".join([clean_text(l.get("value")) for l in c.get("links", []) if l.get("value")])
                        update_d = clean_text(c.get("updateDate"))
                        
                        row_data = [
                            row_count, cname, ccountry, ctax, final_web, clast_date, ctrade_total,
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
            print(f"[{pct:5.1f}%] 已完成第 {p:3d}/{END_PAGE} 页 (本批第 {completed_count:2d}/{total_pages_in_batch} 页) | 累计企业: {company_count} 家 | 真实采购与高管决策人: {contact_count} 位", flush=True)
            
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
    print(f" 【优耐特账号】第 {START_PAGE} ~ {END_PAGE} 页 (共 {company_count} 家企业) 深度修正采集全部完成！")
    print(f" 累计获取真实采购与高管决策人: {contact_count} 位")
    print(f" 耗时: {total_mins:.2f} 分钟 | 输出文件: {os.path.abspath(OUTPUT_EXCEL)}")
    print("=" * 80)

if __name__ == "__main__":
    run()
