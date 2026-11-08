"""
特易资讯 (TopEase) - 墨西哥活跃采购商 (近1年交易>20次)
补爬漏采页码【第 359 ~ 377 页 / 190家企业】并合并入 351-400 页全景 Excel
"""

import os
import sys
import time
import json
import random
from urllib.parse import urlparse
import httpx
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

SEARCH_COMPANY_URL = "https://et.topease.net/api/gt8/search/search"
COMPANY_PROFILE_URL = "https://et.topease.net/api/gt8/search/getCompanyProfile"
SEARCH_KEY_URL = "https://et.topease.net/api/sou/search/searchCompanys"
CONTACTS_URL = "https://et.topease.net/api/sou/search/getCompanyContacts"

MISSING_PAGES = list(range(359, 378))  # 359 ~ 377 共 19 页
PAGE_SIZE = 10

OUTPUT_EXCEL = "墨西哥采购商_近1年高频采购_有效采购决策人_(351-400页_500家公司).xlsx"
PROGRESS_FILE = "crawl_progress_mexico_351_400.json"

AUTH_TOKEN = os.getenv("TOPEASE_TOKEN", "b4ec9b9d711fad7d64b8be8757176d70")

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
    "countryIds": [131],
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
    "countryId": 131,
    "historySearchId": "73953217"
}

COLUMNS_COUNT = 18

EXCLUDED_DOMAINS = {
    "linkedin.com", "lnkd.in", "facebook.com", "fb.com", "meta.com",
    "instagram.com", "twitter.com", "x.com", "youtube.com", "wikipedia.org",
    "creditsafe.com", "dnb.com", "zoominfo.com", "glassdoor.com", "zawya.com",
    "bloomberg.com", "crunchbase.com", "racius.com", "kompass.com",
    "yellowpages.com", "yelp.com", "reddit.com", "github.com"
}

EXCLUDED_EMAIL_DOMAINS = {
    "linkedin.com", "lnkd.in", "fb.com", "facebook.com", "meta.com",
    "instagram.com", "twitter.com", "x.com", "youtube.com", "wikipedia.org",
    "creditsafe.com", "dnb.com", "zoominfo.com", "glassdoor.com", "zawya.com",
    "bloomberg.com", "crunchbase.com", "racius.com", "kompass.com",
    "yellowpages.com", "yelp.com", "google.com", "gmail.com", "yahoo.com", "hotmail.com", "outlook.com",
    "reddit.com", "foundrybrands.com"
}

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def extract_clean_domain(domain_or_url: str) -> str:
    if not domain_or_url:
        return ""
    url = domain_or_url.strip()
    if "://" in url:
        parsed = urlparse(url)
        netloc = parsed.netloc or parsed.path
        return netloc.split(":")[0].strip()
    return url.split("/")[0].split(":")[0].strip()

def is_third_party_domain(domain_or_url: str) -> bool:
    if not domain_or_url:
        return False
    d = domain_or_url.lower().strip()
    for ex in EXCLUDED_DOMAINS:
        if ex in d:
            return True
    return False

def parse_contact_emails(emails_list):
    valid_emails = []
    status_tags = []
    for em in emails_list:
        val = clean_text(em.get("value"))
        if not val or "@" not in val:
            continue
            
        email_domain = val.split("@")[-1].lower().strip()
        if any(email_domain == exc or email_domain.endswith("." + exc) for exc in EXCLUDED_EMAIL_DOMAINS):
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

def fetch_company_profile(client: httpx.Client, company_id: str, retries=3):
    if not company_id:
        return {}
    headers_gt8 = HEADERS.copy()
    headers_gt8["unionkey"] = "gt8"
    for attempt in range(retries):
        try:
            resp = client.post(COMPANY_PROFILE_URL, headers=headers_gt8, json={"id": company_id}, timeout=30.0)
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

def fetch_valid_purchase_and_leader_contacts(client: httpx.Client, website: str, company_name: str, retries=3):
    resolved_web = website or ""
    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"
    
    clean_dom = extract_clean_domain(website)
    
    for attempt in range(retries):
        try:
            search_key = None
            
            if clean_dom and not is_third_party_domain(clean_dom):
                resp_sc = client.get(
                    SEARCH_KEY_URL,
                    params={"pageIndex": 1, "pageSize": 5, "linkType": 0, "country": 0, "keyword": clean_dom, "type": 2, "sourceType": 1},
                    headers=headers_yisou,
                    timeout=30.0
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
            
            if not search_key and company_name:
                resp_sc2 = client.get(
                    SEARCH_KEY_URL,
                    params={"pageIndex": 1, "pageSize": 10, "linkType": 0, "country": 0, "keyword": company_name, "type": 1, "sourceType": 1},
                    headers=headers_yisou,
                    timeout=30.0
                )
                if resp_sc2.status_code == 200:
                    sc_data2 = resp_sc2.json()
                    if sc_data2.get("code") == 1004:
                        print("\n[提示] 易搜 AI 搜索量已达今日账号限额，将仅保留贸易画像！")
                        return [], resolved_web
                    res_list2 = sc_data2.get("data", {}).get("result", [])
                    for res in res_list2:
                        dom = res.get("domain") or res.get("website") or ""
                        if not is_third_party_domain(dom) and res.get("searchKey"):
                            search_key = res.get("searchKey")
                            if not resolved_web:
                                resolved_web = clean_text(dom or res.get("companyWebsite") or "")
                            break
            
            if not search_key:
                return [], resolved_web
            
            time.sleep(random.uniform(0.3, 0.6))
            
            contact_params = {
                "status": 0,
                "pageIndex": 1,
                "pageSize": 50,
                "keyword": "",
                "fuzzyMatch": 2,
                "countryIds": "",
                "searchKey": search_key,
                "sourceType": 1,
                "statusList": "5,3",
                "excludeEmail": ""
            }
            resp_c = client.get(CONTACTS_URL, params=contact_params, headers=headers_yisou, timeout=30.0)
            if resp_c.status_code == 200:
                c_data = resp_c.json()
                raw_contacts = c_data.get("data", {}).get("result", [])
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
            time.sleep(1.0)
    return [], resolved_web

def reindex_excel(filepath):
    wb = openpyxl.load_workbook(filepath)
    ws = wb["采购与企业决策人"]
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=1, value=r - 1)
    wb.save(filepath)

def run():
    print("=" * 80)
    print(" [启动] 墨西哥采购商第 351 ~ 400 页漏采数据补全 (第 359 ~ 377 页，共 19 页 / 190 家企业)")
    print(f" 输出文件: {OUTPUT_EXCEL}")
    print("=" * 80)
    
    wb = openpyxl.load_workbook(OUTPUT_EXCEL)
    ws = wb["采购与企业决策人"]
    
    data_font = Font(name="Microsoft YaHei", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    row_count = ws.max_row - 1
    
    # 统计已有公司名称防止重复追加
    existing_companies = set()
    for r in range(2, ws.max_row + 1):
        cname = ws.cell(r, 2).value
        if cname:
            existing_companies.add(cname)
            
    print(f" 当前表中已有公司: {len(existing_companies)} 家，当前总行数: {row_count}")
    
    backfilled_companies = 0
    backfilled_contacts = 0
    start_time = time.time()
    
    with httpx.Client() as client:
        for idx, p in enumerate(MISSING_PAGES, 1):
            payload = SEARCH_PAYLOAD_TEMPLATE.copy()
            payload["pageIndex"] = p
            payload["pageSize"] = PAGE_SIZE
            
            headers_gt8 = HEADERS.copy()
            headers_gt8["unionkey"] = "gt8"
            
            companies = []
            for attempt in range(5):
                try:
                    resp = client.post(SEARCH_COMPANY_URL, headers=headers_gt8, json=payload, timeout=45.0)
                    data = resp.json()
                    code = data.get("code")
                    msg = str(data.get("msg", ""))
                    if code not in [200, 0]:
                        print(f"\n[警告] 第 {p} 页接口返回代码 {code}: {msg}")
                        if code == 401:
                            print("[停止] Token 已过期！")
                            wb.save(OUTPUT_EXCEL)
                            return
                    companies = data.get("data", {}).get("result", [])
                    if companies:
                        break
                    time.sleep(2.0)
                except Exception as e:
                    print(f"\n[重试] 第 {p} 页网络异常 (尝试 {attempt+1}/5): {e}")
                    time.sleep(3.0)
                    
            if not companies:
                print(f"[错误] 第 {p} 页多次重试后仍未获取到数据！")
                continue
                
            for comp in companies:
                cname = clean_text(comp.get("companyName"))
                if cname in existing_companies:
                    continue
                existing_companies.add(cname)
                backfilled_companies += 1
                
                cid = comp.get("id")
                ccountry = clean_text(comp.get("country"))
                ctax = clean_text(comp.get("companyTaxCode"))
                cweb = clean_text(comp.get("website"))
                clast_date = clean_text(comp.get("lastTradeDate"))
                ctrade_total = clean_text(comp.get("companyTradeTotal"))
                
                profile = fetch_company_profile(client, cid)
                p_products = profile.get("products", "")
                p_hscodes = profile.get("hs_codes", "")
                p_origins = profile.get("origin_countries", "")
                p_suppliers = profile.get("suppliers", "")
                
                time.sleep(random.uniform(0.3, 0.6))
                
                raw_contacts, resolved_web = fetch_valid_purchase_and_leader_contacts(client, cweb, cname)
                cweb = resolved_web or cweb
                valid_contacts_data = []
                for c in raw_contacts:
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
                    row_count += 1
                    row_data = [
                        row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                        p_products, p_hscodes, p_origins, p_suppliers,
                        "", "", "", "", "", "", ""
                    ]
                    ws.append(row_data)
                    for col_idx in range(1, COLUMNS_COUNT + 1):
                        cell = ws.cell(row=row_count + 1, column=col_idx)
                        cell.font = data_font
                        cell.border = thin_border
                        cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 14, 16, 18] else left_align
                else:
                    backfilled_contacts += len(valid_contacts_data)
                    for fname, jtitle, role_type, emails_str, status_str, links, update_d in valid_contacts_data:
                        row_count += 1
                        row_data = [
                            row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                            p_products, p_hscodes, p_origins, p_suppliers,
                            fname, jtitle, role_type, emails_str, status_str, links, update_d
                        ]
                        ws.append(row_data)
                        for col_idx in range(1, COLUMNS_COUNT + 1):
                            cell = ws.cell(row=row_count + 1, column=col_idx)
                            cell.font = data_font
                            cell.border = thin_border
                            cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 14, 16, 18] else left_align
                
                time.sleep(random.uniform(0.8, 1.5))
                
            print(f"[{idx:2d}/{len(MISSING_PAGES)}] 已补齐第 {p} 页 | 累计补齐企业: {backfilled_companies} 家 | 补齐决策人: {backfilled_contacts} 位", flush=True)
            wb.save(OUTPUT_EXCEL)
            time.sleep(random.uniform(1.5, 2.5))
            
    # 重新编序序号
    reindex_excel(OUTPUT_EXCEL)
    
    # 更新 json 进度文件
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "last_completed_page": 400,
            "total_companies": len(existing_companies),
            "total_contacts": 4644 + backfilled_contacts,
            "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
        }, f, indent=2, ensure_ascii=False)
        
    print("=" * 80)
    print(f" [完成] 补全全部完成！当前总企业数: {len(existing_companies)} 家，新增决策人: {backfilled_contacts} 位")
    print(f" 总耗时: {(time.time() - start_time) / 60:.2f} 分钟")
    print("=" * 80)

if __name__ == "__main__":
    run()
