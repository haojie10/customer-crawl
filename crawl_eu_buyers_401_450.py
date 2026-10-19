#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
欧盟高频采购商（近1年>20次交易）深度调研与联系人提取脚本 (第401~450页)
=============================================================================
针对特易资讯 (TopEase / 8号仓) 进行自动化深度检索与全景贸易画像清洗：
1. 筛选条件：欧盟 27 国 + 近 1 年交易频次 > 20 次 + 仅含中国采购经验 + 排除微型/货代/银行。
2. 贸易画像：最近交易日期、总交易单数、主营产品 Top3、HS Code Top3、原产国、合作供应商。
3. 联系人画像：
   - 提取 Leader（企业负责人/高管）与 Purchase（采购决策人）双标签联系人 (`statusList=1,5`)
   - 筛选 ALL 状态联系人 (`status=0`)
   - 单独展示每个联系人的邮箱验证状态：Valid (有效) / Risky (可能有效) / 未验证
4. 防控机制：严格 50 页单批次上限，拟人化动态延迟，分卷自动归档与断点续传。
"""

import os
import sys
import time
import json
import random
import re
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# 确保控制台 UTF-8 输出
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# API 端点
SEARCH_COMPANY_URL = "https://et.topease.net/api/gt8/search/search"
COMPANY_PROFILE_URL = "https://et.topease.net/api/gt8/search/getCompanyProfile"
SEARCH_KEY_URL = "https://et.topease.net/api/sou/search/searchCompanys"
CONTACTS_URL = "https://et.topease.net/api/sou/search/getCompanyContacts"

# 本次任务配置
START_PAGE = 401
END_PAGE = 450         # 第 401 ~ 450 页，共 500 家企业
PAGE_SIZE = 10         # 每页 10 家公司
SAVE_INTERVAL = 5      # 每 5 页保存一次

OUTPUT_EXCEL = "欧盟采购商_近1年高频采购_有效采购决策人_(401-450页_500家公司).xlsx"
PROGRESS_FILE = "crawl_progress_eu_401_450.json"

# 认证与请求头
AUTH_TOKEN = "22ff1382f77524bf76bf3e4c15f799f0"

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7",
    "authorization": AUTH_TOKEN,
    "cookie": f"topease_access_token={AUTH_TOKEN}",
    "language": "en",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=1&wlf=CompanyIndex",
    "unionkey": "gt8",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
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
    "historySearchId": "72965984"
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
    ("数据更新时间", 14)
]

def clean_str(val):
    """清洗非法 XML 字符"""
    if val is None:
        return ""
    s = str(val).strip()
    return ILLEGAL_CHARACTERS_RE.sub("", s)

def init_excel_workbook(file_path):
    """初始化或恢复已有的 Excel 工作簿"""
    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            return wb, ws
        except Exception as e:
            print(f"[警告] 打开已有工作簿异常: {e}，将新建文件")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "欧盟采购决策人"

    # 表头样式：暗蓝背景 + 白字加粗 + 居中
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    ws.row_dimensions[1].height = 26
    headers = [col[0] for col in COLUMNS]
    ws.append(headers)

    for col_idx, (col_name, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"
    wb.save(file_path)
    return wb, ws

def load_progress():
    """读取已完成进度"""
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "last_completed_page": 400,
        "total_companies": 0,
        "total_contacts": 0,
        "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
    }

def save_progress(page: int, companies: int, contacts: int):
    """保存断点进度"""
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "last_completed_page": page,
            "total_companies": companies,
            "total_contacts": contacts,
            "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
        }, f, ensure_ascii=False, indent=2)

def fetch_company_profile(client: httpx.Client, company_id: int):
    """提取采购商全景贸易画像"""
    headers = HEADERS.copy()
    headers["unionkey"] = "gt8"
    try:
        resp = client.post(COMPANY_PROFILE_URL, json={"id": company_id}, headers=headers, timeout=15.0)
        if resp.status_code == 200:
            data = resp.json()
            cards = data.get("data", {}).get("cards", [])
            profile = {
                "products": [],
                "hscodes": [],
                "origins": [],
                "suppliers": []
            }
            for card in cards:
                c_type = card.get("type")
                items = card.get("items", [])
                names = [item.get("name", "").strip() for item in items if item.get("name")]
                if c_type == "purchaseProduct":
                    profile["products"] = names[:3]
                elif c_type == "purchaseHsCode":
                    profile["hscodes"] = names[:3]
                elif c_type == "purchaseOriginCountry":
                    profile["origins"] = names[:3]
                elif c_type == "supplier":
                    profile["suppliers"] = names[:3]
            return profile
    except Exception:
        pass
    return {"products": [], "hscodes": [], "origins": [], "suppliers": []}

def get_email_verification_status(em_obj: dict) -> str:
    """根据邮箱对象判断验证状态"""
    v = em_obj.get("valid")
    g = em_obj.get("grade")
    if v is True or g == "A":
        return "Valid (有效)"
    elif v is False or (g and g in ["B", "C", "D", "F"]):
        return "Risky (可能有效)"
    return "未验证"

def fetch_leader_and_purchase_contacts(client: httpx.Client, website: str):
    """
    通过易搜 AI 搜索企业关键联系人：
    1. 提取 Leader 与 Purchase 两个标签的联系人 (`statusList=1,5`)
    2. 筛选 ALL 状态 (`status=0`)
    """
    if not website:
        return []
    
    clean_domain = website.replace("https://", "").replace("http://", "").split("/")[0].strip()
    if not clean_domain:
        return []

    headers_yisou = HEADERS.copy()
    headers_yisou["unionkey"] = "yisou3"

    for attempt in range(2):
        try:
            # 步骤 1: 根据官网域名获取 searchKey
            params_s = {
                "keyword": clean_domain,
                "type": 2,
                "sourceType": 1
            }
            resp_s = client.get(SEARCH_KEY_URL, params=params_s, headers=headers_yisou, timeout=15.0)
            if resp_s.status_code != 200:
                continue
            
            s_data = resp_s.json()
            if s_data.get("code") == 1004:
                print("\n[提示] 易搜 AI 搜索量已达今日账号限额，将仅保留贸易画像！\n", flush=True)
                return []
                
            result_list = s_data.get("data", {}).get("result", [])
            if not result_list:
                return []
            search_key = result_list[0].get("searchKey")
            if not search_key:
                return []
            
            # 拟人化微停顿
            time.sleep(random.uniform(0.3, 0.6))
            
            # 步骤 2: 精准请求 statusList='1,5' (Leader + Purchase 职位) & status=0 (ALL 筛选)
            contact_params = {
                "status": 0,          # ALL 验证状态
                "pageIndex": 1,
                "pageSize": 50,
                "keyword": "",
                "fuzzyMatch": 2,
                "countryIds": "",
                "searchKey": search_key,
                "sourceType": 1,
                "statusList": "1,5",  # Leader (1) + Purchase (5) 双标签
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
    last_page = prog.get("last_completed_page", 400)
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
    print(" 筛选规则: 欧盟27国 + 近1年交易>20次 + 含中国采购经验")
    print(" 联系人规则: Leader (高管/负责人) + Purchase (采购决策人) | 状态筛选: ALL (全状态) | 独立验证列")
    print(f" 防封保护: 拟人化随机延迟 (0.8~2.0s/企业) | 输出文件: {OUTPUT_EXCEL}")
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
                    resp = client.post(SEARCH_COMPANY_URL, json=payload, headers=headers_gt8, timeout=20.0)
                    if resp.status_code == 200:
                        res_json = resp.json()
                        if res_json.get("code") == 401 or "过期" in str(res_json.get("msg", "")):
                            print("\n[错误] Token 已过期或未授权 (Code 401)，正在安全保存数据并退出...")
                            wb.save(OUTPUT_EXCEL)
                            save_progress(p - 1, company_count, contact_count)
                            return
                        companies = res_json.get("data", {}).get("result", [])
                        break
                except Exception as e:
                    time.sleep(1.5)
            
            if not companies:
                print(f"[提示] 第 {p} 页未返回企业数据，跳过")
                time.sleep(1.5)
                continue
            
            for comp in companies:
                company_count += 1
                cid = comp.get("id")
                cname = clean_str(comp.get("companyName"))
                ccountry = clean_str(comp.get("country"))
                ctax = clean_str(comp.get("companyTaxCode"))
                cweb = clean_str(comp.get("website"))
                clast_date = clean_str(comp.get("lastTradeDate"))
                ctrade_total = clean_str(comp.get("companyTradeTotal"))
                
                # 1. 深度贸易画像
                profile = fetch_company_profile(client, cid)
                p_products = clean_str(", ".join(profile["products"]))
                p_hscodes = clean_str(", ".join(profile["hscodes"]))
                p_origins = clean_str(", ".join(profile["origins"]))
                p_suppliers = clean_str(", ".join(profile["suppliers"]))
                
                # 2. 深度联系人画像 (Leader + Purchase 标签，ALL 状态)
                contacts = fetch_leader_and_purchase_contacts(client, cweb)
                
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
                        cell.alignment = center_align if col_idx in [1, 3, 4, 6, 7, 14, 16, 18] else left_align
                else:
                    for contact in contacts:
                        contact_count += 1
                        row_count += 1
                        
                        fname = clean_str(contact.get("fullName"))
                        jtitle = clean_str(contact.get("jobTitle"))
                        
                        # 角色类型判断
                        is_boss = bool(contact.get("IsBoss"))
                        is_purchase = bool(contact.get("IsPurchase"))
                        if is_boss and is_purchase:
                            role_type = "采购 + 企业负责人 (Leader)"
                        elif is_boss:
                            role_type = "企业负责人/高管 (Leader)"
                        elif is_purchase:
                            role_type = "采购负责人 (Purchase)"
                        else:
                            role_type = "关键决策人 (Decision Maker)"
                        
                        # 邮箱提取与验证状态判定
                        raw_emails = contact.get("emails", [])
                        valid_emails = []
                        statuses = []
                        for em in raw_emails:
                            val = em.get("value")
                            if val and val.strip():
                                valid_emails.append(val.strip())
                                st = get_email_verification_status(em)
                                if st not in statuses:
                                    statuses.append(st)
                        
                        emails = clean_str(", ".join(valid_emails))
                        verify_status = clean_str(", ".join(statuses))
                        
                        # 社媒链接
                        links_list = [clean_str(l.get("url")) for l in contact.get("links", []) if l.get("url")]
                        links = clean_str(", ".join(links_list))
                        
                        # 更新时间取年月
                        update_d = clean_str(contact.get("updateDate", ""))
                        if len(update_d) >= 7:
                            update_d = update_d[:7]
                        
                        row_data = [
                            row_count, cname, ccountry, ctax, cweb, clast_date, ctrade_total,
                            p_products, p_hscodes, p_origins, p_suppliers,
                            fname, jtitle, role_type, emails, verify_status, links, update_d
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
            pct = (p / END_PAGE) * 100
            print(f"[{pct:5.1f}%] 已完成第 {p:3d}/{END_PAGE} 页 | 累计企业: {company_count} 家 | 获取决策人/联系人: {contact_count} 位", flush=True)
            
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
    print(f" 累计获取决策人/联系人: {contact_count} 位")
    print(f" 耗时: {total_mins:.2f} 分钟 | 输出文件: {os.path.abspath(OUTPUT_EXCEL)}")
    print("=" * 80)

if __name__ == "__main__":
    run()
