"""
广交会买家清单 (Canton Fair Buyers) - 全量多卷自动采集系统
特性：
1. 自动从第 2 卷运行至第 18 卷（全量 173 万条数据）
2. 每卷 1,000 页（10 万条）自动生成独立 Excel，并定期存盘（每 50 页存盘）
3. 严格的登录态/Token 失效监测：一旦检测到超时退出或鉴权失败，立即安全保存已爬数据并暂停
4. 完善的断点续传（通过 crawl_progress_all.json 记录精确页码）
5. 全面兼容 Windows 控制台编码
"""

import os
import sys
import time
import json
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

API_URL = "https://et.topease.net/api/gt8/search/getCantonFairPage"
PROGRESS_FILE = "crawl_progress_all.json"

TOTAL_PAGES = 17313  # 总共 17,313 页 (约 1,731,267 条)
PAGE_SIZE = 100
PAGES_PER_VOLUME = 1000  # 每卷 1000 页 (10 万条)
SAVE_INTERVAL = 50  # 每 50 页自动落盘一次

HEADERS = {
    "authorization": "797cfb43125f309cbc8a828084d5aab8",
    "cookie": "topease_access_token=797cfb43125f309cbc8a828084d5aab8",
    "language": "en",
    "origin": "https://et.topease.net",
    "referer": "https://et.topease.net/gt8/company/result?where=search&tabCheck=9&wlf=CompanyIndex",
    "unionkey": "gt8",
    "content-type": "application/json",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36",
    "website": "https://et.topease.net"
}

BASE_PAYLOAD = {
    "autoTranslate": True,
    "buyFromArea": 0,
    "buyerCompanys": [],
    "companyNames": [],
    "companyNumbers": [],
    "companyWebsites": [],
    "containContact": "",
    "countryIds": [0],
    "dataSourceCountryId": 43,
    "excludeBank": False,
    "excludeLogistics": False,
    "excludeMicroAgent": False,
    "fullAddressMatchs": [],
    "hsCodes": [],
    "industryCategorys": [],
    "industryTyCodes": [],
    "industryType": "",
    "linkTypes": [],
    "metric": "",
    "pageSize": PAGE_SIZE,
    "paixunt": False,
    "productDescriptions": [],
    "productTypes": [],
    "searchId": "",
    "showContactIcon": False,
    "sort": "score",
    "sortChange": False,
    "sortType": "desc",
    "supplierCompanys": [],
    "surnames": [],
    "tradeCount": -1,
    "tradeDate": 3,
    "tradeShowNameAndSessions": [],
    "tradeShowNames": [],
    "type": 9
}

COLUMNS = [
    ("#", 8),
    ("Company", 35),
    ("Nation", 18),
    ("Exhibition Session", 20),
    ("Industry", 28),
    ("Product", 25),
    ("contacts", 22),
    ("Tel", 20),
    ("Fax", 20),
    ("Address", 40),
    ("Zip Code", 14),
    ("Email", 30),
    ("Website", 30),
]

import re
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

def clean_text(val):
    if val is None:
        return ""
    s = str(val).strip()
    s = ILLEGAL_CHARACTERS_RE.sub("", s)
    return s

def get_volume_info(page: int):
    """根据页码计算对应的分卷信息"""
    vol_num = ((page - 1) // PAGES_PER_VOLUME) + 1
    vol_start_page = (vol_num - 1) * PAGES_PER_VOLUME + 1
    vol_end_page = min(vol_num * PAGES_PER_VOLUME, TOTAL_PAGES)
    vol_filename = f"广交会买家_第{vol_num:02d}卷_({vol_start_page}-{vol_end_page}页).xlsx"
    return vol_num, vol_start_page, vol_end_page, vol_filename

def init_excel_workbook(filepath):
    """初始化或加载分卷 Excel 文件"""
    if os.path.exists(filepath):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        return wb, ws
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Canton Fair Buyers"
    
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
    # 默认从第 1000 页已完成开始（即从 1001 页开始）
    return {"last_completed_page": 1000, "total_records": 100000}

def save_progress(page: int, total_records: int):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "last_completed_page": page,
            "total_records": total_records,
            "update_time": time.strftime("%Y-%m-%d %H:%M:%S")
        }, f, indent=2, ensure_ascii=False)

def fetch_page(client: httpx.Client, page_index: int, retry_count: int = 3):
    """
    抓取单页数据
    返回: (status_code, items, is_auth_error, error_msg)
    """
    payload = BASE_PAYLOAD.copy()
    payload["pageIndex"] = page_index
    
    for attempt in range(retry_count):
        try:
            resp = client.post(API_URL, headers=HEADERS, json=payload, timeout=25.0)
            if resp.status_code == 200:
                data = resp.json()
                code = data.get("code")
                msg = str(data.get("msg", "")).lower()
                
                # 检查是否返回登录超时或 Token 失效
                if any(k in msg for k in ["token", "auth", "login", "expired", "expire", "未登录", "登录超时", "无权限"]):
                    return 200, [], True, f"登录态已过期: {data.get('msg')}"
                
                if code == 200 or code == 0 or "data" in data:
                    res_obj = data.get("data", {})
                    items = res_obj.get("result", [])
                    return 200, items, False, ""
                else:
                    return 200, [], False, f"接口返回异常代码 {code}: {data.get('msg')}"
                    
            elif resp.status_code in [401, 403]:
                return resp.status_code, [], True, f"HTTP {resp.status_code} 认证失败/登录过期"
            else:
                print(f"[重试] 第 {page_index} 页 HTTP {resp.status_code}，正在重试 ({attempt+1}/{retry_count})...")
        except Exception as e:
            print(f"[重试] 第 {page_index} 页网络异常: {e}，正在重试 ({attempt+1}/{retry_count})...")
        time.sleep(1.5)
        
    return 500, [], False, "多次重试失败"

def run_all_volumes():
    prog = load_progress()
    last_completed = prog.get("last_completed_page", 1000)
    total_records = prog.get("total_records", last_completed * 100)
    
    start_page = last_completed + 1
    if start_page > TOTAL_PAGES:
        print(f"[提示] 所有 17,313 页数据已全部采集完毕！")
        return
    
    print("=" * 70)
    print(f" 广交会买家数据多卷全自动采集系统启动")
    print(f" 起始页码: 第 {start_page} 页 | 目标总页数: {TOTAL_PAGES} 页")
    print(f" 当前已采集: {total_records} 条数据")
    print(f" 登录超时监控: 启用 (一旦检测到登录失效立即安全落盘退出)")
    print("=" * 70)
    
    data_font = Font(name="Microsoft YaHei", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    current_vol_num = None
    current_wb = None
    current_ws = None
    current_vol_filename = None
    
    start_time = time.time()
    
    with httpx.Client() as client:
        for p in range(start_page, TOTAL_PAGES + 1):
            vol_num, vol_start, vol_end, vol_filename = get_volume_info(p)
            
            # 切换分卷文件
            if vol_num != current_vol_num:
                if current_wb and current_vol_filename:
                    current_wb.save(current_vol_filename)
                    print(f"\n[分卷完成] 第 {current_vol_num:02d} 卷已完整保存至: {current_vol_filename}")
                
                current_vol_num = vol_num
                current_vol_filename = vol_filename
                current_wb, current_ws = init_excel_workbook(vol_filename)
                print(f"\n[进入分卷] 正在采集第 {vol_num:02d} 卷 (第 {vol_start} ~ {vol_end} 页) -> {vol_filename}")
            
            # 请求当前页
            status, records, is_auth_err, err_msg = fetch_page(client, p)
            
            # 登录超时安全退出判定
            if is_auth_err:
                print("\n" + "!" * 70)
                print(f"[安全停止] 检测到网站登录态已失效或超时退出！")
                print(f"原因: {err_msg}")
                print(f"已停止在: 第 {p} 页 (第 {vol_num:02d} 卷)")
                if current_wb and current_vol_filename:
                    current_wb.save(current_vol_filename)
                    save_progress(p - 1, total_records)
                    print(f"已将当前分卷数据安全保存至: {current_vol_filename}")
                    print(f"进度已记录在: {PROGRESS_FILE}")
                print(f"请您在浏览器中重新登录后，提供新的 Token，程序将无缝继续采集！")
                print("!" * 70)
                return
            
            if not records:
                print(f"\n[跳过] 第 {p} 页未获取到有效记录: {err_msg}")
                continue
            
            row_index = current_ws.max_row
            for item in records:
                row_index += 1
                total_records += 1
                
                row_data = [
                    total_records,
                    clean_text(item.get("companyName")),
                    clean_text(item.get("companyCountry")),
                    clean_text(item.get("numberSessions")),
                    clean_text(item.get("companyIndustry")),
                    clean_text(item.get("purchaseProduct")),
                    clean_text(item.get("contacts")),
                    clean_text(item.get("tel")),
                    clean_text(item.get("fax")),
                    clean_text(item.get("address")),
                    clean_text(item.get("zip")),
                    clean_text(item.get("email")),
                    clean_text(item.get("website")),
                ]
                
                current_ws.append(row_data)
                
                for col_idx in range(1, len(COLUMNS) + 1):
                    cell = current_ws.cell(row=row_index, column=col_idx)
                    cell.font = data_font
                    cell.border = thin_border
                    if col_idx in [1, 3, 4, 11]:
                        cell.alignment = center_align
                    else:
                        cell.alignment = left_align
            
            # 进度打印
            if p % 10 == 0 or p == TOTAL_PAGES or p == vol_end:
                elapsed = time.time() - start_time
                pages_done = p - start_page + 1
                speed = pages_done / max(1, elapsed)
                remaining_pages = TOTAL_PAGES - p
                eta_minutes = (remaining_pages / max(0.01, speed)) / 60
                percent = (p / TOTAL_PAGES) * 100
                print(f"[{percent:5.2f}%] 第{vol_num:02d}卷 ({p:5d}/{vol_end}) | 累计总条数: {total_records:7d} | 速度: {speed:.1f}页/秒 | 预计剩余: {eta_minutes:.1f}分钟", flush=True)
            
            # 定期存盘
            if p % SAVE_INTERVAL == 0 or p == vol_end or p == TOTAL_PAGES:
                if current_wb and current_vol_filename:
                    current_wb.save(current_vol_filename)
                save_progress(p, total_records)
            
            # 防风控平滑延时
            time.sleep(0.2)
            
    # 全量结束保存
    if current_wb and current_vol_filename:
        current_wb.save(current_vol_filename)
        save_progress(TOTAL_PAGES, total_records)
    
    total_time = (time.time() - start_time) / 60
    print("\n" + "=" * 70)
    print(f"[完成] 全量 {TOTAL_PAGES} 页（共 {total_records} 条数据）已全部采集并分卷保存完毕！")
    print(f"总耗时: {total_time:.2f} 分钟")
    print("=" * 70)

if __name__ == "__main__":
    run_all_volumes()
