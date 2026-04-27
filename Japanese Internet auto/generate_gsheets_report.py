import requests
from requests.auth import HTTPBasicAuth
import re
import urllib3
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import sys

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= 預設設定 =================
DEFAULT_JIRA_URL = "https://pmo-jira.qyrc452.com"
TEMPLATE_SHEET_ID = "1nmDDlSn-GDrs2qvZ86jEgld2UktLBBgOeruSNMWbPkw" 

PRIORITY_MAP = {
    "Highest": "A級",
    "High": "A級",
    "Medium": "B級",
    "Low": "C級",
    "Lowest": "C級"
}

def get_issue_details(issue_key, jira_url, username, password):
    url = f"{jira_url}/rest/api/2/issue/{issue_key}"
    res = requests.get(url, auth=HTTPBasicAuth(username, password), verify=False)
    if res.status_code == 200:
        return res.json()
    return None

def extract_module(summary):
    tags = re.findall(r'\[(.*?)\]', summary)
    custom_tags = [t for t in tags if t not in ['JPCafe', 'UAT', 'x'] and '.' not in t]
    if custom_tags:
        return custom_tags[-1]
    return "未分類"

def clean_jira_description(raw_desc):
    if not raw_desc: return "無測試項目描述"
    for cutoff_word in ["【本次無須測試項目】", "本次無須測試項目", "【驗測時間】"]:
        idx = raw_desc.find(cutoff_word)
        if idx != -1:
            raw_desc = raw_desc[:idx]
            break
    clean_desc = re.sub(r'\{color.+?\}|\{color\}', '', raw_desc)
    clean_desc = re.sub(r'h\d\.\s*', '', clean_desc)
    return clean_desc.strip()

def write_to_cell_adjacent(worksheet, search_str, content):
    try:
        if isinstance(content, str) and len(content) > 48000:
            content = content[:48000] + "\n...(文章過長，因 Google Sheet 限制已在此截斷)"
            
        cell = worksheet.find(re.compile(search_str, re.IGNORECASE))
        if cell:
            worksheet.update_cell(cell.row, cell.col + 1, content)
            print(f"成功寫入相鄰欄位：「{search_str}」 (行 {cell.row})")
            return True
        else:
            print(f"找不到關鍵字：「{search_str}」，略過寫入。")
    except Exception as e:
        print(f"寫入關鍵字「{search_str}」時發生錯誤: {e}")
    return False

def get_custom_val(fields, key, default=""):
    val = fields.get(key)
    if isinstance(val, dict):
        return val.get("value", default)
    if isinstance(val, list) and len(val) > 0 and isinstance(val[0], dict):
        return val[0].get("value", default)
    return str(val) if val else default

def run_report_logic(configs):
    """
    configs 預期包含:
    - username
    - password
    - target_issue
    - sheet_url
    - is_r2 (bool)
    - jira_url (optional)
    """
    username = configs['username']
    password = configs['password']
    target_issue = configs['target_issue']
    new_sheet_url = configs['sheet_url']
    is_r2 = configs.get('is_r2', False)
    jira_url = configs.get('jira_url', DEFAULT_JIRA_URL)

    if not os.path.exists("credentials.json"):
        print("錯誤：找不到 credentials.json。")
        return False

    print(f"正在撈取 Jira 主單: {target_issue} ...")
    parent_issue = get_issue_details(target_issue, jira_url, username, password)
    if not parent_issue:
        print("無法連線或找不到該單號！請檢查帳號密碼或單號是否正確。")
        return False

    raw_desc = parent_issue.get("fields", {}).get("description", "")
    test_items_text = clean_jira_description(raw_desc)

    links = parent_issue.get("fields", {}).get("issuelinks", [])
    bug_list = []
    
    print(f"開始分析關聯的 {len(links)} 個 Issue...")
    for link in links:
        issue = link.get("inwardIssue") or link.get("outwardIssue")
        if not issue: continue
        
        bug_key = issue["key"]
        issuetype = issue["fields"]["issuetype"]["name"]
        
        bug_detail = get_issue_details(bug_key, jira_url, username, password)
        if not bug_detail: continue
            
        fields = bug_detail["fields"]
        summary = fields.get("summary", "")
        status = fields.get("status", {}).get("name", "Unknown").upper()
        
        if status == "DONE": continue # 略過已完成但非 Bug 的項目？或依需求調整

        if issuetype != "Bug": continue
            
        raw_priority = fields.get("priority", {}).get("name", "Medium")
        mapped_priority = PRIORITY_MAP.get(raw_priority, "B級")
        short_priority = mapped_priority.replace("級", "")
        module = extract_module(summary)
        
        creator = fields.get("creator", {}).get("displayName", "System")
        defect_category = get_custom_val(fields, "customfield_11101", "")
        frequency = get_custom_val(fields, "customfield_10210", "Always")
        
        short_status = "Todo" if "TODO" in status else ("Resolved" if "RESOLVED" in status else status.capitalize())

        summary_lower = summary.lower()
        is_game = ("前台" in summary_lower or "遊戲" in summary_lower or "web" in summary_lower)
        is_admin = ("後台" in summary_lower or "admin" in summary_lower)
        bucket = "遊戲端" if is_game else ("後台端" if is_admin else "後台端")

        is_unresolved = (status != "RESOLVED")
        
        bug_list.append({
            "key": bug_key, "summary": summary, "status": status,
            "priority": mapped_priority, "short_priority": short_priority,
            "module": module, "bucket": bucket, "is_unresolved": is_unresolved,
            "creator": creator, "defect_category": defect_category,
            "frequency": frequency, "short_status": short_status
        })
        print(f"  - 處理完成 Bug: {bug_key} ({module})")

    total_bugs = len(bug_list)
    a_count = sum(1 for b in bug_list if b["priority"] == "A級")
    b_count = sum(1 for b in bug_list if b["priority"] == "B級")
    c_count = sum(1 for b in bug_list if b["priority"] == "C級")
    
    game_bugs = [b for b in bug_list if b["bucket"] == "遊戲端"]
    admin_bugs = [b for b in bug_list if b["bucket"] == "後台端"]
    
    total_unresolved = sum(1 for b in bug_list if b["is_unresolved"])
    game_unresolved = sum(1 for b in game_bugs if b["is_unresolved"])
    admin_unresolved = sum(1 for b in admin_bugs if b["is_unresolved"])

    stat_game = f"總計『{len(game_bugs)}』條 bug，剩餘『{game_unresolved}』條 bug 未修正"
    stat_admin = f"總計『{len(admin_bugs)}』條 bug，剩餘『{admin_unresolved}』條 bug 未修正"
    stat_total = f"【總計】共有『{total_unresolved}』條 bug 未修正"

    print(f"\n正在連線 Google Sheets...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    
    try:
        sheet = client.open_by_url(new_sheet_url)
        template_doc = client.open_by_key(TEMPLATE_SHEET_ID)
        
        initial_worksheets = sheet.worksheets()
        existing_titles = [ws.title.strip() for ws in initial_worksheets]
        
        for template_ws in template_doc.worksheets():
            if template_ws.title.strip() not in existing_titles:
                res = template_ws.copy_to(sheet.id)
                sheet.get_worksheet_by_id(res['sheetId']).update_title(template_ws.title) 
                
        # 清理預設分頁並排序
        current_worksheets = sheet.worksheets()
        for ws in current_worksheets:
            if ws.title in ["工作表1", "Sheet1"]:
                if len(sheet.worksheets()) > 1:
                    try:
                        sheet.del_worksheet(ws)
                        print(f"  - 已自動清理預設的空白分頁：{ws.title}")
                    except Exception: pass
            
            # 將「總結」排到第一位
            if "總結" in ws.title:
                try:
                    ws.update_index(0)
                except Exception: pass
                
        # 掃描頁籤資料
        scanned_tabs_text = []
        game_val_stats = {"pass": 0, "fail": 0, "na": 0, "block": 0, "total": 0}
        all_tabs = sheet.worksheets()
        
        tab_keywords = [("新功能", "需求進行測試", "條"), ("遊戲驗證", "驗證案例", "條"), ("功能優化", "需求進行測試", "需求")]
        for tab_name, type_word, unit_word in tab_keywords:
            ws = next((s for s in all_tabs if tab_name in s.title.strip()), None)
            if ws:
                all_data = ws.get_all_values()
                if not all_data: continue
                t_pass = t_fail = t_na = t_block = 0
                fail_tickets = []
                for row in all_data[1:]:
                    row_has_fail = False
                    for cell in row:
                        val = cell.strip().lower()
                        if val in ["pass", "通過"]: t_pass += 1; break
                        elif val in ["fail", "失敗", "阻塞"]:
                            if val in ["fail", "失敗"]: t_fail += 1
                            else: t_block += 1
                            row_has_fail = True; break
                        elif val in ["na", "n/a", "未執行"]: t_na += 1; break
                    if row_has_fail:
                        for cell in row:
                            m = re.search(r'[A-Za-z]+-\d+', cell)
                            if m: fail_tickets.append(m.group(0)); break
                t_total = t_pass + t_fail + t_na + t_block
                if t_total > 0:
                    if tab_name == "遊戲驗證":
                        game_val_stats.update({"pass": t_pass, "fail": t_fail, "na": t_na, "block": t_block, "total": t_total})
                    fail_str = f"({','.join(set(fail_tickets))})" if (t_fail > 0 and fail_tickets) else ""
                    status_text = "不通過" if (t_fail > 0 or t_block > 0) else "通過"
                    scanned_tabs_text.append({"title": tab_name, "text": [
                        f"● 總計：共【{t_total}】{type_word}，目前通過【{t_pass}】{unit_word}，失敗【{t_fail}】{unit_word}{fail_str}，測試{status_text}。",
                        f"● 詳見頁籤【{tab_name}】\n"
                    ]})

        # 彙整測試總結
        target_summary_clean = re.sub(r'\[.*?\]|【.*?】', '', parent_issue.get("fields", {}).get("summary", "功能驗證")).strip()
        summary_text = []
        if is_r2:
            pass_bugs = [b for b in bug_list if b["status"] in ["PASS", "DONE", "RESOLVED"]]
            closed_bugs = [b for b in bug_list if b["status"] in ["CLOSED"]]
            fail_bugs = [b for b in bug_list if b["status"] not in ["PASS", "DONE", "RESOLVED", "CLOSED"]]
            p_cnt, c_cnt, f_cnt = len(pass_bugs), len(closed_bugs), len(fail_bugs)
            tot_cnt = p_cnt + c_cnt + f_cnt
            res_msg = "UAT R2 測試通過" if f_cnt == 0 else "UAT R2 測試不通過"
            summary_text.extend([f"本次共驗證【{tot_cnt}】個Bug，通過共【{p_cnt}】個，不通過共【{f_cnt}】個，設計如此共【{c_cnt}】個。\n", f"{target_summary_clean}，{res_msg}\n"])
            fail_bugs_keys = ",".join([b['key'] for b in fail_bugs])
            fail_bugs_str = f"({fail_bugs_keys})" if f_cnt > 0 else ""
            summary_text.extend([
                "1. Bug複測：", 
                f"● 驗證【{tot_cnt}】條Bug，{c_cnt}條為設計如此，{p_cnt}條驗證通過，{f_cnt}條驗證不通過{fail_bugs_str}。", 
                "● 詳見頁籤【BUG複測】\n"
            ])
            for idx, tab_info in enumerate(scanned_tabs_text, 2):
                summary_text.append(f"{idx}. {tab_info['title']}：")
                summary_text.extend(tab_info['text'])
        else:
            summary_text = [f"本次共發現【{total_bugs}】個Bug，A級Bug共【{a_count}】個，B級Bug共【{b_count}】個，C級Bug共【{c_count}】個。\n", f"{target_summary_clean}，目前測試發現以下狀況:\n"]
            for idx, tab_info in enumerate(scanned_tabs_text, 1):
                summary_text.append(f"{idx}. {tab_info['title']}：")
                summary_text.extend(tab_info['text'])
            idx = len(scanned_tabs_text) + 1
            modules = {}
            for b in bug_list: modules.setdefault(b["module"], []).append(b)
            for mod, bugs in modules.items():
                summary_text.append(f"{idx}. {mod} (異常Bug詳細)：" if mod in [t['title'] for t in scanned_tabs_text] else f"{idx}. {mod}：")
                summary_text.append(f"● 總計：發現【{len(bugs)}】個 Bug，為「{mod}」模組功能異常")
                for bug in bugs:
                    clean_summary = re.sub(r'\[.*?\]|【.*?】', '', bug["summary"]).strip()
                    summary_text.append(f"  - {clean_summary}（{bug['key']}）")
                summary_text.append(f"● 詳見頁籤\n"); idx += 1

        final_summary_string = "\n".join(summary_text)

        # 寫入主控台
        worksheet = sheet.get_worksheet(0)
        write_to_cell_adjacent(worksheet, "測試項目", test_items_text)
        write_to_cell_adjacent(worksheet, "測試結果", final_summary_string)
        write_to_cell_adjacent(worksheet, "遊戲端 bug", stat_game)
        write_to_cell_adjacent(worksheet, "後台 bug", stat_admin)
        write_to_cell_adjacent(worksheet, r"總計\s*$", stat_total)
        
        # 圓餅圖更新
        try:
            bug_level_cell = worksheet.find(re.compile(r"BUG等級"))
            if bug_level_cell:
                worksheet.update_cell(bug_level_cell.row + 1, bug_level_cell.col + 1, a_count)
                worksheet.update_cell(bug_level_cell.row + 2, bug_level_cell.col + 1, b_count)
                worksheet.update_cell(bug_level_cell.row + 3, bug_level_cell.col + 1, c_count)
                worksheet.update_cell(bug_level_cell.row + 6, bug_level_cell.col + 1, total_bugs)
            
            tc_cell = worksheet.find(re.compile(r"測試用例"))
            if tc_cell:
                worksheet.update_cell(tc_cell.row + 2, tc_cell.col + 1, game_val_stats["pass"])
                worksheet.update_cell(tc_cell.row + 3, tc_cell.col + 1, game_val_stats["block"])
                worksheet.update_cell(tc_cell.row + 4, tc_cell.col + 1, game_val_stats["fail"])
                worksheet.update_cell(tc_cell.row + 5, tc_cell.col + 1, game_val_stats["na"])
                worksheet.update_cell(tc_cell.row + 7, tc_cell.col + 1, game_val_stats["total"])
        except Exception as e: print(f"更新圓餅圖數據時產生錯誤：{e}")

        # Bug 清單詳細分頁寫入
        bug_worksheet = next((ws for ws in all_tabs if "bug" in ws.title.lower() and "清單" in ws.title), None)
        if bug_worksheet:
            print(f"\n找到清單分頁: {bug_worksheet.title}，準備渲染多筆 Bug 詳細資料...")
            def update_bug_table(marker_name, data_list, stat_text):
                print(f"正在處理區塊: {marker_name} (資料筆數: {len(data_list)})")
                marker = bug_worksheet.find(re.compile(marker_name, re.IGNORECASE))
                if not marker:
                    print(f"  [錯誤] 找不到標記: {marker_name}")
                    return
                bug_worksheet.update_cell(marker.row, marker.col + 1, stat_text)
                if not data_list: 
                    print(f"  [資訊] {marker_name} 資料列表為空，僅更新統計文字。")
                    return
                # 3. 如果是遊戲端 BUG，要檢查是否會擠壓到下方的「後台 BUG」
                if "遊戲端" in marker_name:
                    admin_marker = bug_worksheet.find(re.compile(r"後台 BUG", re.IGNORECASE))
                    if admin_marker:
                        # 計算要求的佈局：[數據起始列] + [數據量] + [5列間距]
                        data_start_row = marker.row + 2
                        target_admin_row = data_start_row + len(data_list) + 5
                        
                        # 如果目前的後台標記列位置小於目標列數，則插入足夠的空白列
                        if admin_marker.row < target_admin_row:
                            num_to_insert = target_admin_row - admin_marker.row
                            bug_worksheet.insert_rows([[]] * num_to_insert, admin_marker.row)
                            print(f"偵測到數據較多，已自動插入 {num_to_insert} 行以保持 5 列間距。")
                marker = bug_worksheet.find(re.compile(marker_name, re.IGNORECASE))
                start_row = marker.row + 2
                
                # 僅準備 7 欄資料 (A-G)，避免覆蓋 H 欄以後的現有格式
                sheet_data = [[
                    b["key"], b["summary"], b["short_priority"], 
                    b["defect_category"], b["creator"], b["short_status"], 
                    b["frequency"]
                ] for b in data_list]
                
                # 更新寫法：僅寫入 A-G 欄
                range_label = f"A{start_row}:G{start_row + len(sheet_data) - 1}"
                bug_worksheet.update(range_label, sheet_data)
                print(f"  [成功] 已將 {len(sheet_data)} 筆資料寫入 {range_label}")

            update_bug_table(r"遊戲端 BUG", game_bugs, stat_game)
            update_bug_table(r"後台 BUG", admin_bugs, stat_admin)

        print(f"\n全部任務完成！報告連結：https://docs.google.com/spreadsheets/d/{sheet.id}")
        return True
    except Exception as e:
        import traceback
        print(f"產生錯誤:\n{traceback.format_exc()}")
        return False
