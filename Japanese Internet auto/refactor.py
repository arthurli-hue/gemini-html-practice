import os
import re

with open("generate_gsheets_report.py", "r", encoding="utf-8") as f:
    lines = f.read()

split_str = "    # === 測試總結產生 ==="
if split_str not in lines:
    print("Cannot find split point!")
else:
    parts = lines.split(split_str)
    head = parts[0]

    new_tail = """    # === Google Sheets 連線與初始化 ===
    print(f"\\n正在從模板 (ID: {TEMPLATE_SHEET_ID}) 準備報表...")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    
    try:
        if not NEW_SHEET_URL:
            print("錯誤：您未提供 Google Sheet 網址！")
            return
            
        print(f"\\n正在連線到您提供的 Google Sheet...")
        try:
            sheet = client.open_by_url(NEW_SHEET_URL)
        except Exception as e:
            if "404" in str(e):
                print("\\n【錯誤 404】找不到檔案！")
                print("這通常代表您『忘記將這個檔案共用給 Service Account』了。")
                print("請回到您的檔案 -> 點擊右上角「共用」 -> 把 report-bot@qa-auto-report.iam.gserviceaccount.com 設為編輯者！")
                return
            elif "400" in str(e) and "not supported" in str(e):
                print("\\n【錯誤 400】不支援此格式！")
                print("您剛才貼上的網址是一份「微軟 Excel (.xlsx) 檔案」，Google Sheets API 無法直接修改這個格式的檔案。")
                print("【解決方案】：")
                print("1. 請先到您該篇「匯入的報告」網頁中。")
                print("2. 點擊左上角的「檔案」 -> 選擇「儲存為 Google 試算表」。")
                print("3. 此時瀏覽器會為您開啟一個『全新且支援操作』的分頁。")
                print("4. 請把『新分頁』的網址共用給 Service Account 後，再次跑這支程式並貼上『新網址』即可！")
                return
            raise e
            
        print("正在檢查與準備分頁結構，這可能需要幾秒鐘...")
        template_doc = client.open_by_key(TEMPLATE_SHEET_ID)
        
        initial_worksheets = sheet.worksheets()
        existing_titles = [ws.title.strip() for ws in initial_worksheets]
        
        for template_ws in template_doc.worksheets():
            if template_ws.title.strip() in existing_titles:
                print(f"  - 分頁「{template_ws.title}」已存在，保留現有內容。")
            else:
                print(f"  - 正在從模板複製分頁：{template_ws.title} ...")
                res = template_ws.copy_to(sheet.id)
                copied_ws = sheet.get_worksheet_by_id(res['sheetId'])
                copied_ws.update_title(template_ws.title) 
                
        for initial_ws in initial_worksheets:
            if initial_ws.title in ["工作表1", "Sheet1"]:
                if len(sheet.worksheets()) > 1:
                    try:
                        sheet.del_worksheet(initial_ws)
                        print(f"  - 已自動清理預設的空白分頁：{initial_ws.title}")
                    except Exception:
                        pass
        print("表單結構確認完畢！準備掃描測試資料...\\n")

        # === 掃描特定頁籤狀態 ===
        tab_keywords = [
            ("新功能", "需求進行測試", "條"),
            ("遊戲驗證", "驗證案例", "條"),
            ("功能優化", "需求進行測試", "需求")
        ]
        
        game_val_stats = {"pass": 0, "fail": 0, "na": 0, "block": 0, "total": 0}
        scanned_tabs_text = []

        all_tabs = sheet.worksheets()
        
        for tab_name, type_word, unit_word in tab_keywords:
            ws = None
            for s in all_tabs:
                if tab_name in s.title.strip():
                    ws = s
                    break
                    
            if ws:
                all_data = ws.get_all_values()
                if not all_data: continue
                # Skip header row and check body
                t_pass = t_fail = t_na = t_block = 0
                fail_tickets = []
                
                for row in all_data[1:]: # 略過標題行
                    row_has_fail = False
                    for cell in row:
                        val = cell.strip().lower()
                        if val in ["pass", "通過"]:
                            t_pass += 1
                            break
                        elif val in ["fail", "失敗", "阻塞"]:
                            if val in ["fail", "失敗"]: t_fail += 1
                            else: t_block += 1
                            row_has_fail = True
                            break
                        elif val in ["na", "n/a", "未執行"]:
                            t_na += 1
                            break
                            
                    if row_has_fail:
                        # 擷取失敗單號 (優先找同行文字中包含的 Jira Key)
                        for cell in row:
                            m = re.search(r'[A-Za-z]+-\\d+', cell)
                            if m:
                                fail_tickets.append(m.group(0))
                                break
                                
                t_total = t_pass + t_fail + t_na + t_block
                
                if tab_name == "遊戲驗證":
                    game_val_stats["pass"] = t_pass
                    game_val_stats["fail"] = t_fail
                    game_val_stats["na"] = t_na
                    game_val_stats["block"] = t_block
                    game_val_stats["total"] = t_total
                
                # 若該分頁有資料，就算沒人標 Pass/Fail，只要有資料都可能列出來，如果不要就加上 if t_total == 0 continue
                if t_total == 0:
                    continue

                fail_str = ""
                if t_fail > 0 and len(fail_tickets) > 0:
                    fail_str = f"({','.join(set(fail_tickets))})"
                    
                status_text = "不通過" if (t_fail > 0 or t_block > 0) else "通過"
                
                # 注意這兩行是依賴 Image 的長相格式
                item_text = [
                    f"● 總計：共【{t_total}】{type_word}，目前通過【{t_pass}】{unit_word}，失敗【{t_fail}】{unit_word}{fail_str}，測試{status_text}。",
                    f"● 詳見頁籤【{tab_name}】\\n"
                ]
                scanned_tabs_text.append({"title": tab_name, "text": item_text})

        # === 測試總結產生 ===
        target_summary_clean = re.sub(r'\\[.*?\\]|【.*?】', '', parent_issue.get("fields", {}).get("summary", "功能驗證")).strip()
        if not target_summary_clean:
            target_summary_clean = "功能驗證"
            
        summary_text = []

        if IS_R2:
            pass_bugs = [b for b in bug_list if b["status"] in ["PASS", "DONE", "RESOLVED"]]
            closed_bugs = [b for b in bug_list if b["status"] in ["CLOSED"]]
            fail_bugs = [b for b in bug_list if b["status"] not in ["PASS", "DONE", "RESOLVED", "CLOSED"]]

            p_cnt = len(pass_bugs)
            c_cnt = len(closed_bugs)
            f_cnt = len(fail_bugs)
            tot_cnt = p_cnt + c_cnt + f_cnt

            if f_cnt == 0:
                summary_text.extend([
                    f"本次共驗證【{tot_cnt}】個Bug，通過共【{p_cnt}】個，不通過共【{f_cnt}】個，設計如此共【{c_cnt}】個。\\n",
                    f"{target_summary_clean}，UAT R2 測試通過\\n"
                ])
            else:
                fail_str = ""
                if f_cnt > 0:
                    fails = ",".join(list(set([b["key"] for b in fail_bugs])))
                    fail_str = f"({fails})"
                
                summary_text.extend([
                    f"本次共驗證【{tot_cnt}】個Bug，設計如此【{c_cnt}】個，通過【{p_cnt}】個，不通過【{f_cnt}】個。\\n",
                    f"{target_summary_clean}，UAT R2測試不通過\\n"
                ])
                
            summary_text.extend([
                f"1. Bug複測：",
                f"● 驗證【{tot_cnt}】條Bug，{c_cnt}條為設計如此，{p_cnt}條驗證通過，{f_cnt}條驗證不通過{fail_str if f_cnt > 0 else ''}。",
                f"● 詳見頁籤【BUG複測】\\n"
            ])
            
            # 追加掃描到的額外頁籤 (2. 3. 4. 等等)
            section_idx = 2
            for tab_info in scanned_tabs_text:
                summary_text.append(f"{section_idx}. {tab_info['title']}：")
                summary_text.extend(tab_info['text'])
                section_idx += 1
        else:
            # 首次測試
            summary_text = [
                f"本次共發現【{total_bugs}】個Bug，A級Bug共【{a_count}】個，B級Bug共【{b_count}】個，C級Bug共【{c_count}】個。\\n",
                f"{target_summary_clean}，目前測試發現以下狀況:\\n"
            ]
            
            section_idx = 1
            
            # 若有新功能、遊戲驗證掃描結果，優先秀出！
            for tab_info in scanned_tabs_text:
                summary_text.append(f"{section_idx}. {tab_info['title']}：")
                summary_text.extend(tab_info['text'])
                section_idx += 1
                
            # 將發現異常的各模組補上
            modules = {}
            for b in bug_list:
                modules.setdefault(b["module"], []).append(b)
                
            for mod, bugs in modules.items():
                if mod in [t['title'] for t in scanned_tabs_text]:
                    summary_text.append(f"{section_idx}. {mod} (異常Bug詳細)：")
                else:
                    summary_text.append(f"{section_idx}. {mod}：")
                    
                summary_text.append(f"● 總計：發現【{len(bugs)}】個 Bug，為「{mod}」模組功能異常")
                for bug in bugs:
                    clean_summary = re.sub(r'\\[.*?\\]|【.*?】', '', bug["summary"]).strip()
                    summary_text.append(f"  - {clean_summary}（{bug['key']}）")
                summary_text.append(f"● 詳見頁籤\\n")
                section_idx += 1

        final_summary_string = "\\n".join(summary_text)

        # === 資料寫入 主控台 ===
        worksheet = sheet.get_worksheet(0)
        write_to_cell_adjacent(worksheet, "測試項目", test_items_text)
        write_to_cell_adjacent(worksheet, "測試結果", final_summary_string)
        write_to_cell_adjacent(worksheet, "遊戲端 bug", stat_game)
        write_to_cell_adjacent(worksheet, "後台 bug", stat_admin)
        write_to_cell_adjacent(worksheet, r"總計\\s*$", stat_total)
        
        # 圓餅圖
        try:
            bug_level_cell = worksheet.find(re.compile(r"BUG等級"))
            if bug_level_cell:
                worksheet.update_cell(bug_level_cell.row + 1, bug_level_cell.col + 1, a_count)
                worksheet.update_cell(bug_level_cell.row + 2, bug_level_cell.col + 1, b_count)
                worksheet.update_cell(bug_level_cell.row + 3, bug_level_cell.col + 1, c_count)
                worksheet.update_cell(bug_level_cell.row + 6, bug_level_cell.col + 1, total_bugs)
                print(f"成功更新「圓餅圖統計表」的數據 (基準行 {bug_level_cell.row})")
        except Exception as e:
            print(f"更新 Bug 圓餅圖數據時產生錯誤：{e}")

        # 【第二部分】Bug 詳細清單分頁寫入
        bug_worksheet = None
        for ws in all_tabs:
            if "bug" in ws.title.lower() and "清單" in ws.title:
                bug_worksheet = ws
                break
                
        if bug_worksheet:
            print(f"\\n找到清單分頁: {bug_worksheet.title}，準備渲染多筆 Bug 詳細資料...")
            def update_bug_table(marker_name, data_list):
                if not data_list: return
                marker = bug_worksheet.find(re.compile(marker_name, re.IGNORECASE))
                if marker:
                    start_row = marker.row + 2
                    sheet_data = []
                    for bug in data_list:
                        sheet_data.append([
                            bug["key"], bug["summary"], bug["short_priority"], 
                            bug["defect_category"], bug["creator"], bug["short_status"], 
                            bug["frequency"], "", "", ""
                        ])
                    # 依賴 gspread 6
                    bug_worksheet.update(sheet_data, f"A{start_row}:J{start_row + len(sheet_data) - 1}")
                    print(f"成功批量寫入 {len(sheet_data)} 筆 {marker_name}。")
            update_bug_table(r"遊戲端 BUG", game_bugs)
            update_bug_table(r"後台 BUG", admin_bugs)
        else:
            print("\\n警告：找不到名為「bug 清單」的分頁，略過 Bug 明細填寫。")

        # 測試用例圓餅圖更新 (依照 game_val_stats)
        try:
            tc_cell = worksheet.find(re.compile(r"測試用例"))
            if tc_cell:
                if game_val_stats["total"] == 0:
                    worksheet.update_cell(tc_cell.row + 2, tc_cell.col + 1, 0)
                    worksheet.update_cell(tc_cell.row + 3, tc_cell.col + 1, 0)
                    worksheet.update_cell(tc_cell.row + 4, tc_cell.col + 1, 0)
                    worksheet.update_cell(tc_cell.row + 5, tc_cell.col + 1, 0)
                    worksheet.update_cell(tc_cell.row + 7, tc_cell.col + 1, 0)
                else:
                    worksheet.update_cell(tc_cell.row + 2, tc_cell.col + 1, game_val_stats["pass"])
                    worksheet.update_cell(tc_cell.row + 3, tc_cell.col + 1, game_val_stats["block"])
                    worksheet.update_cell(tc_cell.row + 4, tc_cell.col + 1, game_val_stats["fail"])
                    worksheet.update_cell(tc_cell.row + 5, tc_cell.col + 1, game_val_stats["na"])
                    worksheet.update_cell(tc_cell.row + 7, tc_cell.col + 1, game_val_stats["total"])
                    print("成功更新「測試用例」圓餅圖統計表！")
        except Exception as e:
            print(f"計算遊戲驗證圓餅圖時發生錯誤：{e}")

        print("\\n=============================================")
        print("全部任務完成！這是您的全新報告專屬連結：")
        print(f"https://docs.google.com/spreadsheets/d/{sheet.id}")
        print("=============================================\\n")
            
    except gspread.exceptions.APIError as e:
        print(f"Google API 錯誤: {e}")
    except Exception as e:
        import traceback
        print(f"未預期的腳本錯誤:\\n{traceback.format_exc()}")

if __name__ == "__main__":
    main()
"""

    with open("generate_gsheets_report.py", "w", encoding="utf-8") as f:
        f.write(head + new_tail)

    print("Patch applied successfully.")
