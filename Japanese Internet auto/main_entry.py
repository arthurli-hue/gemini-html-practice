import sys
import os
import msvcrt
import argparse
from generate_gsheets_report import run_report_logic

def get_password(prompt="請輸入 PASSWORD: "):
    sys.stdout.write(prompt)
    sys.stdout.flush()
    pw = ""
    while True:
        c = msvcrt.getch()
        if c in (b'\r', b'\n'):
            sys.stdout.write('\n')
            break
        elif c == b'\x08': # backspace
            if len(pw) > 0:
                sys.stdout.write('\b \b')
                sys.stdout.flush()
                pw = pw[:-1]
        elif c == b'\x03': # ctrl+c
            raise KeyboardInterrupt
        else:
            try:
                char = c.decode('utf-8')
                sys.stdout.write('*')
                sys.stdout.flush()
                pw += char
            except UnicodeDecodeError:
                pass
    return pw

def main():
    parser = argparse.ArgumentParser(description="Jira QA 報表自動化工具")
    parser.add_argument("-u", "--user", help="Jira 帳號")
    parser.add_argument("-p", "--password", help="Jira 密碼")
    parser.add_argument("-i", "--issue", help="Jira 單號或完整網址")
    parser.add_argument("-s", "--sheet", help="Google Sheet 網址")
    parser.add_argument("--r2", action="store_true", help="是否為 R2 複測報告")
    parser.add_argument("--url", help="Jira 基礎網址 (選填)")

    args = parser.parse_args()

    # --- 互動式引導流程 (Fallback) ---
    print("=== Jira QA 報表工具 (引導模式) ===")
    
    username = args.user or input("請輸入 USERNAME: ").strip()
    
    password = args.password
    if not password:
        password = get_password("請輸入 PASSWORD: ").strip()
    
    raw_issue = args.issue or input("請輸入 TARGET_ISSUE (輸入單號或直接貼上 Jira 網址):\n").strip()
    target_issue = raw_issue.split('/')[-1] if '/' in raw_issue else raw_issue
    
    sheet_url = args.sheet or input("http:輸入sheets.new 貼上您建立的全新 Google Sheet 網址\n(請確認已到共用權限「加給 report-bot@qa-auto-report.iam.gserviceaccount.com 編輯者權限」)\n確認後請在下方貼上網址並按 Enter: \n").strip()
    
    is_r2 = args.r2
    if not args.r2 and not args.issue: # 如果沒有參數傳入，互動式詢問 R1/R2
        r2_input = input("請問本次產出的報告為 (1) 首次測試報告 還是 (2) R2 複測報告？\n(請輸入 1 或 2，預設為 1): ").strip()
        is_r2 = (r2_input == '2')

    configs = {
        'username': username,
        'password': password,
        'target_issue': target_issue,
        'sheet_url': sheet_url,
        'is_r2': is_r2
    }
    
    if args.url:
        configs['jira_url'] = args.url

    # --- 執行邏輯 ---
    print("\n[系統資訊] 正在準備啟動報告產出流程...")
    success = run_report_logic(configs)
    
    if success:
        print("\n[成功] 報告已成功產出。")
    else:
        print("\n[失敗] 處理流程中斷，請檢查上方錯誤訊息。")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n使用者取消操作。")
        sys.exit(0)
