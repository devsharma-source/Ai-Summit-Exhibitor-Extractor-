import pandas as pd
import requests
import json
import time
import os

TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJraW5kIjoiYXBwIiwiaWQiOiIyNDMzNCIsImV2ZW50X2lkIjoiMTEwIiwiZW1haWwiOiJkZXYuc2hhcm1hQGdyb3dsZWFkcy5pbyIsImZpcnN0X25hbWUiOiJEZXYiLCJsYXN0X25hbWUiOiJTaGFybWEiLCJmdWxsX25hbWUiOiJEZXYgU2hhcm1hIiwiZGVzY3JpcHRpb24iOm51bGwsInBob25lX2UxNjQiOm51bGwsImNvbXBhbnlfbmFtZSI6IiIsImpvYl90aXRsZSI6IiIsInByb2ZpbGVfcGljdHVyZV91cmwiOm51bGwsImhhc19wcm9maWxlX3VwZGF0ZSI6MSwic3RhdHVzIjoxLCJ0aW1lem9uZSI6IkFzaWEvS29sa2F0YSIsImlzX211dGVfbm90aWZpY2F0aW9ucyI6MCwiaWF0IjoxNzcyMDEyNjYzLCJleHAiOjE3NzI2MTc0NjN9.VJEUIP-pF3pNFWLzQg7vU5wKTXIqBBI-8EbgZ6G6dbY"

API_URL = "https://apis.event-copilot.ai/app/v1.0/ODgzNDM1/ai/ask/0"

RESULTS_FILE = "F:/Growleads/New folder/exhibitor_representatives.xlsx"
PROGRESS_FILE = "F:/Growleads/New folder/progress.json"


def query_copilot(exhibitor_name):
    """Send a question to the copilot and parse the streaming response."""
    question = f"who were representatives of {exhibitor_name}"
    files = {
        "isAudio": (None, "0"),
        "question": (None, question),
        "history": (None, "[]"),
    }
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "text/event-stream",
        "Origin": "https://aiimpactexpo.event-copilot.ai",
        "Referer": "https://aiimpactexpo.event-copilot.ai/",
    }
    try:
        resp = requests.post(API_URL, headers=headers, files=files, timeout=90, stream=True)
        if resp.status_code == 401:
            print("ERROR: Token expired! Please re-authenticate.")
            return None
        if resp.status_code not in (200, 201):
            print(f"  HTTP {resp.status_code} for '{exhibitor_name}'")
            return f"ERROR: HTTP {resp.status_code}"

        # Parse SSE stream - each d: line is a chunk, concatenate them all
        accumulated = ""
        for line in resp.iter_lines(decode_unicode=True):
            if not line:
                continue
            if line.startswith("d:"):
                accumulated += line[2:]
            elif line.startswith("z:"):
                break

        if accumulated:
            try:
                parsed = json.loads(accumulated)
                return parsed.get("entity_summary", accumulated)
            except json.JSONDecodeError:
                return accumulated
        return "No response received"
    except requests.exceptions.Timeout:
        return "ERROR: Request timed out"
    except Exception as e:
        return f"ERROR: {str(e)}"


def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_progress(progress):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)


def save_to_excel(progress):
    rows = []
    for name, answer in progress.items():
        rows.append({"exhibitors_name": name, "representatives": answer})
    result_df = pd.DataFrame(rows)
    try:
        result_df.to_excel(RESULTS_FILE, index=False)
    except PermissionError:
        alt_file = RESULTS_FILE.replace(".xlsx", "_backup.xlsx")
        print(f"  WARNING: Cannot write to {RESULTS_FILE} (file open?). Saving to {alt_file}")
        result_df.to_excel(alt_file, index=False)


def main():
    df = pd.read_excel("F:/Growleads/New folder/Finding websites... (484_871).xlsx")
    exhibitors = df["exhibitors_name"].tolist()
    total = len(exhibitors)

    progress = load_progress()
    print(f"Total exhibitors: {total}")
    print(f"Already completed: {len(progress)}")
    print(f"Remaining: {total - len(progress)}")
    print("-" * 60)

    for i, name in enumerate(exhibitors):
        name_str = str(name).strip()
        if name_str in progress:
            continue

        print(f"[{i+1}/{total}] Querying: {name_str}")
        answer = query_copilot(name_str)

        if answer is None:
            print("Stopping due to auth error. Re-run after re-authenticating.")
            break

        progress[name_str] = answer
        save_progress(progress)

        if (i + 1) % 10 == 0:
            save_to_excel(progress)
            print(f"  ... saved progress ({len(progress)}/{total})")

        time.sleep(1.5)

    save_to_excel(progress)
    print(f"\nDone! Results saved to {RESULTS_FILE}")
    print(f"Total results: {len(progress)}")


if __name__ == "__main__":
    main()
