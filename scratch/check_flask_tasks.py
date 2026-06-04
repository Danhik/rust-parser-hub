import requests
import json

try:
    r = requests.get("http://localhost:5050/api/tasks")
    if r.status_code == 200:
        tasks = r.json()
        output_lines = [f"Total tasks: {len(tasks)}"]
        for t in tasks:
            output_lines.append(f"\nTask ID: {t.get('id')}")
            output_lines.append(f"  Parser: {t.get('parser_id')}")
            output_lines.append(f"  Status: {t.get('status')}")
            output_lines.append(f"  Cmd: {t.get('cmd')}")
            output_lines.append(f"  PID: {t.get('pid')}")
            log = t.get("log", [])
            output_lines.append(f"  Log lines count: {len(log)}")
            output_lines.append("  Log:")
            for line in log:
                output_lines.append(f"    {line}")
        
        with open("C:\\Users\\Daniil\\.gemini\\antigravity-ide\\brain\\43a76f28-018f-4141-8b75-11f4e3af6d5d\\scratch\\flask_tasks_debug.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(output_lines))
        print("Success")
    else:
        print(f"Error status code: {r.status_code}")
except Exception as e:
    print(f"Error querying Flask API: {e}")
