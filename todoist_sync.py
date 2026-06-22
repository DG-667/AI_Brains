import urllib.request
import json
import sys

API_TOKEN = "92fdc3dcbe16f0a219b8ac81fff32cd9f70bd9b5"
HEADERS = {
    "Authorization": f"Bearer {API_TOKEN}",
    "Content-Type": "application/json"
}
BASE_URL = "https://api.todoist.com/api/v1"

def make_request(method, endpoint, data=None):
    url = f"{BASE_URL}{endpoint}"
    req = urllib.request.Request(url, method=method, headers=HEADERS)
    if data:
        req.data = json.dumps(data).encode('utf-8')
    try:
        with urllib.request.urlopen(req) as response:
            if response.status in [200, 204]:
                if response.status == 200:
                    return json.loads(response.read().decode('utf-8'))
                return True
    except urllib.error.HTTPError as e:
        print(f"HTTP Error: {e.code} - {e.reason}")
        print(e.read().decode('utf-8'))
    except Exception as e:
        print(f"Error: {e}")
    return None

def get_tasks():
    tasks = make_request("GET", "/tasks")
    if tasks is not None:
        print(f"Found {len(tasks)} active tasks:")
        for t in tasks:
            # handle cases where tasks might have different dict formats based on v1 API
            print(f"- [{t.get('id', 'N/A')}] {t.get('content', '')}")
        return tasks
    return []

def add_task(content):
    res = make_request("POST", "/tasks", {"content": content})
    if res:
        print(f"Successfully added task: {content}")

def close_task(task_id):
    res = make_request("POST", f"/tasks/{task_id}/close")
    if res:
        print(f"Successfully closed task {task_id}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python todoist_sync.py [get|add|close] [args]")
        sys.exit(1)
        
    command = sys.argv[1]
    if command == "get":
        get_tasks()
    elif command == "add":
        add_task(" ".join(sys.argv[2:]))
    elif command == "close":
        close_task(sys.argv[2])
    else:
        print("Unknown command")
