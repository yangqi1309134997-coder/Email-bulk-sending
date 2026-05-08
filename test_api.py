import urllib.request
import json
import time
import sys

BASE = "http://localhost:8000"
token = None

def api(method, path, data=None):
    url = BASE + path
    body = json.dumps(data).encode() if data else None
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        return {"error": e.code, "detail": json.loads(e.read().decode())}

# 1. Health
r = api("GET", "/health")
assert r.get("status") == "ok", f"Health failed: {r}"
print("1. Health OK")

# 2. Login
r = api("POST", "/api/auth/login", {"username": "admin", "password": "admin123"})
token = r.get("access_token")
assert token, f"Login failed: {r}"
print("2. Login OK")

# 3. Get me
r = api("GET", "/api/auth/me")
assert r.get("username") == "admin", f"Me failed: {r}"
print(f"3. Me: {r['username']} ({r['role']})")

# 4. Create sender
r = api("POST", "/api/senders", {
    "email": "test@qq.com", "password": "authcode123",
    "smtp_server": "smtp.qq.com", "smtp_port": 587,
    "use_tls": True, "sender_type": "QQ邮箱",
    "enabled": True, "weight": 50, "daily_quota": 500
})
assert r.get("email") == "test@qq.com", f"Create sender failed: {r}"
sender_id = r.get("id")
print(f"4. Created sender: {r['email']} (id={sender_id})")

# 5. List senders
r = api("GET", "/api/senders")
assert len(r) >= 1, f"List senders failed: {r}"
print(f"5. Senders count: {len(r)}")

# 6. Test sender (will fail but should return response)
r = api("POST", f"/api/senders/{sender_id}/test")
print(f"6. Sender test: {r}")

# 7. Create template
r = api("POST", "/api/templates", {
    "name": "测试模板", "subject": "你好{name}",
    "body": "<p>你好{name}</p>", "variables": ["name", "email"]
})
assert r.get("name") == "测试模板", f"Create template failed: {r}"
template_id = r.get("id")
print(f"7. Created template: {r['name']} (id={template_id})")

# 8. List templates
r = api("GET", "/api/templates")
assert len(r) >= 1, f"List templates failed: {r}"
print(f"8. Templates count: {len(r)}")

# 9. Duplicate template
r = api("POST", f"/api/templates/{template_id}/duplicate")
print(f"9. Duplicated template: {r.get('name')}")

# 10. Dashboard stats
r = api("GET", "/api/dashboard/stats")
print(f"10. Stats: today_sent={r.get('today_sent')}, today_success={r.get('today_success')}")

# 11. List tasks
r = api("GET", "/api/tasks")
print(f"11. Tasks count: {len(r)}")

print("\n=== ALL BACKEND TESTS PASSED! ===")
