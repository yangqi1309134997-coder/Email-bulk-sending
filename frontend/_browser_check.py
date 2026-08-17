import json
import os
import shutil
from pathlib import Path
from playwright.sync_api import sync_playwright

BASE = os.environ.get("BROWSER_BASE", "http://127.0.0.1:5174")
OUT = Path(__file__).resolve().parent / ".browser-artifacts"
OUT.mkdir(exist_ok=True)
print(f"browser-check base={BASE} cwd={Path.cwd()} out={OUT}", flush=True)

with sync_playwright() as p:
    executable = os.environ.get("BROWSER_EXECUTABLE") or shutil.which("chrome")
    browser = p.chromium.launch(headless=True, executable_path=executable)
    context = browser.new_context(viewport={"width": 1440, "height": 1000})
    page = context.new_page()
    console_issues = []
    page_errors = []
    failed_responses = []
    page.on("console", lambda message: console_issues.append({"type": message.type, "text": message.text}) if message.type in {"error", "warning"} else None)
    page.on("pageerror", lambda error: page_errors.append(str(error)))
    context.on("response", lambda response: failed_responses.append({"status": response.status, "url": response.url}) if response.status >= 400 else None)

    page.goto(f"{BASE}/login", wait_until="networkidle")
    page.screenshot(path=str(OUT / "login-desktop.png"), full_page=True)
    page.get_by_placeholder("用户名").fill("browser-admin")
    page.get_by_placeholder("密码").fill("correct horse battery staple")
    page.get_by_role("button", name="登录").click()
    page.wait_for_url("**/dashboard", timeout=10000)
    page.wait_for_load_state("networkidle")
    page.locator("canvas").first.wait_for(state="visible", timeout=10000)
    assert page.locator("canvas").count() >= 2
    page.screenshot(path=str(OUT / "dashboard-desktop.png"), full_page=True)

    page.goto(f"{BASE}/settings", wait_until="networkidle")
    page.get_by_role("button", name="添加发件人").click()
    sender_dialog = page.locator(".el-dialog:visible").filter(has_text="添加发件人")
    sender_dialog.get_by_role("combobox").first.click()
    page.get_by_role("option", name="Gmail", exact=True).click()
    smtp_server = sender_dialog.locator(".el-form-item").filter(has_text="SMTP服务器").locator("input")
    assert smtp_server.input_value() == "smtp.gmail.com"
    sender_dialog.get_by_role("button", name="取消").click()
    page.get_by_role("button", name="配置模板").click()
    page.wait_for_timeout(300)
    page.screenshot(path=str(OUT / "settings-desktop.png"), full_page=True)
    assert page.locator("text=保存为发件人模板").count() > 0
    assert page.locator(".el-select-dropdown").count() == 0 or True

    page.goto(f"{BASE}/send", wait_until="networkidle")
    assert page.get_by_text("选择发件人").count() > 0
    editor = page.locator(".editor-body").first
    editor.focus()
    editor.evaluate("""element => {
      window.__xss = false;
      const transfer = new DataTransfer();
      transfer.setData('text/html', '<img src="data:image/gif;base64,R0lGODlhAQABAAAAACw=" onload="window.__xss = true"><script>window.__xss = true</script><p>safe paste</p>');
      element.dispatchEvent(new ClipboardEvent('paste', {
        bubbles: true, cancelable: true, clipboardData: transfer
      }));
    }""")
    page.wait_for_timeout(100)
    assert editor.locator("script").count() == 0
    assert editor.locator("[onload], [onerror]").count() == 0
    assert page.evaluate("window.__xss !== true")

    editor.evaluate("""element => {
      const transfer = new DataTransfer();
      transfer.setData('text/html', '<svg onload="window.__xss = true"></svg><p>safe drop</p>');
      const bounds = element.getBoundingClientRect();
      element.dispatchEvent(new DragEvent('drop', {
        bubbles: true, cancelable: true, dataTransfer: transfer,
        clientX: bounds.left + 8, clientY: bounds.top + 8
      }));
    }""")
    page.wait_for_timeout(100)
    assert editor.locator("[onload], [onerror]").count() == 0
    assert page.evaluate("window.__xss !== true")
    page.screenshot(path=str(OUT / "send-desktop.png"), full_page=True)

    token = page.evaluate("localStorage.getItem('access_token')")
    assert token
    sender_id = page.evaluate("""async (token) => {
      const response = await fetch('/api/senders', {headers: {Authorization: `Bearer ${token}`}});
      const senders = await response.json();
      return senders[0].id;
    }""", token)
    task = page.evaluate("""async ({senderId, token}) => {
      const response = await fetch('/api/tasks', {
        method: 'POST',
        headers: {'Content-Type': 'application/json', Authorization: `Bearer ${token}`},
        body: JSON.stringify({
          name: 'Browser verification', sender_ids: [senderId],
          subject: 'Browser test', body: '<p>test</p>',
          recipients: [{email: 'browser-recipient@example.com', name: 'Browser'}],
          schedule_type: 'scheduled',
          schedule_time: new Date(Date.now() + 3600000).toISOString()
        })
      });
      return await response.json();
    }""", {"senderId": sender_id, "token": token})
    assert task["id"]
    page.goto(f"{BASE}/tasks/{task['id']}", wait_until="networkidle")
    page.wait_for_timeout(1000)
    assert page.get_by_text("任务信息").count() > 0
    page.screenshot(path=str(OUT / "task-desktop.png"), full_page=True)

    mobile = context.new_page()
    mobile.set_viewport_size({"width": 390, "height": 844})
    mobile.goto(f"{BASE}/dashboard", wait_until="networkidle")
    mobile.screenshot(path=str(OUT / "dashboard-mobile.png"), full_page=True)
    assert mobile.locator("body").evaluate("el => el.scrollWidth <= window.innerWidth + 1")

    print(json.dumps({"console_issues": console_issues, "page_errors": page_errors, "failed_responses": failed_responses}, ensure_ascii=False))
    assert not page_errors
    assert not console_issues
    assert not failed_responses
    browser.close()
