import os
import json
import time
import urllib.parse
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright, Error as PlaywrightError, TimeoutError as PlaywrightTimeoutError

# File & Folder Configuration
EXCEL_INPUT = "contacts.xlsx"
USER_DATA_DIR = "whatsapp_user_data"  # Persists login session (QR code scan)
SCREENSHOTS_DIR = "sent_screenshots"
REPORT_JSON = "whatsapp_report.json"
REPORT_EXCEL = "whatsapp_report.xlsx"

os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

def setup_browser(playwright):
    """
    Launches persistent Chrome browser context to preserve WhatsApp session state.
    """
    browser_context = playwright.chromium.launch_persistent_context(
        user_data_dir=USER_DATA_DIR,
        channel="chrome",
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        viewport={"width": 1280, "height": 800}
    )
    return browser_context

def handle_login(page):
    """
    Navigates to WhatsApp Web and waits for QR code login if not already authenticated.
    """
    print("Navigating to WhatsApp Web...")
    page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
    
    try:
        # Wait up to 60 seconds for the chat list side-panel to be visible
        page.wait_for_selector("#pane-side", timeout=60000)
        print("Logged in successfully!")
    except PlaywrightTimeoutError:
        print("\n" + "="*50)
        print("ACTION REQUIRED: Please scan the QR code on your screen.")
        print("Waiting up to 120 seconds for manual login...")
        print("="*50 + "\n")
        page.wait_for_selector("#pane-side", timeout=120000)
        print("QR code scanned successfully!")

def send_message_and_extract(page, name, phone, message_template):
    """
    Sends personalized WhatsApp message, verifies delivery bubble, captures screenshot,
    and extracts last 3 messages.
    """
    # Clean phone number (remove +, spaces, hyphens)
    clean_phone = "".join(filter(str.isdigit, str(phone)))
    
    # Personalize message template
    personalized_msg = message_template.replace("{name}", str(name)) if pd.notna(message_template) else f"Hello {name}"
    encoded_msg = urllib.parse.quote(personalized_msg)

    report_entry = {
        "name": name,
        "phone": str(phone),
        "message_sent": personalized_msg,
        "status": "Failed",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "screenshot_path": None,
        "last_3_messages": [],
        "error": None
    }

    try:
        # 1. Open chat via direct WhatsApp URL scheme
        chat_url = f"https://web.whatsapp.com/send?phone={clean_phone}&text={encoded_msg}"
        page.goto(chat_url, wait_until="domcontentloaded")
        
        # 2. Check for invalid phone number popup dialog
        time.sleep(3.0)
        invalid_dialog = page.locator("div[role='dialog']")
        if invalid_dialog.count() > 0 and "invalid" in invalid_dialog.inner_text().lower():
            raise Exception(f"Phone number {clean_phone} is invalid or not registered on WhatsApp.")

        # 3. Wait for message input box
        msg_input_selector = "footer div[contenteditable='true']"
        page.wait_for_selector(msg_input_selector, timeout=25000)
        msg_box = page.locator(msg_input_selector).first

        # 4. Fallback: Type text manually if URL auto-fill failed to populate
        time.sleep(1.5)
        current_text = msg_box.inner_text().strip()
        if not current_text:
            msg_box.click()
            page.keyboard.type(personalized_msg)
            time.sleep(1.0)

        # 5. Click Send button or fallback to Enter key
        send_btn_selector = "button span[data-icon='send'], button[aria-label='Send']"
        if page.locator(send_btn_selector).count() > 0:
            page.locator(send_btn_selector).first.click()
        else:
            msg_box.focus()
            page.keyboard.press("Enter")

        # 6. VERIFICATION: Wait for outgoing message bubble to render in chat DOM
        print(f"Verifying message dispatch for {name}...")
        outgoing_bubble_selector = "div.message-out"
        page.wait_for_selector(outgoing_bubble_selector, timeout=15000)

        time.sleep(2.0)  # Short pause for tick marks / rendering

        # 7. Take Screenshot of sent message window
        screenshot_filename = f"{clean_phone}_{int(time.time())}.png"
        screenshot_filepath = os.path.join(SCREENSHOTS_DIR, screenshot_filename)
        page.screenshot(path=screenshot_filepath)
        report_entry["screenshot_path"] = screenshot_filepath

        # 8. Extract Last 3 Messages from chat container
        msg_elements = page.locator("div.message-in span.selectable-text, div.message-out span.selectable-text")
        msg_count = msg_elements.count()

        extracted_msgs = []
        if msg_count > 0:
            start_index = max(0, msg_count - 3)
            for i in range(start_index, msg_count):
                txt = msg_elements.nth(i).inner_text().strip()
                if txt:
                    extracted_msgs.append(txt)

        report_entry["last_3_messages"] = extracted_msgs
        report_entry["status"] = "Success"
        print(f"[SUCCESS] Verified message sent to {name} ({phone})")

    except PlaywrightTimeoutError:
        error_msg = f"Failed to send or verify message for {phone} (Timeout)."
        print(f"[ERROR] {error_msg}")
        report_entry["error"] = error_msg
    except Exception as e:
        error_msg = f"Error processing {phone}: {str(e)}"
        print(f"[ERROR] {error_msg}")
        report_entry["error"] = error_msg

    return report_entry

def run_automation():
    if not os.path.exists(EXCEL_INPUT):
        print(f"Error: Could not find '{EXCEL_INPUT}'. Please ensure the Excel file exists.")
        return

    df = pd.read_excel(EXCEL_INPUT)
    reports = []

    with sync_playwright() as p:
        context = setup_browser(p)
        page = context.new_page()

        # Step 1: Login Check
        try:
            handle_login(page)
        except PlaywrightTimeoutError:
            print("Login timed out. Exiting...")
            context.close()
            return

        # Step 2: Process Contacts from Excel
        for idx, row in df.iterrows():
            name = row.get("Name", f"Contact_{idx+1}")
            phone = row.get("Phone", "")
            msg_template = row.get("Message", "Hello {name}")

            if pd.isna(phone):
                print(f"Skipping row {idx+1}: Missing phone number.")
                continue

            print(f"\nProcessing ({idx+1}/{len(df)}): {name} [{phone}]...")
            result = send_message_and_extract(page, name, phone, msg_template)
            reports.append(result)

            # Pause between contacts to prevent rate-limit bans
            time.sleep(3.0)

        context.close()

    # Step 3: Export Execution Reports
    print("\n" + "="*50)
    print("Generating Reports...")

    # Save JSON Report
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(reports, f, indent=4, ensure_ascii=False)
    print(f"Saved JSON report to '{REPORT_JSON}'")

    # Save Excel Report
    excel_data = []
    for r in reports:
        excel_data.append({
            "Name": r["name"],
            "Phone": r["phone"],
            "Status": r["status"],
            "Message Sent": r["message_sent"],
            "Timestamp": r["timestamp"],
            "Screenshot": r["screenshot_path"],
            "Last 3 Messages": " | ".join(r["last_3_messages"]),
            "Error Details": r["error"]
        })

    report_df = pd.DataFrame(excel_data)
    report_df.to_excel(REPORT_EXCEL, index=False)
    print(f"Saved Excel report to '{REPORT_EXCEL}'")
    print("="*50)

if __name__ == "__main__":
    run_automation()