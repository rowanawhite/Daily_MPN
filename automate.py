import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from playwright.sync_api import sync_playwright
from datetime import datetime, timedelta
import pandas as pd
import imaplib
import email as email_lib
import re
import time
import requests

OTP_EMAIL = os.environ["OTP_EMAIL"]        # new secret: the Outlook address that receives the code

EMAIL_ADDRESS = os.environ["EMAIL_ADDRESS"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"] #yqgk-nmye-vlpo-tcjt
EBUILDER_USERNAME = os.environ["EBUILDER_USERNAME"]
EBUILDER_PASSWORD = os.environ["EBUILDER_PASSWORD"]

def click_button():
    with sync_playwright() as p:
        print("Launching browser...")
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        try:
            print("Navigating to login page...")
            page.goto("https://app.e-builder.net", timeout=60000)
            page.wait_for_load_state("networkidle", timeout=60000)
            print(f"Page title: {page.title()}")
            print(f"Page URL: {page.url}")
            page.screenshot(path="screenshot_1_login.png")
            print("Screenshot 1 saved.")
        except Exception as e:
            print(f"FAILED at login page: {e}")
            page.screenshot(path="screenshot_error.png")
            browser.close()
            raise

        try:
            print("Filling in login form...")

            # Step 1: Username
            page.fill("input[name='username']", EBUILDER_USERNAME)
            print("Username filled")
            page.wait_for_timeout(2000)
            page.click("text=Continue")
            print("Clicked Continue after username")
            page.wait_for_load_state("networkidle", timeout=30000)
            page.wait_for_timeout(2000)
            page.screenshot(path="screenshot_2_after_username.png")
            
            # Step 2: Email entry field
            page.wait_for_selector("input#username-field", state="visible", timeout=15000)
            page.fill("input#username-field", OTP_EMAIL)  # sets value all at once, no character delay
            page.evaluate("""
                const input = document.querySelector("input#username-field");
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
            """)
            page.wait_for_timeout(1000)  # brief pause for the button to react to the events
            page.wait_for_selector("button#enter_username_submit:not([disabled])", state="visible", timeout=15000)
            page.click("button#enter_username_submit")
            page.wait_for_timeout(2000)
            page.screenshot(path="screenshot_3_after_email.png")
            print("Email submitted")

            # Step 3: Password — wait for password step to become active/visible
            page.wait_for_selector("input[name='password']", state="visible", timeout=15000)
            page.fill("input[name='password']", EBUILDER_PASSWORD)
            page.wait_for_timeout(1000)
            page.click("button[name='password-submit']")
            page.wait_for_timeout(5000)
            page.screenshot(path="screenshot_4_after_password.png")
            print("Password submitted, waiting for OTP...")

            # Step 3: Fetch OTP from personal iCloud email (forwarded from work Outlook)
            print("Waiting for OTP email to arrive and forward to iCloud...")
            time.sleep(15)  # Extra buffer for Outlook auto-forward delay

            otp_code = None
            for attempt in range(10):
                print(f"Polling iCloud for OTP email, attempt {attempt + 1}/10...")
                try:
                    mail = imaplib.IMAP4_SSL("imap.mail.me.com", 993)
                    mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                    mail.select("INBOX")

                    # Search for unread emails from Trimble's no-reply address
                    cutoff = (datetime.now() - timedelta(minutes=5)).strftime("%d-%b-%Y")
                    _, search_data = mail.search(None, f'(UNSEEN FROM "no-reply@account.trimble.com" SINCE "{cutoff}")')

                    email_ids = search_data[0].split()
                    if email_ids:
                        # Grab the most recent one
                        _, msg_data = mail.fetch(email_ids[-1], "(RFC822)")
                        raw_email = msg_data[0][1]
                        parsed = email_lib.message_from_bytes(raw_email)

                        # Extract body text
                        body = ""
                        if parsed.is_multipart():
                            for part in parsed.walk():
                                if part.get_content_type() == "text/plain":
                                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                                    break
                                elif part.get_content_type() == "text/html" and not body:
                                    body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        else:
                            body = parsed.get_payload(decode=True).decode("utf-8", errors="ignore")

                        match = re.search(r'\b(\d{6})\b', body)
                        if match:
                            otp_code = match.group(1)
                            # Mark as read so it won't be picked up on future runs
                            mail.store(email_ids[-1], '+FLAGS', '\\Seen')
                            print(f"OTP retrieved: {otp_code}")
                            mail.logout()
                            break

                    mail.logout()

                except Exception as e:
                    print(f"IMAP attempt {attempt + 1} failed: {e}")

                time.sleep(8)  # Wait 8s between polls (~80s total window across 10 attempts)

            if not otp_code:
                raise Exception("Failed to retrieve OTP from iCloud after retries.")

            # Step 4: Enter OTP
            page.fill("input[name='code']", otp_code)
            page.wait_for_timeout(2000)
            page.click("button#enter_verification_code_submit")
            page.wait_for_load_state("networkidle", timeout=60000)
            page.wait_for_timeout(3000)
            page.screenshot(path="screenshot_5_after_login.png")
            print("Login complete")

        # Step 5: Password
            

        except Exception as e:
            print(f"FAILED at login form: {e}")
            page.screenshot(path="screenshot_error.png")
            browser.close()
            raise

        try:
            print("Navigating to report...")
            page.goto("https://app.e-builder.net/da2/Reports/ReportResults.aspx?ReportID={0769ebc1-6c4f-4f04-8858-20f33be1c6df}", timeout=60000)
            page.wait_for_timeout(5000)
            print(f"Report URL: {page.url}")
            page.screenshot(path="screenshot_6_report.png")
            print("Screenshot 6 saved.")
        except Exception as e:
            print(f"FAILED navigating to report: {e}")
            page.screenshot(path="screenshot_error.png")
            raise

        try:
            print("Clicking Print View and capturing download...")
            page.wait_for_timeout(3000)
            with page.expect_download(timeout=30000) as download_info:
                page.evaluate("document.getElementById('ctl00_ctl00_ContentPlaceHolder1_contentSection_btnPrintView1').click()")
            download = download_info.value
            
            # Save as html first since that's what it really is
            html_path = "report_raw.html"
            download.save_as(html_path)
            print("Raw file saved.")
        
            file_path = f"{datetime.now().strftime('%Y%m%d')} Systemwide Daily Material Placement Summary.xls"
            download.save_as(file_path)
            print(f"File saved as: {file_path}")
            
        except Exception as e:
            print(f"FAILED clicking button: {e}")
            page.screenshot(path="screenshot_error.png")
            browser.close()
            raise

        try:
            print("Logging out...")
            page.click("div.sign-out")
            page.wait_for_load_state("networkidle", timeout=30000)
            print("Logged out.")
        except Exception as e:
            print(f"Logout failed (non-critical): {e}")

        browser.close()
        return file_path

def send_email(file_path):
    tomorrow = datetime.now() + timedelta(days=1)
    today = tomorrow.strftime("%m/%d/%Y")
    date = tomorrow.strftime("%A, %B %-d, %Y")

    recipients = [
        "rwhite@bravocoeng.com"
    ]

    msg = MIMEMultipart()
    msg["Subject"] = f"Systemwide Daily Material Placement Summary for {today}"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = ", ".join(recipients)

    body = MIMEText(f"Good evening,\nAttached please find a systemwide daily material placement summary for {date}. The summary includes MPN's submitted by 5:30 PM.")
    msg.attach(body)

    with open(file_path, "rb") as f:
        attachment = MIMEBase("application", "vnd.ms-excel")
        attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", f"attachment; filename={file_path}")
        msg.attach(attachment)

    with smtplib.SMTP("smtp.mail.me.com", 587) as server:
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        # server.sendmail(EMAIL_ADDRESS, recipients, msg.as_string())

if __name__ == "__main__":
    file_path = click_button()
    send_email(file_path)
