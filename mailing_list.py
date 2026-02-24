# reads from `form.xlsx`
# parses emails and year of graduation
# adds to mailing_list.json if year of graduation <= current year
# example: 
"""
{
    
    "mailing_list": [
        {
            "name": "Eddy",
            "email": "eddy12597@163.com",
            "year_of_graduation": "2027"
        },
        {
            "name": "Smith",
            "email": "smith@example.com",
            "year_of_graduation": "2026"
        }
    ]
}
"""
# sends emails and invitation in send()

import json
import jsonschema
from email import encoders
from email.mime.base import MIMEBase
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import pandas as pd
import json5
import re
import datetime
from colorama import Fore, Style
from html import unescape
import csv
from dotenv import load_dotenv
import tqdm
from config import *

# send_email functions directly copied from main.py
type html_str = str
def html_to_markdown(html: str) -> str:
    """Convert HTML email template into markdown/plaintext preview."""
    text = html
    text = re.sub(r"<\s*(p|div|br)\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"<\s*b\s*>(.*?)<\s*/b\s*>", r"**\1**", text, flags=re.I)
    text = re.sub(r'<a\s+href="(.*?)">(.*?)</a>', r"[\2](\1)", text, flags=re.I)
    text = re.sub(r"<(style|script).*?>.*?</\1>", "", text, flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", "", text)
    text = unescape(text)
    text = re.sub(r"\n\s*\n\s*", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = text.strip()

    return text
def send_email(to: str, body_html: html_str, server: smtplib.SMTP, subject: str = "BIPHMUN Registration Confirmation", sender: str = CONFIG['school-email'], attachments: list[str] | None = None) -> bool:
    # Generate plain-text fallback from HTML
    body_text = html_to_markdown(body_html)

    # Create multipart/alternative container
    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg['Cc'] = CONFIG['my-email']

    # Attach plain text first, then HTML
    msg.attach(MIMEText(body_text, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    
    # Attach attachments
    if attachments:
        for f in attachments:
            try:
                with open(f, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {f}'
                )
                msg.attach(part)
            except FileNotFoundError:
                print(f"{Fore.RED}File {f} not found.{Style.RESET_ALL}")
                if input("Skip? [Y/n]") in ("n", "no", "0"):
                    raise SystemExit
    

    # Show preview in console (plaintext)
    print(f"{Fore.MAGENTA}Sending email to {to}:{Style.RESET_ALL}\n---\n{Style.BRIGHT}Subject: {subject}{Style.RESET_ALL}\n\nPreview:\n\n{body_text}\n")
    if attachments:
        print(f"{Fore.MAGENTA}Attachments filenames: \n{Style.RESET_ALL}\t- {"\n\t- ".join(attachments)}")

    if DEBUG:
        outbox_dir = Path("./outbox")
        outbox_dir.mkdir(exist_ok=True)
        outbox_file = outbox_dir / f"{to.replace('@','_at_')}.eml"
        with open(outbox_file, "w", encoding="utf-8") as f:
            f.write(msg.as_string())
        print(f"{Fore.GREEN}Saved to outbox: {outbox_file}{Style.RESET_ALL}")
        return True  # CHANGED: Return success in debug mode

    confirm = input("Send this email? [y/N] ").lower().strip()
    if confirm == "y":
        try:
            server.send_message(msg) # type: ignore
            print(f"{Fore.GREEN}Email sent!{Style.RESET_ALL}")
            return True  # CHANGED: Return True on success
        except smtplib.SMTPException as e:
            print(f"{Fore.RED}{Style.BRIGHT}Error while sending email: {Style.NORMAL}{e}{Style.RESET_ALL}")
            
            failed_email = {
                'to': to,
                'subject': subject,
                'error': str(e)
            }
            
            choice = input(f"{Fore.YELLOW}Email failed. [S]kip, [R]etry, [A]bort? ").lower().strip()
            if choice == 'r':
                # Retry logic
                try:
                    server.send_message(msg)
                    print(f"{Fore.GREEN}Retry successful!{Style.RESET_ALL}")
                    return True  # CHANGED: Return True on retry success
                except smtplib.SMTPException as retry_error:
                    print(f"{Fore.RED}Retry also failed: {retry_error}{Style.RESET_ALL}")
                    return False  # CHANGED: Return False on retry failure
            elif choice == 'a':
                print(f"{Fore.RED}Aborting email sending process.{Style.RESET_ALL}")
                raise SystemExit
            else:
                print(f"{Fore.YELLOW}Skipping this email and continuing...{Style.RESET_ALL}")
                return False  # CHANGED: Return False when skipping
    else:
        print(f"{Fore.BLUE}Canceled{Style.RESET_ALL}")
        return False  # CHANGED: Return False when canceled

schema = {
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "required": ["mailing_list"],
  "properties": {
    "mailing_list": {
      "type": "array",
      "description": "List of mailing list entries",
      "items": {
        "type": "object",
        "required": ["email"],
        "properties": {
          "name": {
            "type": "string",
            "description": "Name of Recipient"  
          },
          "email": {
            "type": "string",
            "format": "email",
            "description": "Email address of the contact"
          },
          "year_of_graduation": {
            "type": "integer",
            "minimum": 1900,
            "maximum": 3000
          }
        },
        "additionalProperties": False
      }
    }
  },
  "additionalProperties": False
}

DEBUG = CONFIG['DEBUG'] not in ("False", "false")

curyear = datetime.datetime.now().year

df = pd.read_excel(CONFIG["delegate-registration-spreadsheet-path"])

if not password: 
    print(f"{Fore.YELLOW}Email password format incorrect in .env file {Style.RESET_ALL}\nShould be: '{CONFIG['school-email-password-question']}=<password>'")
    raise SystemExit

DEBUG = CONFIG['DEBUG'] not in ("False", "false")

if DEBUG:
    print(f"{Fore.GREEN}Running in DEBUG mode{Style.RESET_ALL}")
else:
    print(f"{Fore.RED}NOT IN DEBUG MODE{Style.RESET_ALL}")
    if input("Proceed? [N/y] ") == "y":
        print(f"{Fore.RED}Continuing in ACTUAL mode{Style.RESET_ALL}")
    else:
        print(f"{Fore.GREEN}Changed to DEBUG mode{Style.RESET_ALL}")
        DEBUG = True

# extracts email and grad year from df, returns mailing list json
def parse(df: pd.DataFrame = df, curyear = curyear) -> dict:
    maillist = {
        "mailing_list": [
            
        ]
    }
    for idx, row in df.iterrows():
        if (em := row[CONFIG["email-address-question-name"]]) and (gr := str(row[CONFIG["year-of-grad-question-name"]])) and (n := row[CONFIG["preferred-name-question-name"]]):
            if CONFIG["year-of-grad-type"] == "grade-level":
                gr = gr[gr.lower().index("g") + 1:]
                gr = int(gr)
                gr = curyear + (12 - gr)
            maillist["mailing_list"].append({
                "email": em,
                "year_of_graduation": int(gr),
                "name": n
            })
    return maillist

inv_html = \
"""\
<div id="main" style="font-family: Arial; font-size: 16px; display: flex; align-items: center; flex-direction: column;background-color: #0a4d7f; color: white">
    <div id="graphics" style="font-family: Arial; font-size: 18px; width: 80%">
        <img alt="logo" id="logo" src="__LOGO_URL__" style="width: 60%; margin-left: -3%; border-radius: 8px; margin-top: 15%">
    </div>
    <div id="letter" style="font-family: Arial; font-size: 18px; margin-left: 10%; margin-right: 10%">
        <p>Dear __NAME__,</p>
        <p>You were a member of BIPHMUN before, and we are writing to invite you to the next edition of the conference.</p>
        <p><b>BIPHMUN __YEAR__</b> will take place in <b>__CONF_DUR__</b>, and delegate registration is currently open. The deadline to register is <b>__REG_DDL__</b>.</p>
        <b>Conference Details:</b>
        <ul>
            <li>Dates: __CONF_DUR__</li>
            <li>Delegate Fee: __DEL_FEE__</li>
            <li>Website: <a href="__WEBSITE__" style="color: orange">__WEBSITE__</a></li>
            <li>Committees: __WEBSITE__/committees</li>
        </ul>
        <p >Registration Form (delegates): <a style="color: orange" href="__FORM__">__FORM__</a></p>
        <p>If you are interested in joining us again this year, we strongly encourage you to register as soon as possible, as the deadline is approaching.
        </p>
        <p>
            We appreciate the time and commitment you gave to BIPHMUN in the past, and we hope to welcome you again this __MONTH__.
        </p>
        
        <p>Best Regards,</p>
        <p>BIPH MUN Team</p>
        <a style="color: orange; font-size: 18px" href="__WEBSITE__">__WEBSITE__</a>
        <p style="margin-bottom: 150px;"></p>
    </div>
</div>
""".replace("__WEBSITE__", CONFIG["official-website-url-base"])\
    .replace("__YEAR__", CONFIG["year"])\
    .replace("__REG_DDL__", CONFIG["signup-ddl-delegate-text"])\
    .replace("__CONF_DUR__", CONFIG["conference-duration-text"])\
    .replace("__DEL_FEE__", CONFIG["delegate-fee"])\
    .replace("__FORM__", CONFIG["delegate-form-url"])\
    .replace("__MONTH__", CONFIG["conference-month"])\
    .replace("__LOGO_URL__", CONFIG.get("logo-url",\
        CONFIG.get("official-website-url-base", "https://biphmun.org") + "/assets/logos/logo-rect-text-white.png"
    ))


def send(maillist: dict, curyear = curyear) -> int:
    try:
        jsonschema.validate(instance=maillist, schema=schema)
    except jsonschema.ValidationError as ve:
        print(f"Format error in mailing list json: {ve.__repr__()}")
        raise
    
    rawlist: list[tuple[str, str]] = []
    
    for mailobj in maillist["mailing_list"]:
        if mailobj["year_of_graduation"] >= curyear:
            rawlist.append((mailobj["name"], mailobj["email"]))
    
    server = smtplib.SMTP('smtp.office365.com', 587)
    if not DEBUG:
        try:
            print("Logging in...")
            server.starttls()
            
            server.login(CONFIG['school-email'], password)
            print("Login Successful")
        except smtplib.SMTPAuthenticationError:
            print(f"{Fore.MAGENTA}Incorrect Credentials. Did you forget to enter the password in {Style.BRIGHT}.env{Style.NORMAL}?{Style.RESET_ALL}")
            raise SystemExit
    cnt = 0
    for name, email in rawlist:
        res = send_email(email, body_html=inv_html.replace("__NAME__", name), subject=f"BIPHMUN {CONFIG['year']} Invitation", server=server)
        if res:
            cnt += 1
    if not DEBUG: server.quit()
    return cnt


if __name__ == "__main__":
    mailing_list = parse()
    with open('mailing_list.json', "r", encoding="utf-8") as f:
        original = json.load(f)
    # Merge and remove duplicates based on email
    all_entries = mailing_list["mailing_list"] + original["mailing_list"]

    # Use a dictionary to keep unique emails
    unique_by_email = {}
    for entry in all_entries:
        email = entry["email"]
        if email not in unique_by_email:
            unique_by_email[email] = entry

    merged = {
        "mailing_list": list(unique_by_email.values())
    }
    with open('mailing_list.json', "w", encoding="utf-8") as f:
        json.dump(merged, f)
    
    print(merged)
    
    num = send(merged)
    print(f"Sent {num} emails")