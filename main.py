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
from typing import Literal

from config import *

# Cutoff in excel
CUTOFF_LENGTH=199

for k, v in CONFIG.items():
    if len(v) > 199:
        CONFIG[k] = v[0:CUTOFF_LENGTH-3] + "..."
    
def raiseerror(s: str = ""):
    raise RuntimeError(s)

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
        
year = int(CONFIG['year'])

if datetime.datetime.now().year != year:
    print(f"Year and current date do not match! Please check source file: {os.path.abspath(__file__)}")
    if input(f"Proceed with {datetime.datetime.now().year}? [N/y] ") != "y":
        print(f"{Fore.RED}Please change source file: {os.path.abspath(__file__)}{Style.RESET_ALL}")
        raise SystemExit
    else:
        year = datetime.datetime.now().year
        print(f'{Fore.BLUE}Updated to {year}{Style.RESET_ALL}')

server: smtplib.SMTP | None = None
if not DEBUG:
    try:
        print("Logging in...")
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(CONFIG['school-email'], password)
        print("Login Successful")
    except (smtplib.SMTPAuthenticationError, ConnectionError) as e:
        print(f"{Fore.MAGENTA}Login Failed: {e}. Check .env credentials.{Style.RESET_ALL}")
        raise SystemExit


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

type html_str=str

def send_email(to: str, body_html: html_str, server: smtplib.SMTP | None, subject: str = "BIPHMUN Registration Confirmation", sender: str = CONFIG['school-email'], attachments: list[str] | None = None) -> bool:
    body_text = html_to_markdown(body_html)
    msg = MIMEMultipart("alternative")
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = to
    msg['Cc'] = CONFIG['my-email']

    msg.attach(MIMEText(body_text, "plain", "utf-8"))
    msg.attach(MIMEText(body_html, "html", "utf-8"))
    
    if attachments:
        for f in attachments:
            try:
                with open(f, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename= {f}')
                    msg.attach(part)
            except FileNotFoundError:
                print(f"{Fore.RED}File {f} not found.{Style.RESET_ALL}")
                if input("Skip? [Y/n]") in ("n", "no", "0"):
                    raise SystemExit
    
    print(f"{Fore.MAGENTA}Sending email to {to}:{Style.RESET_ALL}\n---\n{Style.BRIGHT}Subject: {subject}{Style.RESET_ALL}\n\nPreview:\n\n{body_text}\n")
    
    if DEBUG or server is None:
        outbox_dir = Path("./outbox")
        outbox_dir.mkdir(exist_ok=True)
        outbox_file = outbox_dir / f"{to.replace('@','_at_')}.eml"
        with open(outbox_file, "w", encoding="utf-8") as f:
            f.write(msg.as_string())
        print(f"{Fore.GREEN}Saved to outbox: {outbox_file}{Style.RESET_ALL}")
        return True 

    confirm = input("Send this email? [y/N] ").lower().strip()
    if confirm == "y":
        try:
            server.send_message(msg)
            print(f"{Fore.GREEN}Email sent!{Style.RESET_ALL}")
            return True
        except smtplib.SMTPException as e:
            print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
            choice = input(f"{Fore.YELLOW}Email failed. [S]kip, [R]etry, [A]bort? ").lower().strip()
            if choice == 'r':
                try:
                    server.send_message(msg)
                    return True
                except:
                    return False
            elif choice == 'a':
                raise SystemExit
            return False
    return False

def get_email_html(name: str, committee: str, country: str, school: str | None = None) -> html_str:
    html = """
<div id="main" style="font-family: Arial; font-size: 20px; display: flex; align-items: center; flex-direction: column;background-color: #0a4d7f">
    <div id="graphics" style="font-family: Arial; font-size: 20px; width: 80%">
        <img alt="logo" id="logo" src="__LOGO_URL__" style="width: 50%; margin-left: -3%; border-radius: 8px; margin-top: 15%">
    </div>
    <div id="letter" style="font-family: Arial; font-size: 20px; margin-left: 10%; margin-right: 10%">
        <p style="color: #fff">Dear __NAME__,</p>
        <p style="color: #fff">Your registration for BIPHMUN __YEAR__ is confirmed.</p>
        <p style="color: #fff">Your committee is <b>__COMMITTEE__</b>, and your country is <b>__COUNTRY__</b>.</p>
        <p style="color: #fff">Please visit our official MUN website for extra references and resources: </p>
            <a href="__WEBSITE__" style="color: orange">__WEBSITE__</a>
        __OPTIONS__
        <p style="color: #fff">Also, please be noted that transportation to and from the event <b>will not be provided</b>. We apologize for any inconveniences caused and your understanding.</p>
        <p style="color: #fff">Best Regards,</p>
        <p style="color: #fff; margin-bottom: 150px;">BIPH MUN Team</p>
    </div>
</div>
"""
    return html\
        .replace("__NAME__", name)\
        .replace("__COMMITTEE__", committee)\
        .replace("__COUNTRY__", country)\
        .replace("__YEAR__", str(year))\
        .replace("__WEBSITE__", CONFIG['official-website-url-base'])\
        .replace("__LOGO_URL__", CONFIG.get("logo-url", CONFIG.get("official-website-url-base", "https://biphmun.org") + "/assets/logos/logo-rect-text-white.png"))\
        .replace("__OPTIONS__", "" if school is None else '<p style="color: #fff">For BIPH students: Training will be provided on March 11st and March 18th, in A207. All participants are welcome. We will cover basic procedures and guidance on research.</p>')

# --- Processing Starts ---
delegate_reg_path = Path(CONFIG['delegate-registration-spreadsheet-path'])
df = pd.read_excel(str(delegate_reg_path))

email_address_question_name = CONFIG.get("email-address-question-name", "Email Address")
if email_address_question_name not in df.columns:
    raiseerror("Missing 'Email Address' column")

# Deduplicate
dupes = df[df.duplicated(subset=[email_address_question_name], keep=False)]
if not dupes.empty:
    for email, group in dupes.groupby(email_address_question_name):
        choice = input(f"Select row to KEEP for {email} (1=first, 2=last, s=all): ").strip().lower()
        if choice == "1": df = df.drop(group.index[1:])
        elif choice == "2": df = df.drop(group.index[:-1])

# --- PRIORITY LOGIC ---
# Create a helper column to flag if they have a country preference
pref_col = CONFIG['preferred-country-question-name']
df['has_preference'] = df[pref_col].apply(lambda x: pd.notna(x) and str(x).strip().lower() != "nan" and str(x).strip() != "")

# Sort: True (1) comes before False (0) when using ascending=False
df = df.sort_values(by='has_preference', ascending=False).reset_index(drop=True)
print(f"{Fore.CYAN}Prioritizing {len(df[df['has_preference']])} delegates with country preferences.{Style.RESET_ALL}")

configpath = Path(CONFIG['config-committee-and-country-list-path'])
with open(str(configpath), 'r', encoding="utf-8") as file:
    config_json = json5.load(file)

committees_remaining = {com: [c.strip().upper() for c in config_json[com] if str(c).strip()] for com in config_json.keys()}

log_file = Path("assignments_log.csv")
if not log_file.exists():
    with open(log_file, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Name", "Email", "Committee", "CountryCode", "CountryName", "WeChatID", "Status"])

with open(str(Path(CONFIG['config-country-to-code'])), "r", encoding='utf-8') as f:
    convert = json5.load(f)

invalid_committees = {"Security Council"} 

for index, row in df.iterrows():
    # Withdraw check
    withdraw_col = CONFIG.get('withdraw-question-name', 'withdraw')
    if withdraw_col in df.columns and str(row[withdraw_col]).strip().lower() in ("yes", "true", "1"):
        continue
    
    # Parse committee choices
    committees_raw = str(row[CONFIG['committee-preference-question-name']]).strip().split(";")
    committee_choices = []
    for comr in committees_raw:
        match = re.match(r"^(.*?)\s*\(", comr.strip())
        name = match.group(1).strip() if match else comr.strip()
        if name and name not in invalid_committees and name in config_json:
            committee_choices.append(name)

    # Parse country preferences
    preferred_raw = str(row.get(pref_col, '')).strip()
    preferred_list = [p.strip().upper() for p in re.split(r"[;,]", preferred_raw) if p.strip() and p.lower() != "nan"][:3]

    assigned_committee = None
    assigned_code = None

    for choice in committee_choices:
        if committees_remaining.get(choice): 
            assigned_committee = choice
            temp_code = None
            
            # Try preference
            for pref_code in preferred_list:
                if pref_code in committees_remaining[assigned_committee]:
                    temp_code = pref_code
                    committees_remaining[assigned_committee].remove(pref_code)
                    break
            
            # Fallback
            if not temp_code:
                temp_code = committees_remaining[assigned_committee].pop(0)
            
            assigned_code = temp_code
            break
            
    if assigned_committee:
        assigned_country_name = convert.get(assigned_code, assigned_code)
        email_success = send_email(
            row[email_address_question_name],
            get_email_html(row[CONFIG['preferred-name-question-name']], assigned_committee, assigned_country_name, row[CONFIG['school-question-name']]),
            attachments=[s.strip() for s in CONFIG['attachments-filenames-list']],
            server=server
        )
        
        with open(log_file, "a", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow([row[CONFIG['preferred-name-question-name']], row[email_address_question_name], assigned_committee, assigned_code, assigned_country_name, row[CONFIG['wechat-id-question-name']], "SENT" if email_success else "FAILED"])

if server: server.quit()