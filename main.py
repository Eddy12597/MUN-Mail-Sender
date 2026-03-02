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

    if DEBUG or server is None:
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


def get_email_html(name: str, committee: str, country: str) -> html_str:
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
        .replace("__LOGO_URL__", CONFIG.get("logo-url",\
            CONFIG.get("official-website-url-base", "https://biphmun.org") + "/assets/logos/logo-rect-text-white.png"
        ))


delegate_reg_path = Path(CONFIG['delegate-registration-spreadsheet-path'])
df = pd.read_excel(str(delegate_reg_path))

print(df.columns)

email_address_question_name = CONFIG.get("email-address-question-name", "Email Address")

if email_address_question_name not in df.columns:
    raiseerror("Missing 'Email Address' column in registration spreadsheet")

# --- Deduplicate registrations by Email ---
dupes = df[df.duplicated(subset=[email_address_question_name], keep=False)]

if not dupes.empty:
    print(f"{Fore.YELLOW}Found duplicate registrations by email:{Style.RESET_ALL}")
    cols_to_show = [
        "Completion time",
        CONFIG['preferred-name-question-name'],
        CONFIG['wechat-id-question-name'],
        CONFIG['committee-preference-question-name'],
    ]

    for email, group in dupes.groupby(email_address_question_name):
        print(f"\n{Fore.CYAN}Email: {email}{Style.RESET_ALL}")

        for idx, row in group.iterrows():
            completion_time = row[cols_to_show[0]]
            name = row[cols_to_show[1]]
            wechat = row[cols_to_show[2]]
            committees = row[cols_to_show[3]]
            
            print(f"{Style.BRIGHT}{Fore.BLUE}{completion_time}{Style.RESET_ALL} | Name: {name}{Style.RESET_ALL} | {Fore.GREEN}WX: {wechat}{Style.RESET_ALL}")
            print(f"{Fore.MAGENTA}Committees: {Style.RESET_ALL}")

            if pd.notna(committees):
                committee_list = committees.split(";")
                for committee in committee_list:
                    print(f"- {committee.strip()}")
            else:
                print("- No committees listed")
            print() 

        choice = None
        while choice not in ("1", "2", "s"):
            choice = input(
                f"Select which row to KEEP for {email} "
                f"(1 = first, 2 = last, s = skip/keep all): "
            ).strip().lower()

        if choice == "1":
            df = df.drop(group.index[1:])  # keep first, drop the rest
        elif choice == "2":
            df = df.drop(group.index[:-1])  # keep last, drop the rest
        elif choice == "s":
            print(f"{Fore.YELLOW}Keeping all entries for {email}{Style.RESET_ALL}")

# Reset index so iterrows works cleanly later
df = df.reset_index(drop=True)


configpath = Path(CONFIG['config-committee-and-country-list-path']) # for committees' country lists
with open(str(configpath), 'r', encoding="utf-8") as file:
    config = json5.load(file)
committees_list: list[str] = list(config.keys())

# Create a per-committee remaining country list (config values are ISO2 codes now)
committees_remaining: dict[str, list[str]] = {}
for com in committees_list:
    # make a shallow copy of the list so we can remove assigned countries
    committees_remaining[com] = [c.strip().upper() for c in config[com] if str(c).strip()]


log_file = Path("assignments_log.csv")
if not log_file.exists():
    with open(log_file, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Name", "Email", "Committee", "CountryCode", "CountryName", "WeChatID", "Status"])


with open(str(Path(CONFIG['config-country-to-code'])), "r", encoding='utf-8') as f:
    convert = json5.load(f)

# Note: This is the stripped result (no parenthesis)
invalid_committees = {"Security Council"} 

for index, row in df.iterrows():
    # --- Skip & Log Withdrawn Participants ---
    withdraw_col = CONFIG.get('withdraw-question-name', 'withdraw')
    if withdraw_col in df.columns:
        val = str(row[withdraw_col]).strip().lower()
        if val in ("yes", "true", "1"):
            name = row[CONFIG['preferred-name-question-name']]
            email = row[email_address_question_name]
            print(f"{Fore.CYAN}Skipping withdrawn participant: {name} ({email}){Style.RESET_ALL}")
            
            with open(log_file, "a", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow([name, email, "N/A", "N/A", "N/A", row[CONFIG['wechat-id-question-name']], "WITHDRAWN"])
            continue
    
    # 1) parse committee preference list
    committees_raw: list[str] = str(row[CONFIG['committee-preference-question-name']]).strip().split(";")
    
    if len(committees_raw) == 1:
        committees_raw = committees_raw[0].split("；")
        if len(committees_raw) == 1 and not committees_raw[0]:
             continue

    committee_choices: list[str] = []
    for comr in committees_raw:
        comr = comr.strip()
        if not comr:
            continue
        match = re.match(r"^(.*?)\s*\(", comr.strip())
        name = match.group(1).strip() if match else comr.strip()
        if name not in config:
            raiseerror(f"Committee not recognized: {name}\n\tat index {index}")
        committee_choices.append(name)

    # 2) parse preferred countries column
    preferred_raw = str(row.get(CONFIG['preferred-country-question-name'], '')).strip()
    preferred_list: list[str] = []
    if preferred_raw:
        prefs = [p.strip().upper() for p in re.split(r"[;,]", preferred_raw) if p.strip()]
        preferred_list = prefs[:3]

    # 3) assign committee (Filtering out invalid ones)
    assigned_committee = None
    for choice in committee_choices:
        # Check if the committee is canceled
        if choice in invalid_committees:
            print(f"{Fore.YELLOW}Skipping canceled committee '{choice}' for {row[CONFIG['preferred-name-question-name']]}{Style.RESET_ALL}")
            continue
            
        # Check if the committee has spots left
        if committees_remaining.get(choice): 
            assigned_committee = choice
            break
            
    if assigned_committee is None:
        print(f"{Fore.RED}No available or valid committees for {row[CONFIG['preferred-name-question-name']]} at row {index}{Style.RESET_ALL}")
        # Log failure and move to next person instead of crashing the whole script
        with open(log_file, "a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([row[CONFIG['preferred-name-question-name']], row[email_address_question_name], "NONE", "N/A", "N/A", row[CONFIG['wechat-id-question-name']], "FAILED_NO_COMMITTEE"])
        continue
        
    # 4) assign country
    assigned_code = None
    if preferred_list:
        for code in preferred_list:
            if code in committees_remaining[assigned_committee]:
                assigned_code = code
                committees_remaining[assigned_committee].remove(code)
                print(f"{Fore.GREEN}Assigned preferred country: {code}{Style.RESET_ALL}")
                break
                
    if not assigned_code:
        if not committees_remaining[assigned_committee]:
             continue # Safety check
        assigned_code = committees_remaining[assigned_committee].pop(0)
        print(f"{Fore.BLUE}Assigned first available country: {assigned_code}{Style.RESET_ALL}")

    assigned_country_name = convert.get(assigned_code, assigned_code)

    # Send email
    email_success = send_email(
        row[CONFIG['email-address-question-name']],
        get_email_html(row[CONFIG['preferred-name-question-name']], assigned_committee, assigned_country_name),
        attachments=[s.strip() for s in CONFIG['attachments-filenames-list']],
        server=server
    )
    
    status = "DEBUG" if DEBUG else ("SENT" if email_success else "FAILED")
    
    with open(log_file, "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            row[CONFIG['preferred-name-question-name']],
            row[CONFIG['email-address-question-name']],
            assigned_committee,
            assigned_code,
            assigned_country_name,
            row[CONFIG['wechat-id-question-name']],
            status
        ])

if server:
    server.quit()