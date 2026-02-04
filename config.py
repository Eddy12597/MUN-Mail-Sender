import json5
from pathlib import Path
from colorama import Fore, Style
from dotenv import load_dotenv
import os

CONFIG: dict[str, str] = {}
with open(str(Path('./settings.jsonc')), encoding="utf-8") as f:
    CONFIG = json5.load(f)

if len(CONFIG.keys()) == 0:
    print(f"{Fore.RED}{Style.BRIGHT}CONFIG NOT LOADED. Exiting...{Style.RESET_ALL}")
    raise SystemExit

if not (load_dotenv(CONFIG['school-email-password-path'])):
    print(f"{Fore.YELLOW}Did you forget to add a .env file with the email password?{Style.RESET_ALL}")
    raise SystemExit

password: str | None = os.getenv(CONFIG['school-email-password-question'])
if not password: 
    print(f"{Fore.YELLOW}Email password format incorrect in .env file {Style.RESET_ALL}\nShould be: '{CONFIG['school-email-password-question']}=<password>'")
    raise SystemExit