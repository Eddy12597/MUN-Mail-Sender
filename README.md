# BIPHMUN Registration Automation

Open this in a Markdown Viewer or online (https://markdownlivepreview.com/), then paste all of the following into the viewer.

## Description

This script automates the process of sending confirmation emails to MUN delegates. It reads registration data from a spreadsheet, assigns committees and countries based on preferences and availability, and sends personalized emails.

Files:
- `main.py`: the core code
- `country_code.json`: maps 2-letter country codes to full country name
- `settings.jsonc`: configuration settings for the code
- `README.md`: the file you're reading right now
- `assignments_log.csv`: the assignments (committee + country) log with email status
- `wechat-qr-code.jpg`: the wechat qr code. **Important**: remember to update when you send the emails.

## Notes

1. The registration form **should minimally have the following components** (doesn't necessarily have to be worded the same; you can change that in `config.json`) for this to work:
- Completion Time
- Preferred Name
- Email Address
- WeChat ID
- Committee Preference
- Preferred Country

2. You can add **attachments** by modifying `settings.jsonc`.

```json
{
    ...,
    // --- CHANGE WECHAT GROUPCHAT QR CODE FILENAME ---
    "attachments-filenames-list": ["wechat-qr-code.jpg", "some-other-file.pdf"], // etc, if needed
    ...,
}
```

3. The system tracks whether each email is successfully sent or not, stored in `assignments_log.csv`. If successfully sent, it will say `SENT`; if running in debug mode, it will say `DEBUG`, and if failed, it will say `FAILED`.

## Installation / Run

0. Have python installed

You can go to ```https://python.org/downloads``` to install the correct python on your device. **Important: choose a version between 3.11 and 3.13: 3.11.0 <= v <= 3.13.2.** The code uses newer python features that may not work on older versions, and python 3.14 could bring breaking changes that may cause this script to not work. If you would like to help with refactoring for newer versions, raise a pull request or email eddy12597@163.com.

1. Open Terminal
- Windows: `Win + R` -> type `powershell` -> `cd /path/to/mun-mail-sender`
- MacOS: `Cmd + Space` -> type `Terminal` -> `cd /path/to/mun-mail-sender`

2. Check python is working & Install necessary packages
```bash
python --version # should print 'Python 3.13.x' or something
pip --version # should print 'pip 25.x from /path/to/Python/Python313/...' or something

# install dependencies
pip install pandas colorama json5 python-dotenv
```

3. Change configuration (settings) info if needed in `.env` and `settings.jsonc`.

**NOTE: NEVER PUT THE PASSWORD TO `settings.jsonc` or any file other than `.env`.**

Create a new file called `.env` under the current folder, and enter the school email password:

```bash
# Windows
notepad .env # then hit 'Enter'

# MacOS
textedit .env
```
Then enter the password in `.env`:
```env
EMAIL_PASSWORD=<email-password>
```
*Note*: there are no spaces around `=`.

Also edit `settings.jsonc`:
```json
{
    // --- CONTROLS WHETHER IT ACTUALLY RUNS OR ITS A TEST RUN. True = test run, False = ACTUAL run
    "DEBUG": "True",
    // --- CHANGE THE YEAR AND YOUR EMAIL ---
    "year": "2025",
    "my-email": "eddy12597@163.com", // your email would be Cc'ed

    // --- CHANGE WECHAT GROUPCHAT QR CODE FILENAME ---
    "attachments-filenames-list": ["wechat-qr-code.jpg"],

    // --- YOU WILL NEED TO INPUT THE PASSWORD IN THE .env FILE. If you do not know the password, ask me or ask the club sponsor ---
    
    
    // --- CHANGE THIS TO YOUR PATH ---
    "delegate-registration-spreadsheet-path": "./test.xlsx",//"d:/Downloads/BIPHMUN 2025 Individual Registration Form(1-48) (2).xlsx",
    

    // --- CHANGE THESE IF REG FORM IS CHANGED ---
    "preferred-name-question-name": "Preferred Name",
    "email-address-question-name": "Email Address",
    "wechat-id-question-name": "Wechat ID",
    "committee-preference-question-name": "Please rank your committee preference",
    "preferred-country-question-name": "Preferred country",

    // --- USUALLY YOU DONT NEED TO CHANGE THESE ---
    "config-committee-and-country-list-path": "./config.json",
    "config-country-to-code": "./country_code.json",
    "official-website-url-base": "https://biphmun.netlify.app", 
    "school-email": "mun-biph@basischina.com",
    "school-email-password-path": "./.env",
    "school-email-password-question": "EMAIL_PASSWORD"
}
```

4. Modify committee and/or country list in `config.json`

The countries should be in **ISO 2-Letter Country Code.** This could be found online or via LLMs, or via `country_code.json`.

5. Assign preferred committees to delegates in club

If delegates in club have preferred committees, add a new column in the registration spreadsheet named 'Preferred country'. Add the preferred countr(ies) in **ISO 2-Letter Country Code.** Multiple preferred countries are separated by semicolons (`;`).

### Test run

6. Run the program in debug mode to test
    
Just run the following command. This will default to DEBUG mode.
```bash
python main.py
```
It should print, in green, "Running in DEBUG mode".

When prompted to resolve duplicated delegates, choose as appropriate. When running in ACTUAL mode, if the duplicates have same committee preferences, choose the second one (later in time). If they have different committee preferences, add the delegate's WeChat and confirm if needed.

Then, clean up the log files.

```bash
rm -r outbox
rm assignments_log.csv
```

7. Test (Highly Recommended)

*If this breaks, send me (eddy12597@163.com) the terminal output and the output in `./assignments_log.csv` (will appear after you run it). If there isn't much time, send emails manually. The confirmation email should be sent around 1 month in prior to the conference.*

Modify the `test.xlsx` in the current folder so it sends to your email and your WeChat ID, etc.

Then change `settings.jsonc`. Do not replace the actual portion with this snippet. Just change the `delegate-registration-spreadsheet-path` and the `DEBUG` variable. Remember to add the commas.
```json
{
    "DEBUG": "False", // from "True"
    ...,
    // change as needed
    "delegate-registration-spreadsheet-path": "./test.xlsx",
    ..., 
}
```

Then run the following command:
```bash
python main.py
```

You will be prompted to confirm. For every email it attempts to send, it gets confirmation from you. See if it only sends the email to you. If it does and you receive it, see if it renders correctly. You should see a Logo, a blue background, and a confirmation text.

Now check the country and committee. Since there is only one delegate in the test spreadsheet, it should assign you to your chosen committee and your chosen country.

Then, remove the log files:

```bash
rm -r outbox
rm assignments_log.csv
```

### Actual Run

8. ACTUAL RUN

Change the `settings.jsonc` file. Again, do not replace the actual portion with this snippet. Just change the `delegate-registration-spreadsheet-path`. Remember to add the comma after it.

```json
{
    ...,
    "delegate-registration-spreadsheet-path": "/path/to/the/spreadsheet",
    ...,
}
```

And run the program:

```bash
python main.py
```

When prompted to resolve duplicate delegates, choose as appropriate.

If an email fails to send, it wil prompt you to either:
- [S]kip: continue to the next email
- [R]etry: attempt to resend the email
- [A]bort: stop the entire email sending process

Check the `assignments_log.csv` file to see if all emails are successfully sent. 

## Troubleshooting

### Email features

If emails fail to send, check:
1. Internet connection
2. Typo in email (handle manually; add the delegate's wechat)
3. Toggle VPN status (turn off if on, turn on if off)
4. Server status with our school email system
5. `assignments_log.csv` for details

### Retry process

For failed emails in `assignments_log.csv`:
- If there aren't a lot of emails sent yet, you can try re-running the script
- If most of the emails are sent, you should manually send the emails in the log file.


Good Luck!


---

Creator: [Eddy Zhang](https://github.com/Eddy12597)

Our MUN GitHub Account: [BIPHMUN](https://github.com/BIPHMUN)