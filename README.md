# Microsoft Rewards Bot

[![License](https://img.shields.io/badge/license-MIT-green.svg?style=flat)](LICENSE)
[![Version](https://img.shields.io/badge/version-v2.0-blue.svg?style=flat)](#)

An automated solution for earning daily Microsoft Rewards points using Python and Selenium.

## Installation

- Clone the repo.
- Install requirements with the following command :

```
pip install -r requirements.txt
```

- Make sure you have Chrome installed.
- Edit the accounts.json.sample with your accounts credentials and rename it by removing .sample at the end. If you want to add more than one account, the syntax is the following (mobile_user_agent is optional and goal is only required if you enable auto-redeem or it will default to Amazon):

```json
[
  {
    "username": "Your Email",
    "password": "Your Password",
    "mobile_user_agent": "Your Preferred Mobile User Agent",
    "goal": "Amazon"
  },
  {
    "username": "Your Email",
    "password": "Your Password",
    "mobile_user_agent": "Your Preferred Mobile User Agent",
    "goal": "Xbox Game Pass Ultimate"
  }
]
```

- Edit email.json.sample with your GMAIL email credentials. Visit https://myaccount.google.com/apppasswords after enabling 2FA to get the required password and rename it by removing .sample at the end. You can also disable certain alerts in this file if you want to. The syntax is the following:

```json
[
  {
    "sender": "sender@example.com",
    "password": "GoogleAppPassword",
    "receiver": "receiver@example.com",
    "withdrawal": "true",
    "lock": "true",
    "ban": "true",
    "phoneverification": "true",
    "proxyfail": "false"
  }
]
```

- Due to limits of Ipapi sometimes it returns error and it causes bot stops. You can define a default language and location to prevent it.
- Run the script.
  - Optional arguments:
    - `--headless ` You can use this argument to run the script in headless mode.
    - `--session ` Use this argument to create session for each account.
    - `--everyday TIME` This argument takes time in 24h format (HH:MM) to run it everyday at the given time by leaving the program open.
    - `--fast` Reduce delays where ever it's possible to make script faster.
    - `--error` Display errors when app fails.
    - `--accounts` Add accounts (email1:password1 email2:password2..).
    - `--proxies` Add proxies (proxy1 proxy2..). Proxies that require authentication should follow this format -> **hostname:port:username:password**.
    - `--authproxies` Use this argument to indicate that your proxies require authentication. **NOTE for Windows users**: headless mode is not supported when using this argument. **NOTE for Linux Server users**: install **xvfb** package if you are running the script on a Linux server, otherwise the script won't run when using this argument.
    - `--privacy` Enable privacy mode.
    - `--emailalerts` Enable GMAIL email alerts.
    - `--redeem` Enable auto-redeem rewards based on accounts.json goals.
  - If you run the script normally it asks you for input instead.

## Features

- Bing searches (Desktop, Mobile and Edge) with User-Agents.
- Complete automatically the daily set.
- Complete automatically punch cards.
- Complete automatically the others promotions.
- Headless Mode.
- Multi-Account Management (Config and command-line).
- Worflow for automatic deployement.
- Modified to be undetectable as bot.
- If it faces an unexpected error, it will reset and try again that account.
- Save progress of bot in a log file and use it to pass completed account on the next start at the same day.
- Detect suspended accounts.
- Detect locked accounts.
- Detect unusual activites.
- Uses time out to prevent infinite loop
- You can assign custom user-agent for mobile like above example.
- Set clock to start it at specific time.
- For Bing search it uses random word at first try and if API fails, then it uses Google Trends.
- Auto-redeem gift cards.
- Email alerts when an account is locked, banned, phone verification is required for a reward or a reward has been automatically redeemed.

## Warning

- Don't use outlook mail with Microsoft Rewards. They are more likely to get banned.
- Don't run the script with more than six accounts per IP.

## Troubleshooting

### If the script does not work as expected, please check the following things before opening a new issue.

- Is Chrome installed? This must also be the case for "Console only" execution of the script.
- Is Python installed? Please install a Python Version 3.10 or higher. Also dont forget to add the new Python version as environment variable too.
- For Systems without GUI, use --headless parameter to run it.
- Don't forget to install the dependencies from the "requirements.txt".

## Credits

- [@charlesbel](https://github.com/charlesbel) The original author of the repo.
- [@Farshadz1997](https://github.com/farshadz1997) For adding a bunch of features.
- [@Prem-ium](https://github.com/Prem-ium) For part of the code that lets the script auto-redeem gift cards.
