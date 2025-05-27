## KTM WhatsApp Billing Reminder Bot

This bot is used to automatically send billing reminder messages for monthly dues and arisan (rotating savings) via WhatsApp, using Google Sheets as the data source.

## Features
1. Automatically send messages via WhatsApp
2. GUI built with Tkinter
3. Google Sheets integration
4. Updates billing status in spreadsheet (Done / No Dues)

## Prerequisites
1. Python version required (e.g., Python 3.8+)
2. Google Cloud setup (how to create and download credentials.json)

## How to Run
1. Install dependencies: pip install -r requirements.txt
2. Add your credentials.json file from Google Cloud
3. Run the app: python main.py

## Build Installer
1. Install pyinstaller
2. Run the following command: pyinstaller --onefile --noconsole --add-data "credentials.json;." --icon=assets/icon.ico main.py