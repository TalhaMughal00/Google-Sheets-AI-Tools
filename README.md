# Google Sheets AI QA Tools
Google Apps Script that adds AI-powered QA utilities to Google Sheets.
Uses OpenRouter with the Xiaomi MiMo V2 Flash model.

![image alt](https://github.com/TalhaMughal00/Google-Sheets-AI-Tools/blob/8f521167e711e05f8780b1800ec0c65d790b76a8/Tool.png)

## Features
- Rewrite QA bug descriptions for clarity
- Generate short titles for bug descriptions
- Sidebar UI inside Google Sheets
- Batch processing for large selections

## Tech Stack
- Google Apps Script
- OpenRouter API
- Model: xiaomi/mimo-v2-flash:free

## Setup
### 1. Create Apps Script Project
- Open Google Sheets
- Extensions → Apps Script
- Paste `Code.gs` content
- Save project

### 2. Set API Key
- Extensions → Apps Script
- Project Settings → Script Properties from sidebar
- Add script property:
- Click Save Script properties

## How to Use
- Select one or more cells with bug descriptions in your sheet.
- Open AI Tools from the menu bar or Toolbar.
- Click Rewrite Cells to improve clarity.
- Click Generate Titles to create short titles in the column to the left.
- Script will start, run all tasks, and finish on its own. It will rewrite the selected cells or generate the titles you chose.

## Getting Your API Key
- Sign in at openrouter.ai
- Generate a new API key in your dashboard.
- In Apps Script, go to Project Settings → Script Properties.
- Add a property with the correct key name:
  - Xiaomi_API_KEY
- Paste the key as the value and save.

**Notes:** This script caches rewritten cells to prevent duplicate processing and works best with clear, descriptive bug texts. Large selections may take a few moments to process, so be patient when handling many cells at once. Make sure your API key is valid to avoid authentication errors, and you can update the model or key at any time in Script Properties.
