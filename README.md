# Sentiment_Analysis_ai
Analyzing the customers support messages and tagging them through ai.
Support Chat Analyzer (Google Sheets + OpenAI)

Per-channel **sentiment**, **tags** (AI-only, constrained taxonomy), and **short summary** (+ solved?).  
Auto-resumable batches, outputs **AI_Chat_Analysis**, rollups **AI_Tag_Summary** and **AI_Query_Subcategories**.

## What it does
- Groups rows by `channel_identifier`, reads `body`
- Calls OpenAI once per channel (`gpt-4o-mini`)  
- Writes: `channel_identifier | sentiment | tags | summary (Solved: yes/no/unclear)`
- Resumes automatically in safe chunks (time-based triggers)
- Summaries: per-tag totals + query subcategories

## Setup (in Google Sheets)
1. Open your sheet with raw data (headers must include `channel_identifier` and `body`).
2. **Extensions → Apps Script** → paste `src/Code.gs` into `Code.gs`.
3. **Project Settings → Script properties**: add `OPENAI_API_KEY = sk-...`
4. In code, set:
   ```js
   const SOURCE_SHEET_NAME = 'RAW_DATA'; // your input tab name
