## Google Apps Script Chatbot Report Generator

This project is a fully automated reporting pipeline that turns raw chatbot logs into an executive-ready monthly report. It runs on a set time trigger, automatically analyzing chat data, generating visualizations, and creating a professional AI-written summary in Google Docs.

The system provides a consistent, data-driven overview of chatbot performance every month, saving hours of manual reporting and producing clear, presentation-ready insights.

### Key Features

- Fully hands-free reporting: The script runs on a scheduled trigger that automatically generates each monthly report.
- AI-driven storytelling: OpenAI summarizes engagement trends with clear, human-like insight.
- Consistent data visualization: Automatically updates charts for hourly, daily, and weekday/weekend patterns.
- End-to-end automation: From data collection to finished report, all within the Google Workspace ecosystem.

### System Overview

```
Google Sheets (Chat History YYYY)
        ↓
Google Apps Script (Data Processing and Chart Generation)
        ↓
Google Drive (Plots and Report Storage)
        ↓
Google Docs (Template-based Monthly Report)
        ↓
OpenAI API (Narrative Generation)
```

### Workflow

#### Stage 1: Analyze and Plot 

- Automatically identifies the previous month to process
- Reads and cleans chat data from the monthly tab
- Calculates metrics such as unique sessions, total questions, and hourly trends
- Generates charts for Hourly, Daily, and Weekday/Weekend sessions
- Stores chart images in /Monthly Report/Plots

#### Stage 2: Synthesize and Report 

- Retrieves all stored metrics and previous month data
- Prompts OpenAI with structured metrics to produce a professional summary
- Copies the Monthly Report Template
- Inserts the generated text and charts into placeholders
- Saves the new report in Google Drive

### Technology Stack

- **Core**: Google Apps Script (JavaScript)  
- **Data Source**: Google Sheets  
- **Output**: Google Docs  
- **Storage**: Google Drive  
- **LLM**: OpenAI API (gpt-4o-mini)  
- **Visualization**: Google Charts Service  
- **Cloud Management**: Google Cloud Platform (for API access and project configuration)
