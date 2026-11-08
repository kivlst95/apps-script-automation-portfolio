# Google Apps Script Automation Portfolio

Production automations for finance, HR, sales and operations teams, built with Google Apps Script, REST APIs, OAuth 2.0 and webhooks. They connect Google Workspace with Xero, HubSpot, ClickUp, Notion, Manatal, PeopleForce and a clinic management system.

## About me

I'm **Kyrylo Vlasiuk**, a finance professional who builds automation. I started in FP&A (budgeting, forecasting, variance analysis, management reporting), then owned customer implementations of finance software as R&D Manager and led a team of three. Today I work as a Google Workspace Developer and build integrations and AI-assisted workflows that business teams actually use.

What I bring to an automation role:
- **Business first:** I start with the process and the people who run it, and measure every solution in hours saved, errors avoided or decisions made possible.
- **Finance depth:** reconciliations, multi-entity consolidation, payment allocation and P&L logic are things I understand, not just move between systems.
- **Production mindset:** secrets in configuration, idempotent webhooks, quota-aware long runs, structured logging and documentation.

**Stack:** Google Apps Script, JavaScript, Python, SQL, BigQuery, Looker Studio, n8n, Zapier, REST APIs, OAuth 2.0, webhooks, LLMs (Claude, OpenAI)
**Certification:** Google Workspace Administrator
**Education:** Master of Management
**Languages:** Ukrainian (native), English (fluent), German (B2)
**Location:** Baden-Württemberg, Germany

[LinkedIn](https://linkedin.com/in/kyr-vlasiuk-63566b145)

## Cases

Every case README starts with a plain-language diagram, followed by the technical details.

| # | Case | Systems | Result |
|---|---|---|---|
| 01 | [Xero multi-entity consolidation](cases/01-xero-multi-entity-consolidation) | Xero (3 organisations), Google Sheets, Drive | One consolidated view, ~5 h/month saved |
| 02 | [Clinic system to Xero sync](cases/02-clinic-system-to-xero-sync) | Clinic system API, Xero | Monthly reconciliation cut from ~3 days to ~2 hours |
| 03 | [Leave and sick-day approval](cases/03-leave-and-sick-day-approval-clickup) | ClickUp webhooks, ClickUp chat | ~3 h/month saved for a 25-person agency |
| 04 | [ATS to HRIS hiring integration](cases/04-ats-to-hris-hiring-integration) | Manatal, PeopleForce, Gmail | ~2 h/day of manual work saved for managers at an outstaffing company |
| 05 | [Meet recordings router](cases/05-meet-recordings-router) | Google Meet, Calendar, Drive, Notion | Manual filing removed, project documentation stays organised |
| 06 | [HubSpot deals to ClickUp](cases/06-hubspot-deals-to-clickup) | HubSpot OAuth, ClickUp | No manual re-entry, consistent data across both systems |
| 07 | [ClickUp client lists from templates](cases/07-clickup-client-lists-from-templates) | ClickUp API, Zapier webhook, Google Sheets | Lists auto-matched to client tasks, analytics always current, no manual setup |
| 08 | [Budget payment request approval](cases/08-budget-payment-request-approval) | Google Sheets, installable triggers | 10 department files, status tracking fully automated |
| – | [Sheets API toolkit](library/sheets-api-toolkit) (IMPORTRANGE replacement) | Sheets advanced service | Handles volumes IMPORTRANGE cannot, reused across projects |

## Conventions used in every project

- **Configuration:** no secrets or environment-specific IDs in the code. Everything comes from Script Properties named `SERVICE_PURPOSE` in UPPER_SNAKE_CASE (`XERO_CLIENT_ID`, `CLICKUP_HR_LIST_ID`, `WEBHOOK_SECRET`). Each project lists its keys in `config.example.json`; structured settings (field maps, stage IDs) are stored as JSON values.
- **Logging:** a shared `Logger.js` writes structured entries (severity, message, context, stack trace) to Executions and Google Cloud Logging; verbosity via `LOG_LEVEL`.
- **Naming:** `camelCase` functions, `UPPER_SNAKE_CASE` constants, a trailing underscore for private helpers (Apps Script convention: hidden from the Run menu and `google.script.run`).
- **Reliability:** webhooks with locks, deduplication and idempotency checks; long runs that persist state and resume by trigger before Apps Script or API limits; batch reads and writes.
- **Deployment:** each project has `appsscript.json` and `.clasp.json.example` for [clasp](https://github.com/google/clasp); GitHub Actions checks the syntax of every file on each push.

## Repository structure

```
cases/<project>/
  README.md              problem, solution, configuration, result
  config.example.json    Script Properties the project expects
  .clasp.json.example    clasp settings with file push order
  src/
    appsscript.json      manifest (runtime, libraries, web app)
    Logger.js            structured logging
    Config.js            Script Properties helpers
    <Project>.js         project code
library/sheets-api-toolkit/
.github/workflows/syntax-check.yml
```

## About the code

All code comes from real client projects and was prepared for publication: client names, people, email addresses, spreadsheet, folder and CRM IDs were removed from the code and replaced by configuration; comments and messages were translated into English; test and debug functions were removed.

## Usage
This repository is published as a portfolio to demonstrate my work.
The code is shared for review only. All rights reserved.