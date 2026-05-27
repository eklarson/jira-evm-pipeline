# jira-ev-pipeline

Automated weekly sync of Forecast Start/Finish dates from Microsoft Project IMS Excel exports into Jira Capabilities.

This keeps your Structure for Jira earned value calculations always up-to-date with the official IMS schedule.

## Quick Start

1. Clone this repo
2. `pip install -e .`
3. Copy `sync_ims_to_jira.py` and update the path to your weekly IMS Excel export
4. Run with `dry_run=True` first
5. When ready, set `dry_run=False` and schedule it (Windows Task Scheduler, cron, or GitHub Actions)

## How it works

- Reads the weekly IMS Excel export from the scheduling team
- Matches on the "IMS ID" custom field on Capability issues
- Updates the Forecast Start and Forecast Finish custom fields

## Dependencies

- jira-integration-wrapper (the reusable Jira client from https://github.com/eklarson/jira-integration-wrapper)
- pandas + openpyxl

## Next steps

- Add NameRUN Excel sync
- Generate EV summary reports
- GitHub Action for weekly run