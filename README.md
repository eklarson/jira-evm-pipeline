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

## XML Parsing Support

Microsoft Project can also export the IMS as XML. This preserves custom field definitions and aliases much better than Excel exports.

Use the new parser:

```bash
python parse_ims_xml.py /path/to/your/IMS.xml --summary-only
python parse_ims_xml.py /path/to/your/IMS.xml --json | jq '.[0].CustomFields'
```

This script correctly maps custom fields (e.g. `188743731`) back to their defined names/aliases (e.g. "IMS ID", "Forecast Start").

A realistic sample IMS XML export is included at `tests/fixtures/sample_ims_export.xml` for development and testing. It contains representative WBS hierarchy, custom field definitions (including "IMS ID", "Forecast Start", "Forecast Finish", "Predecessor IMS IDs", "Successor IMS IDs", plus Baseline dates), structured `<PredecessorLink>` data, and a mix of tasks with and without custom field values.

The parser now extracts:
- Baseline Start / Baseline Finish (for BCWS calculations)
- Structured `<PredecessorLink>` data, automatically enriched with `PredecessorIMSID` via UID→IMS ID lookup
- Clean top-level `Predecessors` and `Successors` lists (these prefer any "Predecessor IMS IDs" / "Successor IMS IDs" custom fields that schedulers maintain, with fallback to resolved link data)
- The full `uid_to_ims_id` mapping is available on the parser instance after calling `get_tasks()`

This makes it very easy to line up Jira issues (that carry the IMS ID custom field) with their logical predecessors and successors from the IMS.

## Development & Testing

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run the test suite (uses the sample IMS fixture)
pytest tests/ -v

# Manual parser checks against the sample
python parse_ims_xml.py tests/fixtures/sample_ims_export.xml --summary-only
python parse_ims_xml.py tests/fixtures/sample_ims_export.xml --has-field "IMS ID"

# Programmatic usage with enriched predecessor/successor data
from parse_ims_xml import MSProjectXMLParser
p = MSProjectXMLParser("your_ims.xml")
p.load()
tasks = p.get_tasks()
print(p.uid_to_ims_id)                 # UID -> IMS ID map
print(tasks[3].get("Predecessors"))    # e.g. ['IMS-1001', 'IMS-1002']
print(tasks[3].get("Successors"))
```

## Next steps

- Add NameRUN Excel sync
- Generate EV summary reports
- GitHub Action for weekly run
- Use XML parser as primary data source instead of Excel