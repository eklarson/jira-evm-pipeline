# sync_ims_to_jira.py
# Weekly sync of IMS Forecast dates into Jira Capabilities

import pandas as pd
import logging
from pathlib import Path

from jira_integration_wrapper import get_jira_client

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class IMSSync:
    def __init__(self, excel_path: str | Path):
        self.excel_path = Path(excel_path)
        self.jira = get_jira_client()

    def load_ims_data(self) -> pd.DataFrame:
        """Load the weekly IMS Excel export from Microsoft Project."""
        if not self.excel_path.exists():
            raise FileNotFoundError(f"IMS Excel file not found: {self.excel_path}")

        logger.info(f"Loading IMS export: {self.excel_path.name}")

        df = pd.read_excel(
            self.excel_path,
            sheet_name=0,
            parse_dates=["Forecast Start", "Forecast Finish"],
        )

        df.columns = [col.strip() for col in df.columns]

        # Flexible column mapping for typical MS Project exports
        required_mapping = {
            "IMS_ID": ["IMS ID", "Unique ID", "Task ID", "ID"],
            "FORECAST_START": ["Forecast Start", "Start", "Planned Start", "Early Start"],
            "FORECAST_FINISH": ["Forecast Finish", "Finish", "Planned Finish", "Early Finish"]
        }

        col_map = {}
        for target, possibles in required_mapping.items():
            for possible in possibles:
                if possible in df.columns:
                    col_map[possible] = target
                    break
            else:
                raise ValueError(f"Could not find column for {target}. Available columns: {list(df.columns)}")

        df = df.rename(columns=col_map)[["IMS_ID", "FORECAST_START", "FORECAST_FINISH"]].copy()

        df["IMS_ID"] = df["IMS_ID"].astype(str).str.strip()
        df = df.dropna(subset=["IMS_ID"]).reset_index(drop=True)

        logger.info(f"Loaded {len(df)} tasks with IMS_ID")
        return df

    def update_jira_capabilities(self, df: pd.DataFrame, dry_run: bool = True):
        """Update Forecast Start and Forecast Finish on Capabilities in Jira."""
        updated = 0
        skipped = 0
        errors = 0

        # === REPLACE THESE WITH YOUR REAL CUSTOM FIELD IDs ===
        FORECAST_START_FIELD = "customfield_15678"   # e.g. Forecast Start
        FORECAST_FINISH_FIELD = "customfield_15679"  # e.g. Forecast Finish

        for _, row in df.iterrows():
            ims_id = row["IMS_ID"]
            if not ims_id or ims_id.lower() == "nan":
                continue

            try:
                jql = f'issuetype = Capability AND "IMS ID" = "{ims_id}"'
                issues = self.jira.search_issues(jql, maxResults=3)

                if not issues:
                    logger.warning(f"No Capability found for IMS_ID: {ims_id}")
                    skipped += 1
                    continue

                capability = issues[0]
                updates = {}

                if pd.notna(row["FORECAST_START"]):
                    updates[FORECAST_START_FIELD] = row["FORECAST_START"].strftime("%Y-%m-%d")

                if pd.notna(row["FORECAST_FINISH"]):
                    updates[FORECAST_FINISH_FIELD] = row["FORECAST_FINISH"].strftime("%Y-%m-%d")

                if not updates:
                    continue

                if dry_run:
                    logger.info(f"DRY RUN → Would update {capability.key} (IMS_ID {ims_id})")
                else:
                    capability.update(fields=updates)
                    logger.info(f"✅ Updated {capability.key} (IMS_ID {ims_id})")

                updated += 1

            except Exception as e:
                logger.error(f"Failed to update IMS_ID {ims_id}: {e}")
                errors += 1

        logger.info("=" * 60)
        logger.info(f"Sync finished | Updated: {updated} | Skipped: {skipped} | Errors: {errors}")
        if dry_run:
            logger.info("Dry run complete. Set dry_run=False when ready.")


if __name__ == "__main__":
    # UPDATE THIS PATH EVERY WEEK
    excel_file = r"C:\\Path\\To\\Your\\IMS_Weekly_Export.xlsx"

    sync = IMSSync(excel_file)
    df = sync.load_ims_data()

    # Always start with dry_run=True
    sync.update_jira_capabilities(df, dry_run=True)

    # When ready:
    # sync.update_jira_capabilities(df, dry_run=False)
