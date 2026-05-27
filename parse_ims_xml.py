#!/usr/bin/env python3
"""
parse_ims_xml.py

Parse a Microsoft Project XML export (IMS) and extract Tasks with custom fields
correctly mapped to their defined names/aliases.

This is especially useful when you export from Microsoft Project as XML instead
of Excel, because the XML preserves the actual custom field definitions and aliases.

Usage examples:

    # Basic usage - pretty print tasks with resolved custom fields
    python parse_ims_xml.py MySchedule.xml

    # Output as JSON (for piping to other tools)
    python parse_ims_xml.py MySchedule.xml --json

    # Only show summary tasks
    python parse_ims_xml.py MySchedule.xml --summary-only

    # Show only tasks that have a specific custom field
    python parse_ims_xml.py MySchedule.xml --has-field "IMS ID"

Requirements:
    Python 3.10+
    (No external dependencies - uses only the standard library)
"""

import argparse
import json
import logging
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class MSProjectXMLParser:
    """Parser for Microsoft Project XML format that resolves custom fields."""

    def __init__(self, xml_path: Path):
        self.xml_path = Path(xml_path)
        self.root: Optional[ET.Element] = None
        self.custom_field_map: Dict[str, Dict[str, str]] = {}

    def load(self) -> None:
        """Load and parse the XML file."""
        if not self.xml_path.exists():
            raise FileNotFoundError(f"XML file not found: {self.xml_path}")

        logger.info(f"Parsing XML file: {self.xml_path.name}")
        tree = ET.parse(self.xml_path)
        self.root = tree.getroot()

        # Build custom field definitions (this is the key part)
        self._build_custom_field_map()

    def _build_custom_field_map(self) -> None:
        """
        Build a mapping from FieldID to human-readable field information.

        Microsoft Project stores custom field definitions under <ExtendedAttributes>.
        Each definition has FieldID, FieldName (e.g. "Text1"), and optionally an Alias.
        """
        if self.root is None:
            return

        self.custom_field_map = {}

        # Look for ExtendedAttributes at the Project level
        extended_attrs = self.root.find("ExtendedAttributes")
        if extended_attrs is None:
            logger.warning("No <ExtendedAttributes> section found. Custom fields may not be resolvable.")
            return

        for ea in extended_attrs.findall("ExtendedAttribute"):
            field_id = ea.findtext("FieldID")
            field_name = ea.findtext("FieldName") or ""
            alias = ea.findtext("Alias") or ""

            if field_id:
                self.custom_field_map[field_id] = {
                    "field_name": field_name,
                    "alias": alias,
                    "display_name": alias if alias else field_name,
                }

        logger.info(f"Found {len(self.custom_field_map)} custom field definitions")

    def _get_resolved_field_name(self, field_id: str) -> str:
        """Return the best human-readable name for a custom field."""
        info = self.custom_field_map.get(field_id)
        if info:
            return info["display_name"]
        return f"UnknownField_{field_id}"

    def get_tasks(self, summary_only: bool = False) -> List[Dict[str, Any]]:
        """
        Extract all tasks with standard fields + resolved custom fields.
        """
        if self.root is None:
            raise RuntimeError("XML has not been loaded. Call load() first.")

        tasks_element = self.root.find("Tasks")
        if tasks_element is None:
            logger.warning("No <Tasks> section found in the XML.")
            return []

        tasks: List[Dict[str, Any]] = []

        for task_elem in tasks_element.findall("Task"):
            task = self._parse_task(task_elem)

            if summary_only and not task.get("IsSummary", False):
                continue

            tasks.append(task)

        logger.info(f"Extracted {len(tasks)} tasks")
        return tasks

    def _parse_task(self, task_elem: ET.Element) -> Dict[str, Any]:
        """Convert a single <Task> element into a clean dictionary."""
        task: Dict[str, Any] = {}

        # Standard fields we care about
        standard_fields = {
            "UID": "UID",
            "ID": "ID",
            "Name": "Name",
            "Start": "Start",
            "Finish": "Finish",
            "Duration": "Duration",
            "Work": "Work",
            "PercentComplete": "PercentComplete",
            "IsSummary": "IsSummary",
            "OutlineLevel": "OutlineLevel",
            "OutlineNumber": "OutlineNumber",
        }

        for xml_name, friendly_name in standard_fields.items():
            value = task_elem.findtext(xml_name)
            if value is not None:
                task[friendly_name] = value.strip() if isinstance(value, str) else value

        # Parse custom fields (ExtendedAttribute)
        custom_fields: Dict[str, Any] = {}
        for ext in task_elem.findall("ExtendedAttribute"):
            field_id = ext.findtext("FieldID")
            value = ext.findtext("Value")

            if field_id and value is not None:
                display_name = self._get_resolved_field_name(field_id)
                custom_fields[display_name] = value.strip()

        if custom_fields:
            task["CustomFields"] = custom_fields

        return task

    def get_tasks_as_dicts(self, summary_only: bool = False) -> List[Dict[str, Any]]:
        """Convenience method that returns tasks with clean structure."""
        return self.get_tasks(summary_only=summary_only)


def main():
    parser = argparse.ArgumentParser(
        description="Parse Microsoft Project XML export and resolve custom fields to their defined names."
    )
    parser.add_argument(
        "xml_file",
        type=Path,
        help="Path to the Microsoft Project XML export file",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output tasks as JSON instead of pretty printing",
    )
    parser.add_argument(
        "--summary-only",
        action="store_true",
        help="Only include summary tasks (WBS elements)",
    )
    parser.add_argument(
        "--has-field",
        metavar="FIELD_NAME",
        help="Only show tasks that have a custom field with this name (after alias resolution)",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        default=True,
        help="Pretty-print output (default)",
    )

    args = parser.parse_args()

    try:
        parser = MSProjectXMLParser(args.xml_file)
        parser.load()
        tasks = parser.get_tasks(summary_only=args.summary_only)

        # Optional filter
        if args.has_field:
            tasks = [
                t for t in tasks
                if "CustomFields" in t and args.has_field in t["CustomFields"]
            ]
            logger.info(f"Filtered to {len(tasks)} tasks containing field: {args.has_field}")

        if args.json:
            print(json.dumps(tasks, indent=2, default=str))
        else:
            # Human-friendly output
            for i, task in enumerate(tasks, 1):
                print(f"\n=== Task {i} ===")
                for key, value in task.items():
                    if key == "CustomFields":
                        print("  Custom Fields:")
                        for cf_name, cf_value in value.items():
                            print(f"    - {cf_name}: {cf_value}")
                    else:
                        print(f"  {key}: {value}")

            print(f"\nTotal tasks extracted: {len(tasks)}")

    except Exception as e:
        logger.error(f"Failed to process XML: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
