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
        self._ns: str = ""

    def _detect_namespace(self) -> str:
        """Extract the default namespace URI from the root tag if present."""
        if self.root is None:
            return ""
        tag = self.root.tag
        if tag.startswith("{"):
            return tag[1 : tag.index("}")]
        return ""

    def _find(self, parent: ET.Element, tag: str) -> Optional[ET.Element]:
        """Namespace-aware find (handles MS Project default namespace)."""
        if self._ns:
            return parent.find(f"{{{self._ns}}}{tag}")
        return parent.find(tag)

    def _findall(self, parent: ET.Element, tag: str):
        """Namespace-aware findall."""
        if self._ns:
            return parent.findall(f"{{{self._ns}}}{tag}")
        return parent.findall(tag)

    def _findtext(self, parent: ET.Element, tag: str) -> Optional[str]:
        """Namespace-aware findtext."""
        if self._ns:
            return parent.findtext(f"{{{self._ns}}}{tag}")
        return parent.findtext(tag)

    def load(self) -> None:
        """Load and parse the XML file."""
        if not self.xml_path.exists():
            raise FileNotFoundError(f"XML file not found: {self.xml_path}")

        logger.info(f"Parsing XML file: {self.xml_path.name}")
        tree = ET.parse(self.xml_path)
        self.root = tree.getroot()
        self._ns = self._detect_namespace()

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

        # Look for ExtendedAttributes at the Project level (namespace-aware)
        extended_attrs = self._find(self.root, "ExtendedAttributes")
        if extended_attrs is None:
            logger.warning("No <ExtendedAttributes> section found. Custom fields may not be resolvable.")
            return

        for ea in self._findall(extended_attrs, "ExtendedAttribute"):
            field_id = self._findtext(ea, "FieldID")
            field_name = self._findtext(ea, "FieldName") or ""
            alias = self._findtext(ea, "Alias") or ""

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

        After initial extraction we perform a second pass to:
        - Build a UID → IMS ID lookup table (for resolving structured links)
        - Enrich PredecessorLinks with the actual IMS ID when available
        - Add convenient top-level "Predecessors" / "Successors" lists
          (these prefer any "Predecessor IMS IDs" / "Successor IMS IDs" custom fields
           that schedulers commonly maintain, falling back to resolved link UIDs)
        """
        if self.root is None:
            raise RuntimeError("XML has not been loaded. Call load() first.")

        tasks_element = self._find(self.root, "Tasks")
        if tasks_element is None:
            logger.warning("No <Tasks> section found in the XML.")
            return []

        # First pass: parse everything
        all_tasks: List[Dict[str, Any]] = []
        for task_elem in self._findall(tasks_element, "Task"):
            all_tasks.append(self._parse_task(task_elem))

        # Build UID → IMS ID map from all tasks (needed even for summary_only filtering)
        uid_to_ims_id: Dict[str, str] = {}
        for t in all_tasks:
            ims_id = t.get("CustomFields", {}).get("IMS ID")
            if ims_id and t.get("UID"):
                uid_to_ims_id[t["UID"]] = ims_id
        self.uid_to_ims_id = uid_to_ims_id

        # Enrich every task with resolved links and convenience predecessor/successor lists
        for task in all_tasks:
            # Enrich structured PredecessorLinks with the IMS ID of the predecessor
            for link in task.get("PredecessorLinks", []):
                uid = link.get("PredecessorUID")
                if uid and uid in uid_to_ims_id:
                    link["PredecessorIMSID"] = uid_to_ims_id[uid]

            cf = task.get("CustomFields", {})

            # Predecessors list: prefer the scheduler-maintained custom field (split on commas)
            if "Predecessor IMS IDs" in cf and cf["Predecessor IMS IDs"]:
                preds = [x.strip() for x in cf["Predecessor IMS IDs"].split(",") if x.strip()]
                if preds:
                    task["Predecessors"] = preds
            else:
                # Fallback: resolved IMS IDs from the structured links
                resolved = [
                    link.get("PredecessorIMSID")
                    for link in task.get("PredecessorLinks", [])
                    if link.get("PredecessorIMSID")
                ]
                if resolved:
                    task["Predecessors"] = resolved

            # Successors list: comes almost exclusively from the custom field
            # (MSPDI does not store outgoing successor links on each task)
            if "Successor IMS IDs" in cf and cf["Successor IMS IDs"]:
                succs = [x.strip() for x in cf["Successor IMS IDs"].split(",") if x.strip()]
                if succs:
                    task["Successors"] = succs

        # Apply summary filter if requested
        if summary_only:
            tasks = [t for t in all_tasks if t.get("IsSummary")]
        else:
            tasks = all_tasks

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
            "BaselineStart": "BaselineStart",
            "BaselineFinish": "BaselineFinish",
            "Duration": "Duration",
            "Work": "Work",
            "PercentComplete": "PercentComplete",
            "IsSummary": "IsSummary",
            "OutlineLevel": "OutlineLevel",
            "OutlineNumber": "OutlineNumber",
        }

        for xml_name, friendly_name in standard_fields.items():
            value = self._findtext(task_elem, xml_name)
            if value is not None:
                # Normalize IsSummary to a proper boolean for reliable filtering
                if friendly_name == "IsSummary":
                    task[friendly_name] = str(value).strip() in ("1", "true", "True")
                else:
                    task[friendly_name] = value.strip() if isinstance(value, str) else value

        # Parse custom fields (ExtendedAttribute)
        custom_fields: Dict[str, Any] = {}
        for ext in self._findall(task_elem, "ExtendedAttribute"):
            field_id = self._findtext(ext, "FieldID")
            value = self._findtext(ext, "Value")

            if field_id and value is not None:
                display_name = self._get_resolved_field_name(field_id)
                custom_fields[display_name] = value.strip()

        if custom_fields:
            task["CustomFields"] = custom_fields

        # Extract structured predecessor links (by UID)
        pred_links = []
        for pl in self._findall(task_elem, "PredecessorLink"):
            pred_links.append({
                "PredecessorUID": self._findtext(pl, "PredecessorUID") or "",
                "Type": self._findtext(pl, "Type") or "1",          # 1=FS, 2=SS, 3=FF, 4=SF
                "LinkLag": self._findtext(pl, "LinkLag") or "0",
            })
        if pred_links:
            task["PredecessorLinks"] = pred_links

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
