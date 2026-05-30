"""Unit tests for the MS Project XML parser using a realistic IMS sample fixture."""

from pathlib import Path

import pytest

from parse_ims_xml import MSProjectXMLParser


SAMPLE_PATH = Path(__file__).parent / "fixtures" / "sample_ims_export.xml"


@pytest.fixture(scope="module")
def parser() -> MSProjectXMLParser:
    """Load the sample IMS export once for the test module."""
    p = MSProjectXMLParser(SAMPLE_PATH)
    p.load()
    return p


@pytest.fixture(scope="module")
def all_tasks(parser: MSProjectXMLParser):
    return parser.get_tasks()


@pytest.fixture(scope="module")
def summary_tasks(parser: MSProjectXMLParser):
    return parser.get_tasks(summary_only=True)


def test_loads_without_error(parser: MSProjectXMLParser):
    assert parser.root is not None
    # We now have 8 fields (original 6 + Predecessor IMS IDs + Successor IMS IDs)
    assert len(parser.custom_field_map) == 8


def test_custom_field_aliases_resolved(parser: MSProjectXMLParser):
    """The key value of the parser: FieldIDs must map to human aliases."""
    cf = parser.custom_field_map

    # Key IMS fields used by the Jira sync
    assert cf["188743731"]["display_name"] == "IMS ID"
    assert cf["188743945"]["display_name"] == "Forecast Start"
    assert cf["188743946"]["display_name"] == "Forecast Finish"

    # Other common IMS fields
    assert cf["188743732"]["display_name"] == "Control Account"
    assert cf["188743733"]["display_name"] == "Work Package"
    assert cf["188743752"]["display_name"] == "Critical IMS"


def test_extracts_all_tasks(all_tasks):
    # Sample contains 12 tasks (UID 0-11)
    assert len(all_tasks) == 12


def test_summary_only_filtering(summary_tasks):
    """--summary-only should return only tasks where IsSummary is True."""
    assert len(summary_tasks) == 4
    for t in summary_tasks:
        assert t.get("IsSummary") is True
        # OutlineLevel should be 0 or 1 for the summaries in our fixture
        assert t.get("OutlineLevel") in ("0", "1")


def test_tasks_with_ims_id_field(all_tasks):
    """Tasks that have an IMS ID should surface it under the resolved alias."""
    tasks_with_ims = [t for t in all_tasks if "CustomFields" in t and "IMS ID" in t["CustomFields"]]
    assert len(tasks_with_ims) == 6

    ims_values = {t["CustomFields"]["IMS ID"] for t in tasks_with_ims}
    assert "IMS-1001" in ims_values
    assert "IMS-2001" in ims_values
    assert "IMS-4001" in ims_values


def test_forecast_dates_present_on_relevant_tasks(all_tasks):
    """Forecast Start/Finish should be present on tasks that declared them."""
    for t in all_tasks:
        cf = t.get("CustomFields", {})
        if "IMS ID" in cf:
            # Every task that has IMS ID in the fixture also has both forecast dates
            assert "Forecast Start" in cf
            assert "Forecast Finish" in cf


def test_task_without_custom_fields_still_parses(all_tasks):
    """A task with zero ExtendedAttribute children must still be returned cleanly."""
    tasks_without_cf = [t for t in all_tasks if "CustomFields" not in t]
    assert len(tasks_without_cf) >= 1

    # The one we intentionally left bare in the fixture
    bare = next((t for t in tasks_without_cf if t["Name"] == "2.3 Interface Control Documents"), None)
    assert bare is not None
    assert bare["OutlineLevel"] == "2"


def test_is_summary_is_boolean(all_tasks):
    """IsSummary must be normalized to real Python booleans, not strings."""
    for t in all_tasks:
        if "IsSummary" in t:
            assert isinstance(t["IsSummary"], bool)


def test_parser_handles_missing_extended_attributes_gracefully(tmp_path):
    """Parser should not crash when a file has no ExtendedAttributes section."""
    bad_xml = tmp_path / "no_extended.xml"
    bad_xml.write_text(
        '<?xml version="1.0"?><Project xmlns="http://schemas.microsoft.com/project">'
        "<Tasks><Task><UID>1</UID><Name>Orphan</Name></Task></Tasks></Project>"
    )
    p = MSProjectXMLParser(bad_xml)
    p.load()
    tasks = p.get_tasks()
    assert len(tasks) == 1
    assert tasks[0]["Name"] == "Orphan"
    assert p.custom_field_map == {}  # no definitions, but no crash


# --- New fields for BCWS and dependency linking (Baseline + Predecessor/Successor IMS IDs) ---


def test_baseline_dates_are_extracted(all_tasks):
    """BaselineStart / BaselineFinish should be available as top-level fields for BCWS calc."""
    tasks_with_baseline = [t for t in all_tasks if t.get("BaselineStart")]
    assert len(tasks_with_baseline) >= 4

    # Spot-check one we explicitly populated
    req_def = next(t for t in all_tasks if t.get("ID") == "5")
    assert req_def["BaselineStart"] == "2026-06-15T08:00:00"
    assert req_def["BaselineFinish"] == "2026-07-31T17:00:00"


def test_predecessor_links_are_parsed(all_tasks):
    """Structured PredecessorLink elements should be extracted as a list of dicts."""
    task_with_preds = next(t for t in all_tasks if t.get("ID") == "5")
    preds = task_with_preds.get("PredecessorLinks", [])
    assert len(preds) == 2
    assert preds[0]["PredecessorUID"] == "2"
    assert preds[0]["Type"] == "1"   # Finish-to-Start
    assert preds[1]["PredecessorUID"] == "3"


def test_predecessor_and_successor_ims_id_custom_fields(all_tasks):
    """
    Custom fields 'Predecessor IMS IDs' and 'Successor IMS IDs' (common pattern in real IMS
    schedules) should be resolved via alias and available for Jira issue linking.
    """
    task = next(t for t in all_tasks if t.get("ID") == "5")
    cf = task.get("CustomFields", {})

    assert cf.get("Predecessor IMS IDs") == "IMS-1001, IMS-1002"
    assert cf.get("Successor IMS IDs") == "IMS-2002"

    # Another task should also have the fields populated
    arch = next(t for t in all_tasks if t.get("ID") == "6")
    cf2 = arch.get("CustomFields", {})
    assert cf2.get("Predecessor IMS IDs") == "IMS-2001"
    assert cf2.get("Successor IMS IDs") == "IMS-3001"


def test_new_custom_fields_are_registered_in_map(parser: MSProjectXMLParser):
    """The two new pred/succ IMS ID fields must be present in the custom field map."""
    cf = parser.custom_field_map
    assert cf["188743734"]["display_name"] == "Predecessor IMS IDs"
    assert cf["188743735"]["display_name"] == "Successor IMS IDs"


def test_predecessor_links_are_enriched_with_ims_id(all_tasks, parser: MSProjectXMLParser):
    """Structured PredecessorLinks should be enriched with PredecessorIMSID via UID lookup."""
    task = next(t for t in all_tasks if t.get("ID") == "5")
    links = task.get("PredecessorLinks", [])

    assert links[0]["PredecessorUID"] == "2"
    assert links[0]["PredecessorIMSID"] == "IMS-1001"

    assert links[1]["PredecessorUID"] == "3"
    assert links[1]["PredecessorIMSID"] == "IMS-1002"

    # The parser should also expose the full lookup map
    assert parser.uid_to_ims_id["2"] == "IMS-1001"
    assert parser.uid_to_ims_id["5"] == "IMS-2001"


def test_predecessors_and_successors_convenience_lists(all_tasks):
    """
    Tasks should get clean top-level Predecessors / Successors lists.
    These prefer the custom "Predecessor IMS IDs" / "Successor IMS IDs" fields
    (the common real-world pattern), but fall back to resolved link data.
    """
    task = next(t for t in all_tasks if t.get("ID") == "5")
    assert task.get("Predecessors") == ["IMS-1001", "IMS-1002"]
    assert task.get("Successors") == ["IMS-2002"]

    # Task 6 should also have the lists from its custom fields
    task6 = next(t for t in all_tasks if t.get("ID") == "6")
    assert task6.get("Predecessors") == ["IMS-2001"]
    assert task6.get("Successors") == ["IMS-3001"]


def test_predecessors_fallback_to_resolved_links(tmp_path):
    """
    If a task has no "Predecessor IMS IDs" custom field but has structured links,
    Predecessors should still be populated from the UID→IMS ID resolution.
    """
    # Minimal XML with one predecessor link but no custom pred/succ fields
    xml = tmp_path / "links_only.xml"
    xml.write_text('''<?xml version="1.0" encoding="UTF-8"?>
<Project xmlns="http://schemas.microsoft.com/project">
  <ExtendedAttributes>
    <ExtendedAttribute>
      <FieldID>188743731</FieldID>
      <FieldName>Text1</FieldName>
      <Alias>IMS ID</Alias>
    </ExtendedAttribute>
  </ExtendedAttributes>
  <Tasks>
    <Task>
      <UID>1</UID><ID>1</ID><Name>Pred Task</Name>
      <ExtendedAttribute><FieldID>188743731</FieldID><Value>IMS-PRED</Value></ExtendedAttribute>
    </Task>
    <Task>
      <UID>2</UID><ID>2</ID><Name>Current Task</Name>
      <PredecessorLink><PredecessorUID>1</PredecessorUID><Type>1</Type></PredecessorLink>
      <ExtendedAttribute><FieldID>188743731</FieldID><Value>IMS-CURR</Value></ExtendedAttribute>
    </Task>
  </Tasks>
</Project>''')

    p = MSProjectXMLParser(xml)
    p.load()
    tasks = p.get_tasks()

    curr = next(t for t in tasks if t["Name"] == "Current Task")
    assert curr.get("Predecessors") == ["IMS-PRED"]
    assert curr["PredecessorLinks"][0]["PredecessorIMSID"] == "IMS-PRED"
    # No Successors because no custom field and no outgoing links in the data
    assert "Successors" not in curr
