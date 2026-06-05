#!/usr/bin/env python3
"""Build the SEC-47 SOC 2 audit concern workbook.

The preferred spreadsheet artifact runtime is not available in this workspace,
so this follows the existing repo pattern and writes a small OpenXML workbook
using only Python's standard library.
"""

from __future__ import annotations

import json
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "source_data"
OUTPUT_DIR = ROOT / "outputs"
SOURCE_PATH = SOURCE_DIR / "sec47_soc2_audit_findings.json"
WORKBOOK_PATH = OUTPUT_DIR / "sec47_soc2_audit_concerns.xlsx"

AUDIT_URL = (
    "https://app.sprinto.com/app/admin/audits/"
    "840b09b2-88b8-450d-9147-2d68f70f8bf2"
    "?tab=evidence&evidenceFilters=%7B%7D&evidenceFiltersCount=0"
)

EXTRACTED_AT = "2026-06-05 11:28:52 PDT"

AUDIT_CONTEXT = {
    "audit_title": "2026 - SOC2 Audit",
    "smooth_profile": "oneleet",
    "evidence_tab_counts": {
        "Controls": 88,
        "Criteria's": 11,
        "Information requested": 0,
        "Ready for audit": 88,
        "Review completed": 0,
    },
    "control_export": {
        "filename": "ControlMappingFile.csv",
        "row_count": 171,
        "columns": [
            "Area Code",
            "Area",
            "Criteria Code",
            "Criteria",
            "Control Number",
            "Control Statement",
            "Evidence met (Count)",
            "Evidence pending (Count)",
            "Total evidence required (Count)",
        ],
        "aggregate_totals": {
            "Evidence met (Count)": 421,
            "Evidence pending (Count)": 22,
            "Total evidence required (Count)": 443,
        },
    },
    "evidence_export_limitation": (
        "The Sprinto `Download evidences` action did not produce an immediate "
        "browser download; Smooth reported that the evidence export is emailed "
        "later, so evidence-level owners, due dates, and item names could not be "
        "verified from the downloaded file in this run."
    ),
}

UI_CONTROL_DETAILS = {
    9: {
        "requirement": "SDC 9: Performance Review",
        "completion": "1/3 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
    23: {
        "requirement": "SDC 23: Internal Audit using Sprinto",
        "completion": "1/1 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
    25: {
        "requirement": "SDC 25: Periodic Review & Update of Cybersecurity & Privacy Program",
        "completion": "1/4 incomplete",
        "status": "Ready for audit",
        "messages": "N/A",
    },
    49: {
        "requirement": "SDC 49: Encrypting Data At Rest",
        "completion": "1/5 incomplete",
        "status": "Ready for audit",
        "messages": "N/A",
    },
    59: {
        "requirement": "SDC 59: Data Backups",
        "completion": "1/6 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
    60: {
        "requirement": "SDC 60: Testing for Reliability and Integrity",
        "completion": "1/1 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
    62: {
        "requirement": "SDC 62: Capacity & Performance Management",
        "completion": "1/7 incomplete",
        "status": "Ready for audit",
        "messages": "N/A",
    },
    97: {
        "requirement": "SDC 97: Testing",
        "completion": "1/1 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
    389: {
        "requirement": "SDC 389: Updates During Installations / Removals",
        "completion": "1/3 incomplete",
        "status": "Ready for audit",
        "messages": "",
    },
}

PENDING_ROWS = [
    {
        "Area Code": "CC1",
        "Area": "Control Environment",
        "Criteria Code": "CC1.5",
        "Criteria": "COSO Principle 5: The entity holds individuals accountable for their internal control responsibilities in the pursuit of objectives.",
        "Control Number": 9,
        "Control Statement": "Entity requires that all employees in client serving, IT, Engineering, and Information Security roles are periodically evaluated regarding their job responsibilities.",
        "Evidence met (Count)": 2,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 3,
    },
    {
        "Area Code": "CC1",
        "Area": "Control Environment",
        "Criteria Code": "CC1.3",
        "Criteria": "COSO Principle 3: Management establishes, with board oversight, structures, reporting lines, and appropriate authorities and responsibilities in the pursuit of objectives.",
        "Control Number": 25,
        "Control Statement": "Entity's Senior Management reviews and approves the state of the Information Security program including policies, standards, and procedures, at planned intervals or if significant changes occur to ensure their continuing suitability, adequacy, and effectiveness.",
        "Evidence met (Count)": 3,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 4,
    },
    {
        "Area Code": "CC1",
        "Area": "Control Environment",
        "Criteria Code": "CC1.2",
        "Criteria": "COSO Principle 2: The board of directors demonstrates independence from management and exercises oversight of the development and performance of internal control.",
        "Control Number": 25,
        "Control Statement": "Entity's Senior Management reviews and approves the state of the Information Security program including policies, standards, and procedures, at planned intervals or if significant changes occur to ensure their continuing suitability, adequacy, and effectiveness.",
        "Evidence met (Count)": 3,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 4,
    },
    {
        "Area Code": "CC4",
        "Area": "Monitoring Activities",
        "Criteria Code": "CC4.1",
        "Criteria": "COSO Principle 16: The entity selects, develops, and performs ongoing and/or separate evaluations to ascertain whether the components of internal control are present and functioning.",
        "Control Number": 23,
        "Control Statement": "Entity uses Sprinto, a continuous monitoring system, to track and report the health of the information security program to the Information Security Officer and other stakeholders.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "CC4",
        "Area": "Monitoring Activities",
        "Criteria Code": "CC4.1",
        "Criteria": "COSO Principle 16: The entity selects, develops, and performs ongoing and/or separate evaluations to ascertain whether the components of internal control are present and functioning.",
        "Control Number": 25,
        "Control Statement": "Entity's Senior Management reviews and approves the state of the Information Security program including policies, standards, and procedures, at planned intervals or if significant changes occur to ensure their continuing suitability, adequacy, and effectiveness.",
        "Evidence met (Count)": 3,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 4,
    },
    {
        "Area Code": "CC4",
        "Area": "Monitoring Activities",
        "Criteria Code": "CC4.1",
        "Criteria": "COSO Principle 16: The entity selects, develops, and performs ongoing and/or separate evaluations to ascertain whether the components of internal control are present and functioning.",
        "Control Number": 389,
        "Control Statement": "Entity periodically updates and reviews the inventory of systems as a part of installations, removals, and system updates.",
        "Evidence met (Count)": 2,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 3,
    },
    {
        "Area Code": "CC4",
        "Area": "Monitoring Activities",
        "Criteria Code": "CC4.2",
        "Criteria": "COSO Principle 17: The entity evaluates and communicates internal control deficiencies in a timely manner to those parties responsible for taking corrective action, including senior management and the board of directors, as appropriate.",
        "Control Number": 23,
        "Control Statement": "Entity uses Sprinto, a continuous monitoring system, to track and report the health of the information security program to the Information Security Officer and other stakeholders.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "CC4",
        "Area": "Monitoring Activities",
        "Criteria Code": "CC4.2",
        "Criteria": "COSO Principle 17: The entity evaluates and communicates internal control deficiencies in a timely manner to those parties responsible for taking corrective action, including senior management and the board of directors, as appropriate.",
        "Control Number": 25,
        "Control Statement": "Entity's Senior Management reviews and approves the state of the Information Security program including policies, standards, and procedures, at planned intervals or if significant changes occur to ensure their continuing suitability, adequacy, and effectiveness.",
        "Evidence met (Count)": 3,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 4,
    },
    {
        "Area Code": "CC5",
        "Area": "Control Activities",
        "Criteria Code": "CC5.2",
        "Criteria": "COSO Principle 11: The entity also selects and develops general control activities over technology to support the achievement of objectives.",
        "Control Number": 23,
        "Control Statement": "Entity uses Sprinto, a continuous monitoring system, to track and report the health of the information security program to the Information Security Officer and other stakeholders.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "CC5",
        "Area": "Control Activities",
        "Criteria Code": "CC5.2",
        "Criteria": "COSO Principle 11: The entity also selects and develops general control activities over technology to support the achievement of objectives.",
        "Control Number": 25,
        "Control Statement": "Entity's Senior Management reviews and approves the state of the Information Security program including policies, standards, and procedures, at planned intervals or if significant changes occur to ensure their continuing suitability, adequacy, and effectiveness.",
        "Evidence met (Count)": 3,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 4,
    },
    {
        "Area Code": "CC6",
        "Area": "Logical and Physical Access Controls",
        "Criteria Code": "CC6.7",
        "Criteria": "The entity restricts the transmission, movement, and removal of information to authorized internal and external users and processes, and protects it during transmission, movement, or removal to meet the entity's objectives.",
        "Control Number": 49,
        "Control Statement": "Entity has set up cryptographic mechanisms to encrypt all production database[s] that store customer data at rest.",
        "Evidence met (Count)": 4,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 5,
    },
    {
        "Area Code": "CC7",
        "Area": "System Operations",
        "Criteria Code": "CC7.1",
        "Criteria": "To meet its objectives, the entity uses detection and monitoring procedures to identify changes to configurations that result in new vulnerabilities and susceptibilities to newly discovered vulnerabilities.",
        "Control Number": 62,
        "Control Statement": "Entity has set up methods to continuously monitor critical assets to generate capacity alerts to ensure optimal performance, meet future capacity requirements, and protect against denial-of-service attacks.",
        "Evidence met (Count)": 6,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 7,
    },
    {
        "Area Code": "CC7",
        "Area": "System Operations",
        "Criteria Code": "CC7.2",
        "Criteria": "The entity monitors system components and the operation of those components for anomalies that are indicative of malicious acts, natural disasters, and errors affecting the entity's ability to meet its objectives.",
        "Control Number": 62,
        "Control Statement": "Entity has set up methods to continuously monitor critical assets to generate capacity alerts to ensure optimal performance, meet future capacity requirements, and protect against denial-of-service attacks.",
        "Evidence met (Count)": 6,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 7,
    },
    {
        "Area Code": "CC7",
        "Area": "System Operations",
        "Criteria Code": "CC7.3",
        "Criteria": "The entity evaluates security events to determine whether they could or have resulted in a failure of the entity to meet its objectives and takes actions to prevent or address such failures.",
        "Control Number": 62,
        "Control Statement": "Entity has set up methods to continuously monitor critical assets to generate capacity alerts to ensure optimal performance, meet future capacity requirements, and protect against denial-of-service attacks.",
        "Evidence met (Count)": 6,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 7,
    },
    {
        "Area Code": "CC7",
        "Area": "System Operations",
        "Criteria Code": "CC7.3",
        "Criteria": "The entity evaluates security events to determine whether they could or have resulted in a failure of the entity to meet its objectives and takes actions to prevent or address such failures.",
        "Control Number": 23,
        "Control Statement": "Entity uses Sprinto, a continuous monitoring system, to track and report the health of the information security program to the Information Security Officer and other stakeholders.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "CC7",
        "Area": "System Operations",
        "Criteria Code": "CC7.4",
        "Criteria": "The entity responds to identified security incidents by executing a defined incident response program to understand, contain, remediate, and communicate security incidents, as appropriate.",
        "Control Number": 23,
        "Control Statement": "Entity uses Sprinto, a continuous monitoring system, to track and report the health of the information security program to the Information Security Officer and other stakeholders.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "C1",
        "Area": "Additional Criteria for Confidentiality",
        "Criteria Code": "C1.1",
        "Criteria": "The entity identifies and maintains confidential information to meet the entity's objectives related to confidentiality.",
        "Control Number": 49,
        "Control Statement": "Entity has set up cryptographic mechanisms to encrypt all production database[s] that store customer data at rest.",
        "Evidence met (Count)": 4,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 5,
    },
    {
        "Area Code": "A1",
        "Area": "Additional Criteria for Availability",
        "Criteria Code": "A1.2",
        "Criteria": "The entity authorizes, designs, develops or acquires, implements, operates, approves, maintains, and monitors environmental protections, software, data back-up processes, and recovery infrastructure to meet its objectives.",
        "Control Number": 59,
        "Control Statement": "Entity backs up relevant user and system data regularly to meet recovery time and recovery point objectives and verifies the integrity of these backups.",
        "Evidence met (Count)": 5,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 6,
    },
    {
        "Area Code": "A1",
        "Area": "Additional Criteria for Availability",
        "Criteria Code": "A1.2",
        "Criteria": "The entity authorizes, designs, develops or acquires, implements, operates, approves, maintains, and monitors environmental protections, software, data back-up processes, and recovery infrastructure to meet its objectives.",
        "Control Number": 60,
        "Control Statement": "Entity tests backup information periodically to verify media reliability and information integrity.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "A1",
        "Area": "Additional Criteria for Availability",
        "Criteria Code": "A1.1",
        "Criteria": "The entity maintains, monitors, and evaluates current processing capacity and use of system components to manage capacity demand and enable additional capacity to meet its objectives.",
        "Control Number": 62,
        "Control Statement": "Entity has set up methods to continuously monitor critical assets to generate capacity alerts to ensure optimal performance, meet future capacity requirements, and protect against denial-of-service attacks.",
        "Evidence met (Count)": 6,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 7,
    },
    {
        "Area Code": "A1",
        "Area": "Additional Criteria for Availability",
        "Criteria Code": "A1.3",
        "Criteria": "The entity tests recovery plan procedures supporting system recovery to meet its objectives.",
        "Control Number": 97,
        "Control Statement": "Entity has procedures to conduct regular tests and exercises that determine the effectiveness and readiness to execute the contingency plan.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
    {
        "Area Code": "A1",
        "Area": "Additional Criteria for Availability",
        "Criteria Code": "A1.3",
        "Criteria": "The entity tests recovery plan procedures supporting system recovery to meet its objectives.",
        "Control Number": 60,
        "Control Statement": "Entity tests backup information periodically to verify media reliability and information integrity.",
        "Evidence met (Count)": 0,
        "Evidence pending (Count)": 1,
        "Total evidence required (Count)": 1,
    },
]


def col_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_xml(ref: str, value: Any, style: int | None = None) -> str:
    attrs = f' r="{ref}"'
    if style is not None:
        attrs += f' s="{style}"'
    if isinstance(value, bool):
        return f'<c{attrs} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f"<c{attrs}><v>{value}</v></c>"
    text = escape("" if value is None else str(value))
    preserve = ' xml:space="preserve"' if text.startswith(" ") or text.endswith(" ") else ""
    return f'<c{attrs} t="inlineStr"><is><t{preserve}>{text}</t></is></c>'


def severity(control: dict[str, Any]) -> str:
    if control["zero_completed_mappings"] or control["ui_completion"] == "1/1 incomplete":
        return "High"
    return "Medium"


def severity_style(value: str) -> int:
    return {"High": 6, "Medium": 7, "Low": 8}.get(value, 4)


def row_styles(values: list[Any], default: int = 4) -> list[int]:
    styles = []
    for value in values:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            styles.append(5)
        elif value == "High":
            styles.append(6)
        elif value == "Medium":
            styles.append(7)
        elif value == "Low":
            styles.append(8)
        else:
            styles.append(default)
    return styles


def build_unique_controls() -> list[dict[str, Any]]:
    grouped: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in PENDING_ROWS:
        grouped[int(row["Control Number"])].append(row)

    controls = []
    for control_number, rows in sorted(grouped.items()):
        ui = UI_CONTROL_DETAILS[control_number]
        criteria_codes = sorted({row["Criteria Code"] for row in rows})
        area_codes = sorted({row["Area Code"] for row in rows})
        met = sum(int(row["Evidence met (Count)"]) for row in rows)
        pending = sum(int(row["Evidence pending (Count)"]) for row in rows)
        required = sum(int(row["Total evidence required (Count)"]) for row in rows)
        zero_completed = sum(1 for row in rows if int(row["Evidence met (Count)"]) == 0)
        control = {
            "control_number": control_number,
            "ui_requirement": ui["requirement"],
            "ui_completion": ui["completion"],
            "ui_status": ui["status"],
            "messages": ui["messages"] or "None visible",
            "criteria_codes": ", ".join(criteria_codes),
            "area_codes": ", ".join(area_codes),
            "mapping_count": len(rows),
            "met_sum": met,
            "pending_sum": pending,
            "required_sum": required,
            "zero_completed_mappings": zero_completed,
            "control_statement": rows[0]["Control Statement"],
        }
        control["severity"] = severity(control)
        if control["severity"] == "High":
            control["why_concerning"] = (
                "One or more SOC 2 criteria mappings have no completed evidence; "
                "auditors may be unable to test this control for the mapped trust criteria."
            )
        else:
            control["why_concerning"] = (
                "The control is marked Ready for audit, but Sprinto still reports one "
                "incomplete evidence set; unresolved pending evidence may trigger an "
                "auditor request or exception."
            )
        control["recommended_remediation"] = (
            "Complete/upload the pending evidence, confirm it covers every mapped "
            "criteria code, rerun the Sprinto control export, and verify Evidence "
            "pending is 0 before auditor review."
        )
        controls.append(control)
    return controls


def build_area_summary() -> list[dict[str, Any]]:
    by_area: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in PENDING_ROWS:
        by_area[(row["Area Code"], row["Area"])].append(row)
    summary = []
    for (area_code, area), rows in sorted(by_area.items()):
        summary.append(
            {
                "area_code": area_code,
                "area": area,
                "pending_mapping_count": len(rows),
                "unique_controls": len({row["Control Number"] for row in rows}),
                "pending_evidence_count": sum(int(row["Evidence pending (Count)"]) for row in rows),
                "zero_completed_mappings": sum(1 for row in rows if int(row["Evidence met (Count)"]) == 0),
                "criteria_codes": ", ".join(sorted({row["Criteria Code"] for row in rows})),
            }
        )
    return summary


def source_dataset() -> dict[str, Any]:
    controls = build_unique_controls()
    areas = build_area_summary()
    return {
        "ticket": "SEC-47",
        "audit_url": AUDIT_URL,
        "extracted_at": EXTRACTED_AT,
        "audit_context": AUDIT_CONTEXT,
        "ui_control_details": UI_CONTROL_DETAILS,
        "pending_control_mapping_rows": PENDING_ROWS,
        "unique_control_findings": controls,
        "area_summary": areas,
        "summary": {
            "pending_mapping_rows": len(PENDING_ROWS),
            "unique_controls_with_pending_evidence": len(controls),
            "high_severity_unique_controls": sum(1 for row in controls if row["severity"] == "High"),
            "medium_severity_unique_controls": sum(1 for row in controls if row["severity"] == "Medium"),
            "areas_with_pending_evidence": len(areas),
        },
        "limitations": [
            AUDIT_CONTEXT["evidence_export_limitation"],
            "The audit table did not show owner, assignee, due date, or last-updated fields for the pending rows.",
            "No failed, rejected, Information requested, or non-Ready-for-audit statuses were visible during the Smooth review.",
        ],
    }


def table_sheet(
    title: str,
    subtitle: str,
    headers: list[str],
    data_rows: list[list[Any]],
    widths: list[int],
) -> dict[str, Any]:
    rows = [[title], [subtitle], [], headers]
    styles = [[1], [2], [0], [3 for _ in headers]]
    for row in data_rows:
        rows.append(row)
        styles.append(row_styles(row))
    return {
        "name": title[:31],
        "rows": rows,
        "styles": styles,
        "widths": widths,
        "auto_filter": f"A4:{col_letter(len(headers))}{len(rows)}",
        "freeze_rows": 4,
    }


def executive_summary_sheet(data: dict[str, Any]) -> dict[str, Any]:
    totals = AUDIT_CONTEXT["control_export"]["aggregate_totals"]
    pending = totals["Evidence pending (Count)"]
    required = totals["Total evidence required (Count)"]
    controls = data["unique_control_findings"]
    area_counts = Counter(row["Area Code"] for row in PENDING_ROWS)
    top_areas = "; ".join(f"{area}: {count}" for area, count in area_counts.most_common())
    pending_rate = f"{pending / required:.1%}"
    rows: list[list[Any]] = [
        ["SEC-47 SOC 2 Audit Concern Report"],
        ["Sprinto audit: 2026 - SOC2 Audit"],
        [],
        ["Overall conclusion", "No failed/rejected statuses were found, but 22 of 443 evidence mappings remain pending across 9 unique controls. These should be closed before auditor review to reduce risk of requests or exceptions."],
        ["Audit URL", AUDIT_URL],
        ["Extracted at", EXTRACTED_AT],
        [],
        ["Metric", "Value", "Comment"],
        ["Evidence tab controls", 88, "Visible Sprinto evidence tab count"],
        ["Ready for audit controls", 88, "All visible controls were marked Ready for audit"],
        ["Information requested", 0, "No information-requested status visible"],
        ["Review completed", 0, "No rows had completed auditor review yet"],
        ["Control mapping rows reviewed", 171, "Rows in ControlMappingFile.csv"],
        ["Evidence required", required, "Aggregate from ControlMappingFile.csv"],
        ["Evidence met", totals["Evidence met (Count)"], "Aggregate from ControlMappingFile.csv"],
        ["Evidence pending", pending, "Aggregate from ControlMappingFile.csv"],
        ["Pending evidence rate", pending_rate, "Pending / required evidence mappings"],
        ["Unique controls with pending evidence", len(controls), "Grouped by SDC control number"],
        ["High severity controls", sum(1 for row in controls if row["severity"] == "High"), "At least one mapped criteria row has 0 completed evidence or the UI shows 1/1 incomplete"],
        ["Medium severity controls", sum(1 for row in controls if row["severity"] == "Medium"), "Partially evidenced controls with one incomplete evidence set"],
        ["Areas affected", top_areas, "Pending mapping rows by SOC 2 area code"],
        [],
        ["Immediate priorities", "SDC 23 Internal Audit using Sprinto; SDC 60 Testing for Reliability and Integrity; SDC 97 Testing", "These have 1/1 incomplete UI evidence or mapped criteria rows with zero completed evidence."],
        ["Next audit action", "Resolve pending evidence, rerun ControlMappingFile.csv, and confirm Evidence pending (Count) is 0.", "Evidence export itself requires email delivery and was not directly inspectable in-browser."],
    ]
    styles = []
    for idx, row in enumerate(rows, start=1):
        if idx == 1:
            styles.append([1])
        elif idx in {2, 4, 23, 24}:
            styles.append([2 for _ in row])
        elif idx == 8:
            styles.append([3 for _ in row])
        elif not row:
            styles.append([0])
        else:
            styles.append(row_styles(row))
    return {
        "name": "Executive Summary",
        "rows": rows,
        "styles": styles,
        "widths": [28, 72, 70],
        "auto_filter": "A8:C21",
        "freeze_rows": 8,
    }


def unique_controls_sheet(data: dict[str, Any]) -> dict[str, Any]:
    headers = [
        "Severity",
        "Control Number",
        "Requirement",
        "UI Evidence Completion",
        "UI Status",
        "Mapped Criteria",
        "Area Codes",
        "Pending Mapping Rows",
        "Zero-Completed Mappings",
        "Why Concerning",
        "Recommended Remediation",
        "Control Statement",
    ]
    rows = []
    for control in data["unique_control_findings"]:
        rows.append(
            [
                control["severity"],
                control["control_number"],
                control["ui_requirement"],
                control["ui_completion"],
                control["ui_status"],
                control["criteria_codes"],
                control["area_codes"],
                control["mapping_count"],
                control["zero_completed_mappings"],
                control["why_concerning"],
                control["recommended_remediation"],
                control["control_statement"],
            ]
        )
    return table_sheet(
        "Unique Control Risks",
        "One row per SDC control with pending evidence.",
        headers,
        rows,
        [14, 14, 42, 20, 18, 30, 22, 18, 22, 62, 62, 72],
    )


def finding_detail_sheet(data: dict[str, Any]) -> dict[str, Any]:
    headers = [
        "Finding ID",
        "Severity",
        "Area Code",
        "Area",
        "Criteria Code",
        "Control Number",
        "Requirement",
        "UI Evidence Completion",
        "Evidence Met",
        "Evidence Pending",
        "Evidence Required",
        "Control Statement",
        "Criteria",
        "Why Concerning",
        "Recommended Remediation",
        "Source",
    ]
    controls_by_number = {row["control_number"]: row for row in data["unique_control_findings"]}
    rows = []
    for index, pending_row in enumerate(PENDING_ROWS, start=1):
        control = controls_by_number[int(pending_row["Control Number"])]
        rows.append(
            [
                f"SEC47-{index:03d}",
                control["severity"],
                pending_row["Area Code"],
                pending_row["Area"],
                pending_row["Criteria Code"],
                pending_row["Control Number"],
                control["ui_requirement"],
                control["ui_completion"],
                pending_row["Evidence met (Count)"],
                pending_row["Evidence pending (Count)"],
                pending_row["Total evidence required (Count)"],
                pending_row["Control Statement"],
                pending_row["Criteria"],
                control["why_concerning"],
                control["recommended_remediation"],
                "Sprinto ControlMappingFile.csv export plus evidence tab UI confirmation",
            ]
        )
    return table_sheet(
        "Finding Detail",
        "Mapping-level rows with Evidence pending (Count) greater than 0.",
        headers,
        rows,
        [14, 12, 12, 28, 14, 14, 44, 20, 12, 14, 14, 70, 78, 62, 62, 44],
    )


def area_summary_sheet(data: dict[str, Any]) -> dict[str, Any]:
    headers = [
        "Area Code",
        "Area",
        "Pending Mapping Rows",
        "Unique Controls",
        "Pending Evidence Count",
        "Zero-Completed Mappings",
        "Criteria Codes",
        "Audit Concern",
    ]
    rows = []
    for area in data["area_summary"]:
        concern = (
            "High concern: includes mappings with zero completed evidence."
            if area["zero_completed_mappings"]
            else "Medium concern: pending evidence remains despite Ready for audit status."
        )
        rows.append(
            [
                area["area_code"],
                area["area"],
                area["pending_mapping_count"],
                area["unique_controls"],
                area["pending_evidence_count"],
                area["zero_completed_mappings"],
                area["criteria_codes"],
                concern,
            ]
        )
    return table_sheet(
        "Area Summary",
        "Pending evidence concentration by SOC 2 area.",
        headers,
        rows,
        [12, 34, 20, 16, 20, 22, 36, 58],
    )


def source_notes_sheet(data: dict[str, Any]) -> dict[str, Any]:
    rows = [
        ["Source Notes and Limitations"],
        ["Audit reviewed", AUDIT_CONTEXT["audit_title"]],
        ["Audit URL", AUDIT_URL],
        ["Smooth profile", AUDIT_CONTEXT["smooth_profile"]],
        ["Extracted at", EXTRACTED_AT],
        [],
        ["Evidence tab count", "Value"],
        *[[key, value] for key, value in AUDIT_CONTEXT["evidence_tab_counts"].items()],
        [],
        ["Control export fact", "Value"],
        ["Filename", AUDIT_CONTEXT["control_export"]["filename"]],
        ["Rows", AUDIT_CONTEXT["control_export"]["row_count"]],
        ["Evidence met total", AUDIT_CONTEXT["control_export"]["aggregate_totals"]["Evidence met (Count)"]],
        ["Evidence pending total", AUDIT_CONTEXT["control_export"]["aggregate_totals"]["Evidence pending (Count)"]],
        ["Evidence required total", AUDIT_CONTEXT["control_export"]["aggregate_totals"]["Total evidence required (Count)"]],
        ["Columns", ", ".join(AUDIT_CONTEXT["control_export"]["columns"])],
        [],
        ["Limitation", "Detail"],
        *[["Limitation", note] for note in data["limitations"]],
        [],
        ["Interpretation note", "The audit table reported 88 controls as Ready for audit. This workbook treats pending evidence counts as report-risk indicators because SOC 2 auditors may request, reject, or exception controls whose supporting evidence remains incomplete."],
    ]
    styles = []
    for idx, row in enumerate(rows, start=1):
        if idx == 1:
            styles.append([1])
        elif len(row) == 2 and row[0] in {"Evidence tab count", "Control export fact", "Limitation", "Interpretation note"}:
            styles.append([3, 3])
        elif not row:
            styles.append([0])
        else:
            styles.append(row_styles(row))
    return {
        "name": "Source Notes",
        "rows": rows,
        "styles": styles,
        "widths": [26, 116],
        "auto_filter": None,
        "freeze_rows": 1,
    }


def sheet_xml(sheet: dict[str, Any]) -> str:
    rows = sheet["rows"]
    styles = sheet["styles"]
    max_cols = max(len(row) for row in rows)
    max_rows = len(rows)
    dimension_ref = f"A1:{col_letter(max_cols)}{max_rows}"
    cols_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(sheet["widths"], start=1)
    )
    row_xml = []
    for row_idx, values in enumerate(rows, start=1):
        style_row = styles[row_idx - 1] if row_idx - 1 < len(styles) else []
        if not values:
            row_xml.append(f'<row r="{row_idx}"/>')
            continue
        height = 28 if row_idx <= sheet.get("freeze_rows", 1) else 58
        if row_idx == 1:
            height = 30
        cells = []
        for col_idx, value in enumerate(values, start=1):
            style = style_row[col_idx - 1] if col_idx - 1 < len(style_row) else 4
            cells.append(cell_xml(f"{col_letter(col_idx)}{row_idx}", value, style))
        row_xml.append(f'<row r="{row_idx}" ht="{height}" customHeight="1">{"".join(cells)}</row>')

    pane_xml = ""
    freeze_rows = sheet.get("freeze_rows", 1)
    if freeze_rows:
        top_left = f"A{freeze_rows + 1}"
        pane_xml = (
            f'<pane ySplit="{freeze_rows}" topLeftCell="{top_left}" '
            'activePane="bottomLeft" state="frozen"/>'
        )
    filter_xml = f'<autoFilter ref="{sheet["auto_filter"]}"/>' if sheet.get("auto_filter") else ""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension_ref}"/>
  <sheetViews><sheetView workbookViewId="0">{pane_xml}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{cols_xml}</cols>
  <sheetData>{''.join(row_xml)}</sheetData>
  {filter_xml}
</worksheet>'''


def styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE599"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="9">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''


def write_workbook(path: Path, sheets: list[dict[str, Any]]) -> None:
    workbook_sheets = []
    rels = []
    content_overrides = []
    for idx, sheet in enumerate(sheets, start=1):
        workbook_sheets.append(
            f'<sheet name="{escape(sheet["name"])}" sheetId="{idx}" r:id="rId{idx}"/>'
        )
        rels.append(
            f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>'
        )
        content_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    styles_rel_id = f"rId{len(sheets) + 1}"
    rels.append(
        f'<Relationship Id="{styles_rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{''.join(workbook_sheets)}</sheets>
</workbook>'''
    workbook_rels_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {''.join(rels)}
</Relationships>'''
    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    content_types_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  {''.join(content_overrides)}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>SEC-47 SOC 2 Audit Concern Report</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''
    titles = "".join(f"<vt:lpstr>{escape(sheet['name'])}</vt:lpstr>" for sheet in sheets)
    app_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheets)}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="{len(sheets)}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>
  <Company>Salient</Company>
</Properties>'''

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", rels_xml)
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/styles.xml", styles_xml())
        for idx, sheet in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_xml(sheet))


def build() -> None:
    data = source_dataset()
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SOURCE_PATH.write_text(json.dumps(data, indent=2) + "\n")
    sheets = [
        executive_summary_sheet(data),
        unique_controls_sheet(data),
        finding_detail_sheet(data),
        area_summary_sheet(data),
        source_notes_sheet(data),
    ]
    write_workbook(WORKBOOK_PATH, sheets)
    print(f"Wrote {WORKBOOK_PATH.relative_to(ROOT)}")
    print(f"Wrote {SOURCE_PATH.relative_to(ROOT)}")


if __name__ == "__main__":
    build()
