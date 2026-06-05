# SEC-46 SOC 2 Auditor Request Simulation

Collected on 2026-06-05 from Sprinto using the `oneleet` smooth browser
profile.

Sprinto audit: `2026 - SOC2 Audit`

Audit URL:
<https://app.sprinto.com/app/admin/audits/840b09b2-88b8-450d-9147-2d68f70f8bf2>

Framework: SOC 2 Security, SOC 2 Confidentiality

Audit period: 28 Sep, 2025 - 28 Sep, 2026

Observed audit status: Evidence review has not begun. Sprinto showed 88
controls, all ready for audit.

## Source Signal

The Sprinto evidence tab showed 3 evidence sets with incomplete evidence counts:

| Evidence | Evidence tab status | Detail status | Criteria |
| --- | --- | --- | --- |
| SDC 23: Internal Audit using Sprinto | 1/1 incomplete | Ready for audit | CC4.1, CC4.2, CC5.2, CC7.3, CC7.4 |
| SDC 25: Periodic Review & Update of Cybersecurity & Privacy Program | 1/4 incomplete | Ready for audit | CC1.2, CC1.3, CC4.1, CC4.2, CC5.2 |
| SDC 389: Updates During Installations / Removals | 1/3 incomplete | Ready for audit | CC4.1 |

Owner, assignee, explicit manual-vs-automated evidence type, and due dates were
not visible in the inspected Sprinto views. No pagination was visible in the
evidence tab. The browser task could not conclusively prove there were no hidden
rows behind unavailable filters, but the visible view reported 3 evidence sets.

## Simulated Auditor Requests

### SDC 23: Internal Audit using Sprinto

Control detail observed in Sprinto: Entity uses Sprinto, a continuous monitoring
system, to track and report the health of the information security program to
the Information Security Officer and other stakeholders.

Request:

Please provide evidence that Sprinto was used during the audit period to monitor
the information security program and report control health to the Information
Security Officer and relevant stakeholders. Include the dated monitoring or
internal audit report/export, the reviewer or approver, the review date, any
exceptions identified, and evidence that exceptions were tracked through
remediation or accepted with rationale.

Expected artifacts:

- Sprinto control health or internal audit export for the audit period.
- Dated review record, meeting notes, or sign-off by the Information Security
  Officer or delegated reviewer.
- Exception/remediation tracker or tickets for any findings identified during
  review.
- Evidence of stakeholder distribution or acknowledgement, if available.

Suggested owner: Information Security Officer or compliance owner.

### SDC 25: Periodic Review & Update of Cybersecurity & Privacy Program

Control detail observed in Sprinto: Entity's Senior Management reviews and
approves the state of the Information Security program including policies,
standards, and procedures, at planned intervals or if significant changes occur
to ensure their continuing suitability, adequacy, and effectiveness.

Request:

Please provide evidence that senior management reviewed and approved the
cybersecurity and privacy program during the audit period, including policies,
standards, procedures, and any material updates. The evidence should show the
review date, attendees or approvers, scope of materials reviewed, decisions
made, policy or program changes approved, and follow-up actions assigned.

Expected artifacts:

- Security/privacy program review deck, agenda, or packet.
- Dated meeting minutes, approval record, or executive sign-off.
- List of policies, standards, and procedures reviewed.
- Change log or action tracker for updates, exceptions, or follow-up items.

Suggested owner: Senior management, Information Security Officer, or compliance
owner.

### SDC 389: Updates During Installations / Removals

Control detail observed in Sprinto: Entity periodically updates and reviews the
inventory of systems as a part of installations, removals, and system updates.

Request:

Please provide evidence that the system inventory was updated and reviewed when
systems were installed, removed, or materially changed during the audit period.
Include the inventory records, related change or decommissioning tickets, dates
of updates, reviewer approval, and any exceptions where inventory updates were
delayed or not required.

Expected artifacts:

- Current system inventory or asset register.
- Inventory change history for installations, removals, and material system
  updates.
- Change management, deployment, or decommissioning tickets tied to inventory
  updates.
- Reviewer sign-off or periodic inventory review evidence.

Suggested owner: Infrastructure, IT operations, or compliance owner.
