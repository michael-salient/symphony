#!/usr/bin/env python3
"""Build the SEC-5 Sprinto to Oneleet risk comparison workbook.

This intentionally avoids external spreadsheet libraries because the preferred
artifact-tool runtime was not available in this workspace. The XLSX writer below
creates a simple OpenXML workbook with one comparison table.
"""

from __future__ import annotations

import csv
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from io import StringIO
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
SOURCE_DIR = ROOT / "source_data"
OUTPUT_DIR = ROOT / "outputs"
WORKBOOK_PATH = OUTPUT_DIR / "sec5_sprinto_oneleet_risk_comparison.xlsx"


SPRINTO_TSV = """risk_id\tname\tcategory\tinherent_score\tresidual_score\tstatus\ttype\towner\ttreatment
SRK 1\tNetwork security controls\tCard Data Security\tHigh: 8\tLow: 4\tComplete\tProcess based\tMT\tAccept
SRK 4\tCompromising cardholder data\tCard Data Security\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 5\tSystems and networks in the cardholder data environment\tCard Data Security\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 6\tSecurity vulnerabilities in cardholder data environment\tCard Data Security\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 7\tUnauthorized access to cardholder data\tCard Data Security\tMed: 4.2\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 8\tManagement of cardholder data\tCard Data Security\tMed: 5.6\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 12\tTransmission of cardholder data to vendors\tCard Data Security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 14\tOutdated card data\tCard Data Security\tHigh: 8\tLow: 4\tComplete\tProcess based\tMT\tAccept
SRK 15\tExploited vulnerability risk\tInfrastructure\tHigh: 8\tLow: 4\tComplete\tProcess based\tMT\tAccept
SRK 16\tIncident and vulnerability not reported\tVulnerabilities and Incidents\tHigh: 6.4\tLow: 3.2\tComplete\tProcess based\tMT\tAccept
SRK 17\tVulnerabilities in code\tVulnerabilities and Incidents\tMed: 4.8\tLow: 2.4\tComplete\tProcess based\tMT\tAccept
SRK 18\tCompromising production systems\tInfrastructure\tMed: 4.8\tLow: 2.4\tComplete\tProcess based\tMT\tAccept
SRK 19\tUnsafe encryption algorithms\tVulnerabilities and Incidents\tMed: 4.8\tLow: 2.4\tComplete\tProcess based\tMT\tAccept
SRK 20\tCorrupted database\tData security\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 21\tNatural disasters or security incidents risk\tVulnerabilities and Incidents\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 22\tUnauthorized people/staff may gain physical access to the production infrastructure or data centers\tAccess control\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 23\tCompromising encryption keys\tInfrastructure\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 24\tEmployee workstation vulnerabilities\tEndpoint\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 25\tData compromisation\tEndpoint\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 26\tMalicious software on endpoints\tEndpoint\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 27\tUnsecured workstations\tEndpoint\tMed: 6\tLow: 2\tComplete\tProcess based\tMT\tAccept
SRK 28\tCapacity limits\tInfrastructure\tLow: 3.2\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 29\tSynchronisation of internal controls\tControl health\tLow: 3.2\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 30\tInternal controls inefficacy\tControl health\tLow: 3.2\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 31\tData in open/public networks\tData security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 32\tUnsecured Confidential Data\tData security\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 33\tAd-hoc use of data\tData security\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 34\tSecurity Incident Detection\tVulnerabilities and Incidents\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 35\tEmployee misrepresentation of information\tFraud\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 36\tSecurity attack on production environment\tInfrastructure\tMed: 4.8\tLow: 1.6\tComplete\tProcess based\tMT\tAccept
SRK 37\tUnrequired staff access\tAccess control\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 38\tAccess to off-boarded staff\tAccess control\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 39\tFinancial fraud\tFraud\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 40\tConflicts of Interest\tFraud\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 41\tBribery of officials\tFraud\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 42\tOpen ports\tInfrastructure\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 43\tNetwork configuration\tVulnerabilities and Incidents\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 44\tSpoofing and attacks\tVulnerabilities and Incidents\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 45\tCredentials check-in\tVulnerabilities and Incidents\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 46\tDevelopers' security impact\tVulnerabilities and Incidents\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 47\tUnauthorized changes\tAccess control\tLow: 3.6\tLow: 1.2\tComplete\tProcess based\tMT\tAccept
SRK 48\tExecutive awareness of internal controls\tControl health\tLow: 2.4\tLow: 0.8\tComplete\tProcess based\tMT\tAccept
SRK 49\tProduction recovery delay\tVulnerabilities and Incidents\tLow: 1.6\tLow: 0.8\tComplete\tProcess based\tMT\tAccept
SRK 50\tUnintended bugs/downtime\tVulnerabilities and Incidents\tLow: 2.4\tLow: 0.8\tComplete\tProcess based\tMT\tAccept
SRK 51\tEmergency code changes\tVulnerabilities and Incidents\tLow: 2.4\tLow: 0.8\tComplete\tProcess based\tMT\tAccept
SRK 52\tData compromisation\tOthers\tLow: 2.4\tLow: 0.8\tComplete\tProcess based\tMT\tAccept
SRK 53\tData deletion request\tData security\tLow: 1.2\tLow: 0.4\tComplete\tProcess based\tMT\tAccept
SRK 54\tEmployee acceptance of gifts or favors from vendors\tFraud\tLow: 1.2\tLow: 0.4\tComplete\tProcess based\tMT\tAccept
SRK 55\tResponse to security incidents or vulnerabilities\tVulnerabilities and Incidents\tLow: 1.2\tLow: 0.4\tComplete\tProcess based\tMT\tAccept
SRK 56\tStaff awareness of policies and procedures\tOthers\tLow: 1.2\tLow: 0.4\tComplete\tProcess based\tMT\tAccept
SRK 57\tUnauthorized access to critical systems\tAccess control\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 58\tClearly defined agreements with customers\tOthers\tLow: 2.8\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 59\tCompromising sensitive cardholder information\tData security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 60\tCompromising financial data\tData security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 61\tMan in the middle security attacks\tData security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
SRK 62\tUnauthorized access to critical data\tData security\tHigh: 7\tLow: 0\tComplete\tProcess based\tMT\tAccept
"""


ONELEET_ACTIVE_TSV = """id\tname\tcategory\tinherent_risk\tresidual_risk\tassessed\tstatus
62\tBreach of stored customer credentials enables account takeovers and attacks on connected systems\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
66\tCloud infrastructure misconfiguration exposes sensitive data\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
70\tCompany data is exposed through lost or stolen mobile devices\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
74\tCompromised identity verification data undermines customer authentication\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
78\tConfidential customer data is publicly released or exploited from a data breach\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
82\tCredentials or API keys are accidentally exposed in public repositories\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
86\tCritical system has single point of failure without redundancy\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
90\tCritical vulnerability in public-facing APIs is exploited\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
94\tDatabase security rules misconfiguration exposes data publicly\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
98\tEmployee accounts are compromised through phishing or social engineering\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
102\tGrowth without security oversight erodes least-privilege controls and expands attack surface\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
106\tInability to locate or delete personal data prevents fulfilling data subject requests\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
110\tInadequate access controls allow unauthorized data access\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
114\tMalicious insider intentionally causes security incident\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
118\tOne of the company's critical 3rd party service providers goes out of business\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
122\tPayment processing fraud results in chargebacks and losses\tFinancial\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
126\tPCI-DSS non-compliance results in fines and loss of payment processing\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
130\tPhysical offices are shut down due to a natural disaster or pandemic\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
134\tProduction data is deleted or corrupted due to negligence during migration process\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
138\tProduction database is encrypted from a ransomware attack\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
142\tProduction defects and downtime erode customer trust and drive churn\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
146\tSelf-managed servers go unpatched creating known vulnerabilities\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
150\tSensitive PII breach attracts class action litigation\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
154\tSupply chain attack compromises development or deployment pipeline\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
158\tVulnerabilities are discovered in core 3rd party libraries\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tActive - needs assessment
"""


ONELEET_ARCHIVED_TSV = """id\tname\tcategory\tinherent_risk\tresidual_risk\tassessed\tstatus
162\tCash runway depletes faster than anticipated\tFinancial\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
167\tCompany decisions are no longer aligned with customer needs or feedback\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
172\tCompany employees use insider knowledge to commit fraud\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
177\tCompany is brought to court due to lawsuit or legal challenge\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
182\tCompetitive pricing pressure erodes margins\tFinancial\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
187\tCore product fails to achieve product-market fit\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
192\tCore product offering is copied by competitors\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
197\tCustomer data exposure damages Salient's brand reputation and erodes trust\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
202\tDecentralized decision-making creates fragmented standards across architecture and security\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
207\tDown round financing dilutes existing shareholders and damages morale\tFinancial\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
212\tEmployees exfiltrate company data for personal gain\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
217\tExpense report fraud through submission of false or inflated business expenses\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
222\tIncumbent in space creates offering that is competitive with core product\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
227\tLack of effective customer support leads to increased churn\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
232\tLegacy systems and technical debt delay feature delivery\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
237\tManagement fraud through financial statement manipulation or misrepresentation\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
242\tMarket expansion into new regions introduces conflicting regulatory requirements\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
247\tProduct development velocity is not sufficient to meet deadlines, budget or required outcomes\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
252\tProduct loses relevance due to shifts in market conditions\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
257\tRapid growth outpaces organizational infrastructure\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
262\tRegulatory reporting fraud through falsified compliance documentation or metrics\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
267\tRevenue is concentrated in a small number of customers\tFinancial\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
272\tSales team fraud through falsified customer contracts or revenue recognition manipulation\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
277\tSalient's production database is encrypted from a ransomware attack\tSecurity\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
282\tSingle-platform dependency creates existential business risk\tStrategic & Market\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
287\tTCPA Violations cause monetary damages to customers and Salient.\tLegal & Compliance\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
292\tToll Fraud could be generate significant financial losses\tFraud\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
297\tUndocumented knowledge and processes create dangerous dependency on founders and key personnel\tOperational\tNot yet assessed\tNot yet assessed\tNot yet assessed\tArchived
"""


ONELEET_LIBRARY_TSV = """name\tcategory\tstatus
Core product is taken down by a distributed-denial-of-service (DDoS) attack\tSecurity\tLibrary only
Confidential customer data is publicly released or exploited from a data breach\tSecurity\tLibrary only
Production database is encrypted from a ransomware attack\tSecurity\tLibrary only
Vulnerabilities are discovered in core 3rd party libraries\tSecurity\tLibrary only
Employee accounts are compromised through phishing or social engineering\tSecurity\tLibrary only
Malicious insider intentionally causes security incident\tSecurity\tLibrary only
Critical vulnerability in public-facing APIs is exploited\tSecurity\tLibrary only
Cloud infrastructure misconfiguration exposes sensitive data\tSecurity\tLibrary only
Cache or CDN misconfiguration leaks user-specific content across users\tSecurity\tLibrary only
Database security rules misconfiguration exposes data publicly\tSecurity\tLibrary only
Self-managed servers go unpatched creating known vulnerabilities\tSecurity\tLibrary only
Supply chain attack compromises development or deployment pipeline\tSecurity\tLibrary only
Credentials or API keys are accidentally exposed in public repositories\tSecurity\tLibrary only
Breach of stored customer credentials enables account takeovers and attacks on connected systems\tSecurity\tLibrary only
Inadequate access controls allow unauthorized data access\tSecurity\tLibrary only
Growth without security oversight erodes least-privilege controls and expands attack surface\tSecurity\tLibrary only
Company data is exposed through lost or stolen mobile devices\tSecurity\tLibrary only
Company hiring objectives are not aligned with market trends\tOperational\tLibrary only
Company has difficulty hiring due to market or technology stack\tOperational\tLibrary only
Inability to scale operations up or down as efficiently as needed\tOperational\tLibrary only
One of the company's critical 3rd party service providers goes out of business\tOperational\tLibrary only
Physical offices are shut down due to a natural disaster or pandemic\tOperational\tLibrary only
Product development velocity is not sufficient to meet deadlines, budget or required outcomes\tOperational\tLibrary only
Production data is deleted or corrupted due to negligence during migration process\tOperational\tLibrary only
Critical system has single point of failure without redundancy\tOperational\tLibrary only
On-premises hardware failure causes extended downtime due to manual recovery\tOperational\tLibrary only
Over-reliance on contractors for core business functions\tOperational\tLibrary only
Remote workforce creates productivity challenges\tOperational\tLibrary only
Inadequate incident response capabilities delay recovery\tOperational\tLibrary only
Infrastructure or application misconfigurations cause production outages\tOperational\tLibrary only
Production defects and downtime erode customer trust and drive churn\tOperational\tLibrary only
Legacy systems and technical debt delay feature delivery\tOperational\tLibrary only
Unsupported legacy systems fail and cause extended downtime\tOperational\tLibrary only
Decentralized decision-making creates fragmented standards across architecture and security\tOperational\tLibrary only
Undocumented knowledge and processes create dangerous dependency on founders and key personnel\tOperational\tLibrary only
Rapid growth outpaces organizational infrastructure\tOperational\tLibrary only
Financial failure due to inability to receive additional fundraising\tFinancial\tLibrary only
Cash runway depletes faster than anticipated\tFinancial\tLibrary only
Revenue is concentrated in a small number of customers\tFinancial\tLibrary only
Competitive pricing pressure erodes margins\tFinancial\tLibrary only
Financial losses from fraud or embezzlement\tFinancial\tLibrary only
Additional regulatory compliance laws are enacted that the company needs to fulfill\tLegal & Compliance\tLibrary only
Market expansion into new regions introduces conflicting regulatory requirements\tLegal & Compliance\tLibrary only
Company is brought to court due to lawsuit or legal challenge\tLegal & Compliance\tLibrary only
Intellectual property or strategy disclosure compromises competitive position\tLegal & Compliance\tLibrary only
GDPR violation results in regulatory fines and enforcement action\tLegal & Compliance\tLibrary only
CCPA/CPRA violation results in penalties and litigation\tLegal & Compliance\tLibrary only
Inability to locate or delete personal data prevents fulfilling data subject requests\tLegal & Compliance\tLibrary only
Compromised identity verification data undermines customer authentication\tSecurity\tLibrary only
Sensitive PII breach attracts class action litigation\tLegal & Compliance\tLibrary only
Protected health information (PHI) is mishandled or exposed resulting in HIPAA penalties and loss of healthcare customers\tLegal & Compliance\tLibrary only
PCI-DSS non-compliance results in fines and loss of payment processing\tLegal & Compliance\tLibrary only
Improper handling of children's data results in COPPA penalties\tLegal & Compliance\tLibrary only
SOC 2 audit findings jeopardize enterprise customer relationships\tLegal & Compliance\tLibrary only
Employment-related litigation from discrimination or wrongful termination claims\tLegal & Compliance\tLibrary only
Inadequate corporate governance creates compliance and control gaps\tLegal & Compliance\tLibrary only
Company decisions are no longer aligned with customer needs or feedback\tStrategic & Market\tLibrary only
Core product fails to achieve product-market fit\tStrategic & Market\tLibrary only
Core product offering is copied by competitors\tStrategic & Market\tLibrary only
Lack of effective customer support leads to increased churn\tStrategic & Market\tLibrary only
Customer data exposure damages brand reputation and erodes trust\tStrategic & Market\tLibrary only
Incumbent in space creates offering that is competitive with core product\tStrategic & Market\tLibrary only
Product loses relevance due to shifts in market conditions\tStrategic & Market\tLibrary only
Single-platform dependency creates existential business risk\tStrategic & Market\tLibrary only
Key talent is recruited away by competitors\tStrategic & Market\tLibrary only
Company employees use insider knowledge to commit fraud\tFraud\tLibrary only
Employees exfiltrate company data for personal gain\tFraud\tLibrary only
Accounts payable fraud through creation of fictitious vendors or invoice manipulation\tFraud\tLibrary only
Expense report fraud through submission of false or inflated business expenses\tFraud\tLibrary only
Payroll fraud through ghost employees or unauthorized salary modifications\tFraud\tLibrary only
Procurement fraud through kickbacks or inflated vendor pricing schemes\tFraud\tLibrary only
Management fraud through financial statement manipulation or misrepresentation\tFraud\tLibrary only
Sales team fraud through falsified customer contracts or revenue recognition manipulation\tFraud\tLibrary only
Regulatory reporting fraud through falsified compliance documentation or metrics\tFraud\tLibrary only
Tax fraud through manipulation of financial records or improper deductions\tFraud\tLibrary only
Down round financing dilutes existing shareholders and damages morale\tFinancial\tLibrary only
Reliance on bridge financing creates unfavorable terms\tFinancial\tLibrary only
Employees mutiny and commandeer the ship\tOperational\tLibrary only
"""


MATCH_TSV = """risk_id\toneleet_name\toneleet_status\tconfidence\tproposed_action\treasoning
SRK 1\tCloud infrastructure misconfiguration exposes sensitive data\tActive - needs assessment\tMedium\talready covered\tSprinto's network-control risk is broader than the Oneleet scenario, but the active cloud-misconfiguration/API exposure risks cover the main infrastructure-control failure mode.
SRK 4\tConfidential customer data is publicly released or exploited from a data breach\tActive - needs assessment\tHigh\talready covered\tCardholder compromise is a specific form of confidential customer data breach, also supported by the active PCI-DSS risk.
SRK 5\tPCI-DSS non-compliance results in fines and loss of payment processing\tActive - needs assessment\tMedium\talready covered\tThe Sprinto CDE systems/network risk is PCI-specific; Oneleet represents it as the broader PCI-DSS compliance scenario.
SRK 6\tCritical vulnerability in public-facing APIs is exploited\tActive - needs assessment\tMedium\talready covered\tCDE vulnerability risk is covered by active public-facing API and third-party library vulnerability scenarios, though Oneleet is less PCI-specific.
SRK 7\tInadequate access controls allow unauthorized data access\tActive - needs assessment\tHigh\talready covered\tUnauthorized access to cardholder data maps directly to the active unauthorized data access scenario.
SRK 8\tPCI-DSS non-compliance results in fines and loss of payment processing\tActive - needs assessment\tHigh\talready covered\tCardholder data management obligations are captured by the active PCI-DSS non-compliance scenario.
SRK 12\tPCI-DSS non-compliance results in fines and loss of payment processing\tActive - needs assessment\tMedium\talready covered\tVendor transmission of cardholder data is not named directly, but it is a PCI obligation and is close enough to the active PCI-DSS scenario.
SRK 14\tInability to locate or delete personal data prevents fulfilling data subject requests\tActive - needs assessment\tMedium\talready covered\tOutdated card data is a retention/governance issue; Oneleet's active deletion/locate personal data risk covers the closest data lifecycle failure.
SRK 15\tCritical vulnerability in public-facing APIs is exploited\tActive - needs assessment\tHigh\talready covered\tThe Sprinto exploited-vulnerability risk maps directly to active API exploitation and library vulnerability scenarios.
SRK 16\tInadequate incident response capabilities delay recovery\tLibrary only\tHigh\tcopy over\tSprinto explicitly tracks failure to report incidents/vulnerabilities; the closest Oneleet library risk is not active or archived and should be copied into the assessment.
SRK 17\tVulnerabilities are discovered in core 3rd party libraries\tActive - needs assessment\tHigh\talready covered\tCode vulnerability exposure is covered by active third-party library and public-facing API vulnerability scenarios.
SRK 18\tProduction database is encrypted from a ransomware attack\tActive - needs assessment\tMedium\talready covered\tCompromised production systems are represented by active ransomware, cloud misconfiguration, and critical vulnerability scenarios.
SRK 19\tConfidential customer data is publicly released or exploited from a data breach\tActive - needs assessment\tMedium\tcopy over\tThe active breach scenario captures the outcome, but unsafe encryption algorithms are a distinct PCI/security weakness worth copying over as a separate risk.
SRK 20\tProduction data is deleted or corrupted due to negligence during migration process\tActive - needs assessment\tHigh\talready covered\tSprinto's corrupted database risk maps directly to the active data deletion/corruption scenario.
SRK 21\tPhysical offices are shut down due to a natural disaster or pandemic\tActive - needs assessment\tMedium\talready covered\tOneleet covers the disaster component directly; security incident aspects are separately covered by active security incident and ransomware scenarios.
SRK 22\tInadequate access controls allow unauthorized data access\tActive - needs assessment\tLow\tcopy over\tOneleet has logical access risks, but physical access to production infrastructure/data centers is not explicitly represented and should be copied if still relevant.
SRK 23\tCredentials or API keys are accidentally exposed in public repositories\tActive - needs assessment\tHigh\talready covered\tCompromised encryption keys fit the same credential/secret exposure family as the active API key risk.
SRK 24\tCompany data is exposed through lost or stolen mobile devices\tActive - needs assessment\tMedium\tcopy over\tThe active device-loss risk partially covers endpoint exposure, but workstation vulnerability management is broader and should be copied over.
SRK 25\tCompany data is exposed through lost or stolen mobile devices\tActive - needs assessment\tMedium\talready covered\tEndpoint data compromise is covered by the active lost/stolen device and customer data breach scenarios.
SRK 26\tEmployee accounts are compromised through phishing or social engineering\tActive - needs assessment\tMedium\tcopy over\tEndpoint malware is only indirectly represented through phishing/account compromise; copying it over preserves the malware-specific Sprinto risk.
SRK 27\tCompany data is exposed through lost or stolen mobile devices\tActive - needs assessment\tMedium\talready covered\tUnsecured workstation exposure is close enough to the active lost/stolen device data exposure scenario.
SRK 28\tCritical system has single point of failure without redundancy\tActive - needs assessment\tHigh\talready covered\tCapacity and resilience limits map to the active single-point-of-failure/no-redundancy scenario.
SRK 29\tDecentralized decision-making creates fragmented standards across architecture and security\tArchived\tMedium\tcopy over\tSprinto control synchronization maps to governance/control fragmentation; the closest Oneleet scenario is archived, so it should be copied or unarchived.
SRK 30\tInadequate corporate governance creates compliance and control gaps\tLibrary only\tMedium\tcopy over\tInternal control inefficacy is not active; the closest Oneleet library risk is corporate governance/control gaps and should be copied over.
SRK 31\tCloud infrastructure misconfiguration exposes sensitive data\tActive - needs assessment\tHigh\talready covered\tData exposed on open/public networks maps to active cloud and database misconfiguration exposure scenarios.
SRK 32\tConfidential customer data is publicly released or exploited from a data breach\tActive - needs assessment\tHigh\talready covered\tUnsecured confidential data maps directly to the active customer data breach scenario.
SRK 33\tInability to locate or delete personal data prevents fulfilling data subject requests\tActive - needs assessment\tLow\tcopy over\tAd-hoc data use is a governance/process issue; the active data request scenario is adjacent but does not cover unauthorized informal use well.
SRK 34\tInadequate incident response capabilities delay recovery\tLibrary only\tHigh\tcopy over\tSecurity incident detection is not represented in the active assessment; the library incident-response risk is close enough to add.
SRK 35\tRegulatory reporting fraud through falsified compliance documentation or metrics\tArchived\tMedium\tcopy over\tEmployee misrepresentation is a fraud/control integrity risk; Oneleet has a close fraud scenario but it is archived.
SRK 36\tSalient's production database is encrypted from a ransomware attack\tArchived\tMedium\tOK to leave archived\tThis archived row is a company-specific duplicate of the active production-database ransomware risk, so leaving it archived is acceptable.
SRK 37\tInadequate access controls allow unauthorized data access\tActive - needs assessment\tHigh\talready covered\tUnrequired staff access maps directly to the active unauthorized data access scenario.
SRK 38\tGrowth without security oversight erodes least-privilege controls and expands attack surface\tActive - needs assessment\tHigh\talready covered\tOffboarded staff access is a least-privilege/access review failure covered by the active growth/least-privilege and unauthorized access risks.
SRK 39\tFinancial losses from fraud or embezzlement\tLibrary only\tHigh\tcopy over\tSprinto tracks broad financial fraud; Oneleet only has specific active payment fraud and archived specific fraud scenarios, so the broad library risk should be copied over.
SRK 40\tCompany employees use insider knowledge to commit fraud\tArchived\tMedium\tcopy over\tConflicts of interest are not active in Oneleet; the closest insider-fraud scenario is archived and should be copied if fraud coverage is desired.
SRK 41\tProcurement fraud through kickbacks or inflated vendor pricing schemes\tLibrary only\tMedium\tcopy over\tBribery/gifts risk is best represented by procurement kickbacks; that scenario is library-only, not active.
SRK 42\tCritical vulnerability in public-facing APIs is exploited\tActive - needs assessment\tMedium\talready covered\tOpen ports are a public attack surface issue covered by the active public-facing API exploitation risk.
SRK 43\tCloud infrastructure misconfiguration exposes sensitive data\tActive - needs assessment\tHigh\talready covered\tNetwork configuration risk maps to active cloud infrastructure and database security misconfiguration risks.
SRK 44\tBreach of stored customer credentials enables account takeovers and attacks on connected systems\tActive - needs assessment\tMedium\talready covered\tSpoofing/account attacks are covered by the active stored-credential breach and account-takeover scenario.
SRK 45\tCredentials or API keys are accidentally exposed in public repositories\tActive - needs assessment\tHigh\talready covered\tCredentials checked into source maps directly to the active public repository credential/API key exposure risk.
SRK 46\tSupply chain attack compromises development or deployment pipeline\tActive - needs assessment\tMedium\talready covered\tDeveloper security impact maps to the active development/deployment pipeline compromise scenario.
SRK 47\tSupply chain attack compromises development or deployment pipeline\tActive - needs assessment\tMedium\talready covered\tUnauthorized changes are close to development/deployment pipeline compromise and change integrity risk.
SRK 48\tInadequate corporate governance creates compliance and control gaps\tLibrary only\tMedium\tcopy over\tExecutive awareness of controls is a governance/control gap that is not active in Oneleet; the library risk should be copied over.
SRK 49\tInadequate incident response capabilities delay recovery\tLibrary only\tHigh\tcopy over\tProduction recovery delay maps directly to incident response/recovery delay; this risk is library-only and should be copied.
SRK 50\tProduction defects and downtime erode customer trust and drive churn\tActive - needs assessment\tHigh\talready covered\tUnintended bugs and downtime map directly to the active production defects/downtime scenario.
SRK 51\tSupply chain attack compromises development or deployment pipeline\tActive - needs assessment\tMedium\talready covered\tEmergency code changes are a change-control/deployment integrity issue covered broadly by the active deployment pipeline compromise scenario.
SRK 52\tCustomer data exposure damages Salient's brand reputation and erodes trust\tArchived\tMedium\tOK to leave archived\tThe archived brand-damage scenario is duplicative of the active confidential customer data breach scenario for this Sprinto data compromise risk.
SRK 53\tInability to locate or delete personal data prevents fulfilling data subject requests\tActive - needs assessment\tHigh\talready covered\tData deletion requests map directly to the active data subject request fulfillment scenario.
SRK 54\tProcurement fraud through kickbacks or inflated vendor pricing schemes\tLibrary only\tMedium\tcopy over\tEmployee gift/favor acceptance is closest to procurement kickback/vendor fraud, which is not active and should be copied if Sprinto fraud risks are retained.
SRK 55\tInadequate incident response capabilities delay recovery\tLibrary only\tHigh\tcopy over\tResponse to security incidents/vulnerabilities maps directly to incident response capability, which is not active.
SRK 56\tEmployee accounts are compromised through phishing or social engineering\tActive - needs assessment\tLow\tcopy over\tSecurity awareness is only indirectly covered by phishing/account compromise; copy the policy/procedure awareness risk if audit evidence expects it.
SRK 57\tInadequate access controls allow unauthorized data access\tActive - needs assessment\tHigh\talready covered\tUnauthorized access to critical systems maps directly to active unauthorized access controls risk.
SRK 58\tCompany is brought to court due to lawsuit or legal challenge\tArchived\tMedium\tcopy over\tCustomer agreement clarity is a legal exposure; the closest Oneleet legal/lawsuit scenario is archived and should be copied if contract risk remains in scope.
SRK 59\tPCI-DSS non-compliance results in fines and loss of payment processing\tActive - needs assessment\tHigh\talready covered\tSensitive cardholder information compromise is covered by active PCI-DSS and confidential customer data breach scenarios.
SRK 60\tConfidential customer data is publicly released or exploited from a data breach\tActive - needs assessment\tHigh\talready covered\tCompromising financial data maps to active confidential customer data breach and sensitive PII litigation risks.
SRK 61\tBreach of stored customer credentials enables account takeovers and attacks on connected systems\tActive - needs assessment\tMedium\tcopy over\tMan-in-the-middle attacks are only indirectly covered by credential/account takeover; copy over the network attack scenario if PCI specificity is needed.
SRK 62\tInadequate access controls allow unauthorized data access\tActive - needs assessment\tHigh\talready covered\tUnauthorized access to critical data maps directly to the active unauthorized data access risk.
"""


HEADERS = [
    "Sprinto Risk ID",
    "Sprinto Risk Name",
    "Sprinto Category",
    "Sprinto Inherent Score",
    "Sprinto Residual Score",
    "Oneleet Risk Name",
    "Oneleet Category",
    "Oneleet Inherent Risk",
    "Oneleet Residual Risk",
    "Oneleet Active/Archived Status",
    "Match Confidence",
    "Proposed Action",
    "Reasoning",
]


@dataclass(frozen=True)
class RiskRow:
    risk_id: str
    name: str
    category: str
    inherent_score: str
    residual_score: str


def parse_tsv(text: str) -> list[dict[str, str]]:
    return list(csv.DictReader(StringIO(text.strip()), delimiter="\t"))


def normalize_name(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip()).lower()


def col_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xlsx_cell(ref: str, value: str, style: int | None = None) -> str:
    attrs = f' r="{ref}" t="inlineStr"'
    if style is not None:
        attrs += f' s="{style}"'
    escaped_value = escape(value or "")
    return f'<c{attrs}><is><t>{escaped_value}</t></is></c>'


def write_minimal_xlsx(path: Path, rows: list[list[str]]) -> None:
    widths = [16, 36, 22, 18, 18, 50, 22, 20, 20, 24, 18, 18, 80]
    sheet_rows = []
    for row_num, values in enumerate(rows, start=1):
        style = 1 if row_num == 1 else 2
        height = 28 if row_num == 1 else 84
        cells = "".join(
            xlsx_cell(f"{col_letter(col_num)}{row_num}", str(value), style)
            for col_num, value in enumerate(values, start=1)
        )
        sheet_rows.append(f'<row r="{row_num}" ht="{height}" customHeight="1">{cells}</row>')

    cols = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate(widths, start=1)
    )
    dimension_ref = f"A1:{col_letter(len(rows[0]))}{len(rows)}"
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="{dimension_ref}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>{cols}</cols>
  <sheetData>{''.join(sheet_rows)}</sheetData>
  <autoFilter ref="{dimension_ref}"/>
</worksheet>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Comparison" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    workbook_rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    content_types_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>SEC-5 Sprinto to Oneleet Risk Comparison</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Comparison</vt:lpstr></vt:vector></TitlesOfParts>
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
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def build() -> None:
    sprinto = parse_tsv(SPRINTO_TSV)
    active = parse_tsv(ONELEET_ACTIVE_TSV)
    archived = parse_tsv(ONELEET_ARCHIVED_TSV)
    library = parse_tsv(ONELEET_LIBRARY_TSV)
    matches = parse_tsv(MATCH_TSV)

    if len(sprinto) != 56:
        raise RuntimeError(f"Expected 56 Sprinto risks, got {len(sprinto)}")
    if len(active) != 25:
        raise RuntimeError(f"Expected 25 active needs-assessment Oneleet risks, got {len(active)}")
    if len(archived) != 28:
        raise RuntimeError(f"Expected 28 archived Oneleet risks, got {len(archived)}")
    if len(matches) != 56:
        raise RuntimeError(f"Expected 56 comparison matches, got {len(matches)}")

    oneleet_by_name_status: dict[tuple[str, str], dict[str, str]] = {}
    for row in active + archived + library:
        row.setdefault("id", "")
        if row["status"] == "Library only":
            row["inherent_risk"] = "N/A - library only"
            row["residual_risk"] = "N/A - library only"
            row["assessed"] = "N/A - library only"
        else:
            row.setdefault("inherent_risk", "Not yet assessed")
            row.setdefault("residual_risk", "Not yet assessed")
            row.setdefault("assessed", "Not yet assessed")
        oneleet_by_name_status[(normalize_name(row["name"]), row["status"])] = row

    sprinto_by_id = {row["risk_id"]: row for row in sprinto}
    comparison_rows: list[dict[str, str]] = []
    workbook_rows = [HEADERS]

    for match in matches:
        sprinto_row = sprinto_by_id[match["risk_id"]]
        oneleet_row = oneleet_by_name_status.get(
            (normalize_name(match["oneleet_name"]), match["oneleet_status"])
        )
        if not oneleet_row:
            raise RuntimeError(f"Missing Oneleet lookup for {match['risk_id']}: {match['oneleet_name']}")

        comparison = {
            "sprinto_risk_id": sprinto_row["risk_id"],
            "sprinto_name": sprinto_row["name"],
            "sprinto_category": sprinto_row["category"],
            "sprinto_inherent_score": sprinto_row["inherent_score"],
            "sprinto_residual_score": sprinto_row["residual_score"],
            "oneleet_name": oneleet_row["name"],
            "oneleet_category": oneleet_row["category"],
            "oneleet_inherent_risk": oneleet_row["inherent_risk"],
            "oneleet_residual_risk": oneleet_row["residual_risk"],
            "oneleet_status": oneleet_row["status"],
            "match_confidence": match["confidence"],
            "proposed_action": match["proposed_action"],
            "reasoning": match["reasoning"],
        }
        comparison_rows.append(comparison)
        workbook_rows.append([comparison[key] for key in comparison])

    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (SOURCE_DIR / "sprinto_risks.json").write_text(json.dumps(sprinto, indent=2) + "\n")
    (SOURCE_DIR / "oneleet_risks.json").write_text(
        json.dumps(
            {
                "assessment": {
                    "name": "April 2026 Risk Assessment",
                    "risk_count": 53,
                    "needs_assessment_count": 25,
                    "extraction_note": "Verified by smooth_task on 2026-06-04: the assessment page showed All risks=53 and Needs assessment=25; the Export button downloaded risks.csv; the assessment active view produced 25 needs-assessment rows; the archived view produced 28 archived rows. Risk Library rows are retained as Library only when they are not present in the active or archived assessment views.",
                },
                "active_needs_assessment": active,
                "archived": archived,
                "library": library,
            },
            indent=2,
        )
        + "\n"
    )
    (SOURCE_DIR / "oneleet_smooth_extract_summary.json").write_text(
        json.dumps(
            {
                "verified_at": "2026-06-04",
                "method": "smooth_task server-side browser extraction",
                "live_url": "https://app.oneleet.com/tenants/553deae1-f34c-4e47-92a6-c8bc9aa236a4/risk-management/assessments/dbad09db-e09e-45ac-8615-a554d7ffc224",
                "login_wall": False,
                "page_title": "April 2026 Risk Assessment",
                "summary": {
                    "All risks": 53,
                    "Needs assessment": 25,
                    "Accepted": 0,
                    "Avoided": 0,
                    "Mitigated": 0,
                    "Transferred": 0,
                },
                "export_files_reported_by_browser": ["/cache/.browser/downloads/risks.csv"],
                "active_rows_collected": len(active),
                "archived_rows_collected": len(archived),
                "library_rows_collected": len(library),
                "note": "The browser task returned the same active and archived row counts represented in source_data/oneleet_risks.json. Library rows are captured as candidate source risks, not current assessment rows.",
            },
            indent=2,
        )
        + "\n"
    )
    (SOURCE_DIR / "sec5_risk_comparison_rows.json").write_text(
        json.dumps(comparison_rows, indent=2) + "\n"
    )
    write_minimal_xlsx(WORKBOOK_PATH, workbook_rows)

    action_counts: dict[str, int] = {}
    status_counts: dict[str, int] = {}
    for row in comparison_rows:
        action_counts[row["proposed_action"]] = action_counts.get(row["proposed_action"], 0) + 1
        status_counts[row["oneleet_status"]] = status_counts.get(row["oneleet_status"], 0) + 1

    print(
        json.dumps(
            {
                "sprinto_count": len(sprinto),
                "oneleet_active_needs_assessment_count": len(active),
                "oneleet_archived_count": len(archived),
                "oneleet_active_plus_archived_count": len(active) + len(archived),
                "oneleet_library_count": len(library),
                "comparison_rows": len(comparison_rows),
                "action_counts": action_counts,
                "matched_status_counts": status_counts,
                "workbook": str(WORKBOOK_PATH),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    build()
