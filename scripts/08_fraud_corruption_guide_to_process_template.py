#!/usr/bin/env python3
"""Reskin 'Reporting Fraud and Corruption - Confidential Internal Management
Guidelines' v1.0 (PDF source) into the CoPP Process template (v2).

Source is a PDF, so the body is transcribed verbatim from a PyMuPDF dump into
the block list below and rebuilt on the template's styles (see _pdf_reskin_lib).

Decisions specific to this source:
- The source's own cover, contact page and typed Contents page are replaced by
  the template's cover, contact page and a live ToC.
- Section 13 'Document history' moves into the template's Document History table
  at the front; the body therefore ends at 12 Definitions.
- The source has no Document Governance table. Only what the source supports is
  filled: date of approval (30/01/2026), next full refresh by the template's
  stated four-year default, Supersedes N/A (version 1.0, no prior version).
  Division, department, owner and approver are not stated and are left empty.
- The two landscape tables (Complaint Types, Roles) reflow to portrait at 10pt.
- Footnotes 1-3 on the Complaint Types table become real Word footnotes.
- The Process Map is the source's own raster flowchart, re-inserted with alt text.
- A zero-width space after the slash in 'behaviour/performance' lets the narrow
  Key Test column wrap there instead of mid-word.
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pdf_reskin_lib import Builder, front_matter

WORK = os.environ.get('RESKIN_WORK', os.getcwd())
TPL = os.path.join(WORK, 'tpl')
BUILD = os.path.join(WORK, 'build8')
OUT = os.path.join(WORK, 'out8.docx')
FLOW = os.path.join(WORK, 'flow-002.png')

TITLE = 'Reporting Fraud and Corruption – Confidential Internal Management Guidelines'
VERSION = '1.0'

FOOTNOTES = {
    '1': 'See definition in section 4 of the PID Act.',
    '2': 'See definition in section 4 of the IBAC Act.',
    '3': 'See the definition in section 3 of the Act.',
}

AUSTLII_S57 = 'https://classic.austlii.edu.au/au/legis/vic/consol_act/ibaca2011479/s57.html'
LEG = 'https://www.legislation.vic.gov.au/in-force/acts/'
IBAC = 'https://www.ibac.vic.gov.au/'
INTRANET = 'https://intranet.portphillip.vic.gov.au/guides-tasks-and-tools/'

KEY_TEST_IMPROPER = ['Does the information provided show or tend to show there is actual or potential improper conduct?',
                     'OR',
                     'Does the discloser believe on reasonable grounds that actual or potential improper conduct has occurred?']
KEY_TEST_DETRIMENT = ['Does the information provided show or tend to show there is actual or potential detrimental action?',
                      'OR',
                      'Does the discloser believe on reasonable grounds that actual or potential detrimental action has occurred?']

COMPLAINT_TYPES = {
    'cols': [1300, 3400, 1800, 2526],
    'header': ['Complaint Type', 'Definition', 'Key Test', 'Examples of Conduct'],
    'cant_split': False,
    'rows': [
        ['Improper conduct{fn:1}',
         [('b', ['Corrupt conduct;',
                 'conduct that constitutes a criminal offence or serious professional misconduct;',
                 'dishonest performance of public functions;',
                 'an intentional or reckless breach of public trust;',
                 'an intentional or reckless misuse of information or material acquired in the course of the performance of the functions of the public officer or public body;',
                 'a substantial mismanagement of public resources;',
                 'a substantial risk to the health or safety of one or more persons;',
                 'a substantial risk to the environment;',
                 'conduct that adversely affects the honest performance by a public officer or public body of their functions as a public officer or public body;',
                 'conduct that intends to affect the functions or powers of the public officer or body adversely\n'
                 'and results in the person or body obtaining:\n'
                 'i) A licence, permit, approval, authority or other statutory entitlement;\n'
                 'ii) An appointment to a statutory office or a board of a public body;\n'
                 'iii) A financial benefit or real or personal property; or\n'
                 'iv) Any other direct or indirect monetary or proprietary gain;',
                 'conduct that could constitute a conspiracy or an attempted conspiracy.']),
          "The definition excludes matters that are 'trivial'."],
         KEY_TEST_IMPROPER,
         [('b', ['Bribery, kickbacks or gifts offered/accepted to influence decisions',
                 'Steering tenders to favoured suppliers; bid-rigging or collusion',
                 'Awarding contracts or business to friends, family or associates',
                 'Falsifying financial records, timesheets or approvals to circumvent controls'])]],
        ['Corrupt conduct{fn:2}',
         ['Conduct:',
          ('b', ['of any person that adversely affects the honest performance by a public officer or public body of his or her or its functions as a public officer or public body; or',
                 'of a public officer or public body that constitutes or involves the dishonest performance of his or her or its functions as a public officer or public body; or',
                 'of a public officer or public body that constitutes or involves knowingly or recklessly breaching public trust; or',
                 'of a public officer or a public body that involves the misuse of information or material acquired in the course of the performance of his or her or its functions as a public officer or public body, whether or not for the benefit of the public officer or public body or any other person.'])],
         'As above.',
         'As above.'],
        ['Detrimental action{fn:3}',
         ['Action which includes:',
          ('b', ['action causing injury, loss or damage;',
                 'intimidation or harassment;',
                 "discrimination, disadvantage or adverse treatment in relation to a person's employment, career, profession, trade or business, including the taking of disciplinary action,"]),
          'in connection with the making of a PID.'],
         KEY_TEST_DETRIMENT,
         [('b', ['Victimising conduct', 'Bullying conduct', 'Threatening conduct'])]],
        ['Public Interest Disclosure (**PID**)',
         "Previously known as 'whistleblowing,' a report made regarding improper conduct or relevant detrimental action by a public officer or body.",
         'Does the disclosure allege improper conduct (as defined above) or detrimental action (reprisals) linked to the disclosure?',
         [('b', ['Reporting systemic misuse of purchasing cards or deliberate over-billing',
                 'Disclosing that records are being destroyed or altered to mislead investigators',
                 'Alleging managers instructed staff to ignore safety, planning or regulatory obligations',
                 'Detrimental action: victimisation, discrimination, demotion, threats, isolation or roster changes due to raising concerns'])]],
        ['HR Complaint',
         'Workplace behaviour, performance or employment issues that do not meet the thresholds for corrupt or improper conduct.',
         'Is the issue primarily about workplace behaviour/\u200bperformance and not about deceit, misuse of position or personal gain?',
         [('b', ['Interpersonal conflict or incivility not tied to a protected disclosure',
                 'Performance management disputes or feedback style concerns',
                 'Policy breaches: inappropriate language, minor misuse of email or social media',
                 'Attendance, punctuality or supervision issues without evidence of fraud'])]],
    ]}

ROLES = {
    'cols': [1750, 2100, 2976, 2200],
    'header': ['Role', 'Function', 'Tasks', 'Responsible For'],
    'rows': [
        ['CEO (Relevant Principal Officer)',
         'Provides strategic direction on integrity matters; fulfils mandatory reporting obligations under s 57 of the IBAC Act',
         [('b', ['Makes s 57 notifications to IBAC',
                 'Must be informed of all integrity matters*',
                 'Receives and approves final decisions on integrity actions*',
                 'Cannot delegate s 57 reporting obligation'])],
         [('b', ['Section 57 IBAC notifications',
                 'Final approval of corruption responses*',
                 'Strategic oversight of integrity framework'])]],
        ['Chief People Officer',
         'Supports s 57 preparation; leads internal HR investigations, leads corruption investigations (where IBAC has authorised Council to investigate)',
         [('b', ['Supports preparation of s 57 notifications*',
                 'Leads internal investigations (or delegates to relevant GM where subject matter expertise is required) subject to IBAC authorisation for Council to investigate',
                 'Provides preliminary assessments*',
                 "Refers matters brought to the officer's attention that may constitute improper conduct to the Public Interest Disclosure Coordinator"])],
         [('b', ['Internal investigations', 'Section 57 support', 'HR-related misconduct matters'])]],
        ['Manager, Communications and Governance',
         'Manages Public Interest Disclosure assessments and processes (PID Coordinator)',
         [('b', ['Assesses potential PIDs using IBAC PID assessment tool',
                 'Engages disclosers',
                 'Manages PID process noting confidentiality provisions under PID Act',
                 'Refers non-PIDs back to Fraud and Corruption Action Group',
                 'Supports CEO with s 57 notifications'])],
         [('b', ['PID assessments', 'PID management', 'Discloser engagement', 'PID Act compliance',
                 'External liaison (IBAC, Victorian Ombudsman etc)'])]],
        ['Fraud and Corruption Action Group',
         'Conducts initial assessment of allegations; determines pathway and next steps',
         [('b', ['Convenes after PID Coordinator has assessed a matter and determined it is not a PID, to assess matters and determine next steps including whether an incident is s 57 reportable',
                 'Identifies where both S57 and PID processes apply and coordinates concurrent pathways',
                 'Uses IBAC decision-making flowchart',
                 'Refers to PID Coordinator if PID considerations'])],
         [('b', ['Initial assessment', 'Pathway determination', 'Section 57 threshold assessment',
                 'HR complaint triage'])]],
        ['Fraud and Corruption Register Manager',
         'Receives referrals from PID Coordinator (for matters assessed as not a PID); conducts preliminary assessment; maintains fraud and corruption register',
         [('b', ['Refers matters to PID coordinator (after PID ruled out)',
                 'Conducts preliminary assessment',
                 'Convenes Fraud and Corruption Action Group if allegation involves potential corrupt conduct (provided IBAC processes are first exhausted)',
                 'Maintains fraud and corruption register',
                 'Assesses whether HR complaints constitute improper or unethical conduct and refers to PID coordinator where appropriate'])],
         [('b', ['Complaint receipt', 'Preliminary assessment', 'Register maintenance'])]],
        ['Person Receiving Complaint',
         'First point of contact; ensures appropriate referral',
         [('b', ['Receives complaint', 'Refers to PID coordinator for assessment',
                 'Does not investigate or take action'])],
         [('b', ['Appropriate referral of complaints', 'Does not investigate or take action'])]],
    ]}

LINK_COLS = [2600, 4826, 1600]

LEGISLATION = {
    'cols': LINK_COLS, 'jc': {2: 'center'},
    'header': ['Legislation', 'Description', 'Link'],
    'rows': [
        ['Independent Broad-based Anti-corruption Commission Act 2011 (Vic)',
         'Establishes IBAC and sets out obligations for reporting suspected corrupt conduct, including mandatory notification requirements under Section 57',
         f'[Link]({LEG}independent-broad-based-anti-corruption-commission-act-2011)'],
        ['Public Interest Disclosures Act 2012 (Vic)',
         'Provides framework for making and handling public interest disclosures, including protections for disclosers',
         f'[Link]({LEG}public-interest-disclosures-act-2012)'],
        ['Occupational Health and Safety Act 2004 (Vic)',
         'Relevant where complaints involve workplace health and safety matters',
         f'[Link]({LEG}occupational-health-and-safety-act-2004)'],
        ['Privacy and Data Protection Act 2014 (Vic)',
         'Governs handling of personal information in corruption and PID matters',
         f'[Link]({LEG}privacy-and-data-protection-act-2014)'],
    ]}

IBAC_RESOURCES = {
    'cols': LINK_COLS, 'jc': {2: 'center'},
    'header': ['Resource', 'Description', 'Link'],
    'rows': [
        ['Information for principal officers',
         'Overview of mandatory notification obligations for CEOs and department heads',
         f'[Link]({IBAC}information-principal-officers)'],
        ['Directions for making mandatory notifications',
         'Detailed guidance on S57 notification requirements, process and forms',
         f'[Link]({IBAC}media/198/download)'],
        ['Mandatory notification forms', 'Forms for submitting S57 notifications to IBAC',
         f'[Link]({IBAC}mandatory-notifications)'],
        ['What is a public interest disclosure?', 'Overview of PID protections and how to make a disclosure',
         f'[Link]({IBAC}public-interest-disclosure)'],
        ['Public Interest Disclosure process', 'Step-by-step guide to the PID process',
         f'[Link]({IBAC}public-interest-disclosure-process)'],
        ['Guidelines for handling public interest disclosures',
         'Comprehensive guidelines for PID Coordinators and agencies',
         f'[Link]({IBAC}publications-and-resources/article/guidelines-for-making-and-handling-protected-disclosures)'],
        ['Resources for PID Coordinators', 'Training materials, templates and guidance for PID Coordinators',
         f'[Link]({IBAC}resources-public-interest-disclosure-coordinators)'],
        ['Managing an internal investigation into misconduct',
         'Practical guide for conducting internal misconduct investigations',
         f'[Link]({IBAC}investigationguide)'],
        ['Examples of mandatory notifications under Section 57',
         'Case studies showing how IBAC has assessed and responded to notifications',
         f'[Link]({IBAC}publications-and-resources/article/case-studies---examples-of-mandatory-notifications-under-section-57-of-the-independent-broad-based-anti-corruption-commission-act-2011)'],
    ]}

INTRANET_DOCS = {
    'cols': [4500, 4526],
    'header': ['Document', 'Location'],
    'rows': [
        ['**Fraud and Corruption Policy**', f'[Link]({INTRANET}governance-integrity-and-legal/fraud-and-corruption/)'],
        ['**Public Interest Disclosures (Whistle Blowing) Procedures**',
         f'[Link]({INTRANET}governance-integrity-and-legal/public-interest-disclosures-whistle-blowing/)'],
        ['**Integrity Framework**', f'[Link]({INTRANET}governance-integrity-and-legal/integrity-framework/)'],
        ['**Policy Bookcase**', f'[Link]({INTRANET}strategies-policies-plans-frameworks-and-guidelines/policy-bookcase/)'],
    ]}

CONTACT_IBAC = {
    'cols': [1800, 7226],
    'header': ['Contact Method', 'Details'],
    'rows': [
        ['**Phone**', '1300 735 135 (10am – 4pm, Monday – Friday)'],
        ['**Email**', '[info@ibac.vic.gov.au](mailto:info@ibac.vic.gov.au)'],
        ['**Website**', '[www.ibac.vic.gov.au](http://www.ibac.vic.gov.au/)'],
        ['**Address**', 'Level 1, North Tower, 459 Collins Street, Melbourne VIC 3000'],
        ['**Post**', 'GPO Box 24234, Melbourne VIC 3001'],
    ]}

INTERNAL_CONTACTS = {
    'cols': [1950, 1150, 2026, 3900],
    'header': ['Role', 'Name', 'Contact For', 'Email'],
    'rows': [
        ['Governance Inbox', '-', 'General enquiries about the reporting corruption process',
         '[governance@portphillip.vic.gov.au](mailto:governance@portphillip.vic.gov.au)'],
        ['Manager, Communications and Governance', 'James Gullan',
         'Public Interest Disclosures (Whistleblower) matters.',
         '[james.gullan@portphillip.vic.gov.au](mailto:james.gullan@portphillip.vic.gov.au)'],
        ['Fraud and Corruption Register Manager', 'Julie Snowden',
         'Reporting suspected corruption; preliminary assessments; fraud and corruption register',
         '[julie.snowden@portphillip.vic.gov.au](mailto:julie.snowden@portphillip.vic.gov.au)'],
        ['Chief People Officer', 'Daniel Lew', 'HR-related integrity matters; internal investigations',
         '[daniel.lew@portphillip.vic.gov.au](mailto:daniel.lew@portphillip.vic.gov.au)'],
    ]}

DEFINITIONS = {
    'cols': [1600, 4426, 3000],
    'header': ['Term', 'Definition', 'Legislative / Policy Reference'],
    'rows': [
        ['Corruption',
         ['See above.',
          'It encompasses improper actions or inactions by public officers, as well as attempts by private individuals to improperly influence Council decisions.'],
         '//IBAC Act 2011 (Vic)//; Council Fraud & Corruption Policy'],
        ['Fraud',
         'Dishonest activity involving deception that causes actual or potential financial loss. This includes theft of money or property, falsification of documents, and the improper use of information or position for personal benefit.',
         '//AS8001:2021 Fraud & Corruption Control//; Council Fraud & Corruption Policy'],
        ['IBAC',
         'The Independent Broad-based Anti-corruption Commission; Victoria’s independent agency responsible for preventing and exposing public sector corruption and police misconduct.',
         '//IBAC Act 2011 (Vic)//'],
        ['IBAC Act',
         "The Independent Broad-based Anti-corruption Commission Act 2011, which establishes IBAC's powers and sets out the mandatory reporting obligations for public sector bodies.",
         '//IBAC Act 2011 (Vic)//'],
        ['Public Interest Disclosure',
         "Previously known as 'whistleblowing,' a report made under the PID Act 2012 regarding improper conduct by a public officer or body, which provides specific legal protections (confidentiality and immunity) to the discloser.",
         '//Public Interest Disclosures Act 2012 (Vic)//'],
        ['Public Interest Disclosure Action',
         'The formal steps taken by the PID Coordinator to assess, manage, and investigate a disclosure in accordance with the confidentiality and protection provisions of the PID Act.',
         '//Public Interest Disclosures Act 2012 (Vic)//'],
        ['Section 57 Obligations',
         'A mandatory requirement under the IBAC Act 2011 for the CEO (as Principal Officer) to notify IBAC if they suspect, on reasonable grounds, that corrupt conduct has occurred.',
         '//IBAC Act 2011 (Vic), Section 57//'],
    ]}

FLOW_TITLE = 'Reporting Corruption Process Flowchart'
FLOW_DESCR = (
    'Reporting Corruption Process Flowchart, City of Port Phillip. Two pathways. '
    'PID entry pathway: the PID Coordinator receives the complaint and assesses whether it is a PID '
    '(does the information show or tend to show improper conduct or detriment). If it is, the discloser is '
    'asked whether they want it treated as a PID, the PID Coordinator notifies IBAC within 28 days, and IBAC '
    'determines how to handle it: IBAC investigates; or refers it back to Council for an internal investigation '
    'under the PID Act, with confidentiality lifted per IBAC direction, and the discloser is informed of the outcome; '
    'or determines it is not a PID, which moves it to the section 57 question. '
    'Non-PID / S57 pathway: the Chief People Officer or Head of Risk and Assurance receives the complaint '
    '(potential PIDs are referred to the PID Coordinator) and asks whether a section 57 notification is potentially '
    'required. If yes, the CEO is briefed and makes the notification to IBAC, and IBAC responds within 45 days and '
    'may refer the matter back. If no, the matter is referred to the Fraud and Corruption Action Group, which '
    'assesses it using the IBAC decision-making flowchart and determines next steps: an HR assessment under '
    'standard HR processes, or an internal process led by HR or the relevant General Manager. The outcome is '
    'determined, the CEO is informed and makes a decision, and the Fraud and Corruption Register is updated by the '
    'Corruption Register Manager. Confidentiality note: the CEO and other staff must not be informed of a PID unless '
    'and until IBAC refers the matter back, lifting confidentiality obligations.')

BLOCKS = [
    ('h1', 'Purpose'),
    ('p', 'The purpose of this document is to provide clear, practical internal management guidance to specific City of Port Phillip officers who are responsible for, have a substantial role in connection with, or otherwise have a critical need to be aware of, the appropriate processes for reporting and managing suspected fraud, corruption and Public Interest Disclosures (**PID**). In the ordinary course, those stakeholders will be the CEO (Principal Officer), Chief People Officer, Manager, Communication and Governance (PID Coordinator), Head of Risk and Assurance and the Fraud and Corruption Action Group (**Integrity Stakeholders**). It also assists officers in assessing whether a HR complaint constitutes potential or suspected improper or unethical conduct that may require escalation through integrity processes.'),
    ('p', "Different areas of the organisation have responsibilities in receiving, assessing and responding to these matters, and this document brings together the relevant roles, responsibilities and processes into a single reference point. It aims to ensure that all relevant senior officers understand who to contact, what steps to follow, and how matters are assessed and escalated in accordance with Council's obligations under the //Independent Broad-based Anti-corruption Commission Act// 2011 (Vic) (**IBAC Act**) and the //Public Interest Disclosures Act// 2012 (Vic) (**PID Act**)."),
    ('p', 'This document does not replace and is not intended to alter the following critical source documents:'),
    ('b', ['the obligations imposed on City of Port Phillip and its officers under relevant legislation (especially the IBAC Act and the PID Act); or',
           'the //Public Interest Disclosures Procedures of Port Phillip City Council// established under section 58 of the PID Act (as varied from time to time).']),
    ('p', 'This document contains non-binding guidelines for Integrity Stakeholders only, and is not for wider distribution among Council officers.'),

    ('h1', 'Scope'),
    ('p', 'This document provides guidance on reporting and managing integrity matters at the City of Port Phillip. It covers two key streams:'),
    ('b', [f'**Suspected corrupt conduct** requiring mandatory notification to IBAC under [Section 57 of the IBAC Act]({AUSTLII_S57}); and',
           '**Public Interest Disclosures (PIDs)** made under the //Public Interest Disclosures Act 2012 (Vic).//']),
    ('p', '__NOTE:__ In some circumstances, both Section 57 notification and PID processes may run concurrently where the same matter meets thresholds for both pathways.'),
    ('p', 'General HR complaints that do not meet the threshold for potential or suspected corrupt conduct or a public interest disclosure are managed by the People Culture and Safety team through standard HR processes.'),
    ('p', 'Specifically, this guide contains:'),
    ('b', ['**Roles and Responsibilities:** A clear breakdown of accountabilities for integrity stakeholders including the CEO (Principal Officer), Chief People Officer, Manager, Communication and Governance (PID Coordinator), Head of Risk and Assurance and the Fraud and Corruption Action Group.',
           '**Assessment Pathways and Process Maps:** A visual workflow and guidance on distinguishing between general HR grievances, suspected or potential corrupt conduct requiring notification under the IBAC Act and disclosures made under the PID Act.',
           '**Operational Resources:** Frequently asked questions, definitions, and direct links to relevant legislation, IBAC resources, and internal policies.']),
    ('p', 'This document does not replace specific HR policies for standard grievances, nor is it an official PID procedure for the purposes of section 58 of the //Public Interest Disclosures Act 2012 (Vic)//.'),
    ('p', 'This document serves as the primary mechanism for handling matters that meet the threshold for potential or suspected corruption, fraud, or public interest disclosures, ensuring compliance with legislative obligations and Council’s Integrity framework.'),

    ('h1', 'Council’s Integrity Framework'),
    ('p', "This guide operates within Council's broader [Integrity Framework](https://cppcity.sharepoint.com/sites/PnP/SitePages/Integrity-Framework.aspx), which establishes the strategic foundation for how we embed integrity across the organisation."),
    ('p', 'The Framework is built on five pillars: ethical culture, leadership and capability development, governance and accountability, integrated policies and fraud prevention, and risk management. Together, these pillars support Council in honest, transparent decision-making that serves the best interests of our community.'),

    ('h1', 'Different Complaint Types'),
    ('table', COMPLAINT_TYPES),

    ('h1', 'Roles and Responsibilities'),
    ('table', ROLES),
    ('p', '* Subject to PID/public interest complaint confidentiality obligations arising under the PID Act in relation to any matter.'),

    ('h1', 'Process Map'),
    ('img', FLOW, FLOW_TITLE, FLOW_DESCR, 9026),

    ('h1', 'Relevant Legislation'),
    ('table', LEGISLATION),
    ('h1', 'IBAC Resources'),
    ('table', IBAC_RESOURCES),
    ('h1', 'Intranet Documents'),
    ('table', INTRANET_DOCS),
    ('h1', 'Contact IBAC'),
    ('table', CONTACT_IBAC),
    ('h1', 'Internal Contacts'),
    ('table', INTERNAL_CONTACTS),
    ('h1', 'Definitions'),
    ('table', DEFINITIONS),
]

GOVERNANCE = [
    '',                         # Responsible division - not stated in source
    '',                         # Responsible department - not stated in source
    '',                         # Policy owner - not stated in source
    '',                         # Final full refresh approver - not stated in source
    '30/01/2026',               # source: date of approval / adoption 30/1/2026
    '30/01/2030',               # template default: four years from approval
    'N/A',                      # version 1.0, no previous instrument
]
HISTORY = [('1.0', '30/01/2026', 'N/A', 'N/A', '')]

if __name__ == '__main__':
    b = Builder(TPL, BUILD, FOOTNOTES)
    body = b.render(BLOCKS)
    pages_path = os.path.join(WORK, 'toc_pages8.json')
    pages = json.load(open(pages_path)) if os.path.exists(pages_path) else {}
    front = front_matter(b, GOVERNANCE, HISTORY, pages)
    b.finish(TITLE, VERSION, front, body, OUT, cover_title_size=int(os.environ.get('COVER_SZ', '72')))
    print('wrote', OUT, os.path.getsize(OUT), 'bytes;', len(b.toc), 'ToC entries')
