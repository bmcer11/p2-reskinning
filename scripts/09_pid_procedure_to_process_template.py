#!/usr/bin/env python3
"""Reskin 'Public Interest Disclosure Procedure 2026' (Draft v2.0, PDF source)
into the CoPP Process template (v2).

Source is a PDF, so the body is transcribed verbatim from a PyMuPDF dump into
the block list below and rebuilt on the template's styles (see _pdf_reskin_lib).

Decisions specific to this source:
- The source's Document Governance table (14 fields) is mapped onto the
  template's 7-field table by matching label; fields the template has no row for
  are listed in the delivery notes rather than kept (legislation and associated
  instruments already appear in section 14).
- 'Responsible department: Governance and Performance' is carried literally;
  Responsible division is left empty rather than guessed.
- Document History keeps the source's 'TBC' approval date for v2.0: the source
  is a draft and the date is not invented.
- Headings 4.1-14.2 sit on the template's numbered Heading 2; the bold run-in
  labels (Oral / Written / Anonymous disclosures) stay bold body text so no
  numbering is added that the source does not have.
- The blue 'Examples of ...' boxes become callouts in the template orange.
- 'See Attachment 1' points at an attachment the source does not contain; the
  text is kept, unlinked, and flagged in the delivery notes.
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _pdf_reskin_lib import Builder, front_matter

WORK = os.environ.get('RESKIN_WORK', os.getcwd())
TPL = os.path.join(WORK, 'tpl')
BUILD = os.path.join(WORK, 'build9')
OUT = os.path.join(WORK, 'out9.docx')

TITLE = 'Public Interest Disclosure Procedure 2026'
VERSION = '2.0'
FOOTNOTES = {}
LEG = 'https://www.legislation.vic.gov.au/in-force/'
ROLES_BM = '_PIDRolesAndResponsibilities'

WHERE_TO_DISCLOSE = {
    'cols': [3000, 6026],
    'header': ['Subject of the disclosure', 'Who can receive the disclosure'],
    'rows': [
        ['**Council, or an employee, officer or contractor of Council**',
         ['Council’s CEO', 'Public Interest Disclosures Coordinator', 'A Public Interest Disclosures Officer',
          'The discloser’s manager or supervisor',
          'The manager or supervisor of the person the disclosure is about',
          'IBAC', 'The Victorian Ombudsman', 'The Victorian Inspectorate']],
        ['**A Councillor**',
         ['IBAC', 'The Victorian Ombudsman', '//(Disclosures about Councillors cannot be made to Council)//']],
    ]}

IBAC_OMBUDSMAN = {
    'cols': [4513, 4513],
    'header': ['IBAC', 'Victorian Ombudsman'],
    'rows': [
        [['//Contact IBAC to make a disclosure about any public body, public officer or Councillor.//',
          '**//All disclosures about Councillors must be made directly to IBAC.//**'],
         ['//Contact the Ombudsman to make a disclosure about a Council officer or Council administration.//',
          '//The Ombudsman can receive disclosures about most Victorian public sector organisations.//']],
        [['Level 1, North Tower\n459 Collins Street\nMelbourne VIC 3000',
          'Phone: 1300 735 135\n[www.ibac.vic.gov.au](http://www.ibac.vic.gov.au/)'],
         ['Level 2, 570 Bourke Street\nMelbourne VIC 3000',
          'Phone: 9613 6222\n[www.ombudsman.vic.gov.au](http://www.ombudsman.vic.gov.au/)']],
    ]}

COORDINATOR = {
    'cols': [4513, 4513],
    'header': ['Name', 'Contact Details'],
    'rows': [
        [['**James Gullan**', 'Manager Communications and Governance'],
         ['M: 0422 060 667', '[james.gullan@portphillip.vic.gov.au](mailto:james.gullan@portphillip.vic.gov.au)']],
    ]}

OFFICERS = {
    'cols': [4513, 4513],
    'header': ['Name', 'Contact Details'],
    'rows': [
        [['**James Gullan**', 'Head of Governance'],
         ['M: 0422 060 667', '[james.gullan@portphillip.vic.gov.au](mailto:james.gullan@portphillip.vic.gov.au)']],
        [['**Alli Griffin**', 'Senior Privacy and FOI Advisor'],
         ['M: 0403 242 260', '[alli.griffin@portphillip.vic.gov.au](mailto:alli.griffin@portphillip.vic.gov.au)']],
        [['**Robyn Borley**', 'Director Governance and Performance'],
         ['M: 0434 911 528', '[robyn.borley@portphillip.vic.gov.au](mailto:robyn.borley@portphillip.vic.gov.au)']],
    ]}

TWO_STEP = {
    'cols': [4513, 4513],
    'header': ['Step 1: Does the information show or tend to show improper conduct or detrimental action?',
               'Step 2: Does the discloser have a reasonable belief?'],
    'rows': [
        [[('b', ['Does the information satisfy the elements of improper conduct or detrimental action as defined in the PID Act?',
                 'Is there sufficient detail and supporting material?',
                 'What is the discloser’s connection to the alleged conduct?',
                 'How did the discloser come to know about the conduct?',
                 'Council may need to seek further information or conduct a discreet initial enquiry.'])],
         [('b', ['Does the discloser actually believe the information shows improper conduct or detrimental action?',
                 'Is the belief based on facts sufficient to make a reasonable person believe there was improper conduct?',
                 'The belief must be more than a mere suspicion that is unsupported by further information, factors or circumstances— it must be probable.',
                 'Simply stating “I know XYZ is corrupt” without supporting facts is not sufficient.',
                 'Consider the credibility and reliability of the information, even if second-hand.'])]],
    ]}

OUTCOME = {
    'cols': [4513, 4513],
    'header': ['If Council considers it may be a public interest disclosure',
               'If Council does not consider it a public interest disclosure'],
    'rows': [
        [[('b', ['Notify IBAC that Council considers the disclosure may be a public interest disclosure and is sending it for assessment under section 21 of the Act.',
                 'Notify the discloser that the disclosure has been sent to IBAC for assessment.',
                 'Provide IBAC with any information obtained during Council’s enquiries.'])],
         [('b', ['Notify the discloser (if they indicated they wish to receive protections) that Council does not consider the disclosure a PID and it has not been sent to IBAC.',
                 "Advise the discloser that the discloser's identity doesn’t have to be kept secret, but protections under Part 6 of the PID Act still apply regardless.",
                 'Consider whether the matter can be dealt with under Council’s normal complaint handling procedures.'])]],
    ]}

PROCESS_OVERVIEW = {
    'cols': [900, 5126, 3000],
    'header': ['Step', 'Action', 'Responsibility'],
    'keep_together': True,
    'rows': [
        ['1', 'Disclosure is received by Council (via person authorised to receive a public interest disclosure)', 'Receiving person'],
        ['2', 'Disclosure is forwarded to the Public Interest Disclosures Coordinator', 'Receiving person'],
        ['3', 'Coordinator assesses whether the disclosure may be a public interest disclosure', 'PID Coordinator'],
        ['4', 'If it may be a public interest disclosure: notify IBAC and the discloser within 28 days (unless discloser has confirmed in writing that the disclosure is not a public interest disclosure). If not: notify discloser within 28 days and consider other processes.', 'PID Coordinator'],
        ['5', 'IBAC assesses whether the disclosure is a public interest complaint', 'IBAC'],
        ['6', 'IBAC advises Council and the discloser of its determination', 'IBAC'],
        ['7', 'If a public interest complaint: IBAC investigates, refers or dismisses. Discloser is advised of outcome.', 'IBAC / Investigating entity'],
    ]}

DEFINITIONS = {
    'cols': [2600, 6426],
    'header': ['Term', 'Definition'],
    'rows': [[f'**{t}**', d] for t, d in [
        ('PID Act', 'Public Interest Disclosures Act 2012 (Vic)'),
        ('Assessable disclosure', 'A disclosure either made directly to IBAC or an appropriate entity, or received by Council and required under the Act to be notified to IBAC for assessment.'),
        ('Co-operator', 'A person who has cooperated or intends to cooperate with an investigation of a public interest complaint.'),
        ('CEO', 'Chief Executive Officer of Port Phillip City Council.'),
        ('Corrupt conduct', 'Conduct involving dishonest performance of functions, knowingly breaching public trust, misusing information acquired in the course of public duties, or adversely affecting the honest performance of public functions, being conduct that would constitute a relevant offence.'),
        ('Council', 'Port Phillip City Council.'),
        ('Discloser', 'A person who makes a disclosure under the Act.'),
        ('FOI Act', 'Freedom of Information Act 1982 (Vic).'),
        ('Guidelines', 'IBAC Guidelines for Handling Public Interest Disclosures (June 2025).'),
        ('IBAC', 'The Independent Broad-based Anti-corruption Commission.'),
        ('IBAC Act', 'Independent Broad-based Anti-corruption Commission Act 2011 (Vic).'),
        ('In private', 'Circumstances where the only persons present or able to listen are the discloser, the person receiving the disclosure, and any legal practitioner representing the discloser.'),
        ('Investigative entity', 'A body authorised to investigate a public interest complaint: IBAC, Victorian Ombudsman, Victoria Police, Victorian Inspectorate, Judicial Commission, Chief Municipal Inspector, Racing Integrity Commissioner, or Information Commissioner.'),
        ('Misdirected disclosure', 'A disclosure made to a body that is not the correct receiving entity, where the discloser honestly believed that body was the appropriate recipient.'),
        ('Public interest complaint', 'A public interest disclosure that has been determined by IBAC to be a public interest complaint under section 26 of the Act.'),
        ('Public interest disclosure', 'A disclosure made in accordance with Part 2 of the Act.'),
        ('Public body', 'A body specified in section 6 of the Act, including Council.'),
        ('Public officer', 'A person specified in section 6 of the Act, including Council employees, contractors and Councillors.'),
        ('Penalty unit', 'A standard measure used in Victorian legislation to express the amount of a fine. The value is indexed annually under the Monetary Units Act 2004 (Vic).'),
        ('PID', 'Public interest disclosure.'),
        ('Regulations', 'Public Interest Disclosures Regulations 2019 (Vic).'),
        ('Welfare Manager', 'A person appointed by the Public Interest Disclosures Coordinator to support and protect a discloser or co-operator.'),
    ]]}

BLOCKS = [
    ('h1', 'Purpose'),
    ('p', 'Port Phillip City Council (**Council**) is committed to the aims and objectives of the Public Interest Disclosures Act 2012 (**PID Act**). Council does not tolerate improper conduct by its employees, officers or Councillors, nor the taking of reprisal action against those who come forward to disclose such conduct.'),
    ('p', 'Council recognises the value of transparency and accountability in its administrative and management practices. It supports the making of disclosures that reveal improper conduct or detrimental action, and will take all reasonable steps to protect people who make such disclosures from any detrimental action in reprisal.'),
    ('p', 'This procedure:'),
    ('alpha', ['establishes how disclosures of improper conduct or detrimental action may be made to Council',
               'sets out how Council will receive, assess and handle disclosures',
               'describes the protections and welfare support available to disclosers, persons who are the subject of disclosures, and co-operators',
               'explains Council’s confidentiality obligations under the Act',
               'outlines the role of the Independent Broad-based Anti-corruption Commission (**IBAC**) in assessing disclosures.']),
    ('p', 'This procedure has been prepared in accordance with relevant legislation and the IBAC Guidelines for Handling Public Interest Disclosures (June 2025) (**Guidelines**).'),

    ('h1', 'Scope'),
    ('p', 'This procedure applies to all Council employees, officers, contractors and Councillors. It is also available to members of the public who wish to make a disclosure about Council, its employees or officers.'),
    ('p', 'Disclosures about the conduct of Councillors cannot be made to Council. They must be made directly to IBAC or the Victorian Ombudsman.'),

    ('h1', 'About the PID Act'),
    ('p', 'The PID Act commenced operation on 10 February 2013 (formerly known as the Protected Disclosure Act 2012). The PID Act provides a framework to encourage and assist people to report improper conduct and detrimental action by public officers and public bodies.'),
    ('p', 'The PID Act aims to:'),
    ('b', ['encourage and assist the reporting of improper conduct and detrimental action taken in reprisal for a public interest disclosure',
           'provide certain protections for people who make a disclosure, or those who may suffer detrimental action in reprisal',
           'ensure that certain information about a disclosure is kept confidential, including the identity of the discloser and the content of the disclosure in certain circumstances.']),
    ('p', 'Council is required to establish and publish these procedures under section 58 of the PID Act and in accordance with the Guidelines issued by IBAC under section 57 of the Act.'),

    ('h1', 'What is a Public Interest Disclosure?'),
    ('p', 'A public interest disclosure is a report made by a person about information that shows or tends to show, or information that the person making the report reasonably believes shows to tends to show:'),
    ('b', ['**improper conduct** (within the meaning of the PID Act) by a public body or public officer (such as corrupt conduct, misuse of public resources, or risks to health, safety or the environment); and/or',
           '**detrimental action** (within the meaning of the PID Act) taken by a public body or public officer against a person in reprisal for them (or another person) having made a public interest disclosure or cooperated with an investigation.']),
    ('p', 'A disclosure can relate to conduct or action that has already taken place, is currently occurring, or may happen in the future. Disclosures can also be made about conduct that occurred before the Act commenced.'),
    ('h2', 'Improper conduct'),
    ('p', 'Improper conduct involves conduct by a public officer or public body in their official capacity. It includes:'),
    ('b', ['corrupt conduct', 'a criminal offence', 'serious professional misconduct',
           'dishonest performance of public functions', 'an intentional or reckless breach of public trust',
           'an intentional or reckless misuse of information', 'a substantial mismanagement of public resources',
           'a substantial risk to the health or safety of one or more persons', 'a substantial risk to the environment.']),
    ('p', 'Improper conduct also includes conduct by any person that adversely affects the honest performance of a public officer’s functions, or that is intended to adversely affect those functions and results in improper gain (such as obtaining a license, appointment, financial benefit or other advantage). Trivial matters do not meet the threshold of improper conduct.'),
    ('callout', 'Examples of improper conduct',
     ['An environmental health officer ignores or conceals evidence of illegal dumping of waste to avoid closure of a local business.',
      'A building inspector tolerates poor practices and structural defects in the work of a local builder.',
      'A Council officer uses their position to access confidential information for personal benefit.',
      'A Council officer substantially mismanages a procurement process, resulting in significant financial loss to Council.']),
    ('h2', 'Detrimental action'),
    ('p', 'It is an offence under the PID Act for a person to take detrimental action against another person in reprisal for a public interest disclosure. Detrimental action includes:'),
    ('b', ['action causing injury, loss or damage', 'intimidation or harassment',
           'discrimination, disadvantage or adverse treatment in relation to a person’s employment, career, profession, trade or business, including the taking of disciplinary action.']),
    ('p', 'A person does not need to have actually taken the detrimental action — threatening to do so, or inciting or permitting someone else to do so, is also covered. The detrimental action does not need to be directed at the discloser; it can be taken against anyone connected with a public interest disclosure.'),
    ('callout', 'Examples of detrimental action',
     ['Council refuses the promotion of a person who made a disclosure, in reprisal for them making the disclosure.',
      'Council demotes, transfers, isolates or changes the duties of a discloser, in reprisal for them making the disclosure.',
      'A person threatens, abuses or harasses the discloser, their family or friends in connection with the disclosure.',
      'Council discriminates against the discloser or their associates in subsequent applications for jobs, permits or tenders, in reprisal for the disclosure.']),
    ('h2', 'What is not a public interest disclosure?'),
    ('p', 'The following __are not__ public interest disclosures under the PID Act:'),
    ('b', ['a disclosure where the discloser expressly states in writing, at the time of making it or within 28 days after it is made, that it is not a disclosure under the PID Act',
           'a disclosure made by an officer or employee of an investigative entity in the course of carrying out their duties, unless they expressly state in writing that it is a disclosure under the Act',
           'matters that do not involve potential improper conduct or detrimental action as defined in the PID Act',
           'matters that are trivial.']),

    ('h1', 'Making a Disclosure'),
    ('h2', 'Who can make a disclosure?'),
    ('p', 'Anyone can make a public interest disclosure, including Council employees, officers, contractors, Councillors and members of the public. Disclosures can be made by individuals or jointly by a group of people. A company or business cannot make a disclosure, but its officers or employees can.'),
    ('p', "A person making a disclosure does not need to identify themselves (disclosures may be made anonymously) or identify the specific person or body the disclosure is about. You also do not need to refer to the PID Act, or refer to the disclosure as a 'public interest disclosure', for it to be treated as a public interest disclosure."),
    ('h2', 'Where should disclosures be made?'),
    ('p', 'Disclosures must be made to an entity authorised to receive them. The table below summarises who can receive different types of disclosures:'),
    ('table', WHERE_TO_DISCLOSE),
    ('p', 'As a general rule, if your disclosure is made to a person or entity that cannot receive a disclosure, your disclosure will **not** be a public interest disclosure and you will **not** be protected under the PID Act.'),
    ('h2', 'How to make a disclosure to Council'),
    ('p', 'When making a disclosure it is helpful to provide:'),
    ('b', ['a description of the alleged improper conduct or detrimental action',
           'who was involved, and where and when the conduct occurred, is occurring or may occur',
           'your grounds for believing the conduct occurred, is occurring or may occur',
           'any supporting documentation.']),
    ('pk', '**Oral disclosures**'),
    ('p', 'An oral disclosure to Council must be made in private and may be made:'),
    ('b', ['in person',
           'by telephone to one of the persons authorised to receive disclosures (including by voicemail)',
           'by some other form of non-written electronic communication (e.g. a voice message sent via WhatsApp, Teams, or similar apps or a video message or call)']),
    ('p', '“In private” means the person making the disclosure reasonably believes the only people present or able to hear are themselves, the person receiving the disclosure, and any legal practitioner representing them.'),
    ('p', 'If a disclosure is made orally, the person receiving it will make notes at the time. Recording of the conversation will only occur with the discloser’s permission or after giving prior warning.'),
    ('pk', '**Written disclosures**'),
    ('p', 'A written disclosure to Council must be:'),
    ('b', ['delivered personally to a person authorised to receive the disclosure, at the Council office',
           'sent by post addressed to the Public Interest Disclosures Coordinator (or the name of any other person authorised to receive the disclosure), 99a Carlisle Street, St Kilda VIC 3182',
           f'sent by email to the Public Interest Disclosures Coordinator or a Public Interest Disclosures Officer (see [Section 6 for contact details](#{ROLES_BM})).']),
    ('p', 'Council recommends that any written disclosure delivered in person or by post be sealed in an envelope clearly marked: “CONFIDENTIAL – For addressee eyes only - Attention Public Interest Disclosures Coordinator”.'),
    ('p', 'Disclosures cannot be made by fax.'),
    ('pk', '**Anonymous disclosures**'),
    ('p', 'Disclosures can be made anonymously. This can be done by:'),
    ('b', ['using an unverifiable email address', 'making an anonymous phone call',
           'attending a face-to-face meeting and refusing to identify yourself, provided the meeting takes place in private.']),
    ('p', 'However, if you make an anonymous disclosure and provide us with no means of contacting you, then we may find it difficult to determine whether your complaint fits the definition of a disclosure and we will not be able to communicate with you about your disclosure.'),
    ('h2', 'How to make a disclosure to IBAC or other bodies'),
    ('p', 'Disclosures about Council, its employees and contractors (and all disclosures about Councillors) may also be made directly to:'),
    ('table', IBAC_OMBUDSMAN),
    ('h2', 'Misdirected disclosures'),
    ('p', 'If Council receives a disclosure that should have been made to another body, and Council is satisfied the discloser honestly believed Council was the appropriate receiving entity, Council may decide to treat it as a misdirected disclosure and notify IBAC for assessment. Otherwise, the discloser will be directed to the correct receiving body.'),
    ('p', 'A public interest disclosure that relates to a Member of Parliament cannot be treated as a misdirected disclosure.'),

    ('h1', 'Roles and Responsibilities', ROLES_BM),
    ('h2', 'Employees, officers and Councillors'),
    ('p', 'All employees, officers and Councillors are encouraged to report known or suspected incidents of improper conduct or detrimental action in accordance with this procedure. They have an important role in supporting those who have made a disclosure and must refrain from any activity that is, or could be perceived to be, victimisation or harassment of a discloser. They must also protect and maintain the confidentiality of anyone they know or suspect has made a disclosure, unless an exception applies under legislation.'),
    ('h2', 'Public Interest Disclosures Coordinator'),
    ('p', 'The Public Interest Disclosures Coordinator has a central role in Council’s internal reporting system and will:'),
    ('b', ['receive all disclosures (including from Public Interest Disclosures Officers)',
           'be the contact point for general advice about the Act and for integrity agencies such as IBAC',
           'impartially assess each disclosure to determine whether it should be notified to IBAC',
           'refer all public interest disclosures to IBAC for assessment',
           'keep the discloser informed of progress',
           'appoint a Welfare Manager where appropriate',
           'establish and manage a confidential records management system',
           'collate and publish statistics on disclosures made',
           'take all necessary steps to protect the identity of disclosers and persons subject to disclosures',
           'liaise with the Chief Executive Officer (CEO).']),
    ('p', 'The Public Interest Disclosures Coordinator is currently:'),
    ('table', COORDINATOR),
    ('h2', 'Public Interest Disclosures Officers'),
    ('p', 'Council has appointed Public Interest Disclosures Officers to receive disclosures and provide general advice about the operation of the PID Act. They will:'),
    ('b', ['make arrangements for a disclosure to be made privately and discreetly',
           'receive disclosures made orally or in writing and commit oral disclosures to writing',
           'forward all disclosures and supporting evidence to the Public Interest Disclosures Coordinator for further assessment and management',
           'take all necessary steps to keep the information and identities confidential',
           'offer to remain a support person for the discloser.']),
    ('p', 'The Public Interest Disclosures Officers are:'),
    ('table', OFFICERS),
    ('h2', 'Managers and supervisors'),
    ('p', 'A manager or supervisor who receives a disclosure will:'),
    ('b', ['immediately bring the matter to the attention of the Public Interest Disclosures Coordinator',
           'commit any oral disclosure to writing',
           'take all necessary steps to keep the information and identities secure, private and confidential.']),
    ('h2', 'Reception and customer service staff'),
    ('p', 'Staff who receive telephone calls, mail or other correspondence that may be a disclosure must not enquire into the circumstances and must immediately refer the matter to the Public Interest Disclosures Coordinator. No details of the disclosure should be recorded in Council’s electronic document management system.'),

    ('h1', 'Handling and Assessing Disclosures'),
    ('h2', 'Receiving a disclosure'),
    ('p', 'When Council receives a complaint, report or allegation of improper conduct or detrimental action, the first step is to determine whether it is a disclosure that Council can properly receive — that is, whether it relates to the conduct of Council or an officer or employee of Council. If it relates to a Councillor, the discloser must be directed to IBAC or the Victorian Ombudsman.'),
    ('h2', 'Assessing a disclosure'),
    ('p', 'The Public Interest Disclosures Coordinator will assess whether the disclosure may be a public interest disclosure using the following two-step process:'),
    ('table', TWO_STEP),
    ('h2', 'Assessment outcome and notification'),
    ('p', 'Council must complete its assessment and provide notification to IBAC and/or the discloser within 28 days of receiving the disclosure.'),
    ('table', OUTCOME),
    ('h2', 'Urgent action'),
    ('p', 'In some circumstances, a disclosure may be about conduct that poses an immediate threat to health and safety, preservation of property, or serious criminal conduct. In these cases, Council can take immediate action while completing its assessment or awaiting IBAC’s decision.'),
    ('p', 'The Act allows Council to disclose the content of the disclosure to the extent necessary for taking lawful action, including disciplinary processes. However, this does not allow the identity of the discloser to be revealed. It may also be appropriate to report criminal conduct to Victoria Police.'),

    ('h1', 'Assessment by IBAC'),
    ('p', 'Once Council notifies IBAC of a disclosure, IBAC must determine whether it is a public interest complaint. IBAC will inform Council and the discloser of its determination in writing and within a reasonable time.'),
    ('p', 'In making its assessment, IBAC may seek additional information from Council or the discloser.'),
    ('h2', 'If IBAC determines the disclosure is a public interest complaint'),
    ('p', 'IBAC will advise the discloser (provided they are not anonymous) in writing that:'),
    ('b', ['the disclosure has been determined to be a public interest complaint',
           'the protections under Part 6 of the Act apply',
           'the discloser has rights, protections and obligations under the Act.']),
    ('p', 'Once determined to be a public interest complaint, the discloser cannot withdraw it. However, IBAC can decide not to investigate if the discloser requests this.'),
    ('p', 'IBAC may then investigate the complaint, refer it to another investigative entity, conduct preliminary inquiries, or dismiss it on grounds set out in the IBAC Act.'),
    ('h2', 'If IBAC determines it is not a public interest complaint'),
    ('p', 'IBAC will advise the discloser in writing that:'),
    ('b', ['the disclosure is not a public interest complaint',
           'the disclosure will not be investigated as a public interest complaint',
           'the confidentiality provisions under Part 7 of the Act no longer apply',
           'the protections under Part 6 of the Act still apply.']),
    ('p', 'IBAC may also advise Council that the matter could be dealt with through other processes, could refer the matter for investigation including by Council, or treat it as a notification under the IBAC Act.'),
    ('h2', 'What happens during an investigation?'),
    ('p', 'If IBAC or another investigative entity investigates a public interest complaint, it may contact Council for information. Council will be able to share information about the complaint with the investigative entity without breaching confidentiality requirements. It may also contact the discloser.'),
    ('p', 'At the conclusion of the investigation, the investigative entity must generally provide the discloser with information about the results.'),
    ('h2', 'Overview of the disclosure process'),
    ('table', PROCESS_OVERVIEW),

    ('h1', 'Protections'),
    ('h2', 'Protections available to disclosers'),
    ('p', 'Part 6 of the PID Act provides protections to anyone who makes a public interest disclosure or a misdirected disclosure. These protections apply from the time the disclosure is made, even if Council does not notify IBAC or IBAC determines it is not a public interest complaint. The protections include:'),
    ('b', ['the discloser is not subject to any civil or criminal liability for making the disclosure',
           'the discloser is not subject to any administrative or disciplinary action for making the disclosure',
           'the discloser is not committing an offence against any law that imposes confidentiality obligations',
           'the discloser is not breaching any other obligation requiring them to maintain confidentiality',
           'the discloser cannot be held liable for defamation in relation to the information disclosed.']),
    ('h2', 'Protections for public officers handling disclosures'),
    ('p', 'A public officer (including a Council officer or employee) who acts in good faith and in accordance with the PID Act, Regulations and IBAC Guidelines does not commit an offence under laws imposing a duty to maintain confidentiality.'),
    ('h2', 'Limitations on protections'),
    ('p', 'Protections do not apply where a discloser:'),
    ('b', ['knowingly provides false or misleading information, intending it to be acted on as a public interest disclosure (maximum penalty: 120 penalty units or 12 months’ imprisonment, or both)',
           'falsely claims a matter is the subject of a public interest disclosure or public interest complaint (maximum penalty: 120 penalty units or 12 months’ imprisonment, or both).']),
    ('p', 'A discloser is not protected against legitimate management action taken by Council, including performance development, conditions of employment, discipline or workplace safety measures, provided the action is not connected to the making of the disclosure.'),
    ('p', 'The PID Act also provides that a person remains liable for their own conduct, even if they disclose that conduct as part of making a public interest disclosure.'),
    ('h2', 'Where a discloser is implicated in misconduct'),
    ('p', 'Where a discloser is implicated in improper conduct, Council will handle the disclosure and protect them from reprisals in accordance with the Act. However, making a disclosure does not shield a person from the reasonable consequences of their own involvement in misconduct.'),
    ('p', 'The CEO will make the final decision, on the advice of the Public Interest Disclosures Coordinator, about whether disciplinary or other action will be taken against a discloser. Where such action relates to the subject matter of the disclosure, it will only be taken after the disclosed matter has been appropriately dealt with.'),
    ('p', 'In all cases, Council will document its decision-making process to demonstrate the action was taken for appropriate and permitted reasons, not in retribution for the disclosure.'),

    ('h1', 'Confidentiality'),
    ('h2', 'General obligation'),
    ('p', 'Part 7 of the PID Act makes it a criminal offence to disclose the content of, or information likely to identify the person who made, an assessable disclosure, where Part 7 applies. Penalties include a maximum fine of 120 penalty units or 12 months’ imprisonment (or both) for an individual, and 600 penalty units for a body corporate (See Attachment 1 for more information on penalty units).'),
    ('p', 'Council will take all reasonable steps to protect the identity of the discloser and the content of the disclosure, where required to do so under the PID Act.'),
    ('h2', 'Steps Council will take'),
    ('p', 'Council will ensure confidentiality by:'),
    ('b', ['storing all files (paper and electronic) securely, accessible only by the Public Interest Disclosures Coordinator, Public Interest Disclosures Officers involved in the matter, and Welfare Managers for welfare matters',
           'marking all printed material as Public Interest Disclosures Act matters with warnings about criminal penalties for unauthorised access',
           'password-protecting all electronic files',
           'conducting all telephone calls and meetings about disclosures in private',
           'not using unsecured email to transmit disclosure-related documents',
           'not delivering hard copy documents by internal mail to generally accessible areas.']),
    ('h2', 'Exceptions to confidentiality'),
    ('p', 'Limited exceptions to the prohibition on disclosure are permitted by the Act, including where:'),
    ('b', ['disclosure is required for Council to exercise its functions under the Act',
           'disclosure is by an investigating entity for the purpose of its functions',
           'disclosure is in accordance with a direction from the investigating entity',
           'disclosure is necessary for taking lawful action (including disciplinary processes) in relation to the conduct that is the subject of the disclosure',
           'IBAC or the Victorian Inspectorate has determined the disclosure is not a public interest disclosure',
           'disclosure is for the purpose of obtaining legal advice',
           'disclosure is to assist the discloser to obtain support from a registered health practitioner, trade union, employee assistance program, the Victorian WorkCover Authority, or the Fair Work Commission',
           'disclosure is necessary for a discloser under 18 (to a parent or guardian), a discloser with insufficient English (to an interpreter), or a discloser with a disability (to an independent person)',
           'the disclosure is assessed by IBAC to not be a public interest complaint.']),
    ('h2', 'Freedom of Information'),
    ('p', 'The Act provides that certain information related to public interest disclosures is exempt from the Freedom of Information Act 1982 (FOI Act). This includes information relating to a disclosure or assessable disclosure, information notified to IBAC, and information likely to lead to the identification of a discloser. Council will contact IBAC before providing any IBAC-originating or PID-related document sought under FOI.'),

    ('h1', 'Welfare management'),
    ('h2', 'Support for disclosers and co-operators'),
    ('p', 'Council is committed to protecting genuine disclosers and co-operators (persons who cooperate or intend to cooperate with an investigation) from detrimental action. Support will be provided regardless of whether the person is an employee or a member of the public, and includes:'),
    ('b', ['keeping the person informed of progress, protections available and any action proposed or taken',
           'providing active support, acknowledging they have done the right thing',
           'managing expectations through early discussion about realistic outcomes',
           'maintaining confidentiality so others cannot infer the person’s identity',
           'proactively assessing and managing the risk of detrimental action',
           'examining immediate welfare and protection needs and fostering a supportive work environment',
           'preventing the spread of gossip and rumours about any investigation',
           'keeping contemporaneous records of all contact and follow-up action.']),
    ('h2', 'Welfare Manager'),
    ('p', 'The Public Interest Disclosures Coordinator may appoint a Welfare Manager to support a discloser or co-operator. The Welfare Manager will monitor their specific needs and provide practical advice and support. The Welfare Manager may be a person from within Council or a third party engaged for that purpose.'),
    ('h2', 'Support for persons who are the subject of a disclosure'),
    ('p', 'Council will also look after the welfare of any person who is the subject of a public interest disclosure. Until a public interest complaint is resolved, the information about the person is only an allegation. Council will:'),
    ('b', ['take all reasonable steps to protect the confidentiality of the person',
           'offer a Welfare Manager or referral to the Employee Assistance Program',
           'provide support and advice regarding their rights and obligations under the Act',
           'maintain confidentiality even where a disclosure is dismissed or not substantiated.']),
    ('p', 'Where allegations are not substantiated, Council will provide full support to the person and, where appropriate, the CEO may consider issuing a statement of support.'),
    ('h2', 'If detrimental action is reported'),
    ('p', 'If any person reports an incident that may amount to detrimental action taken in reprisal for a disclosure, the Welfare Manager or Public Interest Disclosures Coordinator will record the details and advise the person of their rights under the Act. If the person wishes to disclose the incident as detrimental action, it will be treated as a new disclosure and assessed accordingly.'),
    ('p', 'All persons are reminded it is a criminal offence to take detrimental action against another person in reprisal for a public interest disclosure. The maximum penalty is 240 penalty units or 2 years’ imprisonment or both.'),

    ('h1', 'Collating and publishing statistics'),
    ('p', 'Council is required to publish certain information about the Act in its annual report, including how these procedures may be accessed and the number of disclosures notified to IBAC during the financial year.'),
    ('p', 'The Public Interest Disclosures Coordinator will maintain a secure register to record this information. The register will be confidential and will not include any information that may identify a discloser.'),

    ('h1', 'Training'),
    ('p', 'Council will:'),
    ('b', ['ensure all employees, officers and Councillors have access to this procedure',
           'incorporate PID awareness into its training program, including obligations under the Act and relevant codes of conduct',
           'provide additional training to the Public Interest Disclosures Coordinator, Public Interest Disclosures Officers, Welfare Managers and complaint handling staff',
           'train staff with FOI responsibilities to ensure no prohibited information is disclosed',
           'train customer-facing staff to recognise and appropriately handle potential disclosures from external sources.']),

    ('h1', 'Related legislation and documents'),
    ('h2', 'Legislation'),
    ('b', [f'[Public Interest Disclosures Act 2012 (Vic)]({LEG}acts/public-interest-disclosures-act-2012/030)',
           f'[Public Interest Disclosures Regulations 2019 (Vic)]({LEG}statutory-rules/public-interest-disclosures-regulations-2019)',
           '[IBAC Guidelines for Handling Public Interest Disclosures (June 2025)](https://www.ibac.vic.gov.au/media/1104/download)',
           f'[Independent Broad-based Anti-corruption Commission Act 2011 (Vic)]({LEG}acts/independent-broad-based-anti-corruption-commission-act-2011)',
           f'[Local Government Act 2020 (Vic)]({LEG}acts/local-government-act-2020)',
           f'[Freedom of Information Act 1982 (Vic)]({LEG}acts/freedom-information-act-1982)',
           f'[Charter of Human Rights and Responsibilities Act 2006 (Vic)]({LEG}acts/charter-human-rights-and-responsibilities-act-2006/015)',
           f'[Gender Equality Act 2020 (Vic)]({LEG}acts/gender-equality-act-2020/004)',
           f'[Constitution Act 1975 (Vic)]({LEG}acts/constitution-act-1975)',
           f'[Monetary Units Act 2004 (Vic)]({LEG}acts/monetary-units-act-2004/006)']),
    ('h2', 'Supplementary documents'),
    ('b', ['Employee Code of Conduct', 'Councillor Code of Conduct', 'Managing Complaints Policy',
           'Fraud and Corruption Awareness and Prevention Policy']),

    ('h1', 'Child Safe'),
    ('p', 'The City of Port Phillip is a Child Safe Organisation and has a legal and moral responsibility to understand and activate their role in preventing, detecting, responding and reporting any Child Safety concerns. Council has zero tolerance for child abuse and is actively committed to embedding a culture of safety, wellbeing and inclusion for children and young people.'),
    ('p', 'Consideration has been given to the Child Safe Standards in the development of this procedure.'),

    ('h1', 'Gender Equality'),
    ('p', 'Under the Gender Equality Act 2020, Council has a positive duty to advance gender equality in our organisation and our community. This includes assessing the impacts of Council’s policies on people of different genders, backgrounds and identities, and considering how a policy that directly and significantly impacts the community can be changed to better support people of all genders and promote gender equality.'),
    ('p', 'In the case of this procedure, a gender impact assessment was not required.'),

    ('h1', 'Definitions'),
    ('table', DEFINITIONS),
]

GOVERNANCE = [
    '',                                                   # Responsible division - not stated
    'Governance and Performance',                         # source: Responsible department
    'James Gullan, Manager Communications and Governance',  # source: Responsible officer + s6.2
    'Executive Leadership Team',                          # source: Authorised by
    '09/03/2026',                                         # source: Date of adoption
    '22/04/2030',                                         # source: Full review date
    'Public Interest Disclosures Procedures (Version 1.0, January 2020)',  # source: Supersedes
]
HISTORY = [
    ('2.0', 'TBC', 'Full review',
     'Full review and rewrite. Updated to align with IBAC Guidelines (June 2025). Restructured for accessibility. Updated roles and contact details. Revised definitions table.',
     'Executive Leadership Team'),
    ('1.0', 'January 2020', 'N/A', 'Original version.', ''),
]

if __name__ == '__main__':
    b = Builder(TPL, BUILD, FOOTNOTES)
    body = b.render(BLOCKS)
    pages_path = os.path.join(WORK, 'toc_pages9.json')
    pages = json.load(open(pages_path)) if os.path.exists(pages_path) else {}
    front = front_matter(b, GOVERNANCE, HISTORY, pages)
    b.finish(TITLE, VERSION, front, body, OUT, cover_title_size=int(os.environ.get('COVER_SZ', '72')))
    print('wrote', OUT, os.path.getsize(OUT), 'bytes;', len(b.toc), 'ToC entries')
