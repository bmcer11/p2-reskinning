# Copilot prompt: reskin an instrument onto a CoPP template

## Steps

1. In Word, double-click the right template (.dotx) to open a new document from it.
   - Policy: **Policy** template
   - Guideline, procedure or process guide: **Process** template
   - Framework: **Framework** template
   - Anything else: **Other** template
2. Save the new document under the instrument's name.
3. Open Copilot in that document and paste the prompt below. Where it says `[SOURCE]`,
   type `/` and pick the source file. If Copilot won't take a PDF, open the PDF in Word
   first (Word converts it), save it, and pick that file instead.
4. Do the checks under **After Copilot**. Copilot does not do these reliably.

## Prompt

```
Reskin a City of Port Phillip governance instrument into this template.
The source is [SOURCE]. This open document is the template.

Rules:
1. Copy the source text word for word. Do not rewrite, shorten, summarise, correct or add anything.
2. Keep the source's own headings in the source's order. Do not use the template's example headings. Use Heading 1 for main sections and Heading 2 for sub-sections. Remove typed numbers like "1." or "4.1" from headings, because the template numbers headings automatically.
3. Delete all template guidance: the instructions table at the start, every line starting with ▶, every "DELETE:" box, every "Copilot assist" box, every <placeholder> in angle brackets, and every template section the source does not use.
4. On the cover, replace the title placeholder (for example "Process Guide Title") with the source title, and "Version X.X" with the source version. Put the same title and version in the header and footer.
5. Fill the Document Governance table only with facts stated in the source. If the source does not state something, leave that cell empty. Never guess a name, date or team.
6. Copy the source's version history into the Document History table. Then delete the source's own history section from the body.
7. Keep every table row, column and cell. Use the template's table style (coloured header row).
8. Use the template's bullets for bullet lists. Keep (a), (b), (c) lists as numbered lists.
9. Keep every hyperlink. Wherever the source has a picture or flowchart, write a line: [INSERT IMAGE: what it shows].
10. Do not copy the source's cover page, contact page or contents page. The template already has them.

When you finish, tell me in the chat (not in the document): which governance cells are empty, anything you could not place, and anything in the source that looks wrong or inconsistent.
```

## After Copilot

- Insert any picture or flowchart where you see `[INSERT IMAGE: ...]`.
- Right-click the Table of Contents, choose **Update Field**, then **Update entire table**.
- Check every page:
  - no red or orange guidance text and no `< >` placeholders are left
  - every table is complete
  - headings are numbered once (never "1. 1. Purpose")
- Compare the last sentence of each section with the source to catch skipped text.
- Check where the instrument may be published. Some guidelines are restricted to named
  officers and must not go on the open intranet.
