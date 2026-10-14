# SkillBridge – Mini Project Report

- `SkillBridge_Project_Report.docx` – editable report (Word)
- `SkillBridge_Project_Report.pdf` – print-ready copy

The layout follows the K.S.R. College of Engineering Design Studio report format.

## Changing your name, register number, guide or date

Edit the `DETAILS` block at the top of `build_report.js`, then run:

```bash
node report/build_report.js
cd report && soffice --headless --convert-to pdf SkillBridge_Project_Report.docx
python3 toc_pages.py SkillBridge_Project_Report.pdf   # refresh TOC page numbers
cd .. && node report/build_report.js && cd report && soffice --headless --convert-to pdf SkillBridge_Project_Report.docx
```

You can also edit the `.docx` directly in Word: use Find & Replace for `STUDENT NAME` and `REGISTER NUMBER`.

`assets/` holds the college logos, the architecture and workflow diagrams, and screenshots of the prototype (`codestu.docx`).
