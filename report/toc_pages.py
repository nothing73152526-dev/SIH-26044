"""Reads the rendered PDF and writes toc_pages.json with the page number of
every heading, so build_report.js can fill the TABLE OF CONTENT.
Usage: python3 report/toc_pages.py <rendered.pdf>"""
import json, os, re, subprocess, sys

pdf = sys.argv[1]
n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
pages = [subprocess.run(['pdftotext', '-layout', '-f', str(i), '-l', str(i), pdf, '-'],
                        capture_output=True, text=True).stdout for i in range(1, n + 1)]

def roman(k):
    vals = [(10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    out = ''
    for v, s in vals:
        while k >= v:
            out += s; k -= v
    return out

targets = {
    'abstract': 'ABSTRACT', 'abbrev': 'LIST OF ABBREVIATIONS',
    'ch1': 'CHAPTER 1 –', 'ch2': 'CHAPTER 2 –', 'ch3': 'CHAPTER 3 –', 'ch4': 'CHAPTER 4 –', 'ch5': 'CHAPTER 5 –',
    '1.1': '1.1 INTRODUCTION OF THE PROJECT', '1.2': '1.2 NEED OF THE PROJECT', '1.3': '1.3 OBJECTIVE OF THE PROJECT',
    '1.4': '1.4 THEME OF THE PROJECT', '1.5': '1.5 ORGANIZATION', '1.6': '1.6 SUMMARY',
    '2.1': '2.1 INTRODUCTION', '2.2': '2.2 EXISTING SYSTEMS', '2.3': '2.3 SCOPE OF THE PROJECT', '2.4': '2.4 SUMMARY',
    '3.1': '3.1 FUNDAMENTALS', '3.2': '3.2 PROPOSED SYSTEM ARCHITECTURE', '3.3': '3.3 SYSTEM WORKFLOW',
    '3.4': '3.4 IMPLEMENTATION DETAILS', '3.5': '3.5 SUMMARY',
    '4.1': '4.1 FUNCTIONAL ANALYSIS', '4.2': '4.2 EXPERIMENTAL ANALYSIS', '4.3': '4.3 SUMMARY',
    '5.1': '5.1 CONCLUSION', '5.2': '5.2 FUTURE SCOPE', '5.3': '5.3 APPLICATIONS', '5.4': '5.4 SUMMARY',
    'refs': 'REFERENCES',
}
toc_idx = next(i for i, t in enumerate(pages) if 'TABLE OF CONTENT' in t)
abstract_idx = next(i for i, t in enumerate(pages) if re.search(r'^\s*ABSTRACT\s*$', t, re.M))
ch1_idx = next(i for i, t in enumerate(pages) if 'CHAPTER 1 –' in t and i > toc_idx + 1)
result = {}
for key, text in targets.items():
    for i, t in enumerate(pages):
        if i in (toc_idx, toc_idx + 1) or i < abstract_idx:
            continue
        if any(line.strip().startswith(text) for line in t.splitlines()):
            result[key] = str(i - ch1_idx + 1) if i >= ch1_idx else roman(5 + i - abstract_idx)
            break
json.dump(result, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'toc_pages.json'), 'w'), indent=1)
missing = sorted(set(targets) - set(result))
print(result, '\nmissing:', missing)
