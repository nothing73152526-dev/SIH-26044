// Builds the SkillBridge mini project report (DOCX) in the K.S.R. College
// Design Studio report format.  Usage:  node report/build_report.js
// Edit the DETAILS block below (name, register number, guide, date) and re-run.

const fs = require('fs');
const path = require('path');
let docx;
try { docx = require('docx'); } catch (e) { docx = require('/opt/node22/lib/node_modules/docx'); }
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, LevelFormat, Footer, PageNumber,
  NumberFormat, TabStopType, VerticalAlign,
} = docx;

// ----------------------------------------------------------------- DETAILS
const DETAILS = {
  title: 'SKILLBRIDGE – AI BASED ACADEMIA–INDUSTRY COLLABORATION PORTAL',
  shortTitle: 'SkillBridge',
  course: '24CBP34 - DESIGN STUDIO - I LABORATORY',
  studentName: 'STUDENT NAME',
  regNo: 'REGISTER NUMBER',
  monthYear: 'November 2026',
  shortMonthYear: 'Nov 2026',
  hod: 'Mrs. K.Sudha, M.E, Ph.D',
  hodAck: 'Mrs. K.Sudha M.E., Ph.D',
  guide: 'Ms. A.Priyadharshini, M.E',
  guideAck: 'A.Priyadharshini, M.E.',
};
// -------------------------------------------------------------------------

const ASSETS = path.join(__dirname, 'assets');
const OUT = process.argv[2] || path.join(__dirname, 'SkillBridge_Project_Report.docx');
const PAGES_FILE = path.join(__dirname, 'toc_pages.json');
const TOC_PAGES = fs.existsSync(PAGES_FILE) ? JSON.parse(fs.readFileSync(PAGES_FILE, 'utf8')) : {};

const FONT = 'Times New Roman';
const BODY = 26;          // 13 pt (half-points)
const LINE = 340;         // ≈1.2 line spacing (matches template)
const LINE_WIDE = 440;    // ≈1.5 line spacing (abstract, acknowledgement)
const TEXT_WIDTH = 8352;  // A4 width 11906 - margins (1980 + 1574)

// ---------- inline markup: **bold**, *italic*
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: BODY, ...base }));
    const tok = m[0];
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size: BODY, ...base, bold: true }));
    else out.push(new TextRun({ text: tok.slice(1, -1), font: FONT, size: BODY, ...base, italics: true }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: BODY, ...base }));
  return out;
}

// body paragraph (indented first line, justified)
const P = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  indent: o.noIndent ? undefined : { firstLine: 720 },
  spacing: { line: o.line || LINE, after: o.after ?? 160, before: o.before ?? 0 },
  children: runs(text, o.run || {}),
});
// plain paragraph without first-line indent
const PN = (text, o = {}) => P(text, { ...o, noIndent: true });

const CH = (text, first = false) => new Paragraph({
  alignment: AlignmentType.CENTER,
  pageBreakBefore: !first,
  spacing: { before: 0, after: 360 },
  children: [new TextRun({ text, font: FONT, size: 32, bold: true })],
});
const H2 = (text, o = {}) => new Paragraph({
  keepNext: true,
  spacing: { before: o.before ?? 280, after: 160 },
  children: [new TextRun({ text, font: FONT, size: 28, bold: true })],
});
const H3 = (text, o = {}) => new Paragraph({
  keepNext: true,
  spacing: { before: o.before ?? 120, after: 120 },
  indent: o.indent ? { left: o.indent } : undefined,
  children: [new TextRun({ text, font: FONT, size: BODY, bold: true })],
});
const BUL = (text, level = 0, o = {}) => new Paragraph({
  numbering: { reference: 'bullets', level },
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 100 },
  children: runs(text, o.run || {}),
});
// "1. Heading" + indented explanation, as used throughout the template
const POINT = (n, head, body) => [H3(`${n}. ${head}`), P(body)];
const LABEL = (text) => new Paragraph({
  keepNext: true,
  spacing: { before: 240, after: 200 },
  children: [new TextRun({ text, font: FONT, size: 24, bold: true })],
});
const CAPTION = (text) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 100, after: 240 },
  children: [new TextRun({ text, font: FONT, size: 22, italics: true })],
});
const BLANK = (n = 1, size = BODY) => Array.from({ length: n }, () => new Paragraph({ children: [new TextRun({ text: '', size })] }));

function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
function IMG(name, widthPx, o = {}) {
  const file = path.join(ASSETS, name);
  const { w, h } = pngSize(file);
  let width = widthPx, height = Math.round(widthPx * h / w);
  if (o.maxH && height > o.maxH) { height = o.maxH; width = Math.round(o.maxH * w / h); }
  return new ImageRun({ type: 'png', data: fs.readFileSync(file), transformation: { width, height },
    altText: { title: name, description: o.alt || name, name } });
}
const FIG = (name, widthPx, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  keepNext: !!o.keepNext,
  spacing: { before: 120, after: 60 },
  children: [IMG(name, widthPx, o)],
});

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
const LINE_B = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const GREY_B = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const ALL = (b) => ({ top: b, bottom: b, left: b, right: b });

const cellText = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.LEFT,
  spacing: { before: o.before ?? 40, after: o.after ?? 40, line: o.line || 240 },
  children: runs(text, { size: o.size || BODY, ...(o.run || {}) }),
});

// ======================================================= FRONT MATTER
const C = (text, size, o = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: o.before ?? 0, after: o.after ?? 120, line: o.line || 276 },
  children: [new TextRun({ text, font: FONT, size, bold: o.bold !== false, italics: !!o.italics })],
});

const cover = [
  ...BLANK(1),
  C(DETAILS.title, 30, { after: 1300, line: 360 }),
  C(DETAILS.course, 28, { after: 400 }),
  C('A MINI PROJECT REPORT', 28, { after: 1100 }),
  C('Submitted by', 28, { italics: true, after: 800 }),
  C(`${DETAILS.studentName}  -  (${DETAILS.regNo})`, 28, { after: 500 }),
  C('in  partial fulfilment for the award of the degree of', 28, { bold: false, italics: true, after: 800 }),
  C('BACHELOR OF ENGINEERING', 28, { after: 240 }),
  C('in', 28, { italics: true, after: 480 }),
  C('COMPUTER SCIENCE AND ENGINEERING', 28, { after: 0 }),
  C('(CYBER SECURITY)', 28, { after: 1400 }),
  new Table({
    width: { size: TEXT_WIDTH + 1000, type: WidthType.DXA },
    columnWidths: [1900, TEXT_WIDTH + 1000 - 3800, 1900],
    borders: NO_BORDERS,
    alignment: AlignmentType.CENTER,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: 1900, type: WidthType.DXA }, borders: ALL(NONE), verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [IMG('ksr_logo.png', 100)] })] }),
      new TableCell({ width: { size: TEXT_WIDTH + 1000 - 3800, type: WidthType.DXA }, borders: ALL(NONE), verticalAlign: VerticalAlign.CENTER,
        children: [C('K.S.R COLLEGE OF ENGINEERING', 28, { after: 60 }), C('(Autonomous)', 28, { after: 60 }), C('TIRUCHENGODE 637 215', 28, { after: 0 })] }),
      new TableCell({ width: { size: 1900, type: WidthType.DXA }, borders: ALL(NONE), verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [IMG('anna_univ_logo.png', 100)] })] }),
    ] })],
  }),
  C('ANNA UNIVERSITY :: CHENNAI 600 025', 28, { before: 80, after: 160 }),
  C(DETAILS.monthYear, 28),
];

const sigCell = (lines) => new TableCell({
  width: { size: TEXT_WIDTH / 2, type: WidthType.DXA }, borders: ALL(NONE),
  children: lines.map(([t, bold]) => new Paragraph({ spacing: { after: 60, line: 300 },
    children: [new TextRun({ text: t, font: FONT, size: 24, bold: !!bold })] })),
});
const bonafide = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  ...BLANK(1),
  C('K.S.R. COLLEGE OF ENGINEERING (AUTONOMOUS)', 28, { after: 0 }),
  C('TIRUCHENGODE 637215', 28, { after: 300 }),
  C('ANNA UNIVERSITY:: CHENNAI 600025', 28, { after: 360 }),
  C('BONAFIDE CERTIFICATE', 28, { after: 300 }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED, spacing: { line: 400, after: 120 },
    children: [
      new TextRun({ text: 'Certified that this report “', font: FONT, size: 24 }),
      new TextRun({ text: DETAILS.title, font: FONT, size: 24, bold: true }),
      new TextRun({ text: '” is the bonafide work of ', font: FONT, size: 24 }),
      new TextRun({ text: `${DETAILS.studentName} (${DETAILS.regNo})`, font: FONT, size: 24, bold: true }),
      new TextRun({ text: ' who carried out the project work for Design Studio I under my supervision.', font: FONT, size: 24 }),
    ],
  }),
  ...BLANK(5),
  new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: [TEXT_WIDTH / 2, TEXT_WIDTH / 2], borders: NO_BORDERS,
    rows: [new TableRow({ children: [
      sigCell([['SIGNATURE', true], [DETAILS.hod], ['HEAD OF THE DEPARTMENT', true], ['Professor & Head,'],
        ['Department of CSE(Cyber Security)'], ['K.S.R College of Engineering'], ['Tiruchengode - 637 215']]),
      sigCell([['SIGNATURE', true], [DETAILS.guide], ['SUPERVISOR', true], ['Assistant Professor,'],
        ['Department of CSE(Cyber Security)'], ['K.S.R College of Engineering'], ['Tiruchengode – 637 215']]),
    ] })],
  }),
  new Paragraph({ border: { bottom: { style: BorderStyle.DASHED, size: 6, color: '000000', space: 4 } }, spacing: { after: 300 }, children: [] }),
  new Paragraph({ spacing: { after: 700 }, children: [new TextRun({ text: 'Submitted for the project viva-voce examination held on ………………………', font: FONT, size: 24 })] }),
  new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH - 100 }],
    children: [new TextRun({ text: 'Internal Examiner', font: FONT, size: 24, bold: true }),
      new TextRun({ text: '\tExternal Examiner', font: FONT, size: 24, bold: true })],
  }),
];

const ackP = (parts) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 720 }, spacing: { line: LINE_WIDE, after: 160 },
  children: parts.map((p) => new TextRun({ text: p[0], bold: !!p[1], font: FONT, size: 24 })),
});
const acknowledgement = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  C('ACKNOWLEDGEMENT', 28, { after: 300 }),
  ackP([['We feel highly honored to extend our sincere gratitude to our beloved Founder Theivathiru Lion '], ['Dr. K. S. Rangasamy, MJF.', 1],
    [' and our Chairman '], ['Mr.R.Srinivasan, B.B.M., M.I.S.T.E.,', 1], [' for giving us an opportunity and facilities for the completion of this project.']]),
  ackP([['We express our heartfelt gratitude to our respected Dean '], ['Dr. M.Venkatesan, M.E., Ph.D.,', 1],
    [' for the continuous support, valuable guidance, and encouragement extended to us in all our academic endeavors.']]),
  ackP([['We express our deep sense of gratitude to our beloved Principal '], ['Dr. P.Meenakshi Devi, M.E., Ph.D.,', 1],
    [' who supported us in all our endeavors and has been responsible for motivating us throughout our career.']]),
  ackP([['We feel highly elated to thank our respectable Head of the Department '], [DETAILS.hodAck, 1],
    [' who guided us and was a pillar of support for the successful completion of the project and our guide '], [DETAILS.guideAck, 1],
    [' who has guided us by providing fruitful ideas in various stages of our project and helping us in completing this project.']]),
  ackP([['It is a pleasure to express our gratitude to our beloved parents for providing their support and confidence to us for the completion of the project and our heartfelt thanks to our entire department faculty members, non-teaching staff and beloved friends, directly and indirectly who helped us during the tenure of the project.']]),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 200 },
    children: [new TextRun({ text: `${DETAILS.studentName} (${DETAILS.regNo})`, font: FONT, size: 24, bold: true })] }),
];

const declaration = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  C('DEPARTMENT OF COMPUTER SCIENCE AND', 28, { after: 0, line: 360 }),
  C('ENGINEERING (CYBER SECURITY)', 28, { after: 200, line: 360 }),
  C('K.S.R. COLLEGE OF ENGINEERING (Autonomous)', 28, { after: 0, line: 360 }),
  C('TIRUCHENGODE 637 215', 28, { after: 0, line: 360 }),
  C(DETAILS.shortMonthYear, 28, { after: 0, line: 360 }),
  C('DECLARATION', 28, { after: 300, line: 360 }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED, indent: { firstLine: 720 }, spacing: { line: LINE_WIDE },
    children: [
      new TextRun({ text: 'We affirm that the project report titled being “', font: FONT, size: 24 }),
      new TextRun({ text: DETAILS.title, font: FONT, size: 26, bold: true }),
      new TextRun({ text: '” submitted in partial fulfilment of the requirements for the award of the degree of Bachelor of Engineering is the original work carried out by us. It has not formed the part of any other project report or dissertation on the basis of which a degree or award was conferred on an earlier occasion on this or any other candidates.', font: FONT, size: 24 }),
    ],
  }),
  ...BLANK(3),
  new Paragraph({ children: [new TextRun({ text: `${DETAILS.studentName} (${DETAILS.regNo})`, font: FONT, size: 24, bold: true })] }),
  ...BLANK(3),
  new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE_WIDE },
    children: [new TextRun({ text: 'To verify that the declaration made by the above candidates is true to the best of my knowledge.', font: FONT, size: 24 })] }),
  ...BLANK(3),
  new Paragraph({
    tabStops: [{ type: TabStopType.LEFT, position: 2900 }],
    children: [new TextRun({ text: 'Date:', font: FONT, size: 24 }), new TextRun({ text: '\tName and Signature of the supervisor with seal', font: FONT, size: 24 })],
  }),
];

const abstract = [
  ...BLANK(2),
  C('ABSTRACT', 30, { after: 360 }),
  P('India produces several million graduates every year, yet industry continues to report a wide gap between what students learn in classrooms and the skills that employers actually need. Internship and placement opportunities are scattered across multiple portals, skill assessment is rarely structured, faculty members have limited exposure to current industry practice, and institutions lack a single view of how ready their students are for employment. **SkillBridge**, an AI-based academia–industry collaboration portal developed for the Smart India Hackathon problem statement SIH26044, has been designed to address these issues by bringing students, industry partners, academicians and institutions together on one role-based web platform. The portal allows students to take industry-shared skill assessments, view a skill map that compares their abilities against industry requirements, identify skill gaps, and receive AI-matched internships, jobs, courses and mentors. Industry partners can publish opportunities and learning programs and shortlist candidates ranked by skill compatibility. Academicians can discover faculty internships, faculty development programs, consultancy and collaborative research, while institutions monitor department-wise readiness, internship participation and the placement pipeline through an analytics dashboard. The prototype is implemented as a responsive web application using HTML5, CSS3 and JavaScript and works on desktops, tablets and mobile phones. By converting skill data into clear, actionable insights, SkillBridge helps reduce the employability gap, strengthens industry–institution partnerships and supports the goals of the National Education Policy 2020 and the Skill India mission.', { line: LINE_WIDE }),
];

// TOC
const tocEntries = [
  ['', 'ABSTRACT', 'abstract', true],
  ['', 'LIST OF ABBREVIATIONS', 'abbrev', true],
  ['1', 'INTRODUCTION', 'ch1', true],
  ['', '1.1 INTRODUCTION OF THE PROJECT', '1.1'], ['', '1.2 NEED OF THE PROJECT', '1.2'],
  ['', '1.3 OBJECTIVE OF THE PROJECT', '1.3'], ['', '1.4 THEME OF THE PROJECT', '1.4'],
  ['', '1.5 ORGANIZATION', '1.5'], ['', '1.6 SUMMARY', '1.6'],
  ['2', 'LITERATURE REVIEW', 'ch2', true],
  ['', '2.1 INTRODUCTION', '2.1'], ['', '2.2 EXISTING SYSTEMS AND THEIR LIMITATIONS', '2.2'],
  ['', '2.3 SCOPE OF THE PROJECT', '2.3'], ['', '2.4 SUMMARY', '2.4'],
  ['3', 'SYSTEM DESIGN AND DEVELOPMENT', 'ch3', true],
  ['', '3.1 FUNDAMENTALS OF WEB-BASED ACADEMIA–INDUSTRY COLLABORATION SYSTEMS', '3.1'],
  ['', '3.2 PROPOSED SYSTEM ARCHITECTURE', '3.2'], ['', '3.3 SYSTEM WORKFLOW', '3.3'],
  ['', '3.4 IMPLEMENTATION DETAILS', '3.4'], ['', '3.5 SUMMARY', '3.5'],
  ['4', 'RESULT AND DISCUSSION', 'ch4', true],
  ['', '4.1 FUNCTIONAL ANALYSIS', '4.1'], ['', '4.2 EXPERIMENTAL ANALYSIS', '4.2'], ['', '4.3 SUMMARY', '4.3'],
  ['5', 'CONCLUSIONS', 'ch5', true],
  ['', '5.1 CONCLUSION', '5.1'], ['', '5.2 FUTURE SCOPE', '5.2'], ['', '5.3 APPLICATIONS', '5.3'],
  ['', '5.4 SUMMARY', '5.4'], ['', 'REFERENCES', 'refs'],
];
const TOC_W = [1700, 5352, 1300];
const tocRow = (a, b, c, bold) => new TableRow({ children: [a, b, c].map((t, i) => new TableCell({
  width: { size: TOC_W[i], type: WidthType.DXA }, borders: ALL(NONE),
  children: [new Paragraph({ spacing: { before: 100, after: 100, line: 300 },
    children: [new TextRun({ text: t, font: FONT, size: 24, bold })] })],
})) });
const toc = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  C('TABLE OF CONTENT', 28, { after: 300 }),
  new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: TOC_W, borders: NO_BORDERS,
    rows: [tocRow('CHAPTER NO', 'TITLE', 'PAGE NO', true),
      ...tocEntries.map(([n, t, k, bold]) => tocRow(n, t, String(TOC_PAGES[k] ?? ''), !!bold))],
  }),
];

const abbrevs = [
  ['AI', 'Artificial Intelligence'], ['ML', 'Machine Learning'], ['SIH', 'Smart India Hackathon'],
  ['AICTE', 'All India Council for Technical Education'], ['UGC', 'University Grants Commission'],
  ['NEP', 'National Education Policy'], ['NATS', 'National Apprenticeship Training Scheme'],
  ['FDP', 'Faculty Development Program'], ['TPO', 'Training and Placement Officer'],
  ['RBAC', 'Role-Based Access Control'], ['API', 'Application Programming Interface'],
  ['HTML', 'HyperText Markup Language'], ['CSS', 'Cascading Style Sheets'], ['JS', 'JavaScript'],
  ['UI', 'User Interface'], ['UX', 'User Experience'], ['SQL', 'Structured Query Language'],
  ['DBMS', 'Database Management System'], ['TF-IDF', 'Term Frequency – Inverse Document Frequency'],
  ['OWASP', 'Open Worldwide Application Security Project'], ['XSS', 'Cross-Site Scripting'],
  ['DPDP', 'Digital Personal Data Protection'], ['LPA', 'Lakhs Per Annum'],
];
const AB_W = [2300, 5300];
const abRow = (a, b, bold) => new TableRow({ children: [a, b].map((t, i) => new TableCell({
  width: { size: AB_W[i], type: WidthType.DXA }, borders: ALL(GREY_B),
  margins: { left: 100, right: 100 },
  children: [new Paragraph({ spacing: { before: bold ? 120 : 30, after: bold ? 200 : 30 },
    children: [new TextRun({ text: t, font: FONT, size: 24, bold })] })],
})) });
const abbreviations = [
  new Paragraph({ pageBreakBefore: true, children: [] }),
  C('LIST OF ABBREVIATIONS', 28, { after: 400 }),
  new Table({ width: { size: 7600, type: WidthType.DXA }, columnWidths: AB_W, alignment: AlignmentType.CENTER,
    rows: [abRow('TERMS', 'EXPANSION', true), ...abbrevs.map(([a, b]) => abRow(a, b, false))] }),
];

// ======================================================= CHAPTER 1
const ch1 = [
  CH('CHAPTER 1 – INTRODUCTION', true),
  H2('1.1 INTRODUCTION OF THE PROJECT', { before: 0 }),
  P('Higher education is the foundation of India’s growing knowledge economy. Every year engineering colleges, universities and professional institutions across the country produce a very large number of graduates, and the industry depends on this talent pool to build products, services and research. However, a persistent mismatch exists between the skills taught in academic programs and the skills demanded by employers. Students often graduate without practical exposure, faculty members have limited opportunities to interact with industry, and companies spend considerable time and money on training fresh recruits before they become productive.'),
  P('To address these challenges, the project **SkillBridge** introduces an AI-based Academia–Industry Collaboration Portal. The platform connects four stakeholders – students, industry partners, academicians and institutions – through a single role-based web application. Students take skill assessments shared by industry, view their skill profile and skill gaps, and receive matched internships, jobs, courses and mentors. Industry partners post opportunities, run learning programs and shortlist candidates who are ranked according to skill compatibility. Faculty members discover industry immersion programs, faculty development programs, consultancy and research collaborations, and institutions monitor skill readiness and placement outcomes through an analytics dashboard.'),
  P('The purpose of this project is to bridge the information and skill gap between academia and industry. SkillBridge aims to make the process of skill development, internship discovery and placement transparent, data-driven and continuous rather than a one-time event at the end of a degree program. The platform is designed to be simple, accessible across devices and usable by students, recruiters and faculty with different levels of technical familiarity.'),
  P('In essence, this project acts as a digital collaboration bridge, delivering personalized guidance to every stakeholder in the education-to-employment journey. By harnessing modern web technologies and AI-driven matching, SkillBridge strives to improve employability, strengthen industry–institution partnerships and contribute to the vision of a skilled and future-ready workforce for India.'),

  H2('1.2 NEED OF THE PROJECT'),
  P('The education-to-employment ecosystem in India faces several challenges that directly affect the employability of graduates, the productivity of industry and the reputation of institutions. Despite the availability of many job portals and online courses, a large number of students still depend on informal information, word of mouth and last-minute campus drives for internships and placements. This often leads to missed opportunities, poorly matched roles and long periods of unemployment or under-employment.'),
  PN('The need for this project arises from the following critical gaps in the current system:'),
  ...POINT(1, 'Widening Skill Gap', 'Employability surveys such as the India Skills Report repeatedly show that only around half of Indian graduates are considered job-ready. Students are not aware of which specific skills employers expect, and curricula change much more slowly than technology.'),
  ...POINT(2, 'Fragmented Internship and Placement Channels', 'Opportunities are spread across government portals, private job sites, company career pages and placement cell notices. Students and institutions have no single place to discover, apply for and track opportunities.'),
  ...POINT(3, 'Absence of Structured Skill Assessment', 'Most students are evaluated only through academic marks. There is no standard, industry-validated mechanism to measure practical and soft skills and to compare them against role requirements.'),
  ...POINT(4, 'Limited Industry Exposure for Faculty', 'Faculty members have few channels to undertake industry internships, consultancy or joint research, which makes it difficult to bring current industry practice into the classroom.'),
  ...POINT(5, 'Curriculum Lag Behind Industry Demand', 'Institutions do not receive timely data on which skills are in demand. As a result, new technologies are introduced into the curriculum late, widening the gap further.'),
  ...POINT(6, 'No Unified Analytics for Institutions and Policymakers', 'Placement and skill data are maintained in spreadsheets and separate records. Institutions cannot easily see department-wise readiness, internship participation or placement trends to plan interventions.'),
  ...POINT(7, 'Unverified Credentials and Data Security', 'Certificates and project records are usually shared as scanned documents that are difficult to verify and easy to forge. A secure, role-based system with a verified digital portfolio is required to build trust between students and recruiters.'),

  H2('1.3 OBJECTIVE OF THE PROJECT'),
  P('The primary objective of the SkillBridge project is to develop an AI-driven academia–industry collaboration platform that empowers students, industry, academicians and institutions with real-time, data-based insights for skill development, internships and placements. The project aims to improve employability, reduce hiring effort and strengthen institution–industry partnerships.'),
  PN('The major objectives of the project are as follows:'),
  ...POINT(1, 'To Provide Structured Skill Assessment', 'To allow industry partners to share assessments with students and to convert the responses into a clear skill profile for each student.'),
  ...POINT(2, 'To Perform Skill Mapping and Gap Analysis', 'To compare each student’s skill levels with industry requirements and highlight the exact gaps that need to be closed.'),
  ...POINT(3, 'To Deliver AI-Matched Opportunities', 'To recommend internships, jobs, courses and mentors that best fit a student’s current profile and career goals, and to rank candidates for recruiters by skill compatibility.'),
  ...POINT(4, 'To Enable Industry Participation', 'To let companies post opportunities, publish certification programs, mentorship sessions and live projects, and shortlist candidates directly on the platform.'),
  ...POINT(5, 'To Support Faculty Development', 'To provide academicians with access to industry immersion programs, faculty development programs, consultancy assignments, guest lectures and collaborative research.'),
  ...POINT(6, 'To Provide Institutional Analytics', 'To give institutions a dashboard of students assessed, average skill match, active internships, placement rate and department-wise readiness, along with trend insights for curriculum planning.'),
  ...POINT(7, 'To Maintain a Verified Digital Portfolio', 'To store certifications and project work as verified records and to track every application from submission to offer.'),
  ...POINT(8, 'To Create a Secure and User-Friendly Platform', 'To provide a simple, responsive interface with role-based access so that every stakeholder sees only the information and actions relevant to them.'),

  H2('1.4 THEME OF THE PROJECT'),
  P('The theme of the project SkillBridge revolves around leveraging modern technologies such as Artificial Intelligence, data analytics and web platforms to build a strong and continuous connection between academia and industry. The project addresses the fundamental need for skill-based, transparent and data-driven career development in a country with a very large and diverse student population.'),
  PN('At its core, the project emphasizes **Education–Industry Integration**, enabling students to receive personalized skill insights and opportunities while giving industry and institutions the data they need to collaborate effectively.'),
  PN('This project aligns with national priorities like:'),
  BUL('**National Education Policy (NEP) 2020**'),
  BUL('**Skill India Mission**'),
  BUL('**Digital India**'),
  BUL('**AICTE Internship Policy**'),
  BUL('**Industry 4.0 Workforce Readiness**'),
  PN('By combining technology with practical academic and industry knowledge, the theme focuses on:', { before: 120 }),
  H3('• Empowering Students with Skill Insights'), PN('Providing clear skill profiles, gap analysis and recommendations that guide students towards employable skills.'),
  H3('• Strengthening Industry Partnerships'), PN('Offering a direct channel for companies to engage with students and faculty through internships, programs and research.'),
  H3('• Continuous Faculty Development'), PN('Connecting academicians with industry immersion, FDPs and consultancy so that teaching stays current.'),
  H3('• Data-Driven Institutional Planning'), PN('Using real placement and readiness data to update curricula and plan training programs.'),
  H3('• Trust and Security'), PN('Ensuring role-based access, a secure document vault and verified credentials so that every stakeholder can trust the data on the platform.'),

  H2('1.5 ORGANIZATION'),
  PN('This report is organized into five chapters:'),
  BUL('**Chapter 1** introduces the project, need, objectives, and theme.'),
  BUL('**Chapter 2** reviews existing internship, placement and skill platforms, related research on job recommender systems, and their limitations.'),
  BUL('**Chapter 3** details the system design, proposed architecture, workflow and implementation methodology.'),
  BUL('**Chapter 4** presents the results, functional analysis, experimental analysis and discussion.'),
  BUL('**Chapter 5** concludes the report and highlights possible future improvements and applications.'),

  H2('1.6 SUMMARY'),
  P('Chapter 1 presented an overview of **SkillBridge – AI Based Academia–Industry Collaboration Portal**, highlighting the motivation, relevance and foundational concepts behind the project. The chapter introduced the purpose of the platform and explained why an intelligent, data-driven collaboration system is essential for students, industry, faculty and institutions across India.'),
  PN('The discussion covered the need for the project, emphasizing major challenges such as the skill gap, fragmented opportunity channels, the absence of structured assessment and the lack of institutional analytics. The objectives were clearly defined to demonstrate how the system aims to empower every stakeholder through skill mapping, AI-based matching and analytics.'),
  PN('The theme of the project focused on education–industry integration supported by modern digital technologies. The organizational structure of the report was also outlined, giving readers a clear understanding of how the subsequent chapters are arranged.'),
  PN('Overall, this chapter set the foundation for understanding the importance and vision of the proposed system, preparing the reader for the detailed technical and analytical sections ahead.'),
];

// ======================================================= CHAPTER 2
const LIM = (title, sub, items) => [
  H3(title),
  ...(sub ? [PN(sub, { after: 80 })] : []),
  H3('Limitations:', { before: 40 }),
  ...items.map((t) => BUL(t)),
];
const ch2 = [
  CH('CHAPTER 2 – LITERATURE REVIEW'),
  H2('2.1 INTRODUCTION', { before: 0 }),
  P('A literature review is essential for understanding the existing research, technologies and digital platforms related to internships, placements, skill development and academia–industry collaboration. It provides insights into how modern portals, recommender systems and skill assessment tools have evolved, and where current systems fall short. This chapter evaluates previously developed systems, research studies and AI-based career platforms to identify technological gaps and opportunities.'),
  PN('In recent years, several digital tools have been introduced for students and job seekers, such as government internship portals, private job and internship sites, online learning platforms and institutional placement management systems. While these systems contribute valuable services, they usually address only one part of the problem – either discovery of jobs, or learning, or record-keeping – and rarely connect students, industry, faculty and institutions in one place. As technology changes rapidly, the need for an integrated, skill-centric and personalized collaboration platform has grown significantly.'),
  PN('This chapter reviews existing literature on job and course recommender systems, government and private internship portals and institutional placement tools. It establishes the foundation for understanding how **SkillBridge** advances beyond these models by using skill mapping, gap analysis, AI-based matching and role-based analytics tailored for the Indian higher-education ecosystem.'),

  H2('2.2 EXISTING SYSTEMS AND THEIR LIMITATIONS'),
  P('Over the years, various platforms have been developed to support students with internships, jobs and skill development. Some of the common systems include government portals, private job and internship platforms, online learning platforms, institutional placement cells and research-based recommender systems. The key existing systems and their limitations are summarized below:'),
  ...LIM('1. Government Portals (e.g., AICTE Internship Portal, National Apprenticeship Training Scheme, Skill India Digital Hub)', null, [
    'Focus on listing opportunities rather than measuring individual skill gaps',
    'Limited personalization and AI-based matching for each student',
    'Faculty development and research collaboration are handled separately',
  ]),
  ...LIM('2. Private Job and Internship Platforms (Internshala, Naukri, LinkedIn)', null, [
    'Matching is largely keyword-based and depends on self-reported skills',
    'No direct involvement of the student’s institution or faculty',
    'Premium features and visibility often require paid subscriptions',
  ]),
  ...LIM('3. Online Learning Platforms', '(NPTEL / SWAYAM, Coursera and similar MOOCs)', [
    'Provide courses and certificates but are not linked to specific job requirements',
    'Learners must decide on their own which courses close which skill gaps',
    'Course completion does not automatically reach recruiters as a verified credential',
  ]),
  ...LIM('4. Institutional Placement Cells', '(Training and Placement Office records, spreadsheets and e-mail notices)', [
    'Manual, time-consuming and prone to data-entry errors',
    'No real-time view of student readiness or department-wise skill gaps',
    'Opportunities depend on the personal network of the placement officer',
  ]),
  ...LIM('5. Research-Based Job Recommender Systems', '(Academic models using content-based, collaborative and hybrid filtering)', [
    'Usually evaluated on offline datasets and not deployed for Indian institutions',
    'Focus only on candidate–job matching and ignore faculty and institutional needs',
    'Rarely provide an explanation of the skill gap behind a recommendation',
  ]),
  H3('Key Limitations Across All Existing Systems'),
  BUL('Not skill-gap centric'),
  BUL('Not connected across students, industry, faculty and institutions'),
  BUL('No structured, industry-validated skill assessment'),
  BUL('Limited verification of credentials and portfolios'),
  BUL('No unified analytics for institutions and policymakers'),
  PN('These limitations justify the need for a comprehensive, AI-driven academia–industry collaboration platform like **SkillBridge**.', { before: 120 }),

  H2('2.3 SCOPE OF THE PROJECT'),
  P('The scope of SkillBridge extends across multiple dimensions of skill development, internships, placements and academic collaboration:', { noIndent: true }),
  ...[
    ['Multi-Stakeholder Platform', 'Designed to serve students, industry partners, academicians and institutions through dedicated role-based views.'],
    ['Skill Assessment and Mapping', 'Uses industry-shared assessments to build skill profiles and compare them with role requirements.'],
    ['AI-Powered Matching', 'Ranks internships, jobs, courses and mentors for students, and candidates for recruiters, based on skill compatibility.'],
    ['Faculty Collaboration', 'Covers faculty internships, faculty development programs, consultancy, guest lectures and collaborative research.'],
    ['Institutional Analytics', 'Provides department-wise readiness, placement pipeline and trend insights to support curriculum and policy decisions.'],
    ['Expandable Architecture', 'Future modules can include resume parsing, verifiable digital credentials, chatbot guidance, multilingual support and integration with national portals.'],
  ].flatMap(([h, b], i) => [H3(`${i + 1}. ${h}`), PN(b)]),
  PN('The scope ensures that the system is not limited to a single college or department but is scalable to institutions across the country.', { before: 200 }),

  H2('2.4 SUMMARY'),
  P('This chapter reviewed existing internship, placement, learning and recommender systems and identified major limitations that hinder their effectiveness. While current tools provide useful services, they lack integration across stakeholders, skill-gap analysis, verified credentials and institutional analytics. The literature highlights the need for a platform capable of delivering personalized, skill-based recommendations while keeping industry and institutions in the loop.'),
  PN('The evaluation clearly shows how **SkillBridge** addresses these gaps by combining modern technologies – AI-based matching, skill mapping, role-based access and analytics – to create a robust, user-friendly digital solution. The insights gained from this chapter provide the foundation for the system design and methodology described in the next chapter.'),
];

// ======================================================= CHAPTER 3
const ch3 = [
  CH('CHAPTER 3 – SYSTEM DESIGN AND DEVELOPMENT'),
  H2('3.1 FUNDAMENTALS OF WEB-BASED ACADEMIA–INDUSTRY COLLABORATION SYSTEMS', { before: 0 }),
  P('Web-based academia–industry collaboration systems are platforms designed to connect educational institutions with employers by providing skill information, opportunities and decision-support tools via the internet. These systems operate on the principle of presenting complex data – such as skill levels, role requirements, applications and placement statistics – in a simple and actionable format for each type of user.'),
  PN('Fundamental characteristics include:'),
  ...POINT(1, 'User-Centric, Role-Based Access', 'Students, recruiters, faculty and institution administrators access the same platform but see different dashboards and actions. The browser-based approach removes the need for installation and ensures multi-platform compatibility.'),
  ...POINT(2, 'Centralized Data Processing', 'The backend maintains student profiles, skills, opportunities, applications and certifications in one place, ensuring consistency, scalability and real-time updates for every stakeholder.'),
  H3('3. Dynamic Decision Support'),
  P('These systems integrate matching algorithms and analytics to generate recommendations such as:'),
  BUL('**Suitable internships and jobs**'), BUL('**Courses to close skill gaps**'), BUL('**Mentors and learning programs**'), BUL('**Ranked candidate lists for recruiters**'),
  H3('4. Integration with External Sources'),
  P('Modern collaboration platforms can rely on multiple data sources such as:'),
  BUL('**Industry assessments and job descriptions**'), BUL('**Learning and certification platforms**'), BUL('**National internship and apprenticeship portals**'),
  PN('The web platform acts as a unified interface bringing these data streams together.', { before: 80 }),
  H3('5. Visualization & User Interface'),
  PN('A key part of such systems is presenting results in a user-friendly manner:'),
  BUL('**Dashboard statistics**'), BUL('**Skill progress and gap bars**'), BUL('**Clear status of applications**'), BUL('**Responsive layout for all devices**'),
  PN('SkillBridge follows all these fundamentals and enhances them with AI-driven matching and analytics.', { before: 120 }),

  H2('3.2 PROPOSED SYSTEM ARCHITECTURE'),
  P('The proposed architecture of SkillBridge is designed using modular components to ensure scalability, security and smooth flow of data between stakeholders. The architecture consists of the following layers:'),
  H3('1. User Interface Layer'),
  BUL('Web-based front-end with four role views: Student, Industry, Academician and Institution'),
  BUL('Built for easy navigation through a single role switcher'),
  BUL('Responsive layout for desktop, tablet and mobile'),
  BUL('Displays statistics, skill maps, opportunities and analytics clearly'),
  H3('2. Application Layer'),
  PN('This layer handles:'),
  BUL('Role-based access control and session handling'), BUL('Processing assessment responses and opportunity postings'),
  BUL('Tracking applications, shortlists and offers'), BUL('Formatting results for display'),
  H3('3. AI Processing Layer'),
  PN('The intelligent decision-making component:'),
  BUL('Student skills + Role requirements = Skill Match Score and Skill Gap'),
  BUL('Ranks opportunities for students and candidates for recruiters'),
  BUL('Recommends courses and mentors that close the largest gaps'),
  H3('4. Data Layer'),
  PN('Stores and manages:'),
  BUL('Student profiles and skill scores'), BUL('Job, internship and learning program data'),
  BUL('Verified certifications and project records'), BUL('Applications and placement records'),
  H3('5. Output & Analytics Layer'),
  PN('Produces:'),
  BUL('Matched opportunities and recommendations'), BUL('Verified digital portfolio'),
  BUL('Department-wise readiness and placement pipeline'), BUL('Trend insights for curriculum planning'),
  PN('This architecture ensures secure, accurate and personalized collaboration between academia and industry.', { before: 120 }),
  LABEL('SYSTEM ARCHITECTURE DIAGRAM:'),
  FIG('arch.png', 560, { keepNext: true }),
  CAPTION('Figure 3.1 Layered architecture of SkillBridge'),

  H2('3.3 SYSTEM WORKFLOW'),
  H3('The workflow of SkillBridge follows a sequential, systematic process:', { before: 0 }),
  H3('Step 1: User Login and Role Selection'),
  BUL('The user signs in and selects or is assigned a role – Student, Industry, Academician or Institution.'),
  BUL('The interface loads only the dashboard and actions permitted for that role.'),
  H3('Step 2: Skill Assessment'),
  BUL('The student answers an assessment shared by industry covering technical skills (for example Python, data analysis, SQL) and soft skills.'),
  BUL('Responses are converted into skill levels for the student profile.'),
  H3('Step 3: Skill Mapping and Gap Analysis'),
  BUL('**Each skill is compared with the requirement of the target role.**'),
  BUL('**The system calculates:**'),
  BUL('Current skill level', 1), BUL('Required skill level', 1), BUL('Skill gap', 1), BUL('Overall skill match percentage', 1),
  H3('Step 4: AI Matching and Recommendation'),
  BUL('**The system outputs:**'),
  BUL('Matched internships and jobs with matched and missing skills', 1), BUL('Courses to close the gaps', 1),
  BUL('Suitable mentors', 1), BUL('Ranked candidates for recruiters', 1),
  H3('Step 5: Action and Tracking'),
  BUL('Students apply, industry shortlists, faculty enrol in programs, and every status change is tracked.'),
  BUL('Verified certifications and projects are added to the student’s digital portfolio.'),
  H3('Step 6: Institutional Analytics'),
  BUL('Aggregated data is shown to the institution as readiness, placement pipeline and trend insights.'),
  LABEL('WORK FLOW DIAGRAM:'),
  FIG('workflow.png', 470, { keepNext: true }),
  CAPTION('Figure 3.2 Workflow of SkillBridge'),

  H2('3.4 IMPLEMENTATION DETAILS'),
  P('This section describes the technologies, modules and methods used to build the SkillBridge platform.'),
  H3('Frontend Development'),
  BUL('HTML5, CSS3 and JavaScript (no external frameworks, a single lightweight page)'),
  BUL('CSS custom properties for a consistent colour theme and CSS Grid for layout'),
  BUL('Media query at 760 px that collapses multi-column grids into a single column for mobile devices'),
  BUL('Four role views rendered in the same page and switched without reloading'),
  H3('Prototype Modules (JavaScript Functions)'),
  BUL('**setRole()** – switches between Student, Industry, Academician and Institution views and highlights the active role'),
  BUL('**pick()** – records the selected option for each assessment question'),
  BUL('**runAssessment()** – generates the skill profile and displays AI recommendations (course, internship and mentor)'),
  BUL('**applyJob()** – submits an application and updates the button state to “Applied”'),
  BUL('**postJob()** – publishes a new opportunity with title, required skills and details, marking it as live with students notified'),
  H3('AI & Matching Implementation'),
  BUL('Skill match score: Match % = Σ (wᵢ × min(sᵢ, rᵢ) / rᵢ) / Σ wᵢ × 100, where sᵢ is the student’s level, rᵢ the required level and wᵢ the weight of skill i'),
  BUL('Skill gap for each skill = rᵢ − sᵢ (shown as the red gap bar)'),
  BUL('Opportunities ranked by match score; matched skills shown in green and missing skills in red'),
  BUL('Planned model: TF-IDF / cosine-similarity matching of skills and job descriptions using Python and scikit-learn'),
  H3('Database (Planned)'),
  BUL('Tables for users, roles, skills, assessments, opportunities, applications, certifications and placements'),
  BUL('MySQL or PostgreSQL with access controlled per role'),
  H3('Security Considerations'),
  BUL('Role-based access control so each user sees only their permitted data'),
  BUL('Secure document vault for certificates with verification status'),
  BUL('User input to be escaped before rendering to prevent Cross-Site Scripting (OWASP Top 10)'),
  BUL('Handling of personal data in line with the Digital Personal Data Protection Act, 2023'),
  H3('Hosting'),
  BUL('The prototype is a static web page and can be hosted on any static hosting service or opened directly in a browser'),
  P('This implementation ensures a smooth, fast and secure experience for every stakeholder.', { before: 120 }),

  H2('3.5 SUMMARY'),
  P('This chapter described the system design and development process behind **SkillBridge**. It introduced the fundamentals of web-based academia–industry collaboration systems and presented the proposed layered architecture that integrates user roles, application logic, AI-based matching, data storage and analytics. The workflow and implementation details explained how the system collects assessment data, computes skill matches and gaps, and delivers personalized recommendations and analytics to each user.'),
  PN('The next chapter will present the results, outputs and analysis of system performance.'),
];

// ======================================================= CHAPTER 4
const FA = (n, head, intro, items, close) => [
  H3(`${n}. ${head}`, { before: 100 }),
  ...(intro ? [PN(intro, { after: 60 })] : []),
  ...items.map((t) => BUL(t)),
  ...(close ? [PN(close, { after: 120 })] : []),
];
const TEST_W = [1500, 1900, 1750, 2152, 1050];
const testCell = (t, i, bold) => new TableCell({
  width: { size: TEST_W[i], type: WidthType.DXA }, borders: ALL(LINE_B), margins: { left: 80, right: 80 },
  children: [cellText(t, { size: 22, run: bold ? { bold: true } : {}, align: i === 4 ? AlignmentType.CENTER : AlignmentType.LEFT })],
});
const tests = [
  ['Role Switching Test', 'Switching between the four stakeholder views', 'Student, Industry, Academician, Institution buttons', 'Correct view shown, others hidden, active role highlighted', 'Passed'],
  ['Assessment Test', 'Selecting answers and generating the skill profile', 'One option per question, Generate Skill Profile', 'Selected options highlighted; AI recommendations displayed', 'Passed'],
  ['Skill Mapping Test', 'Displaying skill levels against industry requirement', 'Python, SQL, Communication scores', 'Skill bars and gap bars with correct percentages', 'Passed'],
  ['Application Test', 'Applying to a matched internship', 'Apply button on a job card', 'Status changes to “Applied” and button disabled', 'Passed'],
  ['Opportunity Posting Test', 'Industry publishes a new role', 'Cloud Support Intern; Linux, AWS, Networking', 'New card listed as “Live – students notified”', 'Passed'],
  ['Shortlisting Test', 'Recruiter shortlists a ranked candidate', 'Shortlist on 88% match candidate', 'Button changes to “Sent” and is disabled', 'Passed'],
  ['Responsiveness Test', 'Layout on different screen sizes', '1366 px, 768 px and 360 px widths', 'No horizontal scrolling; grid collapses to one column on mobile', 'Passed'],
  ['Performance Test', 'Page load and view switching speed', '5 page loads per screen size', 'Average load 20–28 ms; role switch < 1 ms', 'Passed'],
  ['Error Test', 'Checking for script errors', 'All interactions on all views', 'No JavaScript errors reported', 'Passed'],
];
const ch4 = [
  CH('CHAPTER 4 – RESULT AND DISCUSSION'),
  H2('4.1 FUNCTIONAL ANALYSIS', { before: 120 }),
  P('The functional analysis focuses on how effectively the SkillBridge system performs its intended operations and fulfils the needs of its users. The prototype was evaluated based on correctness of each feature, responsiveness, usability and the ability to deliver meaningful skill insights and recommendations to every role.'),
  ...FA(1, 'Role-Based Navigation', 'The system correctly separates the four stakeholder views:', ['Student dashboard', 'Industry recruiter dashboard', 'Academician portal', 'Institution analytics'], 'Only one view is visible at a time and the active role is highlighted.'),
  ...FA(2, 'Skill Assessment Handling', 'Functional testing showed that the system successfully:', ['Captures one answer per assessment question', 'Highlights the selected option', 'Generates a skill profile on submission'], 'This confirms correct handling of user input for assessments.'),
  ...FA(3, 'Skill Mapping and AI Recommendation', 'The core functionality is tested for:', ['Correct display of skill levels (78%, 55% and 82%)', 'Correct display of skill gaps (12%, 35% and 8%)', 'Relevant course, internship and mentor recommendations', 'Matched and missing skills on every opportunity'], 'The system consistently presents recommendations that target the largest skill gap (SQL & Databases).'),
  ...FA(4, 'Industry and Faculty Features', 'The interface:', ['Publishes new opportunities instantly', 'Ranks candidates by match percentage and supports shortlisting', 'Lists industry learning programs, FDPs, consultancy and research collaborations'], null),
  ...FA(5, 'Institutional Analytics', 'The dashboard shows:', ['Students assessed, average skill match, active internships and placement rate', 'Department-wise skill readiness', 'Placement pipeline from registration to offers with trend insights'], null),
  ...FA(6, 'Usability', 'The system has been verified to work on:', ['Mobile devices', 'Laptops', 'Tablets'], 'The UI is beginner-friendly and suitable for students, recruiters and faculty with varied digital literacy.'),
  PN('Overall, the functional analysis shows that SkillBridge performs its core tasks effectively and meets the intended project objectives.', { before: 120 }),

  H2('4.2 EXPERIMENTAL ANALYSIS'),
  P('The prototype was tested in the Chromium web browser at three screen sizes. Each feature was exercised through automated browser interaction, and page load times were measured using the browser’s Navigation Timing API with the page loaded from local storage. The results are summarized below:'),
  new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: TEST_W,
    rows: [
      new TableRow({ tableHeader: true, children: ['Test Category', 'Description', 'Input', 'Output', 'Result'].map((t, i) => testCell(t, i, true)) }),
      ...tests.map((r) => new TableRow({ cantSplit: true, children: r.map((t, i) => testCell(t, i, i === 0)) })),
    ],
  }),
  LABEL('PROTOTYPE WEB DESKTOP:'),
  FIG('student_dashboard.png', 560, { keepNext: true }),
  CAPTION('Figure 4.1 Student dashboard with skill assessment and skill mapping'),
  FIG('industry.png', 560, { keepNext: true }),
  CAPTION('Figure 4.2 Industry dashboard – opportunity posting and AI-matched candidates'),
  FIG('faculty.png', 560, { keepNext: true, maxH: 330 }),
  CAPTION('Figure 4.3 Academician portal – faculty internships, FDPs and research'),
  FIG('institution.png', 560, { keepNext: true, maxH: 330 }),
  CAPTION('Figure 4.4 Institution analytics – readiness and placement pipeline'),
  LABEL('PROTOTYPE APP MOBILE:'),
  new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA }, columnWidths: [TEXT_WIDTH / 2, TEXT_WIDTH / 2], borders: NO_BORDERS,
    rows: [new TableRow({ cantSplit: true, children: ['mobile_student.png', 'mobile_reco.png'].map((f) => new TableCell({
      width: { size: TEXT_WIDTH / 2, type: WidthType.DXA }, borders: ALL(NONE),
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [IMG(f, 190)] })],
    })) })],
  }),
  CAPTION('Figure 4.5 Mobile view – student dashboard and AI recommendations'),

  H2('4.3 SUMMARY'),
  P('This chapter presented the results and detailed analysis of the **SkillBridge** system. Functional analysis demonstrated that the platform correctly separates the four stakeholder roles, captures assessment responses, presents skill mapping and gap analysis, and delivers relevant recommendations, postings, shortlists and analytics. The system performed reliably in terms of speed, stability and user interface responsiveness.'),
  PN('The experimental analysis, presented through structured test cases, confirmed that every feature of the prototype works as intended on desktop, tablet and mobile screen sizes, with very fast page loads and no script errors.'),
  PN('Overall, the results highlight that **SkillBridge** is a dependable and user-friendly collaboration tool capable of supporting students, industry, faculty and institutions in making informed decisions. The outcomes of Chapter 4 provide the foundation for the concluding reflections presented in the next chapter.'),
];

// ======================================================= CHAPTER 5
const ch5 = [
  CH('CHAPTER 5 – CONCLUSION AND FUTURE SCOPE'),
  H2('5.1 CONCLUSION', { before: 120 }),
  P('**SkillBridge – AI Based Academia–Industry Collaboration Portal** successfully addresses the major challenges faced by students, industry, faculty and institutions in India by providing structured skill assessment, skill mapping, gap analysis and AI-matched opportunities on a single platform. Through role-based dashboards, the system delivers clear insights that help every stakeholder make informed decisions.'),
  PN('The project achieves its primary objective of promoting data-driven skill development, improving employability and strengthening industry–institution collaboration. The system’s simple and responsive interface ensures easy access for users regardless of technical background or device. Overall, SkillBridge is a practical, scalable and impactful solution that contributes to a skilled and future-ready workforce in India.'),
  H2('5.2 FUTURE SCOPE'),
  PN('The system has strong potential for expansion and can be enhanced with several additional features to increase its usefulness for institutions nationwide:'),
  ...[
    ['Machine Learning on Real Placement Data', 'Training matching models on real assessment, application and placement outcomes to improve recommendation accuracy.'],
    ['Secure Backend and Authentication', 'Adding a full backend with database, secure login, multi-factor authentication and audit logs.'],
    ['Verifiable Digital Credentials', 'Integrating with DigiLocker or digitally signed certificates so that recruiters can verify credentials instantly.'],
    ['Resume Parsing with NLP', 'Automatically extracting skills from resumes and project reports to pre-fill student profiles.'],
    ['AI Career Chatbot', '24/7 chatbot guidance for students on skills, courses and opportunities.'],
    ['Multi-Language Support', 'Adding regional languages for better accessibility across states.'],
    ['Mobile Application Version', 'Launching an Android/iOS app with notifications for new matched opportunities.'],
    ['Integration with National Portals', 'Connecting with AICTE internship and apprenticeship portals to widen the pool of opportunities.'],
  ].flatMap(([h, b], i) => [H3(`${i + 1}. ${h}`, { indent: 360 }), P(b, { noIndent: true, run: {} })]),
  PN('These enhancements can significantly improve the system’s impact and inclusion across India.', { before: 120 }),
  H2('5.3 APPLICATIONS'),
  PN('The SkillBridge system can be used in a variety of real-world academic and industry contexts:'),
  ...[
    ['Engineering Colleges and Universities', 'To assess students, close skill gaps and improve placement outcomes.'],
    ['Training and Placement Cells', 'As a single tool to manage opportunities, applications, shortlists and offers.'],
    ['Industry and Recruiters', 'To find job-ready candidates, run certification programs and build a talent pipeline.'],
    ['Faculty Development Cells', 'To organize industry immersion, FDPs, consultancy and collaborative research.'],
    ['Regulatory Bodies and Policymakers', 'To study skill demand and readiness trends for curriculum and policy decisions.'],
    ['Skill Development Agencies and NGOs', 'For use in training, mentoring and employability programs for students.'],
  ].flatMap(([h, b], i) => [PN(`${i + 1}. ${h}`), PN(b)]),
  H2('5.4 SUMMARY'),
  PN('This chapter summarized the achievements, future potential and real-world applications of the **SkillBridge** system. The project successfully demonstrates how technology can strengthen academia–industry collaboration, reduce the skill gap and support data-driven decisions. The conclusion reinforces that the system is scalable, practical and aligned with India’s mission for a skilled and digitally empowered workforce. Future enhancements can make SkillBridge even more impactful and widely accessible to institutions across the nation.'),
];

// ======================================================= REFERENCES
const refs = [
  'Ministry of Education, Government of India, *National Education Policy 2020*, New Delhi, 2020.',
  'All India Council for Technical Education (AICTE), *AICTE Internship Policy: Guidelines & Procedures*, New Delhi, 2019.',
  'All India Council for Technical Education (AICTE), *Model Curriculum for Undergraduate Degree Courses in Engineering & Technology*, 2018.',
  'Ministry of Education’s Innovation Cell & AICTE, *Smart India Hackathon – Problem Statement SIH26044*, Government of India.',
  'Wheebox, *India Skills Report 2024*, in association with AICTE, AIU and CII, 2024.',
  'World Economic Forum, *The Future of Jobs Report 2023*, Geneva, 2023.',
  'Ministry of Skill Development and Entrepreneurship, *Skill India Digital Hub*, Government of India, 2023.',
  'Board of Apprenticeship Training, *National Apprenticeship Training Scheme (NATS) Guidelines*, Ministry of Education, Government of India.',
  'Ricci, F., Rokach, L. & Shapira, B., *Recommender Systems Handbook*, 2nd ed., Springer, 2015.',
  'Siting, Z., Wenxing, H., Ning, Z. & Fan, Y., “Job Recommender Systems: A Survey,” *Proc. 7th International Conference on Computer Science & Education (ICCSE)*, IEEE, 2012.',
  'de Ruijt, C. & Bhulai, S., “Job Recommender Systems: A Review,” *arXiv preprint arXiv:2111.13576*, 2021.',
  'Paparrizos, I., Cambazoglu, B. B. & Gionis, A., “Machine Learned Job Recommendation,” *Proc. 5th ACM Conference on Recommender Systems (RecSys)*, 2011.',
  'Salton, G. & Buckley, C., “Term-Weighting Approaches in Automatic Text Retrieval,” *Information Processing & Management*, vol. 24, no. 5, 1988.',
  'Pedregosa, F. et al., “Scikit-learn: Machine Learning in Python,” *Journal of Machine Learning Research*, vol. 12, 2011.',
  'Sandhu, R. S., Coyne, E. J., Feinstein, H. L. & Youman, C. E., “Role-Based Access Control Models,” *IEEE Computer*, vol. 29, no. 2, 1996.',
  'OWASP Foundation, *OWASP Top 10 – 2021: The Ten Most Critical Web Application Security Risks*, 2021.',
  'Government of India, *The Digital Personal Data Protection Act, 2023*, Ministry of Electronics and Information Technology, 2023.',
  'World Wide Web Consortium (W3C), *Web Content Accessibility Guidelines (WCAG) 2.1*, 2018.',
  'Mozilla Developer Network, *MDN Web Docs: HTML, CSS and JavaScript Reference*, Mozilla Foundation.',
  'National Programme on Technology Enhanced Learning (NPTEL), *SWAYAM–NPTEL Online Certification*, IITs and IISc, Government of India.',
];
const references = [
  new Paragraph({ pageBreakBefore: true, spacing: { after: 300 }, children: [new TextRun({ text: 'REFERENCES', font: FONT, size: 28 })] }),
  ...refs.map((r) => new Paragraph({
    numbering: { reference: 'refs', level: 0 },
    alignment: AlignmentType.JUSTIFIED, spacing: { line: LINE, after: 200 },
    children: runs(r),
  })),
];

// ======================================================= DOCUMENT
const pageFooter = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
  children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 22 })] })] });
const emptyFooter = new Footer({ children: [new Paragraph({ children: [] })] });
const page = (start, fmt) => ({
  size: { width: 11906, height: 16838 },
  margin: { top: 1800, bottom: 1700, left: 1980, right: 1574, footer: 1250 },
  pageNumbers: start ? { start, formatType: fmt } : undefined,
});

const doc = new Document({
  creator: DETAILS.studentName,
  title: DETAILS.title,
  styles: { default: { document: { run: { font: FONT, size: BODY } } } },
  numbering: { config: [
    { reference: 'bullets', levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.BULLET, text: 'o', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 360 } }, run: { font: 'Courier New' } } },
    ] },
    { reference: 'numbers', levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
    ] },
    { reference: 'refs', levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 400 } } } },
    ] },
  ] },
  sections: [
    { properties: { page: page() }, footers: { default: emptyFooter }, children: [...cover, ...bonafide, ...acknowledgement, ...declaration] },
    { properties: { page: page(5, NumberFormat.UPPER_ROMAN) }, footers: { default: pageFooter }, children: [...abstract, ...toc, ...abbreviations] },
    { properties: { page: page(1, NumberFormat.DECIMAL) }, footers: { default: pageFooter }, children: [...ch1, ...ch2, ...ch3, ...ch4, ...ch5, ...references] },
  ],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log('Wrote', OUT); });
