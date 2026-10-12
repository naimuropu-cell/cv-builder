const fs = require('fs');
const path = require('path');
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  ExternalHyperlink,
  UnderlineType,
  convertInchesToTwip
} = require('docx');

const FONT_MAIN = 'Calibri';
const COLOR_PRIMARY = '1E3A8A';   // Deep tech navy
const COLOR_LINK = '2563EB';      // Clean royal blue for links
const COLOR_DARK = '1F2937';      // Dark charcoal for high ATS contrast
const COLOR_MUTED = '4B5563';     // Slate gray

function createLink(text, url) {
  return new ExternalHyperlink({
    children: [
      new TextRun({
        text: text,
        font: FONT_MAIN,
        size: 19,
        bold: true,
        color: COLOR_LINK,
        underline: { type: UnderlineType.SINGLE },
      }),
    ],
    link: url,
  });
}

function createProjectTitleLink(title, url) {
  return [
    new ExternalHyperlink({
      children: [
        new TextRun({
          text: title,
          font: FONT_MAIN,
          size: 21,
          bold: true,
          color: COLOR_PRIMARY,
          underline: { type: UnderlineType.SINGLE },
        }),
      ],
      link: url,
    }),
  ];
}

function createSectionHeading(title) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    border: {
      bottom: {
        color: COLOR_PRIMARY,
        size: 10,
        style: BorderStyle.SINGLE,
      },
    },
    children: [
      new TextRun({
        text: title.toUpperCase(),
        font: FONT_MAIN,
        size: 22,
        bold: true,
        color: COLOR_PRIMARY,
      }),
    ],
  });
}

// 2-Line Hierarchy for Experience and Education
function createTwoLineItemHeading(designation, companyLine, dates) {
  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [
      // Line 1: Designation & Dates
      new TableRow({
        children: [
          new TableCell({
            width: { size: 6800, type: WidthType.DXA },
            margins: { top: 20, bottom: 0, left: 0, right: 60 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: designation,
                    font: FONT_MAIN,
                    size: 21,
                    bold: true,
                    color: COLOR_DARK,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            width: { size: 2500, type: WidthType.DXA },
            margins: { top: 20, bottom: 0, left: 60, right: 0 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: dates,
                    font: FONT_MAIN,
                    size: 19,
                    bold: true,
                    color: COLOR_MUTED,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      // Line 2: Company Name & Type / Location
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9300, type: WidthType.DXA },
            margins: { top: 0, bottom: 20, left: 0, right: 0 },
            columnSpan: 2,
            children: [
              new Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                  new TextRun({
                    text: companyLine,
                    font: FONT_MAIN,
                    size: 19.5,
                    italics: true,
                    color: COLOR_MUTED,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function createProjectHeadingRow(runs, dateText = '2026') {
  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 7700, type: WidthType.DXA },
            margins: { top: 30, bottom: 20, left: 0, right: 60 },
            children: [
              new Paragraph({
                children: runs,
              }),
            ],
          }),
          new TableCell({
            width: { size: 1600, type: WidthType.DXA },
            margins: { top: 30, bottom: 20, left: 60, right: 0 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: dateText,
                    font: FONT_MAIN,
                    size: 19,
                    bold: true,
                    color: COLOR_MUTED,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function createBullet(text, boldLead = '') {
  const children = [];
  if (boldLead) {
    children.push(
      new TextRun({
        text: boldLead + ' ',
        font: FONT_MAIN,
        size: 19.5,
        bold: true,
        color: COLOR_DARK,
      })
    );
  }
  children.push(
    new TextRun({
      text: text,
      font: FONT_MAIN,
      size: 19.5,
      color: COLOR_DARK,
    })
  );

  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 18, after: 18 },
    children: children,
  });
}

function createParagraph(text, isJustified = true) {
  return new Paragraph({
    spacing: { before: 30, after: 50 },
    alignment: isJustified ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    children: [
      new TextRun({
        text: text,
        font: FONT_MAIN,
        size: 19.5,
        color: COLOR_DARK,
      }),
    ],
  });
}

async function buildDoc() {
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.42),
              bottom: convertInchesToTwip(0.42),
              left: convertInchesToTwip(0.65),
              right: convertInchesToTwip(0.65),
            },
          },
        },
        children: [
          // HEADER: Name & Clean Clickable Headings
          new Paragraph({
            spacing: { before: 0, after: 15 },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: 'Md. Naimur Rahman Apu',
                font: FONT_MAIN,
                size: 38, // 19pt
                bold: true,
                color: COLOR_PRIMARY,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 30 },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: 'SOFTWARE QUALITY ASSURANCE (SQA) ENGINEER',
                font: FONT_MAIN,
                size: 21,
                bold: true,
                color: COLOR_MUTED,
              }),
            ],
          }),
          // Contact line with clickable phone & mail
          new Paragraph({
            spacing: { before: 0, after: 20 },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: 'Dhaka, Bangladesh   •   ', font: FONT_MAIN, size: 19, color: COLOR_DARK }),
              createLink('+8801955345331', 'tel:+8801955345331'),
              new TextRun({ text: '   •   ', font: FONT_MAIN, size: 19, color: COLOR_DARK }),
              createLink('naimuropu@gmail.com', 'mailto:naimuropu@gmail.com'),
            ],
          }),
          // Clean Clickable Headings (Portfolio • GitHub • LinkedIn)
          new Paragraph({
            spacing: { before: 0, after: 100 },
            alignment: AlignmentType.CENTER,
            children: [
              createLink('Portfolio', 'https://naimuropu-cell.github.io/'),
              new TextRun({ text: '   •   ', font: FONT_MAIN, size: 19, color: COLOR_DARK }),
              createLink('GitHub', 'https://github.com/naimuropu-cell'),
              new TextRun({ text: '   •   ', font: FONT_MAIN, size: 19, color: COLOR_DARK }),
              createLink('LinkedIn', 'https://www.linkedin.com/in/naimur-rahman-apu/'),
            ],
          }),

          // 1. PROFESSIONAL SUMMARY
          createSectionHeading('Professional Summary'),
          createParagraph(
            'Detail-oriented and proactive Software QA Engineer with hands-on experience in Manual Testing, API Testing (Postman), and automation testing with a focus on Playwright and Selenium WebDriver. Proven track record across enterprise and web applications: designing 3,000+ comprehensive test cases, executing functional, regression, UAT, and cross-browser suites, and filing 220+ detailed bug reports in Jira throughout the STLC. Possesses hands-on frontend development knowledge in JavaScript, React.js, and Vue.js, enabling close, highly effective collaboration with developers to troubleshoot UI/state defects and accelerate root-cause analysis.'
          ),

          // 2. TECHNICAL SKILLS
          createSectionHeading('Technical Skills'),
          createBullet('Manual Testing, Functional, Regression, Smoke, Sanity, Exploratory Testing, UI/UX Testing, Cross-Browser Testing, User Acceptance Testing (UAT), Test Case Design, STLC, SDLC, Agile / Scrum.', 'QA & Testing Methodologies:'),
          createBullet('Playwright, Selenium WebDriver, Postman (REST API Testing & Collections), JMeter (Performance & Load Testing), Chrome DevTools, Jira (Defect Tracking).', 'Automation & QA Tools:'),
          createBullet('Git, GitHub, GitHub Actions (CI/CD workflows), VS Code, Markdown.', 'Version Control & CI/CD:'),
          createBullet('SQL, MySQL, MongoDB, Database Testing, Query Validation, JSON schema verification.', 'Databases & Backend Testing:'),
          createBullet('JavaScript (ES6+), TypeScript, React.js, Vue.js, HTML5, CSS3, Tailwind CSS, Quasar Framework.', 'Frontend & Development:'),

          // 3. PROFESSIONAL WORK EXPERIENCE (2-Line Designation / Company Hierarchy)
          createSectionHeading('Professional Experience'),

          // Experience 1: LastTouch
          createTwoLineItemHeading(
            'Software Test Engineer',
            'LastTouch • Contractual (Remote)',
            'February 2026 – June 2026'
          ),
          createBullet('Conducted end-to-end manual, exploratory, and User Acceptance Testing (UAT) across client-facing web applications to guarantee product stability and UX fidelity.', 'Functional & Manual QA:'),
          createBullet('Developed foundational automated regression test scripts using Playwright and Selenium WebDriver to accelerate sprint release verification.', 'Automation Testing:'),
          createBullet('Validated REST APIs using Postman, ensuring schema compliance, HTTP response codes, query parameters, authentication headers, and error handling.', 'API Testing:'),
          createBullet('Identified, documented, and tracked critical defects in Jira with reproducible steps, expected/actual outcomes, and network logs; collaborated daily with developers during STLC cycles to verify and close bug fixes.', 'Defect Tracking & Developer Collaboration:'),

          // Experience 2: Iching Lab
          createTwoLineItemHeading(
            'Junior Software QA Engineer & Frontend Developer Intern',
            'Iching Lab • Internship (Dhaka, Bangladesh)',
            'October 2025 – February 2026'
          ),
          createBullet('Executed manual, functional, UI, and regression test suites for dynamic web applications in an Agile environment.', 'QA Execution:'),
          createBullet('Partnered directly with development teams to reproduce edge-case UI glitches, inspect DOM elements using Chrome DevTools, and verify code fixes.', 'Cross-Functional Collaboration:'),
          createBullet('Built responsive UI components using React.js and Vue.js, gaining deep understanding of frontend architecture that enhanced test case coverage.', 'Frontend Development:'),
          createBullet('Maintained test documentation, verified bug fixes across sprints, and helped improve release readiness.', 'Bug Lifecycle:'),

          // Experience 3: Nation IT Limited
          createTwoLineItemHeading(
            'Data Entry & Lead Generation Executive',
            'Nation IT Limited • Remote',
            '2024 – October 2025'
          ),
          createBullet('Conducted systematic web research, CRM data curation, and quality verification, consistently maintaining 99%+ accuracy within fast-paced delivery deadlines.', 'Data Quality & Research:'),

          // 4. FEATURED QA PROJECTS
          createSectionHeading('Featured QA Projects'),

          // Project 1: BizPOS Pro
          createProjectHeadingRow(
            createProjectTitleLink(
              'BizPOS Pro – QA Audit & Automation Framework',
              'https://github.com/naimuropu-cell/bizpos-qa-audit-framework'
            ),
            '2026'
          ),
          createBullet('Spearheaded end-to-end QA audit for enterprise POS covering inventory, accounting, cash flow, asset depreciation, customer management, and e-commerce workflows.', 'Scope & Architecture:'),
          createBullet('Authored and managed 1,100+ structured test cases with full Requirements Traceability Matrix (RTM) and documented 120+ defect reports with video/log evidence.', 'Test Coverage & Metrics:'),
          createBullet('Constructed an automated regression framework using Playwright and TypeScript, validating critical business paths and integrating with GitHub Actions CI/CD.', 'Playwright Automation:'),

          // Project 2: School Management System
          createProjectHeadingRow(
            createProjectTitleLink(
              'School Management System Testing',
              'https://github.com/naimuropu-cell/School-Management-System-QA'
            ),
            '2026'
          ),
          createBullet('Executed comprehensive end-to-end testing across 10+ core modules (student enrollment, attendance, gradebook, financial billing, multi-role auth).', 'Core Modules:'),
          createBullet('Designed and executed 2,000+ manual test cases and logged 100+ actionable bug reports with severity/priority categorization.', 'Test Execution:'),
          createBullet('Automated authentication and critical workflows using Playwright and Selenium WebDriver; performed backend SQL database testing and Postman API validation.', 'Automation & DB Testing:'),

          // Project 3: PulseLoad
          createTwoLineItemHeading(
            'PulseLoad – Web Performance Testing Platform',
            'React • TypeScript • FastAPI • JMeter • Postman',
            '2025 – 2026'
          ),
          createBullet('Conducted functional, API, load, stress, and spike testing using JMeter and Postman to evaluate latency, throughput, and concurrent user capacity.', 'Performance QA:'),
          createBullet('Validated real-time analytics dashboards and API endpoints built with React, TypeScript, and FastAPI.', 'Validation:'),

          // Project 4: QA Testing Portfolio
          createTwoLineItemHeading(
            'QA Testing Portfolio (DemoQA & SauceDemo)',
            'Manual Test Suites & Defect Reports',
            '2025'
          ),
          createBullet('Engineered 50+ manual test cases and 20+ bug reports covering complex UI forms, nested frames, modal dialogs, and e-commerce transactions.', 'Manual QA Suite:'),

          // 5. EDUCATION
          createSectionHeading('Education'),
          createTwoLineItemHeading(
            'B.Sc. in Computer Science & Engineering',
            'Green University of Bangladesh (Dhaka, Bangladesh)',
            '2022 – 2026'
          ),
          createBullet('CGPA: 3.16 / 4.00   |   Relevant Coursework: Software Engineering, Database Systems, Object-Oriented Programming, Web Technologies, Algorithms.'),
          
          createTwoLineItemHeading(
            'Higher Secondary Certificate (HSC)',
            'Milestone College Uttara (Dhaka, Bangladesh)',
            '2018 – 2020'
          ),
          createBullet('Result: GPA 5.00 / 5.00 (Golden A+)   |   Group: Science.'),

          // 6. LANGUAGES
          createSectionHeading('Languages'),
          createBullet('Bangla (Native), English (Professional Working Proficiency - Technical Documentation & Team Communication).', 'Proficiency:'),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync('D:\\tryyyy\\Md_Naimur_Rahman_Apu_SQA_Resume.docx', buffer);
  console.log('Successfully updated Md_Naimur_Rahman_Apu_SQA_Resume.docx with 2-line experience hierarchy');
}

buildDoc().catch(err => {
  console.error(err);
  process.exit(1);
});
