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
  ShadingType,
  ImageRun,
  PageBreak,
  convertInchesToTwip
} = require('docx');

const COLOR_ACCENT = 'ED7D31';      // Template warm orange
const COLOR_BLACK = '000000';       // Pure black for first name like template
const COLOR_DARK = '262626';        // Template charcoal
const COLOR_MUTED = '4A5568';       // Clean text gray
const COLOR_BORDER = 'D8DCE3';      // Subtle border

const FONT_HEADING = 'Arial Black';
const FONT_BODY = 'Calibri';

function createHeading(text) {
  return new Paragraph({
    spacing: { before: 140, after: 50 },
    border: {
      left: {
        color: COLOR_ACCENT,
        size: 18,
        style: BorderStyle.SINGLE,
        space: 6,
      },
    },
    children: [
      new TextRun({
        text: ' ' + text,
        font: FONT_HEADING,
        size: 22, // 11pt
        color: COLOR_ACCENT,
        bold: true,
      }),
    ],
  });
}

function createSubHeading(title, subtitle, dates) {
  const children = [
    new TextRun({
      text: title,
      font: FONT_BODY,
      size: 20, // 10pt
      bold: true,
      color: COLOR_DARK,
    }),
  ];
  if (subtitle) {
    children.push(
      new TextRun({
        text: '  —  ' + subtitle,
        font: FONT_BODY,
        size: 19,
        italics: true,
        color: COLOR_MUTED,
      })
    );
  }
  if (dates) {
    children.push(
      new TextRun({
        text: '  (' + dates + ')',
        font: FONT_BODY,
        size: 19,
        bold: true,
        color: COLOR_ACCENT,
      })
    );
  }

  return new Paragraph({
    spacing: { before: 70, after: 35 },
    children: children,
  });
}

function createBullet(text, boldPrefix = '') {
  const children = [];
  if (boldPrefix) {
    children.push(
      new TextRun({
        text: boldPrefix + ' ',
        font: FONT_BODY,
        size: 19.5,
        bold: true,
        color: COLOR_DARK,
      })
    );
  }
  children.push(
    new TextRun({
      text: text,
      font: FONT_BODY,
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

function createNormalParagraph(text, isItalic = false) {
  return new Paragraph({
    spacing: { before: 35, after: 50 },
    children: [
      new TextRun({
        text: text,
        font: FONT_BODY,
        size: 19.5,
        italics: isItalic,
        color: COLOR_DARK,
      }),
    ],
  });
}

function createTable(rowsData) {
  const tableRows = rowsData.map((row, index) => {
    const isHeader = index === 0;
    return new TableRow({
      children: [
        new TableCell({
          width: { size: 3000, type: WidthType.DXA },
          shading: {
            fill: isHeader ? COLOR_ACCENT : (index % 2 === 1 ? 'F7F9FB' : 'FFFFFF'),
            type: ShadingType.CLEAR,
          },
          margins: { top: 50, bottom: 50, left: 120, right: 120 },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: row[0],
                  font: FONT_BODY,
                  size: 18.5,
                  bold: true,
                  color: isHeader ? 'FFFFFF' : COLOR_DARK,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          width: { size: 6300, type: WidthType.DXA },
          shading: {
            fill: isHeader ? COLOR_ACCENT : (index % 2 === 1 ? 'F7F9FB' : 'FFFFFF'),
            type: ShadingType.CLEAR,
          },
          margins: { top: 50, bottom: 50, left: 120, right: 120 },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: row[1],
                  font: FONT_BODY,
                  size: 18.5,
                  bold: isHeader,
                  color: isHeader ? 'FFFFFF' : COLOR_DARK,
                }),
              ],
            }),
          ],
        }),
      ],
    });
  });

  return new Table({
    width: { size: 9300, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
      left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
      right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'ECECEC' },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: tableRows,
  });
}

async function buildDoc() {
  const photoPath = path.join(__dirname, 'juel_photo_square.jpg');
  const photoData = fs.readFileSync(photoPath);

  // Exact Template Header matching screenshot:
  // Photo on left, Name on right with JUEL in Black and RANA in Orange (#ED7D31)
  // Stacked contact lines beneath, NO bottom horizontal line across header
  const headerTable = new Table({
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
          // Left: Portrait Photo
          new TableCell({
            width: { size: 1700, type: WidthType.DXA },
            margins: { top: 0, bottom: 120, left: 0, right: 140 },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new ImageRun({
                    data: photoData,
                    transformation: { width: 92, height: 92 },
                  }),
                ],
              }),
            ],
          }),
          // Right: Exact 2-Tone Name + Stacked Contact details
          new TableCell({
            width: { size: 7600, type: WidthType.DXA },
            margins: { top: 10, bottom: 120, left: 100, right: 0 },
            children: [
              // 2-tone Name like ANTHONY (black) DAVIES (orange)
              new Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                  new TextRun({
                    text: 'JUEL ',
                    font: FONT_HEADING,
                    size: 44, // 22pt Arial Black
                    bold: true,
                    color: COLOR_BLACK,
                  }),
                  new TextRun({
                    text: 'RANA',
                    font: FONT_HEADING,
                    size: 44, // 22pt Arial Black
                    bold: true,
                    color: COLOR_ACCENT,
                  }),
                ],
              }),
              // Subtitle Role
              new Paragraph({
                spacing: { before: 0, after: 60 },
                children: [
                  new TextRun({
                    text: 'Refrigeration & Electrical Systems Technician',
                    font: FONT_BODY,
                    size: 21,
                    bold: true,
                    color: COLOR_MUTED,
                  }),
                ],
              }),
              // Stacked Contact Lines exactly like template (+1 970 888-3866 / mail / link)
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({
                    text: '+880 1680844360  (Phone & WhatsApp)',
                    font: FONT_BODY,
                    size: 19,
                    color: COLOR_DARK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({
                    text: 'juelranabd64@gmail.com',
                    font: FONT_BODY,
                    size: 19,
                    color: COLOR_DARK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({
                    text: 'Manikgonj, Dhaka, Bangladesh',
                    font: FONT_BODY,
                    size: 19,
                    color: COLOR_DARK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                  new TextRun({
                    text: 'Driving Licence: Category B & A (Valid BD — Ready for Cyprus Conversion)',
                    font: FONT_BODY,
                    size: 19,
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
          // ==================== PAGE 1 ====================
          headerTable,

          // 1. PROFESSIONAL SUMMARY
          createHeading('Summary'),
          createNormalParagraph(
            'Dedicated and skilled Refrigeration & Electrical Systems Technician with over 8 years of hands-on experience (2017 – Present) in installation, preventive maintenance, troubleshooting, and repair of commercial catering equipment and cooling systems. Proven technical expertise in commercial convection ovens, deck ovens, cooking ranges, deep fryers, grills, salamanders, commercial refrigerators, freezers, ice makers, cold rooms, espresso/coffee machines, and industrial dishwashers. Holds a Diploma in Electrical Engineering and specialized industrial automation training (PLC, VFD, HMI). Physically fit, equipped with a clean driving licence, and fully prepared for immediate relocation and deployment to Cyprus.'
          ),

          // 2. CORE COMPETENCIES & TECHNICAL SKILLS
          createHeading('Core Competencies & Technical Skills'),
          createBullet(
            'Installation, preventive maintenance, and advanced diagnostics of commercial convection ovens, deck ovens, cooking ranges, heavy-duty deep fryers, flat-top grills, salamanders, and warming cabinets.',
            'Commercial Catering Equipment:'
          ),
          createBullet(
            'Expert in commercial display refrigerators, reach-in freezers, walk-in cold rooms, and automated ice machines. Proficient in refrigerant leak detection, vacuum testing, nitrogen pressure testing, gas charging (R134a, R404a, R410a), and hermetic/semi-hermetic compressor replacements.',
            'Refrigeration & HVAC Systems:'
          ),
          createBullet(
            'Servicing, descaling, and pump maintenance for commercial espresso/coffee machines, high-capacity dishwashers (conveyor and under-counter), water filtration systems, and booster heaters.',
            'Beverage & Sanitation Equipment:'
          ),
          createBullet(
            '3-phase and single-phase electrical wiring, motor control centers, relay logic, PLC, VFD, HMI systems, digital thermostats, heating elements, and reading complex electrical schematics.',
            'Electrical & Control Systems:'
          ),
          createBullet(
            'Preventive maintenance scheduling, rapid on-site troubleshooting, spare parts inventory management, client relations, and strict compliance with EU-standard safety and hygiene protocols.',
            'Field & Workshop Operations:'
          ),

          // 3. PROFESSIONAL WORK EXPERIENCE
          createHeading('Experience'),
          createSubHeading(
            'Lead Field Technician – Refrigeration & Electrical Systems',
            'Independent / Contract Engineering Services | Manikgonj & Dhaka, Bangladesh',
            'January 2017 – Present (8+ Years)'
          ),
          createBullet('Spearhead on-site installation, commissioning, preventive maintenance, and emergency breakdown repairs of commercial catering, kitchen, and refrigeration systems across high-volume restaurants, bakeries, cafes, and food processing plants.'),
          createBullet('Perform comprehensive mechanical and electrical diagnostics on heavy-duty commercial kitchen equipment, including convection ovens, deck/rotary ovens, industrial fryers, griddles, grills, and salamanders.'),
          createBullet('Execute complete overhauls of commercial cooling units: walk-in chillers, cold storage rooms, reach-in freezers, and automated ice-making machines.'),
          createBullet('Conduct precise refrigerant recovery, system flushing, nitrogen pressure testing, vacuuming, and gas charging; systematically replace faulty compressors, condenser fan motors, filter driers, and expansion valves.'),
          createBullet('Diagnose and repair commercial dishwashing systems, including conveyor and under-counter models, handling solenoid valves, rinse pumps, heating elements, and automated chemical dosing dispensers.'),
          createBullet('Inspect and service commercial coffee brewing systems and espresso machines, troubleshooting water pressure irregularities, solenoid valves, heating elements, and scale buildup.'),
          createBullet('Wire and troubleshoot control circuits, contactors, overload relays, digital temperature controllers, timers, and safety sensors to ensure strict operational compliance.'),
          createBullet('Manage direct client communication on-site, estimate spare part requirements, advise kitchen management on equipment care, and maintain meticulous service logs.'),
          createBullet('Utilize service vehicle daily for rapid dispatch to client locations across the region.'),

          // Clean Page Break right here between Work Experience and Training!
          new Paragraph({
            children: [new PageBreak()],
          }),

          // ==================== PAGE 2 ====================

          // 4. INDUSTRIAL ATTACHMENT & SPECIALIZED TRAINING
          createHeading('Industrial Attachment & Specialized Training'),
          createSubHeading(
            'Industrial Automation & Electrical Controls Trainee',
            'SiliTek Bangladesh | Gazipur, Bangladesh',
            'August 2016 – November 2016 (12 Weeks Intensive)'
          ),
          createBullet('Completed intensive practical training focused on Industrial Automation systems, Relay Logic, Programmable Logic Controllers (PLC), Variable Frequency Drives (VFD), and Human Machine Interface (HMI).'),
          createBullet('Assembled, wired, and tested industrial control panels, motor control centers (MCC), and automated sequence wiring setups.'),
          createBullet('Successfully completed a comprehensive final practical project and technical viva, earning an "Excellent" grade.'),

          // 5. EDUCATION & ACADEMIC CREDENTIALS
          createHeading('Education'),
          createSubHeading(
            'Diploma-in-Engineering in Electrical Technology',
            'New Ideal Polytechnic Institute (Under BTEB, Dhaka)',
            'Passing Year: 2017 | Session: 2012–2013'
          ),
          createBullet('Final Semester GPA: 4.00 / 4.00 (A+)   |   Overall CGPA: 3.19 / 4.00'),
          createBullet('Relevant Coursework: Electrical Circuits, AC/DC Machines, Industrial Electronics, Power Systems, Electrical Wiring & Estimation, Instrumentation.'),
          
          createSubHeading(
            'Secondary School Certificate (SSC Vocational) – General Electronics',
            'Rupsha Wahed Ali High School (Under BTEB, Dhaka)',
            'Passing Year: 2012 | Session: 2010–2012'
          ),
          createBullet('Result: GPA 4.77 / 5.00 (First Division)'),
          createBullet('Specialized Trade: Basic Electronics, Circuit Soldering, Component Testing, Domestic Appliance Maintenance.'),

          // 6. PERSONAL & PASSPORT DETAILS
          createHeading('Personal & Passport Details (Official EU / Cyprus Verification)'),
          new Paragraph({ spacing: { after: 40 }, children: [] }),
          createTable([
            ['Field / Parameter', 'Official Candidate Detail'],
            ['Full Name', 'JUEL RANA'],
            ['Date of Birth', '10 May 1995'],
            ['Gender / Marital Status', 'Male'],
            ['Nationality', 'Bangladeshi'],
            ['Passport Number', 'A14239145 (Issued: 25 Feb 2024 | Expires: 24 Feb 2034)'],
            ['National ID (NID)', '8240329915'],
            ['Driving Licence', 'Category B (Light Vehicle) & Category A (Motorcycle) — Ready for Cyprus Conversion'],
            ["Parents' Names", 'Father: Afzal Hossain   |   Mother: Asma Begum'],
            ['Permanent Address', 'Village: Kastoshagra, P.O: Amdala-1850, P.S: Shibalaya, District: Manikganj, Bangladesh'],
          ]),

          // 7. LANGUAGES & SOFT SKILLS (EU CEFR Framework)
          createHeading('Languages & Soft Skills (CEFR Framework)'),
          createBullet('B2 — Professional Working Proficiency (Effective technical & verbal communication)', 'English:'),
          createBullet('C2 — Native / Bilingual', 'Bengali:'),
          createBullet('B1 — Conversational / Spoken', 'Hindi:'),
          createBullet('Advanced diagnostic problem-solving, high composure under pressure, customer-focused communication, strict safety adherence, and rapid adaptability to international work environments.', 'Core Professional Traits:'),

          // 8. DECLARATION
          createHeading('Declaration'),
          createNormalParagraph(
            'I hereby declare that all information provided in this document is true, complete, and authentic to the best of my knowledge and belief, and can be supported by official academic, professional, and legal documentation upon request.',
            true
          ),
          new Paragraph({
            spacing: { before: 100, after: 20 },
            children: [
              new TextRun({
                text: 'JUEL RANA',
                font: FONT_HEADING,
                size: 22,
                bold: true,
                color: COLOR_DARK,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 40 },
            children: [
              new TextRun({
                text: 'Date: 09 October 2026',
                font: FONT_BODY,
                size: 18.5,
                color: COLOR_MUTED,
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync('D:\\tryyyy\\JUEL_RANA_CV.docx', buffer);
  console.log('Successfully generated JUEL_RANA_CV.docx matching exact template header');
}

buildDoc().catch(err => {
  console.error(err);
  process.exit(1);
});
