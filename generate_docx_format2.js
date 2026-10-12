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
  UnderlineType,
  convertInchesToTwip
} = require('docx');

const FONT_PRIMARY = 'Times New Roman';
const COLOR_BLACK = '000000';
const COLOR_DARK = '222222';

// Clean Section Heading with Underline like in template
function createSectionTitle(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: text.toUpperCase(),
        font: FONT_PRIMARY,
        size: 24, // 12pt
        bold: true,
        underline: { type: UnderlineType.SINGLE },
        color: COLOR_BLACK,
      }),
    ],
  });
}

function createLeftSectionTitle(text) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [
      new TextRun({
        text: text,
        font: FONT_PRIMARY,
        size: 23,
        bold: true,
        underline: { type: UnderlineType.SINGLE },
        color: COLOR_BLACK,
      }),
    ],
  });
}

function createBullet(text, boldPrefix = '') {
  const children = [];
  if (boldPrefix) {
    children.push(
      new TextRun({
        text: boldPrefix + ' ',
        font: FONT_PRIMARY,
        size: 21,
        bold: true,
        color: COLOR_BLACK,
      })
    );
  }
  children.push(
    new TextRun({
      text: text,
      font: FONT_PRIMARY,
      size: 21,
      color: COLOR_DARK,
    })
  );

  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 25, after: 25 },
    children: children,
  });
}

function createNormalParagraph(text, isJustified = true) {
  return new Paragraph({
    spacing: { before: 40, after: 80 },
    alignment: isJustified ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    children: [
      new TextRun({
        text: text,
        font: FONT_PRIMARY,
        size: 21,
        color: COLOR_BLACK,
      }),
    ],
  });
}

// Two-column key-value row for Personal Details & Scholastic Records
function createKeyValueTable(rowsData, col1Width = 3200, col2Width = 6100) {
  const tableRows = rowsData.map(row => {
    return new TableRow({
      children: [
        new TableCell({
          width: { size: col1Width, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 60, right: 60 },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: row[0],
                  font: FONT_PRIMARY,
                  size: 21,
                  bold: true,
                  color: COLOR_BLACK,
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          width: { size: col2Width, type: WidthType.DXA },
          margins: { top: 40, bottom: 40, left: 60, right: 60 },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: ': ' + row[1],
                  font: FONT_PRIMARY,
                  size: 21,
                  color: COLOR_DARK,
                }),
              ],
            }),
          ],
        }),
      ],
    });
  });

  return new Table({
    width: { size: col1Width + col2Width, type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: tableRows,
  });
}

async function buildDoc() {
  const photoPath = path.join(__dirname, 'juel_passport_rect.jpg');
  const photoData = fs.readFileSync(photoPath);

  // Top Header Table:
  // Left: Resume Of JUEL RANA + Contacts
  // Right: Passport Photo Box
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
          // Left: Candidate Info
          new TableCell({
            width: { size: 7300, type: WidthType.DXA },
            margins: { top: 0, bottom: 100, left: 0, right: 100 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 20 },
                children: [
                  new TextRun({
                    text: 'Resume Of',
                    font: FONT_PRIMARY,
                    size: 24, // 12pt
                    bold: true,
                    underline: { type: UnderlineType.SINGLE },
                    color: COLOR_BLACK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                  new TextRun({
                    text: 'JUEL RANA',
                    font: FONT_PRIMARY,
                    size: 34, // 17pt
                    bold: true,
                    color: COLOR_BLACK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 20 },
                children: [
                  new TextRun({
                    text: 'Refrigeration & Electrical Systems Technician',
                    font: FONT_PRIMARY,
                    size: 22,
                    bold: true,
                    color: COLOR_DARK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({ text: 'Address : ', font: FONT_PRIMARY, size: 20, bold: true }),
                  new TextRun({ text: 'Kastoshagra, Amdala-1850, Shibalaya, Manikgonj, Bangladesh', font: FONT_PRIMARY, size: 20 }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({ text: 'Mobile    : ', font: FONT_PRIMARY, size: 20, bold: true }),
                  new TextRun({ text: '+880 1680844360 (Phone & WhatsApp)', font: FONT_PRIMARY, size: 20 }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({ text: 'e-mail     : ', font: FONT_PRIMARY, size: 20, bold: true }),
                  new TextRun({ text: 'juelranabd64@gmail.com', font: FONT_PRIMARY, size: 20 }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                  new TextRun({ text: 'Licence   : ', font: FONT_PRIMARY, size: 20, bold: true }),
                  new TextRun({ text: 'Category B & A (Valid BD Licence — Ready for Cyprus Conversion)', font: FONT_PRIMARY, size: 20 }),
                ],
              }),
            ],
          }),
          // Right: Passport Photo Box
          new TableCell({
            width: { size: 2000, type: WidthType.DXA },
            margins: { top: 0, bottom: 100, left: 100, right: 0 },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new ImageRun({
                    data: photoData,
                    transformation: { width: 95, height: 125 }, // 3:4 passport aspect
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });

  // Reference Table: 2 columns
  const referenceTable = new Table({
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
          // Reference 1
          new TableCell({
            width: { size: 4650, type: WidthType.DXA },
            margins: { top: 40, bottom: 40, left: 0, right: 100 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 20 },
                children: [
                  new TextRun({
                    text: 'Engr. Dewan Md. Gias Mahmud',
                    font: FONT_PRIMARY,
                    size: 21,
                    bold: true,
                    color: COLOR_BLACK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({
                    text: 'Sub-Divisional Engineer & In-charge\nManikganj Grid Sub-Station\nPower Grid Company Of Bangladesh (PGCB)\nCell: 01728786917\nAddress: 132/33 kV Manikganj Grid Sub-Station',
                    font: FONT_PRIMARY,
                    size: 20,
                  }),
                ],
              }),
            ],
          }),
          // Reference 2
          new TableCell({
            width: { size: 4650, type: WidthType.DXA },
            margins: { top: 40, bottom: 40, left: 100, right: 0 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 20 },
                children: [
                  new TextRun({
                    text: 'Md. Abbas Uddin',
                    font: FONT_PRIMARY,
                    size: 21,
                    bold: true,
                    color: COLOR_BLACK,
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 15 },
                children: [
                  new TextRun({
                    text: 'Ass. Engineer\nManikganj Grid Sub-Station\nPower Grid Company Of Bangladesh (PGCB)\nCell: 01718828192\nAddress: 132/33 kV Manikganj Grid Sub-Station',
                    font: FONT_PRIMARY,
                    size: 20,
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
              top: convertInchesToTwip(0.55),
              bottom: convertInchesToTwip(0.55),
              left: convertInchesToTwip(0.65),
              right: convertInchesToTwip(0.65),
            },
          },
        },
        children: [
          headerTable,

          // 1. CAREER ASPIRATION
          createSectionTitle('Career Aspiration'),
          createNormalParagraph(
            'Dedicated, energetic, and technically proficient Refrigeration & Electrical Systems Technician with over 8 years of hands-on experience (2017 – Present) in installation, preventive maintenance, troubleshooting, and repair of commercial catering equipment and commercial cooling systems. Proven technical capability across commercial ovens, deep fryers, cooking ranges, commercial refrigerators, reach-in freezers, walk-in cold rooms, automated ice makers, industrial dishwashers, and espresso coffee machines. Holds a 4-year Diploma in Electrical Engineering and intensive industrial automation training (PLC, VFD, HMI). Prepared to accept challenging responsibilities in a dynamic international organization, demonstrate high capability, commitment, and dedication, and fully ready for immediate relocation and deployment to Cyprus.'
          ),

          // 2. MAJOR SKILLS LEARNED / CORE COMPETENCIES
          createSectionTitle('Major Skills & Core Technical Competencies'),
          createBullet('Commercial convection ovens, deck ovens, rotary ovens, cooking ranges, heavy-duty deep fryers, flat-top grills, salamanders, and warming cabinets.', 'Commercial Catering Equipment:'),
          createBullet('Commercial display refrigerators, reach-in freezers, walk-in cold rooms, automated ice machines. Proficient in leak detection, vacuuming, nitrogen pressure testing, gas charging (R134a, R404a, R410a), and compressor overhaul.', 'Refrigeration & HVAC Systems:'),
          createBullet('Servicing, descaling, pump repair for commercial espresso machines, conveyor & under-counter dishwashers, and water filtration systems.', 'Beverage & Sanitation Equipment:'),
          createBullet('3-phase & single-phase electrical wiring, motor control centers (MCC), relay logic, PLC, VFD, HMI systems, digital temperature controllers, and circuit schematics.', 'Electrical & Control Systems:'),
          createBullet('Microsoft Windows (XP, 7, 8, 10, 11), MS Office (Word, Excel), Technical Documentation & Service Logs.', 'Computer & IT Skills:'),

          // 3. WORK EXPERIENCE
          createSectionTitle('Professional Work Experience'),
          new Paragraph({
            spacing: { before: 80, after: 30 },
            children: [
              new TextRun({ text: 'Position                     : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'Lead Field Technician – Refrigeration & Electrical Systems', font: FONT_PRIMARY, size: 21, bold: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 30 },
            children: [
              new TextRun({ text: 'Organization             : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'Independent / Contract Engineering Services', font: FONT_PRIMARY, size: 21 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 30 },
            children: [
              new TextRun({ text: 'Location                    : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'Manikgonj & Dhaka, Bangladesh', font: FONT_PRIMARY, size: 21 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 60 },
            children: [
              new TextRun({ text: 'Duration                    : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'January 2017 – Present (8+ Years)', font: FONT_PRIMARY, size: 21, bold: true }),
            ],
          }),
          createBullet('Spearhead on-site installation, commissioning, preventive maintenance, and emergency breakdown repairs of commercial catering, kitchen, and refrigeration systems across high-volume restaurants, bakeries, cafes, and food processing plants.'),
          createBullet('Perform comprehensive mechanical and electrical diagnostics on heavy-duty commercial kitchen equipment: convection ovens, deck ovens, fryers, griddles, and salamanders.'),
          createBullet('Execute complete overhauls of commercial cooling units: walk-in chillers, cold storage rooms, reach-in freezers, and automated ice machines.'),
          createBullet('Conduct precise refrigerant recovery, system flushing, nitrogen pressure testing, vacuuming, and gas charging; systematically replace compressors, fan motors, filter driers, and expansion valves.'),
          createBullet('Diagnose and repair commercial dishwashing systems (conveyor and under-counter models), solenoid valves, rinse pumps, heating elements, and automated dosing systems.'),
          createBullet('Inspect and service commercial espresso machines, troubleshooting water pressure, solenoid valves, heating elements, and scale buildup.'),
          createBullet('Wire and troubleshoot control circuits, contactors, overload relays, digital temperature controllers, and safety interlocks.'),
          createBullet('Manage client communication on-site, estimate spare parts, advise kitchen management on equipment care, and maintain meticulous service logs.'),
          createBullet('Utilize service vehicle daily for rapid dispatch to client locations across the region.'),

          // 4. INDUSTRIAL ATTACHMENT & SPECIALIZED TRAINING
          createSectionTitle('Industrial Attachment & Specialized Training'),
          new Paragraph({
            spacing: { before: 80, after: 30 },
            children: [
              new TextRun({ text: 'Training Position     : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'Industrial Automation & Electrical Controls Trainee', font: FONT_PRIMARY, size: 21, bold: true }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 30 },
            children: [
              new TextRun({ text: 'Company Name         : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'SiliTek Bangladesh, Gazipur', font: FONT_PRIMARY, size: 21 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 30 },
            children: [
              new TextRun({ text: 'Department              : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: 'Electrical & Industrial Automation', font: FONT_PRIMARY, size: 21 }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 60 },
            children: [
              new TextRun({ text: 'Area of Experience : ', font: FONT_PRIMARY, size: 21, bold: true }),
              new TextRun({ text: '12-Week Intensive Practical Training (August 2016 – November 2016)', font: FONT_PRIMARY, size: 21 }),
            ],
          }),
          createBullet('Completed intensive practical training on Industrial Automation systems, Relay Logic, Programmable Logic Controllers (PLC), Variable Frequency Drives (VFD), and Human Machine Interface (HMI).'),
          createBullet('Assembled, wired, and tested industrial control panels, motor control centers (MCC), and automated sequence wiring setups.'),
          createBullet('Successfully completed comprehensive final practical project and technical viva, earning an "Excellent" grade.'),

          // 5. SCHOLASTIC RECORDS
          createSectionTitle('Scholastic Records'),
          createLeftSectionTitle('1. Diploma in Engineering:'),
          createKeyValueTable([
            ['Achievement', 'Diploma in Engineering Certificate'],
            ['Institute', 'New Ideal Polytechnic Institute, Manikganj'],
            ['Board', 'Bangladesh Technical Education Board (BTEB)'],
            ['Group / Technology', 'Electrical Technology'],
            ['Result', 'CGPA 3.19 (out of 4.00) | Final Semester GPA: 4.00 / 4.00 (A+)'],
            ['Passing Year', '2017 (Session: 2012–2013)'],
          ], 2600, 6700),

          createLeftSectionTitle('2. Secondary School Certificate (SSC):'),
          createKeyValueTable([
            ['Achievement', 'Secondary School Certificate (Vocational)'],
            ['Institute', 'Rupsha Wahed Ali High School, Manikganj'],
            ['Board', 'Bangladesh Technical Education Board (BTEB)'],
            ['Group / Trade', 'General Electronics'],
            ['Result', 'GPA 4.77 (out of 5.00) First Division'],
            ['Passing Year', '2012 (Session: 2010–2012)'],
          ], 2600, 6700),

          // 6. LANGUAGE SKILLS
          createSectionTitle('Language Skills'),
          createBullet('Can read, write, speak and understand (Professional Working Proficiency / CEFR B2)', 'English:'),
          createBullet('Can read, write, speak and understand (Native Language / CEFR C2)', 'Bengali:'),
          createBullet('Can understand and speak (Conversational / CEFR B1)', 'Hindi:'),

          // 7. PERSONAL DETAILS
          createSectionTitle('Personal Details'),
          createKeyValueTable([
            ['Name', 'Juel Rana'],
            ["Father's Name", 'Afzal Hossain (Afjal Hosen)'],
            ["Mother's Name", 'Asma Begum'],
            ['Present Address', 'Village: Kastoshagra, Post office: Amdala-1850, Upazila: Shibalaya, District: Manikgonj.'],
            ['Permanent Address', 'Village: Kastoshagra, Post office: Amdala-1850, Upazila: Shibalaya, District: Manikgonj.'],
            ['Passport Number', 'A14239145 (Date of Issue: 25 Feb 2024 | Date of Expiry: 24 Feb 2034)'],
            ['National ID (NID) No', '8240329915'],
            ['Date of Birth', '10-05-1995 (10 May 1995)'],
            ['Sex / Gender', 'Male'],
            ['Marital Status', 'Married'],
            ['Nationality', 'Bangladeshi (by birth)'],
            ['Religion', 'Islam'],
            ['Driving Licence', 'Category B (Light Vehicle) & Category A (Motorcycle) — Ready for Cyprus Conversion'],
          ], 2600, 6700),

          // 8. REFERENCES
          createSectionTitle('References'),
          referenceTable,

          // 9. DECLARATION (No handwritten sign line as requested)
          createSectionTitle('Declaration'),
          createNormalParagraph(
            'I hereby declare that the information supplied by me on this curriculum vitae is complete, true, and correct in every particular. I acknowledge that the authority may terminate the employment if I have misrepresented my past or present circumstances, and all information can be supported by authentic official documentation upon request.'
          ),
          new Paragraph({
            spacing: { before: 80, after: 20 },
            children: [
              new TextRun({
                text: 'JUEL RANA',
                font: FONT_PRIMARY,
                size: 22,
                bold: true,
                color: COLOR_BLACK,
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 0, after: 60 },
            children: [
              new TextRun({
                text: 'Date: 09 October 2026',
                font: FONT_PRIMARY,
                size: 20,
                color: COLOR_DARK,
              }),
            ],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync('D:\\tryyyy\\JUEL_RANA_CV_2.docx', buffer);
  console.log('Successfully generated JUEL_RANA_CV_2.docx in classic format 2');
}

buildDoc().catch(err => {
  console.error(err);
  process.exit(1);
});
