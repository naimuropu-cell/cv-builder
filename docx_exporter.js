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
const COLOR_PRIMARY = '1E3A8A';
const COLOR_LINK = '2563EB';
const COLOR_DARK = '1F2937';
const COLOR_MUTED = '4B5563';

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

function createSectionHeading(title) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    border: {
      bottom: {
        color: COLOR_PRIMARY,
        space: 2,
        value: BorderStyle.SINGLE,
        size: 8,
      },
    },
    children: [
      new TextRun({
        text: title.toUpperCase(),
        font: FONT_MAIN,
        size: 20,
        bold: true,
        color: COLOR_PRIMARY,
      }),
    ],
  });
}

async function generateDocxBuffer(state) {
  const children = [];

  // 1. Candidate Name
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 40 },
      children: [
        new TextRun({
          text: (state.name || 'Candidate Name').toUpperCase(),
          font: FONT_MAIN,
          size: 30,
          bold: true,
          color: COLOR_PRIMARY,
        }),
      ],
    })
  );

  // 2. Designation
  if (state.title) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 70 },
        children: [
          new TextRun({
            text: state.title.toUpperCase(),
            font: FONT_MAIN,
            size: 20,
            bold: true,
            color: COLOR_MUTED,
          }),
        ],
      })
    );
  }

  // 3. Contact Details
  const contactRuns = [];
  if (state.location) {
    contactRuns.push(new TextRun({ text: state.location, font: FONT_MAIN, size: 19, color: COLOR_DARK }));
  }
  if (state.phone) {
    if (contactRuns.length) contactRuns.push(new TextRun({ text: '  •  ', font: FONT_MAIN, size: 19, color: COLOR_MUTED }));
    contactRuns.push(new TextRun({ text: state.phone, font: FONT_MAIN, size: 19, color: COLOR_DARK }));
  }
  if (state.email) {
    if (contactRuns.length) contactRuns.push(new TextRun({ text: '  •  ', font: FONT_MAIN, size: 19, color: COLOR_MUTED }));
    contactRuns.push(createLink(state.email, `mailto:${state.email}`));
  }
  if (state.portfolio) {
    if (contactRuns.length) contactRuns.push(new TextRun({ text: '  •  ', font: FONT_MAIN, size: 19, color: COLOR_MUTED }));
    contactRuns.push(createLink('Portfolio', state.portfolio));
  }
  if (state.github) {
    if (contactRuns.length) contactRuns.push(new TextRun({ text: '  •  ', font: FONT_MAIN, size: 19, color: COLOR_MUTED }));
    contactRuns.push(createLink('GitHub', state.github));
  }
  if (state.linkedin) {
    if (contactRuns.length) contactRuns.push(new TextRun({ text: '  •  ', font: FONT_MAIN, size: 19, color: COLOR_MUTED }));
    contactRuns.push(createLink('LinkedIn', state.linkedin));
  }

  if (contactRuns.length) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 120 },
        children: contactRuns,
      })
    );
  }

  // 4. Professional Summary
  if (state.summary) {
    children.push(createSectionHeading('Professional Summary'));
    children.push(
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 40, after: 90 },
        children: [
          new TextRun({
            text: state.summary,
            font: FONT_MAIN,
            size: 19,
            color: COLOR_DARK,
          }),
        ],
      })
    );
  }

  // 5. Technical Skills
  if (state.skills && state.skills.length) {
    children.push(createSectionHeading('Technical Skills'));
    state.skills.forEach(s => {
      children.push(
        new Paragraph({
          spacing: { before: 20, after: 30 },
          children: [
            new TextRun({
              text: `•  ${s.category}: `,
              font: FONT_MAIN,
              size: 19,
              bold: true,
              color: COLOR_DARK,
            }),
            new TextRun({
              text: s.items || '',
              font: FONT_MAIN,
              size: 19,
              color: COLOR_DARK,
            }),
          ],
        })
      );
    });
  }

  // 6. Professional Experience
  if (state.experience && state.experience.length) {
    children.push(createSectionHeading('Professional Experience'));
    state.experience.forEach(exp => {
      // Row 1: Role (Left) & Dates (Right)
      children.push(
        new Paragraph({
          spacing: { before: 70, after: 15 },
          tabStops: [{ type: 'right', position: 10800 }],
          children: [
            new TextRun({
              text: exp.role || 'Role',
              font: FONT_MAIN,
              size: 20,
              bold: true,
              color: COLOR_PRIMARY,
            }),
            new TextRun({
              text: `\t${exp.dates || ''}`,
              font: FONT_MAIN,
              size: 19,
              bold: true,
              color: COLOR_MUTED,
            }),
          ],
        })
      );

      // Row 2: Company
      if (exp.company) {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 35 },
            children: [
              new TextRun({
                text: exp.company,
                font: FONT_MAIN,
                size: 19,
                italics: true,
                color: COLOR_DARK,
              }),
            ],
          })
        );
      }

      // Bullets
      (exp.bullets || []).forEach(b => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            spacing: { before: 15, after: 20 },
            children: [
              new TextRun({
                text: b,
                font: FONT_MAIN,
                size: 19,
                color: COLOR_DARK,
              }),
            ],
          })
        );
      });
    });
  }

  // 7. Featured Projects
  if (state.projects && state.projects.length) {
    children.push(createSectionHeading('Featured Projects'));
    state.projects.forEach(p => {
      const titleRuns = [];
      if (p.url) {
        titleRuns.push(
          new ExternalHyperlink({
            children: [
              new TextRun({
                text: `${p.name} ↗`,
                font: FONT_MAIN,
                size: 20,
                bold: true,
                color: COLOR_PRIMARY,
                underline: { type: UnderlineType.SINGLE },
              }),
            ],
            link: p.url,
          })
        );
      } else {
        titleRuns.push(
          new TextRun({
            text: p.name || 'Project',
            font: FONT_MAIN,
            size: 20,
            bold: true,
            color: COLOR_PRIMARY,
          })
        );
      }

      if (p.date) {
        titleRuns.push(
          new TextRun({
            text: `\t${p.date}`,
            font: FONT_MAIN,
            size: 19,
            bold: true,
            color: COLOR_MUTED,
          })
        );
      }

      children.push(
        new Paragraph({
          spacing: { before: 70, after: 20 },
          tabStops: [{ type: 'right', position: 10800 }],
          children: titleRuns,
        })
      );

      if (p.tags && p.tags.length) {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 30 },
            children: [
              new TextRun({
                text: `Key Highlights: ${p.tags.join(' | ')}`,
                font: FONT_MAIN,
                size: 18,
                italics: true,
                color: COLOR_MUTED,
              }),
            ],
          })
        );
      }

      (p.bullets || []).forEach(b => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            spacing: { before: 15, after: 20 },
            children: [
              new TextRun({
                text: b,
                font: FONT_MAIN,
                size: 19,
                color: COLOR_DARK,
              }),
            ],
          })
        );
      });
    });
  }

  // 8. Education
  if (state.education && state.education.length) {
    children.push(createSectionHeading('Education & Credentials'));
    state.education.forEach(edu => {
      children.push(
        new Paragraph({
          spacing: { before: 50, after: 15 },
          tabStops: [{ type: 'right', position: 10800 }],
          children: [
            new TextRun({
              text: edu.degree || 'Degree',
              font: FONT_MAIN,
              size: 20,
              bold: true,
              color: COLOR_PRIMARY,
            }),
            new TextRun({
              text: `\t${edu.date || ''}`,
              font: FONT_MAIN,
              size: 19,
              bold: true,
              color: COLOR_MUTED,
            }),
          ],
        })
      );

      if (edu.institution) {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 20 },
            children: [
              new TextRun({
                text: edu.institution,
                font: FONT_MAIN,
                size: 19,
                italics: true,
                color: COLOR_DARK,
              }),
            ],
          })
        );
      }

      if (edu.details) {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 40 },
            children: [
              new TextRun({
                text: edu.details,
                font: FONT_MAIN,
                size: 18,
                color: COLOR_MUTED,
              }),
            ],
          })
        );
      }
    });
  }

  // 9. Languages
  if (state.languages) {
    children.push(createSectionHeading('Languages'));
    children.push(
      new Paragraph({
        spacing: { before: 30, after: 60 },
        children: [
          new TextRun({
            text: state.languages,
            font: FONT_MAIN,
            size: 19,
            color: COLOR_DARK,
          }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.5),
              bottom: convertInchesToTwip(0.5),
              left: convertInchesToTwip(0.5),
              right: convertInchesToTwip(0.5),
            },
          },
        },
        children: children,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}

module.exports = { generateDocxBuffer };
