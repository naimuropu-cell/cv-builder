# NovaCV Studio 🚀
### Interactive ATS & Multi-Format Resume Builder

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Vanilla JS](https://img.shields.io/badge/JavaScript-ES6+-yellow.svg)](builder.js)
[![HTML5 & CSS3](https://img.shields.io/badge/Style-Vanilla%20CSS3-orange.svg)](builder.css)
[![ATS Optimized](https://img.shields.io/badge/ATS-Optimized-brightgreen.svg)](#-live-ats-optimization-engine)

**NovaCV Studio** is a responsive, real-time resume engineering studio built with Vanilla JavaScript, HTML5, and CSS3. It combines live split-screen editing, pixel-perfect A4 canvas preview, dynamic section repeaters, an automated ATS readiness scoring engine, and multi-format export capabilities (PDF & JSON).

---

## 🌟 Key Highlights

- **⚡ Split-Screen Reactive Studio:** Live two-way binding updates the right-hand A4 sheet instantaneously on every keystroke.
- **🎯 3 Interchangeable Architectures:**
  - **ATS Tech Minimalist (Default):** Single-column, keyword-dense, link-enabled layout engineered for automated applicant tracking systems (ideal for Software QA, Developers, and Tech).
  - **Modern Timeline:** Vertical accent rail with circular photo avatar and two-tone headers.
  - **Classic Executive:** Formal layout featuring rectangular portrait headshot, distinct section dividers, and traditional structure.
- **📊 Real-Time ATS Optimization Meter (0–100%):** Automatically audits candidate info, summary length, quantified metrics in job bullets, skill categories, and proof-of-work links with actionable advice.
- **🔁 Dynamic Field Repeaters:** Add, edit, or remove work experience, featured projects, skill groups, and education records dynamically.
- **🎨 Visual Design Customizer:** 7 curated color palettes, Google Fonts switcher (`Inter`, `Outfit`, `Roboto`, `Times New Roman`), photo upload, and canvas zoom controls (70% to 130%).
- **💾 JSON State Management:** 1-click **Export JSON** to backup your resume state and **Import JSON** to load or switch profiles instantly.
- **📄 Pixel-Perfect PDF Generation:** Dedicated print stylesheets hide editor controls and produce border-accurate A4 documents directly via the browser print dialog.
- **👤 Preloaded Profiles:**
  - **Md. Naimur Rahman Apu** — Software Quality Assurance (SQA) Engineer
  - **Juel Rana** — Refrigeration & Electrical Systems Technician
  - **Empty Template** — Blank canvas for building custom resumes from scratch.

---

## 📸 Preview & Architecture

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                  NovaCV Studio Navbar                                   │
│  [Logo] NovaCV Studio   |   Preset: [Md. Naimur Rahman Apu ▾]   |   [Import] [Export JSON] [Export PDF]  │
├─────────────────────────────────────────┬──────────────────────────────────────────────┤
│          LEFT EDITOR STUDIO             │             RIGHT LIVE PREVIEW               │
│                                         │                                              │
│  Tabs:                                  │  [Toolbar:  −  100%  ＋  Reset]              │
│  🎨 Template & Style                    │                                              │
│  👤 Personal Info                       │  ┌────────────────────────────────────────┐  │
│  📝 Summary & Languages                 │  │             A4 LIVE CANVAS             │  │
│  ⚡ Technical Skills (Repeater)          │  │                                        │  │
│  💼 Experience (Repeater)               │  │  Md. Naimur Rahman Apu                 │  │
│  🚀 Featured Projects (Repeater)        │  │  Software Quality Assurance Engineer   │  │
│  🎓 Education Records (Repeater)        │  │  ------------------------------------  │  │
│  📊 Real-Time ATS Score Meter           │  │  • Professional Experience              │  │
│                                         │  │  • Featured QA Projects                 │  │
│  • Instant 2-way data binding           │  │  • Technical Skills                    │  │
│  • 1-click Delete / Add cards           │  │                                        │  │
│  • Photo upload with base64 preview     │  │  (Updates live on every keystroke)     │  │
│                                         │  └────────────────────────────────────────┘  │
└─────────────────────────────────────────┴──────────────────────────────────────────────┘
```

---

## 🛠️ Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/naimuropu-cell/cv-builder.git
cd cv-builder
```

### 2. Run Locally

#### Option A: Node.js Dev Server (Recommended)
```bash
npm start
# or
node server.js
```
The server will start at `http://localhost:3000/` and automatically launch in your default web browser.

#### Option B: Direct Browser Launch
Open `builder.html` directly in any web browser (Google Chrome, Microsoft Edge, Firefox, Brave):
```bash
start builder.html
```

---

## 📊 Live ATS Optimization Engine

The built-in ATS validator analyzes your resume in real time and scores it across 5 core dimensions:

1. **Contact Completeness (25 pts):** Name, target title, verified phone, email, and location.
2. **Summary Quality (15 pts):** Comprehensive summary (>100 characters) containing role-relevant technical keywords.
3. **Quantified Impact (30 pts):** Multiple roles with metric-backed accomplishments (percentages, numbers, test case volumes).
4. **Skills & Taxonomy (20 pts):** Categorized domain proficiencies and degree credentials.
5. **Proof-of-Work Links (10 pts):** Clickable links to GitHub repositories, live portfolios, and LinkedIn.

---

## 📁 Project Structure

```text
cv-builder/
├── builder.html               # Main application interface & split-screen layout
├── builder.css                # Studio design system, responsive tabs, and print media
├── builder.js                 # Reactive state engine, templates, repeaters & ATS scorer
├── server.js                  # Lightweight zero-dependency Node.js static server
├── package.json               # Project manifest & dev scripts
├── .gitignore                 # Node modules & temporary files exclusion
├── README.md                  # Comprehensive project documentation
└── templates & deliverables/
    ├── Md_Naimur_Rahman_Apu_SQA_Resume.pdf  # Finalized 2-page SQA PDF
    ├── Md_Naimur_Rahman_Apu_SQA_Resume.docx # Finalized SQA Word document
    ├── naimur_sqa_resume.html               # Standalone interactive SQA web resume
    └── generate_naimur_docx.js              # Word docx generation script
```

---

## 🤝 Contributing

Contributions, feedback, and template suggestions are welcome!
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 👤 Author

**Md. Naimur Rahman Apu**
- **GitHub:** [@naimuropu-cell](https://github.com/naimuropu-cell)
- **LinkedIn:** [naimur-rahman-apu](https://www.linkedin.com/in/naimur-rahman-apu/)
- **Portfolio:** [naimuropu-cell.github.io](https://naimuropu-cell.github.io/)
