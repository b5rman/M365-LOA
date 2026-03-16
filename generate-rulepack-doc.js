const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, ExternalHyperlink, TableOfContents,
  LevelFormat
} = require("docx");

// ── Load rule pack ──
const rp = JSON.parse(fs.readFileSync("./LOA_RulePack_M365.json", "utf8"));
const rules = rp.rules;
const docs = rp.docs;

// Build doc lookup
const docLookup = {};
for (const d of docs) docLookup[d.id] = d;

// ── Category ordering & display config ──
const categoryConfig = [
  { key: "Waste",           color: "C0392B", icon: "Waste",               desc: "Direct cost savings from unused, dormant, or redundant licenses that can be reclaimed immediately." },
  { key: "Right-Sizing",    color: "2980B9", icon: "Right-Sizing",        desc: "Optimization opportunities where users can be moved to more appropriate (typically lower-cost) license tiers based on actual usage patterns." },
  { key: "Licensing Check",  color: "E67E22", icon: "Licensing Compliance", desc: "Gaps where features are in use but the required licensing entitlement is missing, creating compliance exposure during Microsoft audits." },
  { key: "Noncompliant",    color: "8E44AD", icon: "Noncompliant",        desc: "Active licensing violations where the current configuration does not meet Microsoft licensing requirements." },
  { key: "Security",        color: "27AE60", icon: "Security",            desc: "Security posture findings related to admin accounts, identity protection, and threat defense coverage." },
  { key: "Governance",      color: "34495E", icon: "Governance",          desc: "Operational hygiene items including account lifecycle, data governance, and license assignment best practices." },
  { key: "Copilot",         color: "3498DB", icon: "Copilot & AI",        desc: "Microsoft 365 Copilot and AI add-on adoption, overlap detection, and prerequisite validation." },
];

// Group rules by category
const grouped = {};
for (const cat of categoryConfig) grouped[cat.key] = [];
for (const r of rules) {
  if (grouped[r.category]) grouped[r.category].push(r);
  else {
    console.warn(`Unknown category: ${r.category} (rule ${r.id})`);
    if (!grouped[r.category]) grouped[r.category] = [];
    grouped[r.category].push(r);
  }
}

// ── Styling constants ──
const BRAND_BLUE = "1B4F72";
const BRAND_ACCENT = "2E86C1";
const LIGHT_BLUE = "D6EAF8";
const LIGHT_GRAY = "F2F3F4";
const WHITE = "FFFFFF";
const PAGE_WIDTH = 12240; // US Letter
const CONTENT_WIDTH = 9360; // 1" margins
const border = { style: BorderStyle.SINGLE, size: 1, color: "BDC3C7" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

// ── Confidence badge color ──
function confidenceColor(level) {
  switch (level) {
    case "High": return "27AE60";
    case "Medium": return "F39C12";
    case "Low": return "95A5A6";
    default: return "BDC3C7";
  }
}

// ── Build rule table for one rule ──
function buildRuleBlock(rule) {
  const elements = [];
  const conf = rule.confidence ? rule.confidence.base : "Medium";
  const confColor = confidenceColor(conf);
  const action = rule.recommendation ? rule.recommendation.action : "Review";
  const scope = rule.scope || "User";
  const autoLevel = rule.automationLevel || "AUTO";

  // Rule title with ID badge
  elements.push(new Paragraph({
    spacing: { before: 240, after: 80 },
    children: [
      new TextRun({ text: rule.id, bold: true, font: "Consolas", size: 18, color: WHITE,
        shading: { type: ShadingType.CLEAR, fill: BRAND_ACCENT, color: BRAND_ACCENT } }),
      new TextRun({ text: "  " }),
      new TextRun({ text: rule.title, bold: true, size: 22, font: "Arial", color: "2C3E50" }),
    ]
  }));

  // Metadata row as compact table
  const metaRow = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [2340, 2340, 2340, 2340],
    rows: [
      new TableRow({
        children: [
          metaCell("Scope", scope, LIGHT_BLUE),
          metaCell("Action", action, LIGHT_BLUE),
          metaCell("Confidence", conf, confColor, true),
          metaCell("Detection", autoLevel === "AUTO" ? "Automated" : autoLevel === "SEMI" ? "Semi-Auto" : "Manual", LIGHT_BLUE),
        ]
      })
    ]
  });
  elements.push(metaRow);

  // Recommendation text
  if (rule.recommendation && rule.recommendation.textTemplate) {
    elements.push(new Paragraph({
      spacing: { before: 100, after: 40 },
      children: [
        new TextRun({ text: "Recommendation: ", bold: true, size: 20, font: "Arial", color: "2C3E50" }),
        new TextRun({ text: rule.recommendation.textTemplate.replace(/\{[^}]+\}/g, "[...]"), size: 20, font: "Arial", color: "555555" }),
      ]
    }));
  }

  // Notes
  if (rule.recommendation && rule.recommendation.notes && rule.recommendation.notes.length > 0) {
    for (const note of rule.recommendation.notes) {
      elements.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        indent: { left: 360 },
        children: [
          new TextRun({ text: "Note: ", bold: true, italics: true, size: 18, font: "Arial", color: "7F8C8D" }),
          new TextRun({ text: note, italics: true, size: 18, font: "Arial", color: "7F8C8D" }),
        ]
      }));
    }
  }

  // Manual check steps
  if (rule.manualCheck && rule.manualCheck.length > 0) {
    elements.push(new Paragraph({
      spacing: { before: 80, after: 40 },
      children: [
        new TextRun({ text: "Manual Verification Steps:", bold: true, size: 20, font: "Arial", color: "2C3E50" }),
      ]
    }));
    for (const step of rule.manualCheck) {
      elements.push(new Paragraph({
        spacing: { before: 20, after: 20 },
        indent: { left: 360 },
        numbering: { reference: "manual-steps", level: 0 },
        children: [
          new TextRun({ text: step, size: 19, font: "Arial", color: "555555" }),
        ]
      }));
    }
  }

  // References
  if (rule.references && rule.references.length > 0) {
    const refChildren = [
      new TextRun({ text: "References: ", bold: true, size: 18, font: "Arial", color: "7F8C8D" }),
    ];
    for (let i = 0; i < rule.references.length; i++) {
      const ref = rule.references[i];
      const doc = docLookup[ref];
      if (doc) {
        refChildren.push(new ExternalHyperlink({
          children: [new TextRun({ text: doc.title, style: "Hyperlink", size: 18, font: "Arial" })],
          link: doc.url,
        }));
      } else {
        refChildren.push(new TextRun({ text: ref, size: 18, font: "Arial", color: "7F8C8D" }));
      }
      if (i < rule.references.length - 1) {
        refChildren.push(new TextRun({ text: "  |  ", size: 18, font: "Arial", color: "BDC3C7" }));
      }
    }
    elements.push(new Paragraph({ spacing: { before: 60, after: 40 }, children: refChildren }));
  }

  // Separator line
  elements.push(new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 1, color: "D5D8DC", space: 1 } },
    children: []
  }));

  return elements;
}

function metaCell(label, value, fill, isConfidence = false) {
  const textColor = isConfidence ? WHITE : "2C3E50";
  const cellFill = isConfidence ? fill : LIGHT_GRAY;
  return new TableCell({
    borders,
    width: { size: 2340, type: WidthType.DXA },
    shading: { fill: cellFill, type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: label + ": ", bold: true, size: 16, font: "Arial", color: isConfidence ? WHITE : "7F8C8D" }),
        new TextRun({ text: value, bold: true, size: 17, font: "Arial", color: textColor }),
      ]
    })]
  });
}

// ── Build document sections ──
const sections = [];

// ── Title Page ──
sections.push({
  properties: {
    page: {
      size: { width: PAGE_WIDTH, height: 15840 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    }
  },
  children: [
    new Paragraph({ spacing: { before: 3600 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "EASI", bold: true, size: 56, font: "Arial", color: BRAND_BLUE })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: "Microsoft 365", size: 48, font: "Arial", color: BRAND_ACCENT })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: "License Optimization Audit", bold: true, size: 44, font: "Arial", color: "2C3E50" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 600 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND_ACCENT, space: 1 } },
      children: [new TextRun({ text: "Detection Scenario Reference Guide", size: 28, font: "Arial", color: "7F8C8D" })]
    }),
    new Paragraph({ spacing: { before: 800 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [
        new TextRun({ text: `Version ${rp.version}`, size: 24, font: "Arial", color: "7F8C8D" }),
        new TextRun({ text: `    |    ${rules.length} Detection Rules    |    ${categoryConfig.length} Categories`, size: 24, font: "Arial", color: "7F8C8D" }),
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: `Generated: ${new Date().toISOString().split("T")[0]}`, size: 22, font: "Arial", color: "95A5A6" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 1200 },
      children: [new TextRun({ text: "CONFIDENTIAL", bold: true, size: 20, font: "Arial", color: "C0392B" })]
    }),
  ]
});

// ── TOC + Overview Section ──
const overviewChildren = [
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { after: 200 },
    children: [new TextRun({ text: "Table of Contents", bold: true, size: 32, font: "Arial", color: BRAND_BLUE })]
  }),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: [new PageBreak()] }),

  // About section
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { after: 200 },
    children: [new TextRun({ text: "About This Document", bold: true, size: 32, font: "Arial", color: BRAND_BLUE })]
  }),
  new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({
      text: "This document describes every detection scenario evaluated by the EASI Microsoft 365 License Optimization Audit (LOA). Each rule represents a specific licensing pattern that the audit tool identifies and reports on, enabling IT administrators and licensing specialists to take targeted action.",
      size: 22, font: "Arial", color: "2C3E50"
    })]
  }),
  new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({
      text: "Rules are organized into seven categories. Each rule includes the detection logic scope, recommended action, confidence level, and links to relevant Microsoft documentation.",
      size: 22, font: "Arial", color: "2C3E50"
    })]
  }),

  // Rules vs Messages explanation
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text: "Rules vs. Recommendation Messages", bold: true, size: 26, font: "Arial", color: BRAND_ACCENT })]
  }),
  new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({
      text: "The audit engine evaluates 146 distinct detection paths, but many of those are context-specific variants of the same logical scenario. This document groups related variants into ",
      size: 22, font: "Arial", color: "2C3E50"
    }),
    new TextRun({ text: `${rules.length} rules`, bold: true, size: 22, font: "Arial", color: "2C3E50" }),
    new TextRun({
      text: " — one per scenario — so each finding is explained once rather than repeating near-identical entries.",
      size: 22, font: "Arial", color: "2C3E50"
    })]
  }),
  new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({
      text: "Examples of how variants map to a single rule:",
      bold: true, size: 21, font: "Arial", color: "2C3E50"
    })]
  }),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [3200, 900, 5260],
    rows: [
      new TableRow({ children: [
        headerCell("Rule", 3200),
        headerCell("Paths", 900),
        headerCell("Variant Examples", 5260),
      ]}),
      ...([
        ["Copilot Reclaim", "3", "Has usage report + zero activity, not in report + zero activity, no report available"],
        ["Conditional Access Licensing", "4", "Targeted CA no P1, tenant-wide CA no P1, targeted risk-based CA no P2, tenant-wide risk-based CA no P2"],
        ["Disabled Account", "4", "With hold + expensive SKU, with hold + cheap SKU, EXO not connected, no hold"],
        ["Automation Account", "5", "Unlicensed admin, unlicensed service pattern, licensed admin, licensed service pattern, never signed in"],
        ["Frontline Blocked", "2", "Archive mailbox prevents F3 downgrade, multi-PC activations prevent F3 downgrade"],
        ["Shelfware", "6+", "Visio, Project, Teams Premium, web-only product, generic suite, per-product type"],
        ["Mailbox Storage Warning", "3", "Plan 1 approaching 50 GB, Plan 2 approaching 100 GB, Kiosk approaching 2 GB"],
      ]).map((row, i) => new TableRow({ children: [
        new TableCell({
          borders, width: { size: 3200, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: row[0], bold: true, size: 19, font: "Arial", color: "2C3E50" })] })]
        }),
        new TableCell({
          borders, width: { size: 900, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: row[1], bold: true, size: 19, font: "Arial", color: BRAND_ACCENT })] })]
        }),
        new TableCell({
          borders, width: { size: 5260, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: row[2], size: 18, font: "Arial", color: "555555" })] })]
        }),
      ]}))
    ]
  }),
  new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [new TextRun({
      text: "This means: 121 rules = 146 recommendation messages grouped by logical scenario. The per-user audit report shows the specific variant message; this reference guide explains the scenario once.",
      italics: true, size: 20, font: "Arial", color: "7F8C8D"
    })]
  }),

  // How to read
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text: "How to Read Each Rule", bold: true, size: 26, font: "Arial", color: BRAND_ACCENT })]
  }),
];

// Legend table
const legendRows = [
  ["Scope", "Whether the rule evaluates individual users (User) or the entire tenant (Tenant)."],
  ["Action", "The recommended remediation: Remove, Downgrade, Assign, Validate, Review, Consolidate, or Upgrade."],
  ["Confidence", "How certain the finding is: High (act immediately), Medium (validate before acting), or Low (informational, requires manual verification)."],
  ["Detection", "Automated (script detects automatically), Semi-Auto (partially automated), or Manual (requires portal verification)."],
];
overviewChildren.push(new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [2200, 7160],
  rows: legendRows.map((row, i) => new TableRow({
    children: [
      new TableCell({
        borders, width: { size: 2200, type: WidthType.DXA },
        shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ children: [new TextRun({ text: row[0], bold: true, size: 20, font: "Arial", color: BRAND_BLUE })] })]
      }),
      new TableCell({
        borders, width: { size: 7160, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [new Paragraph({ children: [new TextRun({ text: row[1], size: 20, font: "Arial", color: "555555" })] })]
      }),
    ]
  }))
}));

// Category overview table
overviewChildren.push(new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 300, after: 120 },
  children: [new TextRun({ text: "Category Overview", bold: true, size: 26, font: "Arial", color: BRAND_ACCENT })]
}));

overviewChildren.push(new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [2000, 1000, 6360],
  rows: [
    new TableRow({
      children: [
        headerCell("Category", 2000),
        headerCell("Rules", 1000),
        headerCell("Description", 6360),
      ]
    }),
    ...categoryConfig.map((cat) => new TableRow({
      children: [
        new TableCell({
          borders, width: { size: 2000, type: WidthType.DXA },
          shading: { fill: cat.color, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: cat.icon, bold: true, size: 20, font: "Arial", color: WHITE })] })]
        }),
        new TableCell({
          borders, width: { size: 1000, type: WidthType.DXA },
          shading: { fill: LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: String(grouped[cat.key].length), bold: true, size: 22, font: "Arial", color: "2C3E50" })]
          })]
        }),
        new TableCell({
          borders, width: { size: 6360, type: WidthType.DXA },
          shading: { fill: WHITE, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: cat.desc, size: 19, font: "Arial", color: "555555" })] })]
        }),
      ]
    }))
  ]
}));

function headerCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: BRAND_BLUE, type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: WHITE })] })]
  });
}

sections.push({
  properties: {
    page: {
      size: { width: PAGE_WIDTH, height: 15840 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    }
  },
  headers: {
    default: new Header({
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: BRAND_ACCENT, space: 1 } },
        children: [
          new TextRun({ text: "EASI M365 License Optimization Audit", italics: true, size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ text: "  |  Detection Scenario Reference", italics: true, size: 16, font: "Arial", color: "BDC3C7" }),
        ]
      })]
    })
  },
  footers: {
    default: new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 1, color: "D5D8DC", space: 1 } },
        children: [
          new TextRun({ text: "Page ", size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ text: `    |    Version ${rp.version}    |    CONFIDENTIAL`, size: 16, font: "Arial", color: "BDC3C7" }),
        ]
      })]
    })
  },
  children: overviewChildren
});

// ── Category sections ──
for (const cat of categoryConfig) {
  const catRules = grouped[cat.key];
  if (catRules.length === 0) continue;

  // Sort rules by ID
  catRules.sort((a, b) => a.id.localeCompare(b.id));

  const children = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { after: 80 },
      children: [new TextRun({ text: `${cat.icon} (${catRules.length} Rules)`, bold: true, size: 32, font: "Arial", color: cat.color })]
    }),
    new Paragraph({
      spacing: { after: 200 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: cat.color, space: 1 } },
      children: [new TextRun({ text: cat.desc, size: 22, font: "Arial", color: "555555" })]
    }),
  ];

  for (const rule of catRules) {
    children.push(...buildRuleBlock(rule));
  }

  sections.push({
    properties: {
      page: {
        size: { width: PAGE_WIDTH, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: cat.color, space: 1 } },
          children: [
            new TextRun({ text: "EASI M365 LOA", italics: true, size: 16, font: "Arial", color: "95A5A6" }),
            new TextRun({ text: `  |  ${cat.icon}`, italics: true, size: 16, font: "Arial", color: cat.color }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 1, color: "D5D8DC", space: 1 } },
          children: [
            new TextRun({ text: "Page ", size: 16, font: "Arial", color: "95A5A6" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "95A5A6" }),
            new TextRun({ text: `    |    Version ${rp.version}    |    CONFIDENTIAL`, size: 16, font: "Arial", color: "BDC3C7" }),
          ]
        })]
      })
    },
    children
  });
}

// ── Appendix: Microsoft Documentation References ──
const appendixChildren = [
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { after: 200 },
    children: [new TextRun({ text: "Appendix: Microsoft Documentation References", bold: true, size: 32, font: "Arial", color: BRAND_BLUE })]
  }),
  new Paragraph({
    spacing: { after: 200 },
    children: [new TextRun({
      text: "The following Microsoft Learn articles are referenced by the detection rules in this document.",
      size: 22, font: "Arial", color: "555555"
    })]
  }),
];

appendixChildren.push(new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: [2800, 6560],
  rows: [
    new TableRow({
      children: [
        headerCell("Reference ID", 2800),
        headerCell("Document Title & Link", 6560),
      ]
    }),
    ...docs.map((d, i) => new TableRow({
      children: [
        new TableCell({
          borders, width: { size: 2800, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: d.id, bold: true, size: 18, font: "Consolas", color: BRAND_ACCENT })] })]
        }),
        new TableCell({
          borders, width: { size: 6560, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? WHITE : LIGHT_GRAY, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({
            children: [new ExternalHyperlink({
              children: [new TextRun({ text: d.title, style: "Hyperlink", size: 19, font: "Arial" })],
              link: d.url,
            })]
          })]
        }),
      ]
    }))
  ]
}));

sections.push({
  properties: {
    page: {
      size: { width: PAGE_WIDTH, height: 15840 },
      margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    }
  },
  headers: {
    default: new Header({
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: BRAND_ACCENT, space: 1 } },
        children: [
          new TextRun({ text: "EASI M365 LOA", italics: true, size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ text: "  |  Appendix", italics: true, size: 16, font: "Arial", color: "BDC3C7" }),
        ]
      })]
    })
  },
  footers: {
    default: new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 1, color: "D5D8DC", space: 1 } },
        children: [
          new TextRun({ text: "Page ", size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: "95A5A6" }),
          new TextRun({ text: `    |    Version ${rp.version}    |    CONFIDENTIAL`, size: 16, font: "Arial", color: "BDC3C7" }),
        ]
      })]
    })
  },
  children: appendixChildren
});

// ── Create document ──
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: BRAND_BLUE },
        paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BRAND_ACCENT },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: "manual-steps",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  sections
});

// ── Write to file ──
const outPath = "./LOA_RulePack_M365_Reference.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log(`Document generated: ${outPath}`);
  console.log(`  Rules: ${rules.length}`);
  console.log(`  Categories: ${categoryConfig.length}`);
  console.log(`  Docs: ${docs.length}`);
}).catch(err => {
  console.error("Error generating document:", err);
  process.exit(1);
});
