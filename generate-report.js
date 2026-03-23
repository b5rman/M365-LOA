#!/usr/bin/env node
// generate-report.js — Produces a branded EASI Word document from LOA summary data
// Usage: node generate-report.js <input.json> <output.docx> [--cover <cover.jpg>] [--header <header.jpg>]

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak, PageNumber,
  LevelFormat
} = require('docx');

// ── EASI Brand Constants ────────────────────────────────────────────────────
const PURPLE = '412888';
const DARK   = '120C3A';
const LIGHT_BG = 'F3F0FA';  // light purple tint for table headers
const ACCENT_BG = 'E8E0F5'; // slightly stronger purple tint for callouts
const FONT  = 'Axiforma';
const FONT_FALLBACK = 'Calibri';
const GREY  = '6A6A8E';

// ── CLI Args ────────────────────────────────────────────────────────────────
const args = process.argv.slice(2);
const inputPath  = args[0];
const outputPath = args[1];

let coverPath  = path.join(__dirname, 'assets', 'easi-cover.jpg');
let headerPath = path.join(__dirname, 'assets', 'easi-header.jpg');

for (let i = 2; i < args.length; i++) {
  if (args[i] === '--cover'  && args[i+1]) coverPath  = args[++i];
  if (args[i] === '--header' && args[i+1]) headerPath = args[++i];
}

if (!inputPath || !outputPath) {
  console.error('Usage: node generate-report.js <input.json> <output.docx>');
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(inputPath, 'utf8'));

// ── Helper Functions ────────────────────────────────────────────────────────
const txt = (text, opts = {}) => new TextRun({
  text, font: FONT, color: opts.color || DARK, size: opts.size || 20,
  bold: opts.bold || false, italics: opts.italics || false,
  ...opts
});

const heading = (text, level = HeadingLevel.HEADING_1) => new Paragraph({
  heading: level,
  spacing: { before: level === HeadingLevel.HEADING_1 ? 360 : 240, after: 120 },
  children: [txt(text, { bold: true, color: PURPLE, size: level === HeadingLevel.HEADING_1 ? 28 : 24 })]
});

const para = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after || 120 },
  alignment: opts.align || AlignmentType.LEFT,
  children: Array.isArray(text) ? text : [txt(text, opts)]
});

const emptyPara = () => new Paragraph({ children: [] });

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0 },
  bottom: { style: BorderStyle.NONE, size: 0 },
  left: { style: BorderStyle.NONE, size: 0 },
  right: { style: BorderStyle.NONE, size: 0 }
};

const cell = (text, opts = {}) => new TableCell({
  borders: opts.noBorder ? noBorders : borders,
  width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
  shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [para(text, { bold: opts.bold, size: opts.size || 20, color: opts.color })]
});

const headerCell = (text, width) => cell(text, {
  bold: true, shading: LIGHT_BG, color: PURPLE, width, size: 18
});

// Build a simple two-column key-value table
function kvTable(items, labelWidth = 5500, valueWidth = 3860) {
  const totalWidth = labelWidth + valueWidth;
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    rows: items.filter(([k, v]) => v != null && v !== '').map(([k, v]) => new TableRow({
      children: [
        cell(k, { width: labelWidth, bold: true, size: 18, shading: LIGHT_BG }),
        cell(String(v), { width: valueWidth, size: 18 })
      ]
    }))
  });
}

// Build a table with headers and data rows
function dataTable(headers, rows, widths) {
  const totalWidth = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({
        children: headers.map((h, i) => headerCell(h, widths[i]))
      }),
      ...rows.map(row => new TableRow({
        children: row.map((val, i) => cell(String(val || ''), { width: widths[i], size: 18 }))
      }))
    ]
  });
}

// ── Parse sections from data ────────────────────────────────────────────────
const d = data;

// ── Build Document Sections ─────────────────────────────────────────────────
const children = [];

// ── COVER PAGE ──────────────────────────────────────────────────────────────
const coverChildren = [];

// Cover image
if (fs.existsSync(coverPath)) {
  coverChildren.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    children: [new ImageRun({
      type: 'jpg', data: fs.readFileSync(coverPath),
      transformation: { width: 500, height: 200 },
      altText: { title: 'EASI', description: 'EASI logo banner', name: 'cover' }
    })]
  }));
}

coverChildren.push(emptyPara(), emptyPara(), emptyPara());

coverChildren.push(para([
  txt('Microsoft 365', { bold: true, color: PURPLE, size: 48 })
], { align: AlignmentType.CENTER }));

coverChildren.push(para([
  txt('License Assessment Report', { bold: true, color: PURPLE, size: 40 })
], { align: AlignmentType.CENTER }));

coverChildren.push(emptyPara(), emptyPara());

if (d.customerName) {
  coverChildren.push(para([
    txt('Prepared for: ', { color: GREY, size: 24 }),
    txt(d.customerName, { bold: true, color: DARK, size: 28 })
  ], { align: AlignmentType.CENTER }));
}

coverChildren.push(para([
  txt(d.reportDate || '', { color: GREY, size: 22 })
], { align: AlignmentType.CENTER }));

coverChildren.push(emptyPara());

coverChildren.push(para([
  txt('Confidential', { color: GREY, size: 18, italics: true })
], { align: AlignmentType.CENTER }));

coverChildren.push(new Paragraph({ children: [new PageBreak()] }));

// ── DISCLAIMER ──────────────────────────────────────────────────────────────
coverChildren.push(para([
  txt('Disclaimer', { bold: true, color: PURPLE, size: 22 })
]));
coverChildren.push(para([
  txt('All cost figures in this report are indicative estimates based on vendor list prices (EUR). Actual costs may differ due to EA, CSP, or volume pricing agreements. Copilot usage and Cloud PC analytics rely on Microsoft Graph beta APIs; these sections may show limited results until the APIs become generally available. All assessments are advisory and should be validated with stakeholders before any changes are made. Usage data is based on the last 90 days and may not reflect seasonal patterns.', {
    color: GREY, size: 18, italics: true
  })
]));
coverChildren.push(emptyPara());
coverChildren.push(para([
  txt('About the figures in this report', { bold: true, color: GREY, size: 18 })
]));
coverChildren.push(para([
  txt('This report and the interactive HTML dashboard use primary-category attribution: each user is counted once under their most significant finding, providing a clear, non-overlapping view of the optimization landscape. The Excel workbook contains the full per-user detail where a user may appear under multiple findings. Totals in the Excel executive summary may therefore differ slightly from this report and the dashboard.', {
    color: GREY, size: 18, italics: true
  })
]));

coverChildren.push(new Paragraph({ children: [new PageBreak()] }));

// ── EXECUTIVE FINANCIAL SUMMARY ─────────────────────────────────────────────
coverChildren.push(heading('Executive Financial Summary'));

// Big numbers callout
if (d.totalAnnualSpend || d.optimizationPotential) {
  coverChildren.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: noBorders,
          width: { size: 4680, type: WidthType.DXA },
          shading: { fill: ACCENT_BG, type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 200, right: 200 },
          children: [
            para([txt('Total Annual Spend', { color: GREY, size: 16 })]),
            para([txt(d.totalAnnualSpend || '—', { bold: true, color: DARK, size: 36 })])
          ]
        }),
        new TableCell({
          borders: noBorders,
          width: { size: 4680, type: WidthType.DXA },
          shading: { fill: ACCENT_BG, type: ShadingType.CLEAR },
          margins: { top: 200, bottom: 200, left: 200, right: 200 },
          children: [
            para([txt('Estimated Optimization Potential', { color: GREY, size: 16 })]),
            para([txt(d.optimizationPotential || '—', { bold: true, color: PURPLE, size: 36 })]),
            para([txt(d.optimizationPct || '', { color: GREY, size: 16 })])
          ]
        })
      ]
    })]
  }));

  // Compliance cost callout (if available)
  if (d.complianceCost) {
    coverChildren.push(emptyPara());
    coverChildren.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: noBorders,
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: 'FFF0EB', type: ShadingType.CLEAR },
            margins: { top: 150, bottom: 150, left: 200, right: 200 },
            children: [
              para([
                txt('Estimated Compliance Remediation Cost: ', { color: GREY, size: 18 }),
                txt(d.complianceCost, { bold: true, color: 'D94070', size: 22 })
              ]),
              para([txt('This represents the estimated annual cost of add-on licenses needed to close identified compliance gaps (Conditional Access, Defender for Office 365, Privileged Identity Management). Alternatively, affected users or mailboxes can be excluded from the relevant policies to avoid this cost.', { color: GREY, size: 16, italics: true })])
            ]
          })
        ]
      })]
    }));
  }

  coverChildren.push(emptyPara());
}

// Tier 1 Quick Wins
if (d.tier1 && d.tier1.items && d.tier1.items.length > 0) {
  coverChildren.push(heading('Tier 1 — Quick Wins', HeadingLevel.HEADING_2));
  coverChildren.push(para('Potential license cost that may be recoverable by reviewing dormant, disabled, or unused licenses.', { color: GREY, size: 18 }));
  coverChildren.push(dataTable(
    ['Category', 'Potential Annual Savings', 'Users'],
    d.tier1.items.map(i => [i.category, i.amount, i.count]),
    [5000, 2500, 1860]
  ));
  if (d.tier1.subtotal) {
    coverChildren.push(para([
      txt('Tier 1 Subtotal: ', { bold: true, size: 20 }),
      txt(d.tier1.subtotal, { bold: true, color: PURPLE, size: 20 })
    ], { after: 200 }));
  }
}

// Tier 2 Right-Sizing
if (d.tier2 && d.tier2.items && d.tier2.items.length > 0) {
  coverChildren.push(heading('Tier 2 — Right-Sizing Opportunities', HeadingLevel.HEADING_2));
  coverChildren.push(para('Potential savings identified through SKU downgrades, consolidation, and plan alignment. These opportunities should be reviewed with stakeholders before implementation.', { color: GREY, size: 18 }));
  coverChildren.push(dataTable(
    ['Opportunity', 'Potential Annual Savings', 'Users'],
    d.tier2.items.map(i => [i.category, i.amount, i.count]),
    [5000, 2500, 1860]
  ));
  if (d.tier2.subtotal) {
    coverChildren.push(para([
      txt('Tier 2 Subtotal: ', { bold: true, size: 20 }),
      txt(d.tier2.subtotal, { bold: true, color: PURPLE, size: 20 })
    ], { after: 200 }));
  }
}

// Pool waste
if (d.poolWaste && d.poolWaste.items && d.poolWaste.items.length > 0) {
  coverChildren.push(heading('Unassigned License Pool', HeadingLevel.HEADING_2));
  coverChildren.push(para('Paid licenses with unassigned seats. Consider reducing seat count at next renewal or assigning to users who need them.', { color: GREY, size: 18 }));
  coverChildren.push(dataTable(
    ['SKU', 'Unassigned / Total', 'Potential Reclaim'],
    d.poolWaste.items.map(i => [i.sku, i.seats, i.amount]),
    [4500, 2500, 2360]
  ));
  if (d.poolWaste.total) {
    coverChildren.push(para([
      txt('Pool Waste Total: ', { bold: true, size: 20 }),
      txt(d.poolWaste.total, { bold: true, color: PURPLE, size: 20 })
    ], { after: 200 }));
  }
}

coverChildren.push(new Paragraph({ children: [new PageBreak()] }));

// ── ACCOUNT & ROLE OVERVIEW ─────────────────────────────────────────────────
coverChildren.push(heading('Account & Role Overview'));
if (d.accountFlags && d.accountFlags.length > 0) {
  coverChildren.push(kvTable(d.accountFlags));
}

coverChildren.push(emptyPara());

// ── ACTIVITY-BASED FINDINGS ─────────────────────────────────────────────────
coverChildren.push(heading('Activity-Based Findings'));
if (d.activityFlags && d.activityFlags.length > 0) {
  coverChildren.push(kvTable(d.activityFlags));
}

coverChildren.push(emptyPara());

// ── QUICK WINS DETAIL ───────────────────────────────────────────────────────
coverChildren.push(heading('Quick Wins Detail'));
if (d.quickWins && d.quickWins.length > 0) {
  coverChildren.push(kvTable(d.quickWins));
}

coverChildren.push(emptyPara());

// ── RIGHT-SIZING DETAIL ─────────────────────────────────────────────────────
coverChildren.push(heading('Right-Sizing Opportunities'));
if (d.rightSizing && d.rightSizing.length > 0) {
  coverChildren.push(kvTable(d.rightSizing, 6000, 3360));
}

coverChildren.push(new Paragraph({ children: [new PageBreak()] }));

// ── COMPLIANCE GAPS ─────────────────────────────────────────────────────────
coverChildren.push(heading('Licensing Compliance'));
if (d.compliance && d.compliance.length > 0) {
  coverChildren.push(kvTable(d.compliance));
  coverChildren.push(emptyPara());
  coverChildren.push(para([
    txt('Note: ', { bold: true, color: GREY, size: 18 }),
    txt('Compliance gaps represent licenses that may need to be added to maintain policy coverage. Alternatively, exclude affected users from the policy to avoid the compliance cost.', { color: GREY, size: 18 })
  ]));
}

coverChildren.push(emptyPara());

// ── COPILOT ADOPTION (conditional) ──────────────────────────────────────────
if (d.copilot && d.copilot.totalHolders > 0) {
  coverChildren.push(heading('Copilot Adoption'));

  // Adoption metrics callout
  const cpTotal = d.copilot.totalHolders;
  const cpActive = d.copilot.activeCount || 0;
  const cpWatchlist = d.copilot.watchlistCount || 0;
  const cpReclaim = d.copilot.reclaimCount || 0;
  const adoptionPct = cpTotal > 0 ? Math.round(cpActive / cpTotal * 100) : 0;
  const annualInvest = cpTotal * 312; // €26/mo standard
  const costPerActive = cpActive > 0 ? Math.round(annualInvest / cpActive) : 0;

  coverChildren.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: noBorders, width: { size: 3120, type: WidthType.DXA },
          shading: { fill: ACCENT_BG, type: ShadingType.CLEAR },
          margins: { top: 150, bottom: 150, left: 150, right: 150 },
          children: [
            para([txt('Adoption Rate', { color: GREY, size: 14 })]),
            para([txt(adoptionPct + '%', { bold: true, color: PURPLE, size: 32 })]),
            para([txt(cpActive + ' of ' + cpTotal + ' holders', { color: GREY, size: 14 })])
          ]
        }),
        new TableCell({
          borders: noBorders, width: { size: 3120, type: WidthType.DXA },
          shading: { fill: ACCENT_BG, type: ShadingType.CLEAR },
          margins: { top: 150, bottom: 150, left: 150, right: 150 },
          children: [
            para([txt('Annual Investment', { color: GREY, size: 14 })]),
            para([txt('\u20AC' + annualInvest.toLocaleString('de-DE'), { bold: true, color: DARK, size: 28 })]),
            para([txt('\u20AC26/user/month', { color: GREY, size: 14 })])
          ]
        }),
        new TableCell({
          borders: noBorders, width: { size: 3120, type: WidthType.DXA },
          shading: { fill: ACCENT_BG, type: ShadingType.CLEAR },
          margins: { top: 150, bottom: 150, left: 150, right: 150 },
          children: [
            para([txt('Cost per Active User', { color: GREY, size: 14 })]),
            para([txt(costPerActive > 0 ? '\u20AC' + costPerActive.toLocaleString('de-DE') + '/yr' : '\u2014', { bold: true, color: DARK, size: 28 })]),
            para([txt(cpReclaim > 0 ? cpReclaim + ' reclaim candidate(s)' : cpWatchlist > 0 ? cpWatchlist + ' on watchlist' : 'All holders active', { color: GREY, size: 14 })])
          ]
        })
      ]
    })]
  }));
  coverChildren.push(emptyPara());

  // Pipeline detail
  coverChildren.push(kvTable(d.copilot.items || []));
  coverChildren.push(emptyPara());
}

// ── CLOUD PC (conditional) ──────────────────────────────────────────────────
if (d.cloudPc && d.cloudPc.length > 0) {
  coverChildren.push(heading('Cloud PC Utilization'));
  coverChildren.push(kvTable(d.cloudPc));
  coverChildren.push(emptyPara());
}

// ── SUBSCRIPTION ALERTS (conditional) ───────────────────────────────────────
if (d.subscriptionAlerts && d.subscriptionAlerts.length > 0) {
  coverChildren.push(heading('Subscription Alerts'));
  coverChildren.push(para('The following subscriptions are expiring within 90 days or have non-standard status.', { color: GREY, size: 18 }));
  coverChildren.push(dataTable(
    ['SKU', 'Status', 'Days Remaining'],
    d.subscriptionAlerts.map(a => [a.sku, a.status, a.days]),
    [5500, 2000, 1860]
  ));
  coverChildren.push(emptyPara());
}

// ── COST BY DEPARTMENT ──────────────────────────────────────────────────────
if (d.departments && d.departments.length > 0) {
  coverChildren.push(heading('Cost by Department'));
  coverChildren.push(dataTable(
    ['Department', 'Users', 'Annual Cost'],
    d.departments.map(dep => [dep.name, dep.users, dep.cost]),
    [5000, 1500, 2860]
  ));
  coverChildren.push(emptyPara());
}

// ── RECOMMENDATION DISTRIBUTION ─────────────────────────────────────────────
if (d.recDistribution && d.recDistribution.length > 0) {
  coverChildren.push(new Paragraph({ children: [new PageBreak()] }));
  coverChildren.push(heading('Assessment Distribution'));
  coverChildren.push(para('Primary assessment category assigned to each user. Users may have multiple secondary findings.', { color: GREY, size: 18 }));
  coverChildren.push(dataTable(
    ['Category', 'Users'],
    d.recDistribution.map(r => [r.category, r.count]),
    [7000, 2360]
  ));
}

// ── HEADER / FOOTER ─────────────────────────────────────────────────────────
const headerChildren = [];
if (fs.existsSync(headerPath)) {
  headerChildren.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    children: [new ImageRun({
      type: 'jpg', data: fs.readFileSync(headerPath),
      transformation: { width: 100, height: 12 },
      altText: { title: 'EASI', description: 'EASI header logo', name: 'header-logo' }
    })]
  }));
}

const footerContent = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      txt('Confidential', { color: GREY, size: 14, italics: true }),
      txt('  |  ', { color: GREY, size: 14 }),
      txt('Prepared by EASI', { color: GREY, size: 14 }),
      txt('  |  Page ', { color: GREY, size: 14 }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, color: GREY, size: 14 })
    ]
  })]
});

// ── CREATE DOCUMENT ─────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, color: DARK, size: 20 }
      }
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: PURPLE },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 }
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: FONT, color: PURPLE },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
      }
    },
    headers: {
      default: new Header({ children: headerChildren.length > 0 ? headerChildren : [new Paragraph({ children: [] })] })
    },
    footers: { default: footerContent },
    children: coverChildren
  }]
});

// ── WRITE OUTPUT ────────────────────────────────────────────────────────────
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Report written to: ${outputPath}`);
}).catch(err => {
  console.error('Error generating report:', err);
  process.exit(1);
});
