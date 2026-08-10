const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, LevelFormat, convertInchesToTwip,
} = require("docx");
const fs = require("fs");

const INK = "1A1A1A";
const MUTED = "555555";
const GAP = "C00000";

const g = (t) => new TextRun({ text: t, color: GAP, bold: true });
const t = (text, opts = {}) => new TextRun({ text, color: INK, ...opts });

const sectionHeading = (text) =>
  new Paragraph({
    spacing: { before: 250, after: 105 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 4 } },
    children: [new TextRun({ text, bold: true, size: 21, color: INK, allCaps: true, characterSpacing: 30 })],
  });

const role = (children) => new Paragraph({ spacing: { before: 165, after: 20 }, children });
const meta = (children) =>
  new Paragraph({ spacing: { after: 70 }, children: children.map((c) =>
    typeof c === "string" ? new TextRun({ text: c, italics: true, size: 19, color: MUTED }) : c) });
const bullet = (children) =>
  new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 55 },
    children: (Array.isArray(children) ? children : [children]).map((c) =>
      typeof c === "string" ? t(c) : c) });
const body = (children, opts = {}) =>
  new Paragraph({ spacing: { after: 90 }, ...opts,
    children: (Array.isArray(children) ? children : [children]).map((c) =>
      typeof c === "string" ? t(c) : c) });

const doc = new Document({
  creator: "Marc Lebreton",
  title: "Marc Lebreton — CV",
  description: "CV targeted at Microsoft Senior Go-to-Market Manager, AI Business Solutions, Brazil",
  numbering: {
    config: [{
      reference: "b",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: convertInchesToTwip(0.24), hanging: convertInchesToTwip(0.16) } } },
      }],
    }],
  },
  styles: {
    default: { document: { run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 262 } } } },
  },
  sections: [{
    properties: { page: { margin: { top: 880, bottom: 880, left: 1000, right: 1000 } } },
    children: [

      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "Go-to-Market & Commercial Strategy — Brazil and Latin America", size: 21, color: MUTED })] }),
      new Paragraph({ spacing: { after: 30 }, children: [
        t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
        g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[linkedin.com/in/…]"),
      ]}),
      new Paragraph({ spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [ t("Portuguese ", { size: 19 }), g("[native / fluent]"),
          t("  ·  English (fluent)  ·  Spanish ", { size: 19 }), g("[fluent]"),
          t("  ·  Based in São Paulo — available for 3 days/week on site", { size: 19, color: MUTED }) ]}),

      sectionHeading("Profile"),
      body([
        t("Go-to-market and commercial strategy leader with "),
        g("[X]"),
        t("+ years translating global product strategy into local execution in Brazil and Latin America. Built and ran the market-entry plan that took Crypto.com's LATAM business from zero to "),
        t("6M+ users and the #1 app position in the region", { bold: true }),
        t(", orchestrating across marketing, product, partnerships and operations in a matrixed global organisation. Owns business planning, performance frameworks and field enablement end to end, and turns market and customer insight into prioritised, measurable action. Deep working knowledge of the Brazilian market — "),
        g("[name the industries you genuinely know: Financial Services, Retail, Public Sector, Oil & Gas — Microsoft lists these four explicitly]"),
        t(" — and an established local network across "),
        g("[partners, regulators, media, ecosystem]"),
        t("."),
      ]),

      sectionHeading("Selected Outcomes"),
      bullet([ t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
        t(" — owned the regional go-to-market plan for Crypto.com's market entry across "), g("[N]"), t(" countries "), g("[over __ months]") ]),
      bullet([ t("0 → 130,000", { bold: true }), t(" — built Sui Foundation's regional channel from a standing start "), g("[and the adoption or pipeline outcome it produced]") ]),
      bullet([ g("[Revenue / usage / adoption metric you owned and moved: from __ to __ over __ months]") ]),
      bullet([ g("[A business plan you authored that was adopted — what changed as a result]") ]),
      bullet([ g("[A cross-functional programme you led across __ teams / __ countries, and its measurable result]") ]),
      bullet([ g("[An executive-level review or recommendation that changed a regional investment or priority]") ]),

      sectionHeading("Experience"),

      role([ t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ t("Built the regional channel from "), t("0 to 130,000", { bold: true }), t(" — "), g("[which markets, and what it fed: adoption, developer/partner acquisition, pipeline]") ]),
      bullet([ g("[Cross-functional orchestration: which teams — product, marketing, partnerships, ecosystem — and what you aligned them around]") ]),
      bullet([ g("[Insight → recommendation: a market or competitive read you produced that changed the plan]") ]),
      bullet([ g("[Partner or ecosystem programmes you designed and ran, with enablement materials and measured uptake]") ]),

      role([ t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ t("Owned the Latin America go-to-market plan end to end — market prioritisation, segment strategy, launch sequencing and local execution — taking the region from zero to "),
        t("6M+ users", { bold: true }), t(" and "), t("#1 app position", { bold: true }) ]),
      bullet([ g("[Business planning: how you built the regional plan, what you committed to, and how it laddered up to global objectives. Microsoft calls this \"OU business plan\" — mirror that language.]") ]),
      bullet([ g("[Performance management: the KPIs and scorecards you owned, the review cadence you ran, and who you reported them to. Microsoft calls this the \"ROB\" — rhythm of business.]") ]),
      bullet([ g("[Field enablement: the plays, materials, training or partner enablement you built, and adoption of them]") ]),
      bullet([ g("[Market intelligence: competitive and regulatory reads you produced, and the decisions they drove — LATAM crypto regulation is a strong, specific example if you worked on it]") ]),
      bullet([ g("[Working in a matrixed global org: which HQ functions you had to influence without authority, and how]") ]),

      role([ g("[Earlier role — company]"), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ g("[2–3 bullets. Anything enterprise, B2B, consulting, or field-sales adjacent should be surfaced prominently — it is the scarcest evidence for this role.]") ]),

      sectionHeading("Capabilities"),
      body([ t("Business planning & strategy  ", { bold: true }),
        t("Regional business plans aligned to global objectives · market and segment prioritisation · turning strategy into measurable execution plans · "),
        g("[P&L or budget ownership, if any]") ]),
      body([ t("Go-to-market execution  ", { bold: true }),
        t("Launch strategy and sequencing · solution and campaign plays · acquisition, expansion and adoption motions · partner and channel go-to-market") ]),
      body([ t("Business health & performance  ", { bold: true }),
        t("KPI frameworks and scorecards · business reviews and reporting cadence · pipeline and funnel analysis · "),
        g("[tools: Power BI, Excel, Salesforce, Amplitude, Looker — list what you've actually used]") ]),
      body([ t("Cross-functional leadership  ", { bold: true }),
        t("Influencing without authority across marketing, product, engineering, finance, partners and operations · matrixed global/regional organisations · executive-level written and verbal communication") ]),
      body([ t("Market intelligence  ", { bold: true }),
        t("Competitive, regulatory and customer insight synthesised into prioritised recommendations · "),
        g("[Brazilian industries you can credibly speak to: Financial Services, Retail, Public Sector, Oil & Gas]") ]),
      body([ t("AI  ", { bold: true }),
        g("[Microsoft expects AI fluency in every role. Name concretely how you use it — market and competitive research, campaign and enablement content at scale, data analysis, or a workflow you automated. One specific example beats a list of tools.]") ]),

      sectionHeading("Education"),
      body([ g("[Degree — Institution — Year]") ]),
      body([ g("[MBA or postgraduate qualification, if you have one]") ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Marc_Lebreton_CV_Microsoft.docx", buf);
  console.log("written:", buf.length, "bytes");
});
