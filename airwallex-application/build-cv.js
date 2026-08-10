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
  description: "CV targeted at Airwallex Head of Sales, SME & Growth, Brazil",
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
        new TextRun({ text: "Commercial & Growth Leadership — Fintech, Brazil and Latin America", size: 21, color: MUTED })] }),
      new Paragraph({ spacing: { after: 30 }, children: [
        t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
        g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[linkedin.com/in/…]"),
      ]}),
      new Paragraph({ spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [ t("English (fluent)  ·  Portuguese ", { size: 19 }), g("[native / fluent]"),
          t("  ·  Spanish ", { size: 19 }), g("[fluent]"),
          t("  ·  São Paulo based, available on site", { size: 19, color: MUTED }) ]}),

      sectionHeading("Profile"),
      body([
        t("Commercial and growth leader in Brazilian fintech, with a zero-to-one track record in a hyper-growth environment. Built and ran the market-entry plan that took Crypto.com's Latin America business from nothing to "),
        t("6M+ customers and the #1 app position in the region", { bold: true }),
        t(", owning the full customer lifecycle — acquisition, activation, retention and expansion — across "),
        g("[N]"),
        t(" markets. "),
        g("[IF TRUE: Built and led a commercial team of __ across sales, account management and partnerships. THIS SENTENCE IS THE ROLE'S HARD MINIMUM — see the note in my message. If it isn't true, delete it and do not imply it.]"),
        t(" Established network across the Brazilian fintech and payments ecosystem — "),
        g("[name real relationships: PSPs, acquirers, banks, regulators, fintech operators]"),
        t(" — and working knowledge of local rails including PIX and TED."),
      ]),

      sectionHeading("Selected Outcomes"),
      bullet([ t("0 → 6,000,000+ customers", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
        t(" — Crypto.com regional launch across "), g("[N]"), t(" markets in "), g("[__ months]") ]),
      bullet([ g("[Revenue / ARR / monetisation: took the region from $__ to $__ — Airwallex will look for a revenue number, not just a user number]") ]),
      bullet([ g("[Team: built a commercial team of __ across SDR / AE / partnerships; __ hired directly; __ promoted]") ]),
      bullet([ g("[Quota or target attainment: __% against a $__ target, over __ quarters]") ]),
      bullet([ g("[Multi-channel acquisition: which channels you built, CAC by channel, and payback period]") ]),
      bullet([ t("0 → 130,000", { bold: true }), t(" — Sui Foundation regional channel "),
        g("[keep only if you can attach a commercial outcome; otherwise cut from this CV]") ]),
      bullet([ g("[Deals you personally closed as an IC — Airwallex explicitly wants a player/coach, so name the largest and the segment]") ]),

      sectionHeading("Experience"),

      role([ t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ g("[Lead with the commercial outcome — revenue, partnerships closed, institutional or B2B accounts activated]") ]),
      bullet([ t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" — "), g("[which markets, and what it converted into]") ]),
      bullet([ g("[Partnerships you originated and closed, and their contribution]") ]),
      bullet([ g("[Cross-functional work with product and marketing — Airwallex names both explicitly]") ]),

      role([ t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ t("Owned the Latin America commercial and go-to-market strategy end to end, scaling from zero to "),
        t("6M+ customers", { bold: true }), t(" and "), t("#1 app position", { bold: true }), t(" across "), g("[which countries]") ]),
      bullet([ g("[Team leadership — the single most important bullet on this page. How many people, in which functions (SDR, AE, partnerships, marketing), how many you hired, and over how long. Airwallex requires 5+ years of this as a MINIMUM qualification.]") ]),
      bullet([ g("[Multi-channel customer acquisition: the channels you built and scaled, with CAC and volume by channel]") ]),
      bullet([ g("[Payments and local rails: PIX, TED, boleto, card acquiring, local banking partners — what you integrated, negotiated or prioritised]") ]),
      bullet([ g("[Regulatory and licensing work in Brazil — genuinely differentiating for a fintech role, if you touched it]") ]),
      bullet([ g("[Customer lifecycle beyond acquisition: activation, retention, upsell — Airwallex names all four]") ]),
      bullet([ g("[Representing the company externally: events, press, industry panels in the Brazilian market]") ]),

      role([ g("[Earlier role — company]"), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ g("[2–3 bullets. Any B2B, SME-segment, consulting or team-management experience belongs here and should be surfaced hard — those are the three things this CV is thinnest on.]") ]),

      sectionHeading("Capabilities"),
      body([ t("Commercial leadership  ", { bold: true }),
        t("Brazil commercial strategy · vertical and segment positioning · "),
        g("[building and scaling SDR / AE / partnership teams — only if true]"),
        t(" · player/coach operating style · pipeline and forecast ownership") ]),
      body([ t("Growth  ", { bold: true }),
        t("Multi-channel customer acquisition · full lifecycle ownership (acquisition, activation, retention, expansion) · growth experimentation · cohort and CAC/LTV analysis · partner and channel-led growth") ]),
      body([ t("Fintech & payments  ", { bold: true }),
        t("PIX · TED · "), g("[boleto · card acquiring · FX · cross-border · treasury]"),
        t(" · Brazilian fintech regulation · "), g("[named ecosystem relationships you can actually activate]") ]),
      body([ t("Markets  ", { bold: true }),
        t("Brazil · "), g("[Mexico · Argentina · Colombia · Chile — only where you operated]") ]),
      body([ t("Tools  ", { bold: true }),
        g("[Salesforce / HubSpot · Amplitude or Mixpanel · Looker or Power BI · SQL — list only what you've used]") ]),

      sectionHeading("Education"),
      body([ g("[Degree — Institution — Year]") ]),
      body([ g("[MBA, if you have one — 52% of applicants to this posting reportedly do]") ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Marc_Lebreton_CV_Airwallex.docx", buf);
  console.log("written:", buf.length, "bytes");
});
