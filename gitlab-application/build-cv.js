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
    spacing: { before: 260, after: 110 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 4 } },
    children: [new TextRun({ text, bold: true, size: 21, color: INK, allCaps: true, characterSpacing: 30 })],
  });

const role = (children) => new Paragraph({ spacing: { before: 170, after: 20 }, children });
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
  description: "CV targeted at GitLab Regional Sales Director, Brazil",
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
    default: { document: { run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 264 } } } },
  },
  sections: [{
    properties: { page: { margin: { top: 900, bottom: 900, left: 1000, right: 1000 } } },
    children: [

      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "Regional Go-to-Market & Revenue Leader — Brazil & LATAM", size: 21, color: MUTED })] }),
      new Paragraph({ spacing: { after: 30 }, children: [
        t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
        g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[linkedin.com/in/…]"),
      ]}),
      new Paragraph({ spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [ t("Portuguese ", { size: 19 }), g("[native / fluent]"),
          t("  ·  Spanish ", { size: 19 }), g("[fluent]"), t("  ·  English (fluent)", { size: 19 }) ]}),

      sectionHeading("Profile"),
      body([
        t("Go-to-market leader who has taken global technology products into Brazil and Hispanic LATAM from a standing start. Led Crypto.com's LATAM market entry to "),
        t("6M+ users and the #1 app position in the region", { bold: true }),
        t(", owning regional strategy, local partnerships and the commercial plan across "),
        g("[N]"),
        t(" markets. "),
        g("[X]"),
        t(" years operating in Brazil with an established network across "),
        g("[enterprise buyers / partners / channel / regulators — name what's real]"),
        t(". "),
        g("[IF TRUE: Built and led a team of __ , carrying $__ in annual bookings against a $__ target. IF NOT TRUE, DELETE THIS SENTENCE and see the note below.]"),
      ]),

      sectionHeading("Performance Snapshot"),
      body([ t("For a sales-leadership CV this block decides the screen. GitLab's recruiter is looking for a number, a target, and a team size. Replace every red field or delete the row.", { italics: true, size: 18, color: MUTED }) ]),
      bullet([ g("[Quota / bookings owned: $__ ARR or R$__ annual]"), t("  —  "), g("[Attainment: __% of target, __ consecutive quarters/years]") ]),
      bullet([ g("[Team led: __ quota-carrying reps + __ SDR/SE; hired __, promoted __, managed out __]") ]),
      bullet([ g("[Pipeline: $__ built / __x coverage; win rate __%; average deal size $__; sales cycle __ days]") ]),
      bullet([ g("[Growth: region grew from $__ to $__ over __ months (__% YoY)]") ]),
      bullet([ t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in LATAM", { bold: true }), t(" — Crypto.com regional market entry "), g("[state the commercial outcome this drove: revenue, transacting accounts, or partner-sourced volume]") ]),
      bullet([ t("0 → 130,000", { bold: true }), t(" — Sui Foundation regional channel "), g("[keep only if you can tie it to a commercial result; otherwise cut it from THIS CV]") ]),

      sectionHeading("Experience"),

      role([ t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ g("[Lead with a commercial outcome — revenue, partnerships signed, ecosystem accounts activated — not audience size.]") ]),
      bullet([ t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" — "), g("[and what it produced downstream]") ]),
      bullet([ g("[Partnerships or enterprise/institutional relationships you personally originated and closed]") ]),
      bullet([ g("[Cross-functional work: marketing, product, ops — GitLab asks for this explicitly]") ]),

      role([ t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ t("Owned LATAM market entry and go-to-market strategy, taking the business from zero to "), t("6M+ users", { bold: true }), t(" and "), t("#1 app in the region", { bold: true }), t(" across "), g("[which countries, over what period]") ]),
      bullet([ g("[Territory / market planning: how you segmented the region, prioritised markets, and what you chose NOT to do. GitLab lists territory planning as a core responsibility.]") ]),
      bullet([ g("[Team leadership: how many people, in what functions, hired vs inherited. If you have never managed a quota-carrying team, say so honestly in the cover letter rather than blurring it here.]") ]),
      bullet([ g("[Forecasting and reporting: what metrics you owned, what cadence, to whom, in what system]") ]),
      bullet([ g("[Named partners, logos or institutional accounts you can reference]") ]),

      role([ g("[Earlier role — company]"), t(" — "), g("[Title]") ]),
      meta([ g("[Location]"), " · ", g("[Start – End]") ]),
      bullet([ g("[2–3 bullets. Any quota-carrying, partnership-closing, or team-management work belongs here and should be surfaced prominently — it is the scarcest evidence for this role.]") ]),

      sectionHeading("Capabilities"),
      body([ t("Regional leadership  ", { bold: true }),
        t("Brazil and LATAM go-to-market strategy · market entry from zero · territory and market segmentation · "),
        g("[hiring, coaching and performance management — only if true]"),
        t(" · cross-functional alignment with marketing, product and operations") ]),
      body([ t("Commercial  ", { bold: true }),
        g("[Forecasting · pipeline management · negotiation · executive-level presentation · partner and channel development — keep what's genuinely yours]") ]),
      body([ t("Analytics  ", { bold: true }),
        t("Interpreting performance data to adjust regional plans · cohort and funnel analysis · "),
        g("[reporting stack you've used]") ]),
      body([ t("Tools  ", { bold: true }),
        g("[Salesforce · Clari · Marketo — GitLab names these three explicitly. List only what you have actually used, and add your real stack: HubSpot, Amplitude, Looker, etc.]") ]),
      body([ t("Markets  ", { bold: true }),
        t("Brazil · "), g("[Mexico · Argentina · Colombia · Chile — only where you've operated]") ]),

      sectionHeading("Education"),
      body([ g("[Degree — Institution — Year]") ]),
      body([ g("[MBA or equivalent, if you have one — 52% of applicants to this posting reportedly hold one, so lead with it if you do]") ]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Marc_Lebreton_CV_GitLab.docx", buf);
  console.log("written:", buf.length, "bytes");
});
