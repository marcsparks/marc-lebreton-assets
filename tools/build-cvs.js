// Clean-placeholder rebuild of the Microsoft and Airwallex CVs.
// Placeholders are short, neutral blanks — the guidance lives outside the CV.
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, LevelFormat, convertInchesToTwip,
} = require("docx");
const fs = require("fs");

const INK = "1A1A1A", MUTED = "555555", GAP = "C00000";
const g = (s) => new TextRun({ text: s, color: GAP, bold: true });
const t = (s, o = {}) => new TextRun({ text: s, color: INK, ...o });

const H = (text) => new Paragraph({
  spacing: { before: 250, after: 105 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 4 } },
  children: [new TextRun({ text, bold: true, size: 21, color: INK, allCaps: true, characterSpacing: 30 })],
});
const role = (c) => new Paragraph({ spacing: { before: 165, after: 20 }, children: c });
const meta = (c) => new Paragraph({ spacing: { after: 70 }, children: c.map((x) =>
  typeof x === "string" ? new TextRun({ text: x, italics: true, size: 19, color: MUTED }) : x) });
const li = (c) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 55 },
  children: (Array.isArray(c) ? c : [c]).map((x) => (typeof x === "string" ? t(x) : x)) });
const p = (c, o = {}) => new Paragraph({ spacing: { after: 90 }, ...o,
  children: (Array.isArray(c) ? c : [c]).map((x) => (typeof x === "string" ? t(x) : x)) });

const numbering = { config: [{ reference: "b", levels: [{
  level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
  style: { paragraph: { indent: { left: convertInchesToTwip(0.24), hanging: convertInchesToTwip(0.16) } } },
}] }] };
const styles = { default: { document: {
  run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 262 } } } } };

// ---- shared header ----
const header = (tagline, langLine) => ([
  new Paragraph({ spacing: { after: 40 }, children: [
    new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })] }),
  new Paragraph({ spacing: { after: 40 }, children: [
    new TextRun({ text: tagline, size: 21, color: MUTED })] }),
  new Paragraph({ spacing: { after: 30 }, children: [
    t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
    g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[LinkedIn]")] }),
  new Paragraph({ spacing: { after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
    children: langLine }),
]);

// ---- shared experience skeleton ----
const experience = (extra) => ([
  H("Experience"),

  role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
  meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
  li([t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" across "), g("[markets]")]),
  li([g("[Second achievement, with a number]")]),
  li([g("[Third achievement, with a number]")]),

  role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
  meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
  li([t("Owned the Latin America go-to-market plan end to end, scaling from zero to "),
      t("6M+ users", { bold: true }), t(" and "), t("#1 app in the region", { bold: true }),
      t(" across "), g("[countries]"), t(" in "), g("[timeframe]")]),
  ...extra,
  li([g("[Team led: size and functions]")]),

  role([g("[Company]"), t(" — "), g("[job title]")]),
  meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
  li([g("[Achievement, with a number]")]),
  li([g("[Achievement, with a number]")]),
]);

// ================= MICROSOFT =================
const microsoft = new Document({
  creator: "Marc Lebreton", title: "Marc Lebreton — CV",
  description: "Microsoft Senior Go-to-Market Manager, AI Business Solutions, Brazil",
  numbering, styles,
  sections: [{ properties: { page: { margin: { top: 880, bottom: 880, left: 1000, right: 1000 } } },
    children: [
      ...header("Go-to-Market & Commercial Strategy — Brazil and Latin America", [
        t("Portuguese ", { size: 19 }), g("[level]"), t("  ·  English (fluent)  ·  Spanish ", { size: 19 }), g("[level]"),
        t("  ·  São Paulo based, available on site", { size: 19, color: MUTED })]),

      H("Profile"),
      p([t("Go-to-market and commercial strategy leader with "), g("[X]"),
         t(" years translating global product strategy into local execution in Brazil and Latin America. Built and ran the market-entry plan that took Crypto.com's LATAM business from zero to "),
         t("6M+ users and the #1 app position in the region", { bold: true }),
         t(", orchestrating across marketing, product, partnerships and operations. Owns business planning, performance frameworks and field enablement end to end, and turns market insight into prioritised, measurable action. Deep working knowledge of the Brazilian market and an established local network across "),
         g("[ecosystem]"), t(".")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — Crypto.com regional launch across "), g("[N]"), t(" markets in "), g("[timeframe]")]),
      li([t("0 → 130,000", { bold: true }), t(" — Sui Foundation regional channel, from a standing start")]),
      li([g("[Revenue, usage or adoption metric you owned: from __ to __]")]),
      li([g("[A business plan you authored, and what changed as a result]")]),
      li([g("[A cross-functional programme you led, and its measurable result]")]),

      ...experience([
        li([g("[The KPIs and review cadence you owned]")]),
        li([g("[Field or partner enablement you built, and its uptake]")]),
      ]),

      H("Capabilities"),
      p([t("Business planning  ", { bold: true }),
         t("Regional business plans aligned to global objectives · market and segment prioritisation · turning strategy into measurable execution plans")]),
      p([t("Go-to-market  ", { bold: true }),
         t("Launch strategy and sequencing · solution and campaign plays · acquisition, expansion and adoption motions · partner and channel go-to-market")]),
      p([t("Business health  ", { bold: true }),
         t("KPI frameworks and scorecards · business reviews and reporting cadence · pipeline and funnel analysis · "), g("[tools]")]),
      p([t("Cross-functional leadership  ", { bold: true }),
         t("Influencing without authority across marketing, product, engineering, finance and partners · matrixed global and regional organisations · executive-level communication")]),
      p([t("Market intelligence  ", { bold: true }),
         t("Competitive, regulatory and customer insight synthesised into prioritised recommendations · "), g("[Brazilian industries you know]")]),
      p([t("AI  ", { bold: true }), g("[How you use AI in your work — one concrete example]")]),

      H("Education"),
      p([g("[Degree — institution — year]")]),
    ] }],
});

// ================= AIRWALLEX =================
const airwallex = new Document({
  creator: "Marc Lebreton", title: "Marc Lebreton — CV",
  description: "Airwallex Head of Sales, SME & Growth, Brazil",
  numbering, styles,
  sections: [{ properties: { page: { margin: { top: 880, bottom: 880, left: 1000, right: 1000 } } },
    children: [
      ...header("Commercial & Growth Leadership — Fintech, Brazil and Latin America", [
        t("English (fluent)  ·  Portuguese ", { size: 19 }), g("[level]"),
        t("  ·  Spanish ", { size: 19 }), g("[level]"),
        t("  ·  São Paulo based, available on site", { size: 19, color: MUTED })]),

      H("Profile"),
      p([t("Commercial and growth leader in Brazilian fintech with a zero-to-one track record in a hyper-growth environment. Built and ran the market-entry plan that took Crypto.com's Latin America business from nothing to "),
         t("6M+ customers and the #1 app position in the region", { bold: true }),
         t(", owning the full customer lifecycle — acquisition, activation, retention and expansion — across "),
         g("[N]"), t(" markets. "), g("[Team you built and led — size and functions]"),
         t(". Established network across the Brazilian fintech and payments ecosystem, and working knowledge of local rails including PIX and TED.")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ customers", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — Crypto.com regional launch across "), g("[N]"), t(" markets in "), g("[timeframe]")]),
      li([g("[Revenue or monetisation: took the region from __ to __]")]),
      li([g("[Team built: __ across SDR / AE / partnerships]")]),
      li([g("[Target attainment: __% against a __ target]")]),
      li([g("[Acquisition channels you built, with CAC and volume]")]),
      li([t("0 → 130,000", { bold: true }), t(" — Sui Foundation regional channel, from a standing start")]),

      ...experience([
        li([g("[Acquisition channels you built and scaled]")]),
        li([g("[Payments and local rails work: PIX, TED, banking partners]")]),
        li([g("[Largest deal or partnership you closed personally]")]),
      ]),

      H("Capabilities"),
      p([t("Commercial leadership  ", { bold: true }),
         t("Brazil commercial strategy · vertical and segment positioning · "), g("[team building and scaling]"),
         t(" · player/coach operating style · pipeline and forecast ownership")]),
      p([t("Growth  ", { bold: true }),
         t("Multi-channel customer acquisition · full lifecycle ownership across acquisition, activation, retention and expansion · growth experimentation · CAC and LTV analysis · partner and channel-led growth")]),
      p([t("Fintech & payments  ", { bold: true }),
         t("PIX · TED · "), g("[other rails]"), t(" · Brazilian fintech regulation · "), g("[ecosystem relationships]")]),
      p([t("Markets  ", { bold: true }), t("Brazil · "), g("[other markets you operated in]")]),
      p([t("Tools  ", { bold: true }), g("[CRM and analytics stack you've used]")]),

      H("Education"),
      p([g("[Degree — institution — year]")]),
    ] }],
});

(async () => {
  for (const [doc, path] of [
    [microsoft, "microsoft-application/Marc_Lebreton_CV_Microsoft.docx"],
    [airwallex, "airwallex-application/Marc_Lebreton_CV_Airwallex.docx"],
  ]) {
    fs.writeFileSync(path, await Packer.toBuffer(doc));
    console.log("wrote", path);
  }
})();
