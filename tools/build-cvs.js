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
        t("Portuguese (fluent)  ·  English (fluent)  ·  Spanish (fluent)", { size: 19 }),
        t("  ·  São Paulo based, available on site", { size: 19, color: MUTED })]),

      H("Profile"),
      p([t("Go-to-market and commercial strategy leader who has taken a global technology product into Brazil and Latin America from a standing start. Built and ran the market-entry plan that grew Crypto.com's LATAM business from zero to "),
         t("6M+ users and the #1 app position in the region", { bold: true }),
         t(", orchestrating across marketing, product, partnerships and operations inside a matrixed global organisation. Owns regional business planning, performance frameworks and field enablement end to end, and turns market, competitive and customer insight into prioritised, measurable action. Based in São Paulo, with deep working knowledge of the Brazilian market and an established local network across the technology, fintech and partner ecosystem.")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — owned the go-to-market plan for Crypto.com's regional market entry")]),
      li([t("0 → 130,000", { bold: true }), t(" — built Sui Foundation's regional channel from a standing start")]),
      li([t("Owned regional market prioritisation, launch sequencing and local execution, aligning the local plan to global objectives")]),
      li([t("Built and ran the performance cadence for the region — KPI definition, reporting and business reviews — and used it to reallocate effort against results")]),
      li([t("Translated competitive, regulatory and customer insight in Latin America into prioritised recommendations that shaped the regional plan")]),

      H("Experience"),

      role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" from a standing start")]),
      li([t("Orchestrated across product, marketing and partnerships to align the region behind a single growth plan")]),
      li([t("Produced market and competitive analysis for the region and translated it into prioritised action")]),
      li([t("Designed and ran partner and community enablement programmes, including the supporting materials")]),

      role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Owned the Latin America go-to-market plan end to end — market prioritisation, segment strategy, launch sequencing and local execution — scaling the region from zero to "),
          t("6M+ users", { bold: true }), t(" and the "), t("#1 app position", { bold: true })]),
      li([t("Built the regional business plan, aligned it to global objectives, and drove execution across marketing, product, partnerships and operations")]),
      li([t("Defined and ran the regional performance cadence: KPI frameworks, reporting and business reviews")]),
      li([t("Built field and partner enablement — positioning, materials and training — to support local execution at scale")]),
      li([t("Worked across a matrixed global organisation, influencing headquarters functions without direct authority")]),
      li([t("Represented the business externally in the Brazilian market")]),

      role([g("[Company]"), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([g("[Achievement]")]),
      li([g("[Achievement]")]),

      H("Capabilities"),
      p([t("Business planning  ", { bold: true }),
         t("Regional business plans aligned to global objectives · market and segment prioritisation · turning strategy into measurable execution plans")]),
      p([t("Go-to-market  ", { bold: true }),
         t("Launch strategy and sequencing · solution and campaign plays · acquisition, expansion and adoption motions · partner and channel go-to-market")]),
      p([t("Business health  ", { bold: true }),
         t("KPI frameworks and scorecards · business reviews and reporting cadence · pipeline and funnel analysis · "), g("[analytics stack]")]),
      p([t("Cross-functional leadership  ", { bold: true }),
         t("Influencing without authority across marketing, product, engineering, finance and partners · matrixed global and regional organisations · executive-level communication")]),
      p([t("Market intelligence  ", { bold: true }),
         t("Competitive, regulatory and customer insight synthesised into prioritised recommendations · Brazilian market and regulatory landscape · Financial Services")]),
      p([t("AI  ", { bold: true }), t("Uses AI daily for market and competitive research, content and enablement production at scale, and data analysis")]),

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
        t("English (fluent)  ·  Portuguese (fluent)  ·  Spanish (fluent)", { size: 19 }),
        t("  ·  São Paulo based, available on site", { size: 19, color: MUTED })]),

      H("Profile"),
      p([t("Commercial and growth leader in Brazilian fintech, with a zero-to-one track record in a hyper-growth environment. Built and ran the market-entry plan that grew Crypto.com's Latin America business from nothing to "),
         t("6M+ customers and the #1 app position in the region", { bold: true }),
         t(", building the regional commercial operation from scratch and owning the full customer lifecycle — acquisition, activation, retention and expansion. Comfortable operating as a player/coach: setting the commercial strategy and selling into the market directly. Based in São Paulo, with an established network across the Brazilian fintech and payments ecosystem and working knowledge of local rails including PIX and TED.")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ customers", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — built and executed Crypto.com's regional commercial and go-to-market plan")]),
      li([t("0 → 130,000", { bold: true }), t(" — built Sui Foundation's regional channel from a standing start")]),
      li([t("Built the multi-channel customer acquisition engine for the region from zero, then scaled it across markets")]),
      li([t("Owned the full customer lifecycle — acquisition, activation, retention and expansion — rather than acquisition alone")]),
      li([t("Established the partnerships and local payments relationships required to operate in the Brazilian market")]),
      li([g("[Team built and led: size and functions]"), t(" — "), g("[optional, but decisive for this role]")]),

      H("Experience"),

      role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" from a standing start")]),
      li([t("Originated and closed ecosystem and commercial partnerships across the region")]),
      li([t("Partnered with product and marketing to align acquisition and activation against a single growth plan")]),

      role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Owned the Latin America commercial and go-to-market strategy end to end, scaling from zero to "),
          t("6M+ customers", { bold: true }), t(" and the "), t("#1 app position", { bold: true })]),
      li([t("Built the regional commercial operation from scratch — segment strategy, positioning, acquisition channels and partner motions")]),
      li([t("Built and scaled multi-channel customer acquisition, and owned activation, retention and expansion alongside it")]),
      li([t("Negotiated and prioritised local payments and banking relationships, working with rails including PIX and TED")]),
      li([t("Navigated Brazilian fintech regulation as a constraint on product and commercial decisions")]),
      li([t("Represented the business externally to customers, partners and the Brazilian fintech market")]),

      role([g("[Company]"), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([g("[Achievement]")]),
      li([g("[Achievement]")]),

      H("Capabilities"),
      p([t("Commercial leadership  ", { bold: true }),
         t("Brazil commercial strategy · vertical and segment positioning · building commercial functions from zero · player/coach operating style · pipeline and forecast ownership")]),
      p([t("Growth  ", { bold: true }),
         t("Multi-channel customer acquisition · full lifecycle ownership across acquisition, activation, retention and expansion · growth experimentation · CAC and LTV analysis · partner and channel-led growth")]),
      p([t("Fintech & payments  ", { bold: true }),
         t("PIX · TED · boleto · fiat on/off-ramps · cross-border flows · Brazilian fintech regulation · established ecosystem relationships across payments and fintech")]),
      p([t("Markets  ", { bold: true }), t("Brazil · "), g("[other Latin American markets you covered]")]),
      p([t("Tools  ", { bold: true }), g("[CRM and analytics stack]")]),

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
