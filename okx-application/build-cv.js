const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, LevelFormat, convertInchesToTwip,
} = require("docx");
const fs = require("fs");

// ---- helpers -------------------------------------------------------------
const INK = "1A1A1A";
const MUTED = "555555";
const GAP = "C00000"; // placeholders Marc must fill — deliberately loud

// a placeholder run
const g = (t) => new TextRun({ text: t, color: GAP, bold: true });
const t = (text, opts = {}) => new TextRun({ text, color: INK, ...opts });

const sectionHeading = (text) =>
  new Paragraph({
    spacing: { before: 260, after: 110 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 4 } },
    children: [new TextRun({ text, bold: true, size: 21, color: INK, allCaps: true, characterSpacing: 30 })],
  });

// role header: Title — Company, then a right-ish line for location/dates
const role = (children) =>
  new Paragraph({ spacing: { before: 170, after: 20 }, children });

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

// ---- document ------------------------------------------------------------
const doc = new Document({
  creator: "Marc Lebreton",
  title: "Marc Lebreton — CV",
  description: "CV targeted at OKX (Senior) Product Director, Local Growth",
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
    default: {
      document: { run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 264 } } },
    },
  },
  sections: [{
    properties: { page: { margin: { top: 900, bottom: 900, left: 1000, right: 1000 } } },
    children: [

      // ---------------- HEADER ----------------
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })],
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [new TextRun({
          text: "Product & Growth Leader — Emerging Markets, Crypto & Fintech",
          size: 21, color: MUTED,
        })],
      }),
      new Paragraph({
        spacing: { after: 30 },
        children: [
          t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
          g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[linkedin.com/in/…]"),
        ],
      }),
      new Paragraph({
        spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [
          t("Portuguese ", { size: 19 }), g("[native / fluent — confirm]"),
          t("  ·  Spanish ", { size: 19 }), g("[fluent — confirm]"),
          t("  ·  English (fluent)", { size: 19 }),
          t("  ·  ", { size: 19, color: MUTED }), g("[Mandarin? — OKX lists it as a strong plus for HQ collaboration]"),
        ],
      }),

      // ---------------- PROFILE ----------------
      sectionHeading("Profile"),
      body([
        t("Product and growth leader with a completed 0-to-1 cycle in Latin America's crypto market. Led Crypto.com's LATAM launch from zero to "),
        t("6M+ users and the #1 app position in the region", { bold: true }),
        t(", owning the funnel from acquisition through activation across "),
        g("[N]"),
        t(" markets. Built Sui Foundation's regional community channel from "),
        t("0 to 130,000", { bold: true }),
        t(". "),
        g("[X]"),
        t(" years building in emerging markets, with an established local network across creators/KOLs, payment providers and "),
        g("[regulatory / exchange / fintech]"),
        t(" contacts in Brazil and Hispanic LATAM, and working knowledge of local rails including PIX"),
        g("[, boleto, SPEI, Mercado Pago — keep only what's true]"),
        t("."),
      ]),

      // ---------------- SELECTED RESULTS ----------------
      sectionHeading("Selected Results"),
      body([
        t("OKX asks candidates to speak to the funnel with data. Fill these in with real figures — this block is the single highest-leverage part of the CV for this role.", { italics: true, size: 18, color: MUTED }),
      ]),
      bullet([t("0 → 6,000,000+ users", { bold: true }), t(" — Crypto.com LATAM launch; "), t("#1 app in the region", { bold: true }), t(".")]),
      bullet([t("0 → 130,000", { bold: true }), t(" — Sui Foundation regional community channel.")]),
      bullet([g("[Registration → first trade conversion: __% → __%]"), t(" — and the product change that moved it.")]),
      bullet([g("[Time-to-first-trade: reduced from __ to __]"), t(" — OKX names this metric explicitly in the JD.")]),
      bullet([g("[CAC: $__ blended / $__ by channel]"), t("  ·  "), g("[LTV: $__]"), t("  ·  "), g("[LTV:CAC ratio]")]),
      bullet([g("[D7 retention: __%  ·  D30 retention: __%]")]),
      bullet([g("[Reactivation: __% of dormant users returned via __]")]),
      bullet([g("[A/B tests run: __ per quarter; __ shipped; largest single lift: __%]")]),

      // ---------------- EXPERIENCE ----------------
      sectionHeading("Experience"),

      role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[Title]")]),
      meta([g("[Location]"), " · ", g("[Start – End]")]),
      bullet([t("Built the regional community channel from "), t("0 to 130,000", { bold: true }), t(" — "), g("[which channel, which markets, and what it fed into: sign-ups, wallet activations, transacting users]")]),
      bullet([g("[Product intervention you drove — a flow you changed, a feature you specced, an integration you shipped. OKX is screening for product work, not campaign work.]")]),
      bullet([g("[Partnership or integration you led — KOLs, payment provider, exchange, dApp]")]),
      bullet([g("[Measurable outcome with a number]")]),

      role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[Title]")]),
      meta([g("[Location]"), " · ", g("[Start – End]")]),
      bullet([t("Led the LATAM market launch from zero to "), t("6M+ users", { bold: true }), t(", reaching "), t("#1 app in the region", { bold: true }), t(" — "), g("[which countries, over what period, and what you personally owned vs the wider team]")]),
      bullet([g("[Fiat on-ramp / local payment work: PIX, boleto, SPEI, local bank rails — what you integrated or prioritised and what it did to deposit conversion]")]),
      bullet([g("[Funnel ownership: the specific stage(s) you owned and the metric movement — onboarding, KYC pass rate, first deposit, first trade]")]),
      bullet([g("[Localisation / regulatory work: what changed in the product for LATAM specifically]")]),
      bullet([g("[Team: did you manage or mentor anyone? Needed to qualify for the SENIOR Director level.]")]),

      role([g("[Earlier role — company]"), t(" — "), g("[Title]")]),
      meta([g("[Location]"), " · ", g("[Start – End]")]),
      bullet([g("[2–3 bullets. If any of this is agency, marketing or community-only work, frame it around funnel metrics and product decisions rather than reach or audience size.]")]),

      // ---------------- SKILLS ----------------
      sectionHeading("Capabilities"),
      body([
        t("Growth  ", { bold: true }),
        t("Full-funnel ownership (acquisition → activation → retention · reactivation) · growth experimentation and A/B testing · cohort and funnel analysis · lifecycle and CRM · channel discovery, integration and scaling · incentive and referral programme design"),
      ]),
      body([
        t("Product  ", { bold: true }),
        t("Product strategy and success-metric definition · roadmap prioritisation · "),
        g("[PRDs / specs — confirm]"),
        t(" · working directly with engineering and design · "),
        g("[experimentation tooling you've used: Amplitude, Mixpanel, Braze, Optimizely, GA4, SQL…]"),
      ]),
      body([
        t("Markets  ", { bold: true }),
        t("Brazil · "),
        g("[Mexico · Argentina · Colombia · Chile — keep only markets you've actually operated in]"),
      ]),
      body([
        t("Payments & crypto  ", { bold: true }),
        t("PIX · "),
        g("[boleto · SPEI · Mercado Pago · local bank transfer]"),
        t(" · fiat on/off-ramp flows · "),
        g("[spot · derivatives · wallet · custody · token listings — keep what's true]"),
      ]),
      body([
        t("Network  ", { bold: true }),
        g("[Name the categories you can genuinely activate: KOLs and creator agencies, payment providers, local exchanges, regulators, media. OKX asks for this explicitly.]"),
      ]),

      // ---------------- EDUCATION ----------------
      sectionHeading("Education"),
      body([g("[Degree — Institution — Year]")]),
      body([g("[Any relevant certification — optional]")]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("Marc_Lebreton_CV_OKX.docx", buf);
  console.log("written:", buf.length, "bytes");
});
