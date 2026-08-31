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

const doc = new Document({
  creator: "Marc Lebreton", title: "Marc Lebreton — CV",
  description: "ARQ Growth Manager, Partnerships, Brazil",
  numbering: { config: [{ reference: "b", levels: [{
    level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: convertInchesToTwip(0.24), hanging: convertInchesToTwip(0.16) } } },
  }] }] },
  styles: { default: { document: {
    run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 262 } } } } },
  sections: [{
    properties: { page: { margin: { top: 880, bottom: 880, left: 1000, right: 1000 } } },
    children: [

      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "Growth & Partnerships — Fintech, Brazil and Latin America", size: 21, color: MUTED })] }),
      new Paragraph({ spacing: { after: 30 }, children: [
        t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
        g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[LinkedIn]")] }),
      new Paragraph({ spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [ t("Portuguese (fluent)  ·  English (fluent)  ·  Spanish (fluent)", { size: 19 }),
          t("  ·  São Paulo based, available in-office", { size: 19, color: MUTED }) ] }),

      H("Profile"),
      p([t("Partnership-led growth operator in Latin American consumer fintech. Built and ran the market-entry plan that took Crypto.com's LATAM business from zero to "),
         t("6M+ users and the #1 app position in the region", { bold: true }),
         t(", acquiring users through performance-based partner channels rather than paid media alone — sourcing partners, negotiating variable commercial terms, driving them live fast, and cutting the ones that didn't produce. Established network across the Brazilian creator, KOL and fintech ecosystem, and comfortable structuring partner deals inside Brazil's LGPD, CONAR and tax requirements. Hands-on, quota-minded, and used to owning a number end to end.")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — Crypto.com regional launch, built substantially on partner and creator acquisition channels")]),
      li([t("0 → 130,000", { bold: true }), t(" — built Sui Foundation's regional channel from a standing start")]),
      li([t("Built partner acquisition channels from zero in a market with no existing programme — sourcing, contracting, activation and ongoing management")]),
      li([t("Negotiated and closed variable, performance-based partner agreements and ran the payment operation behind them")]),
      li([t("Killed underperforming channels and reallocated spend on channel-level data rather than on relationships")]),
      li([g("[Deals closed: __ partners signed in __ months]"), t("  ·  "), g("[Volume: __ activated users / __ % of new signups]")]),

      H("Experience"),

      role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" from a standing start")]),
      li([t("Sourced and closed ecosystem and creator partnerships across the region, then drove them to launch")]),
      li([t("Owned the ongoing partner relationship — activation, troubleshooting, engagement and performance review")]),

      role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Owned Latin America growth end to end, scaling from zero to "),
          t("6M+ users", { bold: true }), t(" and the "), t("#1 app position", { bold: true })]),
      li([t("Built the partner and creator acquisition channel from nothing — prospecting, pitching, negotiating and closing deals with content creators, communities and affiliate operators")]),
      li([t("Structured performance-based compensation — CPA, revenue share and hybrid models — and ran the partner payment process alongside finance")]),
      li([t("Drove newly signed partners to publish and produce activated users quickly, treating time-to-first-activation as the metric that mattered")]),
      li([t("Monitored results partner by partner and channel by channel, discontinuing what underperformed and concentrating budget behind what worked")]),
      li([t("Worked with legal and finance on partner contracting, disclosure obligations and cross-border payment operations")]),

      role([g("[Company]"), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([g("[Achievement]")]),
      li([g("[Achievement]")]),

      H("Capabilities"),
      p([t("Partnerships  ", { bold: true }),
         t("Sourcing and prospecting · commercial negotiation · CPA, revenue-share and hybrid performance models · contracting · onboarding and activation · ongoing relationship and volume management · partner payment operations")]),
      p([t("Channels  ", { bold: true }),
         t("Content creators and KOLs · communities · affiliate operators · fintech and ecosystem partners · "),
         g("[financial advisors · points and miles programmes — add if you have real reach here]")]),
      p([t("Growth  ", { bold: true }),
         t("Full-funnel ownership from acquisition to activation · channel-level performance analysis · cohort and CAC/LTV analysis · rapid kill-or-scale decisions · experimentation")]),
      p([t("Brazil regulatory  ", { bold: true }),
         t("LGPD · CONAR influencer and disclosure rules, including the June 2026 guide and its joint-liability provisions for advertisers · partner contracting, withholding and tax treatment on performance payments")]),
      p([t("Markets  ", { bold: true }), t("Brazil · "), g("[other LATAM markets you covered]")]),
      p([t("Tools  ", { bold: true }), g("[CRM and analytics stack you've used]")]),

      H("Education"),
      p([g("[Degree — institution — year]")]),
    ],
  }],
});

Packer.toBuffer(doc).then((b) => { fs.writeFileSync("Marc_Lebreton_CV_ARQ.docx", b); console.log("written", b.length); });
