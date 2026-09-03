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
  description: "Regional Business Development Lead — Stablecoin Payments, Brazil",
  numbering: { config: [{ reference: "b", levels: [{
    level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: convertInchesToTwip(0.24), hanging: convertInchesToTwip(0.16) } } },
  }] }] },
  styles: { default: { document: {
    run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { line: 260 } } } } },
  sections: [{
    properties: { page: { margin: { top: 860, bottom: 860, left: 1000, right: 1000 } } },
    children: [

      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "MARC LEBRETON", bold: true, size: 40, color: INK, characterSpacing: 20 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [
        new TextRun({ text: "Business Development — Payments & Digital Assets, Brazil and LATAM", size: 21, color: MUTED })] }),
      new Paragraph({ spacing: { after: 30 }, children: [
        t("São Paulo, Brazil  ·  write2marc@gmail.com  ·  ", { size: 19, color: MUTED }),
        g("[phone]"), t("  ·  ", { size: 19, color: MUTED }), g("[LinkedIn]")] }),
      new Paragraph({ spacing: { after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: "1A1A1A", space: 6 } },
        children: [ t("Portuguese (fluent)  ·  English (fluent)  ·  Spanish (fluent)", { size: 19 }),
          t("  ·  Based in Brazil, available to travel", { size: 19, color: MUTED }) ] }),

      H("Profile"),
      p([t("Commercial and business development operator in Brazilian fintech and digital assets, with a completed zero-to-one market build behind me. Took Crypto.com's Latin America business from nothing to "),
         t("6M+ users and the #1 app position in the region", { bold: true }),
         t(", running commercial strategy independently from pipeline through to signed agreement, and negotiating the local partnerships — payment providers, banking rails, distribution channels — that made the product operable in market. Fluent in how stablecoin settlement and treasury flows actually work, and in the Brazilian rails, QR standards and acquiring landscape a merchant acceptance play depends on. Comfortable owning a single market end to end and deciding how it gets built.")]),

      H("Selected Outcomes"),
      li([t("0 → 6,000,000+ users", { bold: true }), t(" and "), t("#1 app in Latin America", { bold: true }),
          t(" — Crypto.com regional market entry, built on locally negotiated partnerships rather than paid acquisition alone")]),
      li([t("0 → 130,000", { bold: true }), t(" — built Sui Foundation's regional channel from a standing start")]),
      li([t("Negotiated and closed the local payment, banking and distribution partnerships required to launch a regulated digital-asset product across multiple Latin American markets")]),
      li([t("Ran commercial strategy independently — sourcing, qualifying, pitching, negotiating and closing, then driving partners to live and transacting")]),
      li([g("[Payments deals: acquirers, PSPs, aggregators or merchants you personally brought in — name them and the volume]")]),
      li([g("[Largest contract closed: value, term, counterparty type]")]),

      H("Experience"),

      role([t("Sui Foundation", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Grew the regional channel from "), t("0 to 130,000", { bold: true }), t(" from a standing start")]),
      li([t("Sourced, negotiated and closed ecosystem and commercial partnerships across the region, then drove them to launch")]),
      li([t("Fed in-market feedback back into product and ecosystem priorities")]),

      role([t("Crypto.com", { bold: true, size: 22 }), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([t("Owned Latin America market entry end to end, scaling from zero to "),
          t("6M+ users", { bold: true }), t(" and the "), t("#1 app position", { bold: true })]),
      li([t("Ran the full commercial cycle — sourcing and qualifying counterparties, pitching, negotiating terms, and closing — across payment providers, banking partners and distribution channels")]),
      li([t("Negotiated and prioritised local payment rails and settlement arrangements, working with PIX and local banking infrastructure")]),
      li([t("Coordinated with legal and compliance on licensing, settlement and regulatory requirements in each market before launch")]),
      li([t("Drove signed partners to integration and live transaction volume, treating time-to-first-transaction as the metric that mattered")]),
      li([t("Channelled merchant and partner feedback into product priorities — flows, settlement currency and reconciliation")]),
      li([g("[Team led, if any: size and functions]")]),

      role([g("[Company]"), t(" — "), g("[job title]")]),
      meta([g("[city]"), " · ", g("[start]"), " – ", g("[end]")]),
      li([g("[Any merchant acquiring, PSP, aggregator or payments-adjacent experience belongs here and should lead — it is the single requirement this CV is thinnest on]")]),
      li([g("[Achievement]")]),

      H("Market Knowledge — Brazilian Payments"),
      p([t("Rails and standards  ", { bold: true }),
         t("PIX, including static and dynamic QR (Pix QR Code), Pix Automático and Pix Parcelado · boleto · TED · card rails and installment (parcelamento) dynamics · settlement and reconciliation flows")]),
      p([t("Acceptance landscape  ", { bold: true }),
         t("Acquirers and sub-acquirers — Cielo, Rede, Getnet, Stone, PagBank, SumUp, InfinitePay · gateways and PSPs — Pagar.me, Adyen, dLocal, EBANX, Zoop, Malga · marketplace and wallet acceptance via Mercado Pago and PicPay")]),
      p([t("Regulatory  ", { bold: true }),
         t("BCB Resolutions 519, 520 and 521, in force 2 February 2026 — foreign-currency stablecoin operations reclassified as FX, bringing IOF-Câmbio into merchant settlement economics, mandatory BCB reporting from 4 May 2026, and PSAV authorisation required to intermediate or custody crypto assets · payment institution (IP) authorisation · LGPD")]),
      p([t("Digital assets  ", { bold: true }),
         t("Stablecoin settlement and treasury flows · multi-chain and multi-wallet operations · on- and off-ramp economics · "),
         g("[specific stablecoin payment products you've worked with, if any]")]),

      H("Capabilities"),
      p([t("Commercial  ", { bold: true }),
         t("Independent pipeline build · sourcing and qualification · commercial negotiation including multi-year and enterprise structures · contracting with legal and compliance in the loop · partner activation and volume management")]),
      p([t("Markets  ", { bold: true }), t("Brazil · "), g("[other LATAM markets you covered]")]),
      p([t("Tools  ", { bold: true }), g("[CRM and analytics stack you've used]")]),

      H("Education"),
      p([g("[Degree — institution — year]")]),
    ],
  }],
});

Packer.toBuffer(doc).then((b) => { fs.writeFileSync("Marc_Lebreton_CV_StablecoinBD.docx", b); console.log("written", b.length); });
