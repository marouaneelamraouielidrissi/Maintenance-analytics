const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
pres.title = "Maintenance Analytics — Guide Utilisateur";
pres.author = "Bureau de Méthode Daoui";

// Color palette
const C = {
  primary: "1e3a8a",    // dark blue
  secondary: "dbeafe",  // light blue
  accent: "f59e0b",     // amber/gold
  white: "FFFFFF",
  lightBg: "f0f9ff",    // placeholder fill
  borderBlue: "93c5fd", // placeholder border
  textMuted: "64748b",  // muted text
  darkText: "1e293b",   // dark text
  midBlue: "3b82f6",    // medium blue for accents
};

const W = 13.33;
const H = 7.5;

// ─────────────────────────────────────────────────────────────
// SLIDE 1 — Cover (dark background)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "1e3a8a" };

  // Top accent bar in amber
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.08,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });

  // OCP branding strip on left
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.08, w: 0.5, h: H - 0.08,
    fill: { color: "172554" }, line: { color: "172554" }
  });

  // Decorative circle (large, background)
  slide.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.5, w: 6, h: 6,
    fill: { color: "1d4ed8", transparency: 70 },
    line: { color: "1d4ed8", transparency: 70 }
  });

  // OCP Label
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.2, w: 1.6, h: 0.45,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });
  slide.addText("OCP GROUP", {
    x: 0.8, y: 1.2, w: 1.6, h: 0.45,
    fontSize: 11, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  // Main title
  slide.addText("Maintenance", {
    x: 0.8, y: 2.0, w: 10, h: 1.1,
    fontSize: 64, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "left", valign: "bottom", margin: 0
  });
  slide.addText("Analytics", {
    x: 0.8, y: 3.1, w: 10, h: 1.1,
    fontSize: 64, fontFace: "Calibri", color: "f59e0b",
    bold: true, align: "left", valign: "top", margin: 0
  });

  // Subtitle
  slide.addText("Guide Utilisateur  ·  Bureau de Méthode Daoui  ·  SAP PM", {
    x: 0.8, y: 4.5, w: 11, h: 0.5,
    fontSize: 16, fontFace: "Calibri", color: "bfdbfe",
    align: "left", margin: 0
  });

  // Separator line
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 4.35, w: 6, h: 0.04,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });

  // Footer
  slide.addText("2026 · Maintenance Analytics · OCP", {
    x: 0.8, y: 6.9, w: 11.5, h: 0.4,
    fontSize: 10, fontFace: "Calibri", color: "93c5fd",
    align: "left", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 2 — Vue d'ensemble: "Une application, 3 modules"
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  // Title
  slide.addText("Une application, 3 modules", {
    x: 0.5, y: 0.35, w: 12, h: 0.75,
    fontSize: 36, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Subtitle
  slide.addText("Maintenance Analytics — Bureau de Méthode Daoui", {
    x: 0.5, y: 1.05, w: 12, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: "64748b",
    align: "left", margin: 0
  });

  // 3 module cards
  const cards = [
    { icon: "📊", title: "Dashboard", desc: "Indicateurs SAP PM mensuels et annuels", color: "1e3a8a", light: "dbeafe" },
    { icon: "📅", title: "Plan de charge", desc: "OT planifiés par semaine", color: "0369a1", light: "e0f2fe" },
    { icon: "⚡", title: "Moteurs Électriques", desc: "Gestion des demandes de moteurs", color: "b45309", light: "fef3c7" },
  ];

  const cardW = 3.6;
  const cardH = 3.8;
  const startX = 0.7;
  const gapX = 0.55;

  cards.forEach((card, i) => {
    const cx = startX + i * (cardW + gapX);
    const cy = 1.65;

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: "FFFFFF" },
      line: { color: card.color, width: 2 },
      shadow: { type: "outer", color: "000000", blur: 10, offset: 3, angle: 135, opacity: 0.1 }
    });

    // Top color band
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 0.7,
      fill: { color: card.color }, line: { color: card.color }
    });

    // Icon circle background
    slide.addShape(pres.shapes.OVAL, {
      x: cx + cardW / 2 - 0.5, y: cy + 0.35, w: 1.0, h: 1.0,
      fill: { color: card.light }, line: { color: card.light }
    });

    // Icon emoji
    slide.addText(card.icon, {
      x: cx + cardW / 2 - 0.5, y: cy + 0.38, w: 1.0, h: 1.0,
      fontSize: 28, align: "center", valign: "middle", margin: 0
    });

    // Module title
    slide.addText(card.title, {
      x: cx + 0.2, y: cy + 1.55, w: cardW - 0.4, h: 0.6,
      fontSize: 20, fontFace: "Calibri", color: card.color,
      bold: true, align: "center", margin: 0
    });

    // Description
    slide.addText(card.desc, {
      x: cx + 0.25, y: cy + 2.3, w: cardW - 0.5, h: 1.0,
      fontSize: 14, fontFace: "Calibri", color: "334155",
      align: "center", valign: "top", margin: 0
    });
  });

  // Footer bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("2 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 3 — Module Dashboard
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  // Module badge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 1.6, h: 0.35,
    fill: { color: "dbeafe" }, line: { color: "dbeafe" }
  });
  slide.addText("📊 DASHBOARD", {
    x: 0.5, y: 0.28, w: 1.6, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  // Title
  slide.addText("Dashboard — Indicateurs SAP PM", {
    x: 0.5, y: 0.72, w: 12, h: 0.7,
    fontSize: 34, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Left column: bullets
  const bullets = [
    "Connexion automatique aux données SAP (Google Sheets)",
    "Indicateurs KPI : taux de réalisation, OT planifiés, OT réalisés",
    "Graphiques mensuels et vue annuelle",
    "Sélection du mois et de l'année",
  ];

  bullets.forEach((b, i) => {
    const by = 1.65 + i * 0.8;
    // Blue square bullet marker
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: by + 0.07, w: 0.15, h: 0.15,
      fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
    });
    slide.addText(b, {
      x: 0.78, y: by, w: 5.4, h: 0.7,
      fontSize: 14, fontFace: "Calibri", color: "1e293b",
      align: "left", valign: "middle", margin: 0
    });
  });

  // KPI highlight boxes
  const kpis = [
    { label: "Taux de réalisation", color: "1e3a8a" },
    { label: "OT Planifiés", color: "0369a1" },
    { label: "OT Réalisés", color: "b45309" },
  ];
  kpis.forEach((k, i) => {
    const bx = 0.5 + i * 2.0;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: 5.0, w: 1.8, h: 0.6,
      fill: { color: k.color }, line: { color: k.color }
    });
    slide.addText(k.label, {
      x: bx, y: 5.0, w: 1.8, h: 0.6,
      fontSize: 10, fontFace: "Calibri", color: "FFFFFF",
      bold: true, align: "center", valign: "middle", margin: 0
    });
  });

  // Right column: screenshot placeholder
  const px = 7.0, py = 1.35, pw = 5.8, ph = 3.8;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: px, y: py, w: pw, h: ph,
    fill: { color: "f0f9ff" },
    line: { color: "93c5fd", width: 1.5, dashType: "dash" }
  });
  slide.addText("Capture d'écran — Dashboard", {
    x: px, y: py + ph / 2 - 0.25, w: pw, h: 0.5,
    fontSize: 13, fontFace: "Calibri", color: "64748b",
    italic: true, align: "center", valign: "middle", margin: 0
  });
  // Camera icon placeholder
  slide.addText("📷", {
    x: px, y: py + ph / 2 - 0.85, w: pw, h: 0.5,
    fontSize: 24, align: "center", valign: "middle", margin: 0
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("3 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 4 — Plan de charge (part 1)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 1.8, h: 0.35,
    fill: { color: "e0f2fe" }, line: { color: "e0f2fe" }
  });
  slide.addText("📅 PLAN DE CHARGE", {
    x: 0.5, y: 0.28, w: 1.8, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "0369a1",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Plan de charge — Consultation des OT", {
    x: 0.5, y: 0.72, w: 12, h: 0.7,
    fontSize: 34, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Left: screenshot placeholder
  const px = 0.5, py = 1.55, pw = 5.5, ph = 3.8;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: px, y: py, w: pw, h: ph,
    fill: { color: "f0f9ff" },
    line: { color: "93c5fd", width: 1.5, dashType: "dash" }
  });
  slide.addText("📷", {
    x: px, y: py + ph / 2 - 0.6, w: pw, h: 0.5,
    fontSize: 24, align: "center", valign: "middle", margin: 0
  });
  slide.addText("Capture d'écran — Plan de charge", {
    x: px, y: py + ph / 2 - 0.1, w: pw, h: 0.5,
    fontSize: 13, fontFace: "Calibri", color: "64748b",
    italic: true, align: "center", valign: "middle", margin: 0
  });

  // Right: bullets
  const bulletsRight = [
    "Sélection de l'année et de la semaine",
    "Chargement automatique depuis SAP",
  ];

  bulletsRight.forEach((b, i) => {
    const by = 1.65 + i * 0.85;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.7, y: by + 0.07, w: 0.15, h: 0.15,
      fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
    });
    slide.addText(b, {
      x: 6.98, y: by, w: 5.8, h: 0.7,
      fontSize: 14, fontFace: "Calibri", color: "1e293b",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Table columns header
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.7, y: 3.45, w: 6.2, h: 0.4,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });
  slide.addText("Colonnes du tableau :", {
    x: 6.7, y: 3.45, w: 6.2, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "left", valign: "middle", margin: 8
  });

  const cols = ["Date", "Poste de travail", "Objet technique", "Ordre", "Avis", "Type d'ordre", "Désignation"];
  cols.forEach((col, i) => {
    const row = i;
    const cy = 3.85 + row * 0.3;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.7, y: cy, w: 6.2, h: 0.3,
      fill: { color: i % 2 === 0 ? "f8fafc" : "FFFFFF" },
      line: { color: "e2e8f0", width: 0.5 }
    });
    slide.addText(col, {
      x: 6.9, y: cy, w: 6.0, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: "334155",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("4 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 5 — Plan de charge (Filtres & Export)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 1.8, h: 0.35,
    fill: { color: "e0f2fe" }, line: { color: "e0f2fe" }
  });
  slide.addText("📅 PLAN DE CHARGE", {
    x: 0.5, y: 0.28, w: 1.8, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "0369a1",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Plan de charge — Filtres & Export", {
    x: 0.5, y: 0.72, w: 12, h: 0.7,
    fontSize: 34, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // 3 feature cards
  const featureCards = [
    {
      icon: "🔍",
      title: "Filtrer",
      items: ["Poste de travail", "Type d'ordre", "Date", "Objet technique"],
      color: "1e3a8a",
      light: "dbeafe"
    },
    {
      icon: "📊",
      title: "Exporter Excel",
      items: ["Fichier .xlsx nommé automatiquement", "Plan_de_charge_YYYY_SNN.xlsx"],
      color: "166534",
      light: "dcfce7"
    },
    {
      icon: "📄",
      title: "Exporter PDF",
      items: ["Impression optimisée A4 paysage", "Avec entête de document"],
      color: "7c2d12",
      light: "ffedd5"
    },
  ];

  const cardW2 = 3.7;
  const cardH2 = 4.3;
  const startX2 = 0.6;
  const gapX2 = 0.5;

  featureCards.forEach((card, i) => {
    const cx = startX2 + i * (cardW2 + gapX2);
    const cy = 1.65;

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW2, h: cardH2,
      fill: { color: "FFFFFF" },
      line: { color: card.color, width: 1.5 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.1 }
    });

    // Top band
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW2, h: 0.6,
      fill: { color: card.light }, line: { color: card.light }
    });

    // Icon + title
    slide.addText(card.icon + "  " + card.title, {
      x: cx + 0.15, y: cy + 0.08, w: cardW2 - 0.3, h: 0.5,
      fontSize: 18, fontFace: "Calibri", color: card.color,
      bold: true, align: "left", valign: "middle", margin: 0
    });

    // Items
    card.items.forEach((item, j) => {
      const iy = cy + 0.85 + j * 0.7;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: cx + 0.2, y: iy + 0.12, w: 0.12, h: 0.12,
        fill: { color: card.color }, line: { color: card.color }
      });
      slide.addText(item, {
        x: cx + 0.42, y: iy, w: cardW2 - 0.55, h: 0.6,
        fontSize: 13, fontFace: "Calibri", color: "1e293b",
        align: "left", valign: "middle", margin: 0
      });
    });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("5 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 6 — Moteurs Électriques — Nouvelle demande
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 2.2, h: 0.35,
    fill: { color: "fef3c7" }, line: { color: "fef3c7" }
  });
  slide.addText("⚡ MOTEURS ÉLECTRIQUES", {
    x: 0.5, y: 0.28, w: 2.2, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "b45309",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Moteurs Électriques — Soumettre une demande", {
    x: 0.5, y: 0.72, w: 12.3, h: 0.7,
    fontSize: 30, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Left: 3 types
  slide.addText("3 types de demande :", {
    x: 0.5, y: 1.58, w: 6.5, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  const types = [
    { label: "Changement", desc: "Remplacement d'un moteur existant", color: "1e3a8a", bg: "dbeafe" },
    { label: "Pose", desc: "Installation d'un nouveau moteur", color: "166534", bg: "dcfce7" },
    { label: "Réparation", desc: "Envoi d'un moteur en réparation", color: "7c2d12", bg: "ffedd5" },
  ];

  types.forEach((t, i) => {
    const ty = 2.1 + i * 0.85;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: ty, w: 6.2, h: 0.72,
      fill: { color: t.bg }, line: { color: t.color, width: 1 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: ty, w: 0.08, h: 0.72,
      fill: { color: t.color }, line: { color: t.color }
    });
    slide.addText(t.label, {
      x: 0.72, y: ty + 0.04, w: 1.4, h: 0.32,
      fontSize: 13, fontFace: "Calibri", color: t.color,
      bold: true, align: "left", valign: "middle", margin: 0
    });
    slide.addText(t.desc, {
      x: 0.72, y: ty + 0.36, w: 5.8, h: 0.3,
      fontSize: 12, fontFace: "Calibri", color: "475569",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Informations requises
  slide.addText("Informations requises :", {
    x: 0.5, y: 4.8, w: 6.5, h: 0.35,
    fontSize: 13, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  const infos = ["Installation", "Objet technique", "Puissance", "Tension", "Vitesse", "Anomalie", "Photos", "Signature"];
  infos.forEach((info, i) => {
    const ix = 0.5 + (i % 4) * 1.6;
    const iy = 5.2 + Math.floor(i / 4) * 0.45;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: ix, y: iy, w: 1.45, h: 0.35,
      fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
    });
    slide.addText(info, {
      x: ix, y: iy, w: 1.45, h: 0.35,
      fontSize: 10, fontFace: "Calibri", color: "FFFFFF",
      bold: false, align: "center", valign: "middle", margin: 0
    });
  });

  // Right: screenshot placeholder
  const px = 7.0, py = 1.4, pw = 5.8, ph = 4.2;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: px, y: py, w: pw, h: ph,
    fill: { color: "f0f9ff" },
    line: { color: "93c5fd", width: 1.5, dashType: "dash" }
  });
  slide.addText("📷", {
    x: px, y: py + ph / 2 - 0.65, w: pw, h: 0.5,
    fontSize: 24, align: "center", valign: "middle", margin: 0
  });
  slide.addText("Capture d'écran — Formulaire", {
    x: px, y: py + ph / 2 - 0.15, w: pw, h: 0.5,
    fontSize: 13, fontFace: "Calibri", color: "64748b",
    italic: true, align: "center", valign: "middle", margin: 0
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("6 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 7 — Moteurs Électriques — Suivi
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 2.2, h: 0.35,
    fill: { color: "fef3c7" }, line: { color: "fef3c7" }
  });
  slide.addText("⚡ MOTEURS ÉLECTRIQUES", {
    x: 0.5, y: 0.28, w: 2.2, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "b45309",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Moteurs Électriques — Suivi des demandes", {
    x: 0.5, y: 0.72, w: 12, h: 0.7,
    fontSize: 30, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Left: screenshot placeholder
  const px = 0.5, py = 1.55, pw = 5.5, ph = 4.0;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: px, y: py, w: pw, h: ph,
    fill: { color: "f0f9ff" },
    line: { color: "93c5fd", width: 1.5, dashType: "dash" }
  });
  slide.addText("📷", {
    x: px, y: py + ph / 2 - 0.6, w: pw, h: 0.5,
    fontSize: 24, align: "center", valign: "middle", margin: 0
  });
  slide.addText("Capture d'écran — Suivi", {
    x: px, y: py + ph / 2 - 0.1, w: pw, h: 0.5,
    fontSize: 13, fontFace: "Calibri", color: "64748b",
    italic: true, align: "center", valign: "middle", margin: 0
  });

  // Right: bullets
  const bullets7 = [
    "Recherche par Installation, Matricule, Demandeur, Statut…",
    "Statuts : Envoyée → Approuvée / Refusée",
    "État de réparation : En attente → En réparation → Réparé",
    "Notifications email automatiques à chaque changement",
  ];

  bullets7.forEach((b, i) => {
    const by = 1.65 + i * 0.88;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 6.7, y: by + 0.07, w: 0.15, h: 0.15,
      fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
    });
    slide.addText(b, {
      x: 6.98, y: by, w: 5.9, h: 0.75,
      fontSize: 13, fontFace: "Calibri", color: "1e293b",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Status flow diagram
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 6.7, y: 5.25, w: 6.2, h: 0.35,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });
  slide.addText("Flux des statuts", {
    x: 6.7, y: 5.25, w: 6.2, h: 0.35,
    fontSize: 11, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  const statuses = [
    { label: "Envoyée", color: "64748b" },
    { label: "Approuvée", color: "166534" },
    { label: "Refusée", color: "991b1b" },
  ];
  statuses.forEach((s, i) => {
    const sx = 6.7 + i * 2.07;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: 5.6, w: 1.8, h: 0.4,
      fill: { color: s.color }, line: { color: s.color }
    });
    slide.addText(s.label, {
      x: sx, y: 5.6, w: 1.8, h: 0.4,
      fontSize: 11, fontFace: "Calibri", color: "FFFFFF",
      bold: true, align: "center", valign: "middle", margin: 0
    });
    if (i < statuses.length - 1) {
      slide.addText("→", {
        x: sx + 1.8, y: 5.6, w: 0.27, h: 0.4,
        fontSize: 14, color: "94a3b8", align: "center", valign: "middle", margin: 0
      });
    }
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("7 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 8 — Administration
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.28, w: 1.5, h: 0.35,
    fill: { color: "dbeafe" }, line: { color: "dbeafe" }
  });
  slide.addText("🔐 ADMINISTRATION", {
    x: 0.5, y: 0.28, w: 1.5, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "center", valign: "middle", margin: 0
  });

  slide.addText("Administration — Panneau de gestion", {
    x: 0.5, y: 0.72, w: 12, h: 0.7,
    fontSize: 34, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });

  // Left column — Admin features
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.6, w: 5.8, h: 0.42,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });
  slide.addText("Fonctionnalités admin", {
    x: 0.5, y: 1.6, w: 5.8, h: 0.42,
    fontSize: 14, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "left", valign: "middle", margin: 8
  });

  const adminFeatures = [
    "Accès sécurisé par mot de passe",
    "Approuver ou refuser les demandes",
    "Modifier l'état de réparation du moteur",
    "Statistiques globales : total, approuvées, refusées, en cours",
  ];

  adminFeatures.forEach((f, i) => {
    const fy = 2.17 + i * 0.72;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: fy, w: 5.8, h: 0.62,
      fill: { color: i % 2 === 0 ? "f8fafc" : "FFFFFF" },
      line: { color: "e2e8f0", width: 0.5 }
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.65, y: fy + 0.22, w: 0.13, h: 0.13,
      fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
    });
    slide.addText(f, {
      x: 0.9, y: fy + 0.08, w: 5.2, h: 0.48,
      fontSize: 13, fontFace: "Calibri", color: "1e293b",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Right column — Workflow
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.0, y: 1.6, w: 5.8, h: 0.42,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });
  slide.addText("Workflow", {
    x: 7.0, y: 1.6, w: 5.8, h: 0.42,
    fontSize: 14, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "left", valign: "middle", margin: 8
  });

  const steps = [
    { step: "1", text: "Demandeur soumet la demande" },
    { step: "2", text: "Admin reçoit un email de notification" },
    { step: "3", text: "Admin approuve ou refuse via le panneau" },
    { step: "4", text: "Demandeur reçoit un email de mise à jour" },
  ];

  steps.forEach((s, i) => {
    const sy = 2.17 + i * 0.72;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 7.0, y: sy, w: 5.8, h: 0.62,
      fill: { color: i % 2 === 0 ? "fffbeb" : "FFFFFF" },
      line: { color: "fde68a", width: 0.5 }
    });
    // Step number circle
    slide.addShape(pres.shapes.OVAL, {
      x: 7.12, y: sy + 0.1, w: 0.4, h: 0.4,
      fill: { color: "f59e0b" }, line: { color: "f59e0b" }
    });
    slide.addText(s.step, {
      x: 7.12, y: sy + 0.1, w: 0.4, h: 0.4,
      fontSize: 12, fontFace: "Calibri", color: "FFFFFF",
      bold: true, align: "center", valign: "middle", margin: 0
    });
    slide.addText(s.text, {
      x: 7.65, y: sy + 0.1, w: 5.0, h: 0.42,
      fontSize: 13, fontFace: "Calibri", color: "1e293b",
      align: "left", valign: "middle", margin: 0
    });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("8 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 9 — Accès mobile
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.07,
    fill: { color: "1e3a8a" }, line: { color: "1e3a8a" }
  });

  slide.addText("Optimisé pour mobile", {
    x: 0.5, y: 0.28, w: 12, h: 0.7,
    fontSize: 36, fontFace: "Calibri", color: "1e3a8a",
    bold: true, align: "left", margin: 0
  });
  slide.addText("Accessible depuis n'importe quel smartphone", {
    x: 0.5, y: 0.95, w: 12, h: 0.4,
    fontSize: 16, fontFace: "Calibri", color: "64748b",
    align: "left", margin: 0
  });

  // 2x2 grid of highlight cards
  const highlights = [
    { icon: "📱", title: "Interface responsive", desc: "S'adapte à toutes les tailles d'écran", color: "1e3a8a", bg: "dbeafe" },
    { icon: "🃏", title: "Vue cartes", desc: "Tableau converti en cartes lisibles sur petit écran", color: "0369a1", bg: "e0f2fe" },
    { icon: "📲", title: "Installable", desc: "Ajout à l'écran d'accueil comme une app native", color: "166534", bg: "dcfce7" },
    { icon: "🔄", title: "Sync", desc: "Données synchronisées en temps réel avec Google Sheets", color: "7c2d12", bg: "ffedd5" },
  ];

  const cardW3 = 5.7;
  const cardH3 = 2.6;
  const marginX3 = 0.6;
  const marginY3 = 1.55;
  const gapX3 = 0.5;
  const gapY3 = 0.35;

  highlights.forEach((h, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = marginX3 + col * (cardW3 + gapX3);
    const cy = marginY3 + row * (cardH3 + gapY3);

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW3, h: cardH3,
      fill: { color: "FFFFFF" },
      line: { color: h.color, width: 1.5 },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.08 }
    });

    // Left color band
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 0.08, h: cardH3,
      fill: { color: h.color }, line: { color: h.color }
    });

    // Icon circle
    slide.addShape(pres.shapes.OVAL, {
      x: cx + 0.2, y: cy + cardH3 / 2 - 0.4, w: 0.8, h: 0.8,
      fill: { color: h.bg }, line: { color: h.bg }
    });
    slide.addText(h.icon, {
      x: cx + 0.2, y: cy + cardH3 / 2 - 0.4, w: 0.8, h: 0.8,
      fontSize: 22, align: "center", valign: "middle", margin: 0
    });

    // Title
    slide.addText(h.title, {
      x: cx + 1.2, y: cy + 0.4, w: cardW3 - 1.4, h: 0.5,
      fontSize: 16, fontFace: "Calibri", color: h.color,
      bold: true, align: "left", valign: "middle", margin: 0
    });

    // Description
    slide.addText(h.desc, {
      x: cx + 1.2, y: cy + 1.0, w: cardW3 - 1.4, h: 0.9,
      fontSize: 13, fontFace: "Calibri", color: "475569",
      align: "left", valign: "top", margin: 0
    });
  });

  // Footer
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: H - 0.35, w: W, h: 0.35,
    fill: { color: "f8fafc" }, line: { color: "e2e8f0", width: 0.5 }
  });
  slide.addText("Maintenance Analytics — OCP  |  Bureau de Méthode Daoui", {
    x: 0.5, y: H - 0.35, w: 12, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "left", valign: "middle", margin: 0
  });
  slide.addText("9 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.35,
    fontSize: 9, fontFace: "Calibri", color: "94a3b8",
    align: "right", valign: "middle", margin: 0
  });
}

// ─────────────────────────────────────────────────────────────
// SLIDE 10 — Conclusion (dark background)
// ─────────────────────────────────────────────────────────────
{
  const slide = pres.addSlide();
  slide.background = { color: "1e3a8a" };

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W, h: 0.08,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });

  // Left decorative bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.08, w: 0.5, h: H - 0.08,
    fill: { color: "172554" }, line: { color: "172554" }
  });

  // Background circle decoration
  slide.addShape(pres.shapes.OVAL, {
    x: 8.5, y: 1.0, w: 6, h: 6,
    fill: { color: "1d4ed8", transparency: 75 },
    line: { color: "1d4ed8", transparency: 75 }
  });

  // Title
  slide.addText("Prêt à utiliser", {
    x: 0.8, y: 0.8, w: 10, h: 0.9,
    fontSize: 52, fontFace: "Calibri", color: "FFFFFF",
    bold: true, align: "left", margin: 0
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 4, h: 0.05,
    fill: { color: "f59e0b" }, line: { color: "f59e0b" }
  });

  // 3 access steps
  const steps10 = [
    { num: "1", text: "Ouvrir l'URL de l'application" },
    { num: "2", text: "Choisir votre module (Dashboard / Plan de charge / Moteurs Électriques)" },
    { num: "3", text: "Les données se chargent automatiquement depuis SAP" },
  ];

  steps10.forEach((s, i) => {
    const sy = 2.1 + i * 1.3;

    // Number circle
    slide.addShape(pres.shapes.OVAL, {
      x: 0.8, y: sy, w: 0.7, h: 0.7,
      fill: { color: "f59e0b" }, line: { color: "f59e0b" }
    });
    slide.addText(s.num, {
      x: 0.8, y: sy, w: 0.7, h: 0.7,
      fontSize: 22, fontFace: "Calibri", color: "1e3a8a",
      bold: true, align: "center", valign: "middle", margin: 0
    });

    // Step text
    slide.addText(s.text, {
      x: 1.7, y: sy + 0.05, w: 10.5, h: 0.65,
      fontSize: 18, fontFace: "Calibri", color: "FFFFFF",
      align: "left", valign: "middle", margin: 0
    });

    // Connector line (except for last step)
    if (i < steps10.length - 1) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 1.12, y: sy + 0.7, w: 0.03, h: 0.6,
        fill: { color: "f59e0b", transparency: 50 },
        line: { color: "f59e0b", transparency: 50 }
      });
    }
  });

  // Footer with contact info
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: H - 1.1, w: W - 1.0, h: 0.7,
    fill: { color: "172554" }, line: { color: "172554" }
  });
  slide.addText("Questions ? Contactez le Bureau de Méthode Daoui", {
    x: 0.5, y: H - 1.1, w: W - 1.0, h: 0.7,
    fontSize: 16, fontFace: "Calibri", color: "bfdbfe",
    align: "center", valign: "middle", margin: 0
  });

  // Page number
  slide.addText("10 / 10", {
    x: 0, y: H - 0.35, w: W - 0.4, h: 0.3,
    fontSize: 9, fontFace: "Calibri", color: "93c5fd",
    align: "right", valign: "middle", margin: 0
  });
}

// Write the file
pres.writeFile({ fileName: "/Users/mounaim/Desktop/mainAna/Guide_Utilisateur_Maintenance_Analytics.pptx" })
  .then(() => {
    console.log("PPTX created successfully!");
  })
  .catch(err => {
    console.error("Error creating PPTX:", err);
    process.exit(1);
  });
