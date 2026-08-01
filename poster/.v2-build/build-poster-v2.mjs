import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "/Users/huyang/Workspace/bmvc-bdmedvis.github.io";
const OUT_DIR = path.join(ROOT, "poster");
const W = 793.7008;
const H = 1122.5197;

const C = {
  ink: "#102329",
  muted: "#607176",
  teal: "#005D66",
  teal2: "#087B83",
  teal3: "#3F9EA0",
  mint: "#E7F1EF",
  pale: "#F3F7F6",
  warm: "#FBF7F2",
  white: "#FFFFFF",
  line: "#CFDAD7",
  red: "#D9232E",
};

async function bytes(rel) {
  return new Uint8Array(await fs.readFile(path.join(ROOT, rel)));
}

function textBox(slide, name, text, x, y, w, h, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: style.typeface ?? "Aptos",
    fontSize: style.fontSize ?? 12,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
    verticalAlignment: style.verticalAlignment ?? "top",
    autoFit: style.autoFit ?? "shrinkText",
    lineSpacing: style.lineSpacing,
    insets: style.insets ?? { top: 0, right: 0, bottom: 0, left: 0 },
  };
  return shape;
}

function rect(slide, name, x, y, w, h, fill, line = C.line, radius = 0) {
  return slide.shapes.add({
    geometry: radius > 0 ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: line, width: line === "none" ? 0 : 1 },
    ...(radius > 0 ? { borderRadius: radius } : {}),
  });
}

function line(slide, name, x, y, w, h, color = C.line, width = 1) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

function eyebrow(slide, name, text, x, y, w, color = C.teal) {
  return textBox(slide, name, text, x, y, w, 18, {
    fontSize: 10.5,
    bold: true,
    color,
    lineSpacing: 0.95,
  });
}

function pill(slide, name, label, x, y, w) {
  rect(slide, `${name}-bg`, x, y, w, 22, C.mint, "none", 11);
  textBox(slide, name, label, x + 5, y + 4, w - 10, 14, {
    fontSize: 8.2,
    bold: true,
    color: C.teal,
    alignment: "center",
    verticalAlignment: "middle",
  });
}

async function addImage(slide, rel, name, x, y, w, h, fit = "contain", geometry = "rect", radius = 0, alt = "") {
  const ext = path.extname(rel).toLowerCase();
  const contentType = ext === ".jpg" || ext === ".jpeg" ? "image/jpeg" : "image/png";
  const image = slide.images.add({
    blob: await bytes(rel),
    contentType,
    alt,
    fit,
    position: { left: x, top: y, width: w, height: h },
    geometry,
    ...(radius > 0 ? { borderRadius: radius } : {}),
  });
  image.name = name;
  return image;
}

function topic(slide, name, label, x, y, w) {
  rect(slide, `${name}-dot`, x, y + 4, 5, 5, C.teal2, "none", 2.5);
  textBox(slide, name, label, x + 11, y, w - 11, 15, {
    fontSize: 8.7,
    bold: true,
    color: C.ink,
    verticalAlignment: "middle",
  });
}

function deadlineTile(slide, name, label, date, x, y, w, h) {
  rect(slide, `${name}-bg`, x, y, w, h, C.white, C.line, 7);
  textBox(slide, `${name}-label`, label.toUpperCase(), x + 10, y + 7, w - 20, 13, {
    fontSize: 7.2,
    bold: true,
    color: C.teal2,
    lineSpacing: 0.9,
  });
  textBox(slide, `${name}-date`, date, x + 10, y + 22, w - 20, h - 26, {
    fontSize: 10.8,
    bold: true,
    color: C.ink,
    verticalAlignment: "middle",
  });
}

function organizer(slide, idx, name, affiliation, x, y, w) {
  textBox(slide, `organizer-${idx}-name`, name, x, y, w, 14, {
    fontSize: 9.3,
    bold: true,
    color: C.ink,
  });
  textBox(slide, `organizer-${idx}-affiliation`, affiliation, x, y + 14, w, 17, {
    fontSize: 7.4,
    color: C.muted,
    lineSpacing: 0.92,
  });
}

async function main() {
  await fs.mkdir(OUT_DIR, { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  presentation.theme.colorScheme = {
    name: "BD-MedVis Academic Teal",
    themeColors: {
      bg1: C.white,
      bg2: C.pale,
      tx1: C.ink,
      tx2: C.muted,
      dk1: C.ink,
      lt1: C.white,
      dk2: C.teal,
      lt2: C.pale,
      accent1: C.teal,
      accent2: C.teal2,
      accent3: C.red,
      accent4: C.mint,
      accent5: C.warm,
      accent6: C.muted,
      hlink: C.teal2,
      folHlink: C.teal,
    },
  };

  const slide = presentation.slides.add();
  slide.background.fill = C.pale;

  // Subtle website-inspired grid.
  for (let x = 18; x < W; x += 36) line(slide, `grid-v-${x}`, x, 0, 0, H, "#E7ECEA", 0.45);
  for (let y = 18; y < H; y += 36) line(slide, `grid-h-${y}`, 0, y, W, 0, "#E7ECEA", 0.45);

  // Header: no circular BD device.
  rect(slide, "header-white", 0, 0, W, 84, C.white, "none");
  rect(slide, "header-accent", 0, 81, W, 3, C.teal, "none");
  rect(slide, "brand-accent", 38, 21, 4, 43, C.red, "none");
  textBox(slide, "brand-name", "BD-MedVis", 54, 15, 250, 32, {
    typeface: "Aptos Display",
    fontSize: 27,
    bold: true,
    color: C.ink,
    verticalAlignment: "middle",
  });
  textBox(slide, "brand-subtitle", "BMVC 2026 WORKSHOP", 55, 50, 220, 17, {
    fontSize: 11,
    bold: true,
    color: C.muted,
  });
  textBox(slide, "header-kicker", "37TH BRITISH MACHINE VISION CONFERENCE", 346, 31, 244, 20, {
    fontSize: 8.4,
    bold: true,
    color: C.teal,
    alignment: "right",
    verticalAlignment: "middle",
  });
  await addImage(slide, "assets/bmvc-2026-logo.png", "bmvc-logo", 614, 18, 141, 44, "contain", "rect", 0, "BMVC 2026 logo");

  // Hero.
  eyebrow(slide, "hero-kicker", "BIOMARKER-DRIVEN MEDICAL VISION INTELLIGENCE", 38, 106, 405);
  textBox(slide, "hero-title", "From pixels to\nclinically credible\nbiomarkers.", 38, 132, 390, 132, {
    typeface: "Aptos Display",
    fontSize: 37,
    bold: true,
    color: C.ink,
    lineSpacing: 0.88,
  });
  textBox(slide, "hero-copy", "Turning image-derived phenotypes into biologically grounded, interpretable evidence for precision medicine.", 40, 274, 380, 44, {
    fontSize: 13.4,
    color: C.muted,
    lineSpacing: 1.03,
  });
  // Use the original website artwork directly, without a decorative frame.
  await addImage(slide, "assets/hero-front-page.jpg", "hero-artwork-original", 462, 107, 294, 196, "contain", "rect", 0, "Medical vision workshop illustration");
  textBox(slide, "hero-triad", "DISCOVER  ·  VALIDATE  ·  TRANSLATE", 467, 312, 284, 16, {
    fontSize: 9.2,
    bold: true,
    color: C.teal,
    alignment: "center",
  });

  // Event strip.
  rect(slide, "event-strip", 0, 344, W, 69, C.teal, "none");
  rect(slide, "event-red", 0, 344, 7, 69, C.red, "none");
  const eventXs = [38, 287, 547];
  const eventLabels = ["DATE", "VENUE", "ONLINE"];
  const eventValues = ["26 November 2026", "Lancaster Town Hall · UK", "bmvc-bdmedvis.github.io"];
  for (let i = 0; i < 3; i++) {
    textBox(slide, `event-label-${i}`, eventLabels[i], eventXs[i], 358, i === 2 ? 205 : 220, 13, {
      fontSize: 7.6,
      bold: true,
      color: "#A8D0CD",
    });
    textBox(slide, `event-value-${i}`, eventValues[i], eventXs[i], 377, i === 2 ? 210 : 235, 19, {
      fontSize: i === 2 ? 11.5 : 14,
      bold: true,
      color: C.white,
      verticalAlignment: "middle",
    });
    if (i < 2) line(slide, `event-divider-${i}`, eventXs[i] + 220, 357, 0, 39, "#3B8389", 1);
  }

  // Why this workshop.
  rect(slide, "why-surface", 27, 429, 740, 105, C.white, C.line, 9);
  eyebrow(slide, "why-kicker", "WHY THIS WORKSHOP", 40, 443, 160);
  textBox(slide, "why-copy", "Medical vision is moving beyond prediction toward evidence linked to biology, clinical endpoints and patient outcomes. BD-MedVis brings together vision, pathology, radiology and multimodal AI to discover, validate and translate trustworthy imaging biomarkers.", 40, 464, 714, 39, {
    fontSize: 10.9,
    color: C.muted,
    lineSpacing: 1.02,
  });
  pill(slide, "pillar-biology", "BIOLOGICAL GROUNDING", 40, 505, 168);
  pill(slide, "pillar-clinical", "CLINICAL CREDIBILITY", 220, 505, 166);
  pill(slide, "pillar-reliable", "RELIABLE & INTERPRETABLE", 398, 505, 180);
  pill(slide, "pillar-translation", "RESPONSIBLE TRANSLATION", 590, 505, 163);

  // Main content columns.
  const leftX = 27;
  const leftW = 445;
  const rightX = 485;
  const rightW = 282;
  rect(slide, "cfp-surface", leftX, 549, leftW, 246, C.white, C.line, 9);
  rect(slide, "speakers-surface", rightX, 549, rightW, 246, C.warm, C.line, 9);

  eyebrow(slide, "cfp-kicker", "CALL FOR PAPERS", 40, 563, 180);
  textBox(slide, "cfp-heading", "Share work that makes medical vision matter.", 40, 583, 405, 27, {
    typeface: "Aptos Display",
    fontSize: 17.8,
    bold: true,
    color: C.ink,
  });
  textBox(slide, "cfp-copy", "We welcome original research, methods, benchmarks and clinical perspectives that connect medical images to robust biomarkers and meaningful evidence.", 40, 612, 406, 31, {
    fontSize: 9.4,
    color: C.muted,
    lineSpacing: 1.02,
  });

  topic(slide, "topic-1", "Imaging biomarkers", 40, 648, 127);
  topic(slide, "topic-2", "Computational pathology", 175, 648, 140);
  topic(slide, "topic-3", "Radiogenomics", 325, 648, 120);
  topic(slide, "topic-4", "Multimodal biomedical AI", 40, 667, 150);
  topic(slide, "topic-5", "Reliability & interpretability", 199, 667, 152);
  topic(slide, "topic-6", "Clinical translation", 358, 667, 96);

  deadlineTile(slide, "deadline-open", "Submissions open", "3 Aug 2026", 40, 695, 124, 43);
  deadlineTile(slide, "deadline-intent", "Intent to submit", "28 Aug 2026", 172, 695, 128, 43);
  deadlineTile(slide, "deadline-full", "Full paper", "4 Sep 2026", 308, 695, 137, 43);
  deadlineTile(slide, "deadline-notify", "Notification", "25 Sep 2026", 40, 746, 124, 38);
  deadlineTile(slide, "deadline-camera", "Camera-ready", "16 Oct 2026", 172, 746, 128, 38);
  deadlineTile(slide, "deadline-aoe", "Deadline time", "23:59 AoE", 308, 746, 137, 38);

  eyebrow(slide, "speakers-kicker", "INVITED SPEAKERS", 499, 563, 170);
  await addImage(slide, "assets/speakers/yalin-zheng.jpg", "speaker-yalin-photo", 501, 593, 64, 64, "cover", "ellipse", 0, "Professor Yalin Zheng");
  textBox(slide, "speaker-yalin-name", "Professor Yalin Zheng", 578, 594, 172, 20, {
    fontSize: 13.3,
    bold: true,
    color: C.ink,
  });
  textBox(slide, "speaker-yalin-role", "University of Liverpool\nAI in healthcare & ophthalmic imaging", 578, 618, 168, 34, {
    fontSize: 8.7,
    color: C.muted,
    lineSpacing: 0.95,
  });
  line(slide, "speaker-divider", 500, 672, 249, 0, C.line, 1);
  await addImage(slide, "assets/speakers/sharib-ali.jpg", "speaker-sharib-photo", 501, 688, 64, 64, "cover", "ellipse", 0, "Dr Sharib Ali");
  textBox(slide, "speaker-sharib-name", "Dr Sharib Ali", 578, 690, 172, 20, {
    fontSize: 13.3,
    bold: true,
    color: C.ink,
  });
  textBox(slide, "speaker-sharib-role", "University of Leeds\nMedical & surgical image analysis", 578, 714, 168, 34, {
    fontSize: 8.7,
    color: C.muted,
    lineSpacing: 0.95,
  });
  textBox(slide, "speaker-note", "Two translational perspectives. One shared question: what makes an image-derived biomarker clinically credible?", 501, 759, 248, 25, {
    fontSize: 7.9,
    bold: true,
    color: C.teal,
    alignment: "center",
    lineSpacing: 0.94,
  });

  // Organizers.
  rect(slide, "organizers-surface", 27, 811, 740, 126, C.white, C.line, 9);
  eyebrow(slide, "organizers-kicker", "ORGANIZING COMMITTEE", 40, 823, 200);
  textBox(slide, "organizers-tagline", "An interdisciplinary team spanning medical imaging, biomedical AI and clinical translation.", 224, 823, 529, 15, {
    fontSize: 8.2,
    color: C.muted,
    alignment: "right",
  });

  const organizers = [
    ["Dr Yang Hu", "University of Leicester"],
    ["Dr Tianyang Zhang", "Oxford / Birmingham"],
    ["Dr He Zhao", "University of Liverpool"],
    ["Dr Jian Chen", "University of Cambridge"],
    ["Dr Dan Dai", "Aston University"],
    ["Dr Yakun Ju", "University of Leicester"],
    ["Dr Le Zhang", "University of Birmingham"],
    ["Dr Shangqi Gao", "University of Cambridge"],
    ["Dr Zheheng Jiang", "University of Leicester"],
  ];
  const orgXs = [40, 278, 516];
  const orgYs = [849, 879, 909];
  let n = 0;
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      const [name, affiliation] = organizers[n];
      organizer(slide, n + 1, name, affiliation, orgXs[col], orgYs[row], 210);
      n += 1;
    }
  }

  // Institution and partner marks, adapted from the reference poster.
  rect(slide, "logo-band", 0, 950, W, 113, C.white, "none");
  line(slide, "logo-band-top", 0, 950, W, 0, C.line, 1);
  eyebrow(slide, "institutions-heading", "ORGANISING INSTITUTIONS", 38, 966, 140);
  const universityLogos = [
    ["assets/universities/leicester.jpg", "University of Leicester"],
    ["assets/universities/oxford.jpg", "University of Oxford"],
    ["assets/universities/liverpool.jpg", "University of Liverpool"],
    ["assets/universities/cambridge.jpg", "University of Cambridge"],
    ["assets/universities/aston.jpg", "Aston University"],
    ["assets/universities/birmingham.jpg", "University of Birmingham"],
  ];
  const logoXs = [183, 276, 369, 462, 555, 648];
  for (let i = 0; i < universityLogos.length; i++) {
    await addImage(slide, universityLogos[i][0], `institution-logo-${i + 1}`, logoXs[i], 958, 80, 34, "contain", "rect", 0, universityLogos[i][1]);
  }

  eyebrow(slide, "partners-heading", "PARTNERS & SUPPORTERS", 38, 1014, 160);
  await addImage(slide, "tmp/pdfs/bdmvi-reference/extracted/asset-001-009.png", "partner-welcome-lancaster", 203, 1002, 74, 48, "contain", "rect", 0, "Welcome to Lancaster graphic");
  await addImage(slide, "tmp/pdfs/bdmvi-reference/extracted/asset-001-047.png", "partner-lancaster-council", 308, 1006, 92, 44, "contain", "rect", 0, "Lancaster City Council logo");
  await addImage(slide, "tmp/pdfs/bdmvi-reference/extracted/asset-001-049.png", "partner-lancaster-crest", 443, 1002, 47, 48, "contain", "rect", 0, "Lancaster crest");
  await addImage(slide, "tmp/pdfs/bdmvi-reference/extracted/asset-001-051.png", "partner-icvl", 538, 1009, 104, 32, "contain", "rect", 0, "ICVL logo");
  textBox(slide, "partner-caption", "In association with BMVC 2026 and the Lancaster host community", 648, 1009, 108, 36, {
    fontSize: 7.2,
    bold: true,
    color: C.muted,
    alignment: "right",
    verticalAlignment: "middle",
    lineSpacing: 0.94,
  });

  // Footer.
  rect(slide, "footer", 0, 1063, W, H - 1063, C.teal, "none");
  textBox(slide, "footer-event", "26 NOVEMBER 2026  ·  LANCASTER TOWN HALL  ·  LANCASTER, UK", 38, 1078, 470, 17, {
    fontSize: 9.2,
    bold: true,
    color: C.white,
    verticalAlignment: "middle",
  });
  textBox(slide, "footer-contact", "bmvc-bdmedvis.github.io\nhy208@leicester.ac.uk", 523, 1074, 232, 30, {
    fontSize: 8.6,
    bold: true,
    color: "#C8E0DD",
    alignment: "right",
    lineSpacing: 0.98,
  });

  slide.speakerNotes.textFrame.setText(
    `[Sources]\n` +
    `- ${path.join(ROOT, "index.html")} (workshop title, speakers, organizers, venue, public copy; accessed 2026-08-01)\n` +
    `- ${path.join(ROOT, "WORKSHOP_SCHEDULE.md")} (confirmed public dates; accessed 2026-08-01)\n` +
    `- ${path.join(ROOT, "assets/hero-front-page.jpg")} (workshop illustration)\n` +
    `- ${path.join(ROOT, "assets/bmvc-2026-logo.png")} (BMVC 2026 logo)\n` +
    `- ${path.join(ROOT, "assets/speakers/yalin-zheng.jpg")} and sharib-ali.jpg (speaker photos)\n` +
    `- ${path.join(ROOT, "assets/universities")} (organising institution marks)\n` +
    `- /Users/huyang/Library/CloudStorage/OneDrive-UniversityofLeicester/2025-2026 Leicester-connecting/2026.5 BMVC workshop/BMVC-BDMVI 2026 2.pdf (visual reference and Lancaster/ICVL marks)\n` +
    `[/Sources]`
  );

  // 3.125 × the 96-dpi canvas produces an A4 image at approximately 300 dpi.
  const png = await presentation.export({ slide, format: "png", scale: 3.125 });
  await fs.writeFile(path.join(OUT_DIR, "BD-MedVis-2026-A4-Poster-v2.png"), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(OUT_DIR, "BD-MedVis-2026-A4-Poster-v2.layout.json"), await layout.text());
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(path.join(OUT_DIR, "BD-MedVis-2026-A4-Poster-Editable-v2.pptx"));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
