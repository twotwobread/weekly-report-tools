#!/usr/bin/env node
/**
 * Weekly Report slide generator using PptxGenJS.
 * Usage: node generate_slides.js --data slides.json --output report.pptx
 *
 * Slide structure (12 slides for 3 members):
 *   1. Cover
 *   2. Executive Summary
 *   3-11. Per member: Last Week | Output Summary+Issues | This Week
 *   12. Team Analysis
 */
"use strict";

const pptxgen = (() => {
  try { return require("pptxgenjs"); } catch {}
  try {
    const { execSync } = require("child_process");
    const npmRoot = execSync("npm root -g 2>/dev/null").toString().trim();
    if (npmRoot) return require(`${npmRoot}/pptxgenjs`);
  } catch {}
  throw new Error("pptxgenjs not found. Run: npm install -g pptxgenjs");
})();
const fs = require("fs");

// ── CLI args ──────────────────────────────────────────────────────────
const args = process.argv.slice(2);
function arg(name) {
  const i = args.indexOf(name);
  return i !== -1 ? args[i + 1] : null;
}
const dataPath = arg("--data");
const outputPath = arg("--output") || "weekly-report.pptx";
if (!dataPath) {
  console.error("Usage: node generate_slides.js --data slides.json --output report.pptx");
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(dataPath, "utf8"));

// ── Design tokens ────────────────────────────────────────────────────
const C = {
  navy:     "1E2761",
  blue:     "3B82F6",
  white:    "FFFFFF",
  ice:      "CADCFC",
  muted:    "64748B",
  dark:     "1E293B",
  body:     "475569",
  rowAlt:   "EEF2FF",
  progBg:   "E2E8F0",
  green:    "22C55E",
  amber:    "F59E0B",
  greenBg:  "DCFCE7",
  greenTxt: "166534",
  amberBg:  "FEF3C7",
  amberTxt: "92400E",
};
const FONT_H = "Georgia";
const FONT_B = "Calibri";

// ── PptxGenJS instance ───────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 × 5.625 inches
pres.title = `Weekly Report ${data.week}`;

// ── Helpers ──────────────────────────────────────────────────────────

/** Convert "YYYY-WXX" to { start: "YYYY/MM/DD", end: "YYYY/MM/DD" } */
function weekToDateRange(weekStr) {
  const [yearStr, wStr] = weekStr.split("-W");
  const y = parseInt(yearStr, 10);
  const w = parseInt(wStr, 10);
  // Jan 4th is always in ISO week 1
  const jan4 = new Date(y, 0, 4);
  const jan4Dow = jan4.getDay() || 7; // Monday=1 … Sunday=7
  const monday = new Date(jan4);
  monday.setDate(jan4.getDate() - jan4Dow + 1 + (w - 1) * 7);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  const fmt = (d) =>
    `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
  return { start: fmt(monday), end: fmt(sunday) };
}

const members = data.members || [];
const dateRange = weekToDateRange(data.week);

// ── Image dimension helpers ──────────────────────────────────────────
function getPNGDimensions(base64Data) {
  const buf = Buffer.from(base64Data, "base64");
  // PNG: 8-byte signature + IHDR chunk (4 len, 4 type, 4 width, 4 height)
  return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
}

function getJPEGDimensions(base64Data) {
  const buf = Buffer.from(base64Data, "base64");
  let i = 2; // skip SOI (FF D8)
  while (i + 3 < buf.length) {
    if (buf[i] !== 0xFF) break;
    const marker = buf[i + 1];
    const segLen = buf.readUInt16BE(i + 2);
    if (marker >= 0xC0 && marker <= 0xC3) {
      return { height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7) };
    }
    i += 2 + segLen;
  }
  return null;
}

function getImageDimensions(base64Data, mimeType) {
  try {
    if (mimeType === "image/png") return getPNGDimensions(base64Data);
    if (mimeType === "image/jpeg" || mimeType === "image/jpg") return getJPEGDimensions(base64Data);
  } catch (_) {}
  return null;
}

// Dynamically compute total slide count
function computeTotalSlides() {
  let total = 2; // Cover + Exec Summary
  for (const m of members) {
    total += 1; // Last Week
    total += 1; // Output (text-only) + Issues
    total += (m.output_summary || []).filter((l) => !!l.image_data).length;
    total += 1; // This Week
  }
  return total + 1; // Team Analysis
}
const TOTAL_SLIDES = computeTotalSlides();

function addFooter(slide, slideNum) {
  slide.addText(
    `NuFi Weekly Report  |  ${dateRange.start} ~ ${dateRange.end}  |  ${slideNum}/${TOTAL_SLIDES}`,
    {
      x: 0.5, y: 5.2, w: 9, h: 0.3,
      fontFace: FONT_B, fontSize: 9, color: C.muted,
      align: "center", valign: "middle",
    }
  );
}

/** Navy header bar (h=0.9). If name+role provided, renders rich text title. */
function addNavyHeader(slide, name, role) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.9,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  if (name) {
    slide.addText(
      [
        { text: name, options: { bold: true, fontSize: 22, color: C.white } },
        { text: "  " + (role || "Developer"), options: { bold: false, fontSize: 14, color: C.ice } },
      ],
      { x: 0.5, y: 0, w: 9, h: 0.9, fontFace: FONT_B, valign: "middle" }
    );
  }
}

/** Colored section badge below the header. */
function addSectionBadge(slide, label) {
  const longLabels = ["주요 산출물", "Output Summary"];
  const badgeW = longLabels.includes(label) ? 2.2 : 1.6;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.05, w: badgeW, h: 0.35,
    fill: { color: C.blue }, line: { color: C.blue },
  });
  slide.addText(label, {
    x: 0.5, y: 1.05, w: badgeW, h: 0.35,
    fontFace: FONT_B, fontSize: 12, color: C.white,
    bold: true, align: "center", valign: "middle",
  });
}

// ── Slide counter ────────────────────────────────────────────────────
let slideCount = 0;

// ── Slide 1: Cover ───────────────────────────────────────────────────
function addCoverSlide() {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.navy };

  // Title
  slide.addText("Weekly Report", {
    x: 0.5, y: 1.0, w: 9, h: 0.8,
    fontFace: FONT_H, fontSize: 44, color: C.white,
    bold: true, align: "left",
  });

  // Blue separator line
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 2.1, w: 10, h: 0.06,
    fill: { color: C.blue }, line: { color: C.blue },
  });

  // Date range
  slide.addText(`${dateRange.start} ~ ${dateRange.end}`, {
    x: 0.5, y: 2.3, w: 9, h: 0.5,
    fontFace: FONT_B, fontSize: 20, color: C.ice, align: "left",
  });

  // Team name
  slide.addText("NuFi Platform Team", {
    x: 0.5, y: 3.2, w: 9, h: 0.4,
    fontFace: FONT_B, fontSize: 16, color: C.muted, align: "left",
  });

  // Member chips
  const chipX = [1.9, 4.1, 6.3];
  members.forEach((m, i) => {
    const cx = chipX[i] !== undefined ? chipX[i] : 0.5 + i * 2.5;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 4.0, w: 1.8, h: 0.45,
      fill: { color: C.navy }, line: { color: C.blue, pt: 1 },
    });
    slide.addText(m.name, {
      x: cx, y: 4.0, w: 1.8, h: 0.45,
      fontFace: FONT_B, fontSize: 13, color: C.ice,
      bold: true, align: "center", valign: "middle",
    });
  });

  addFooter(slide, slideCount);
}

// ── Slide 2: Executive Summary ───────────────────────────────────────
function addExecutiveSummarySlide() {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, null, null);
  slide.addText("Executive Summary", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontFace: FONT_B, fontSize: 22, color: C.white,
    bold: true, valign: "middle",
  });

  const ta = data.team_analysis || {};
  const totalCompleted = ta.total_completed ?? 0;
  const totalPlanned = ta.total_planned ?? 0;
  let inProgress = 0;
  members.forEach((m) => {
    (m.last_week || []).forEach((t) => {
      if (t.status && t.status !== "Completed") inProgress++;
    });
  });

  // 4 stat cards
  const statCards = [
    { number: String(members.length),  label: "팀원",       sub: "Team Members"   },
    { number: String(totalCompleted),  label: "완료",       sub: "Completed Tasks" },
    { number: String(inProgress),      label: "진행중",     sub: "In Progress"    },
    { number: String(totalPlanned),    label: "이번주 계획", sub: "Planned Tasks"  },
  ];
  const cardX = [0.5, 2.85, 5.2, 7.55];
  statCards.forEach((s, i) => {
    const cx = cardX[i];
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: 1.15, w: 2.1, h: 1.3,
      fill: { color: C.white }, line: { color: C.rowAlt, pt: 1 },
    });
    slide.addText(s.number, {
      x: cx, y: 1.2, w: 2.1, h: 0.6,
      fontFace: FONT_H, fontSize: 36, color: C.blue,
      bold: true, align: "center", valign: "middle",
    });
    slide.addText(s.label, {
      x: cx, y: 1.8, w: 2.1, h: 0.3,
      fontFace: FONT_B, fontSize: 11, color: C.dark,
      bold: true, align: "center",
    });
    slide.addText(s.sub, {
      x: cx, y: 2.1, w: 2.1, h: 0.25,
      fontFace: FONT_B, fontSize: 9, color: C.muted,
      align: "center",
    });
  });

  // 주요 성과 section
  slide.addText("주요 성과", {
    x: 0.5, y: 2.7, w: 9, h: 0.35,
    fontFace: FONT_B, fontSize: 14, color: C.dark, bold: true,
  });

  const execSummary = data.executive_summary || {};
  let bullets = execSummary.bullets || [];

  // Auto-derive from completed tasks if not provided
  if (!bullets.length) {
    members.forEach((m) => {
      (m.last_week || [])
        .filter((t) => t.status === "Completed")
        .slice(0, 2)
        .forEach((t) => bullets.push(`${m.name}: ${t.task}`));
    });
  }

  if (bullets.length) {
    slide.addText(
      bullets.map((b, i) => ({
        text: b,
        options: {
          bullet: { type: "bullet" },
          fontFace: FONT_B, fontSize: 11, color: C.body,
          breakLine: i < bullets.length - 1,
          paraSpaceBefore: 4,
        },
      })),
      { x: 0.6, y: 3.05, w: 8.8, h: 1.9, valign: "top" }
    );
  }

  addFooter(slide, slideCount);
}

// ── Member slide: Last Week ──────────────────────────────────────────
function addLastWeekSlide(member) {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, member.name, member.role || "Developer");
  addSectionBadge(slide, "지난 주");

  const rows = member.last_week || [];
  if (!rows.length) {
    slide.addText("내용 없음", {
      x: 0.5, y: 2.5, w: 9, h: 0.5,
      fontFace: FONT_B, fontSize: 14, color: C.muted, align: "center",
    });
    addFooter(slide, slideCount);
    return;
  }

  // Columns: Project | Task | Status | Note
  const colW = [1.1, 4.4, 1.3, 2.2];

  const hCell = (text) => ({
    text,
    options: {
      fontFace: FONT_B, fontSize: 10, bold: true,
      color: C.white, fill: { color: C.navy },
      align: "center", valign: "middle",
      margin: [4, 6, 4, 6],
    },
  });

  const tableData = [
    [hCell("Project"), hCell("Task"), hCell("Status"), hCell("Note")],
  ];

  rows.forEach((row, i) => {
    const isCompleted = row.status === "Completed";
    const isInProgress = !isCompleted && !!row.status;
    const rowBg = i % 2 === 0 ? C.white : C.rowAlt;
    const statusBg  = isCompleted ? C.greenBg  : isInProgress ? C.amberBg  : rowBg;
    const statusTxt = isCompleted ? C.greenTxt : isInProgress ? C.amberTxt : C.body;

    const cell = (text, opts = {}) => ({
      text: String(text || ""),
      options: {
        fontFace: FONT_B, fontSize: 10,
        fill: { color: rowBg }, color: C.body,
        valign: "middle", margin: [3, 6, 3, 6],
        ...opts,
      },
    });

    tableData.push([
      cell(row.project),
      cell(row.task),
      {
        text: String(row.status || ""),
        options: {
          fontFace: FONT_B, fontSize: 10,
          bold: isCompleted || isInProgress,
          fill: { color: statusBg }, color: statusTxt,
          align: "center", valign: "middle",
          margin: [3, 6, 3, 6],
        },
      },
      cell(row.note || ""),
    ]);
  });

  const tableH = Math.min(3.8, 0.45 + rows.length * 0.42);
  slide.addTable(tableData, {
    x: 0.5, y: 1.55, w: 9, h: tableH,
    colW,
    border: { pt: 0.5, color: "D1D5DB" },
    autoPage: false,
  });

  addFooter(slide, slideCount);
}

// ── Member slide: Output Summary (text-only items) + Issues ─────────
function addOutputSummarySlide(member) {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, member.name, member.role || "Developer");
  addSectionBadge(slide, "주요 산출물");

  // Only text-only items (image items get their own slide)
  const links = (member.output_summary || []).filter((l) => !l.image_data);

  const LINK_CARD_H = 0.65;
  const CARD_GAP    = 0.2;

  let cardY = 1.55;
  links.forEach((link) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: cardY, w: 9, h: LINK_CARD_H,
      fill: { color: C.white }, line: { color: C.rowAlt, pt: 1 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: cardY, w: 0.06, h: LINK_CARD_H,
      fill: { color: C.blue }, line: { color: C.blue },
    });
    slide.addText(link.title || "", {
      x: 0.75, y: cardY, w: 3.5, h: LINK_CARD_H,
      fontFace: FONT_B, fontSize: 13, color: C.dark,
      bold: true, valign: "middle",
    });
    if (link.url) {
      slide.addText(link.url, {
        x: 4.3, y: cardY, w: 5.0, h: LINK_CARD_H,
        fontFace: FONT_B, fontSize: 9, color: C.blue,
        hyperlink: { url: link.url },
        valign: "middle",
      });
    }
    cardY += LINK_CARD_H + CARD_GAP;
  });

  // Issues section
  const issues = member.issues || [];
  const issueY = cardY + 0.1;

  slide.addText("Issues", {
    x: 0.5, y: issueY, w: 9, h: 0.35,
    fontFace: FONT_B, fontSize: 13, color: C.dark, bold: true,
  });

  if (!issues.length) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: issueY + 0.4, w: 9, h: 0.4,
      fill: { color: C.greenBg }, line: { color: C.greenBg },
    });
    slide.addText("이슈 없음", {
      x: 0.5, y: issueY + 0.4, w: 9, h: 0.4,
      fontFace: FONT_B, fontSize: 11, color: C.greenTxt,
      align: "center", valign: "middle",
    });
  } else {
    let issueBoxY = issueY + 0.4;
    issues.forEach((issue) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: issueBoxY, w: 9, h: 0.42,
        fill: { color: C.amberBg }, line: { color: C.amberBg },
      });
      slide.addText(issue, {
        x: 0.65, y: issueBoxY, w: 8.7, h: 0.42,
        fontFace: FONT_B, fontSize: 11, color: C.amberTxt,
        valign: "middle",
      });
      issueBoxY += 0.47;
    });
  }

  addFooter(slide, slideCount);
}

// ── Member slide: single image (one per image output item) ──────────
function addImageSlide(member, link) {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, member.name, member.role || "Developer");

  // Image title — below the header bar, clearly as a description label
  slide.addText(link.title || "", {
    x: 0.5, y: 1.0, w: 9, h: 0.38,
    fontFace: FONT_B, fontSize: 14, color: C.dark,
    bold: true, valign: "middle",
  });

  // Available area for the image (below title, above footer)
  const AREA_X = 0.5, AREA_Y = 1.45;
  const AREA_W = 9.0,  AREA_H = 3.6; // inches

  // Compute box size from actual image dimensions to preserve aspect ratio
  const dims = getImageDimensions(link.image_data, link.image_mime || "image/png");
  let boxW = AREA_W, boxH = AREA_H;
  if (dims && dims.width > 0 && dims.height > 0) {
    const ar = dims.width / dims.height;
    const areaAr = AREA_W / AREA_H;
    if (ar > areaAr) {
      boxW = AREA_W;
      boxH = AREA_W / ar;
    } else {
      boxH = AREA_H;
      boxW = AREA_H * ar;
    }
  }

  // Center the box within the available area
  const boxX = AREA_X + (AREA_W - boxW) / 2;
  const boxY = AREA_Y + (AREA_H - boxH) / 2;

  // Outer card frame
  slide.addShape(pres.shapes.RECTANGLE, {
    x: boxX, y: boxY, w: boxW, h: boxH,
    fill: { color: C.white }, line: { color: C.rowAlt, pt: 1 },
  });

  // Image — no sizing transform needed; box is already proportional
  slide.addImage({
    data: `data:${link.image_mime || "image/png"};base64,${link.image_data}`,
    x: boxX, y: boxY, w: boxW, h: boxH,
  });

  addFooter(slide, slideCount);
}

// ── Member slide: This Week ──────────────────────────────────────────
function addThisWeekSlide(member) {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, member.name, member.role || "Developer");
  addSectionBadge(slide, "이번 주");

  const rows = member.this_week || [];
  if (!rows.length) {
    slide.addText("내용 없음", {
      x: 0.5, y: 2.5, w: 9, h: 0.5,
      fontFace: FONT_B, fontSize: 14, color: C.muted, align: "center",
    });
    addFooter(slide, slideCount);
    return;
  }

  // Columns: Project | Task | Est. | Due Date
  const colW = [1.1, 4.9, 1.2, 1.8];

  const hCell = (text) => ({
    text,
    options: {
      fontFace: FONT_B, fontSize: 10, bold: true,
      color: C.white, fill: { color: C.navy },
      align: "center", valign: "middle",
      margin: [4, 6, 4, 6],
    },
  });

  const tableData = [
    [hCell("Project"), hCell("Task"), hCell("Est."), hCell("Due Date")],
  ];

  rows.forEach((row, i) => {
    const bg = i % 2 === 0 ? C.white : C.rowAlt;
    const cell = (text, opts = {}) => ({
      text: String(text || ""),
      options: {
        fontFace: FONT_B, fontSize: 10,
        fill: { color: bg }, color: C.body,
        valign: "middle", margin: [3, 6, 3, 6],
        ...opts,
      },
    });
    tableData.push([
      cell(row.project),
      cell(row.task),
      cell(row.est_days != null ? `${row.est_days}d` : "", { align: "center" }),
      cell(row.due_date || "", { align: "center" }),
    ]);
  });

  const tableH = Math.min(3.8, 0.45 + rows.length * 0.42);
  slide.addTable(tableData, {
    x: 0.5, y: 1.55, w: 9, h: tableH,
    colW,
    border: { pt: 0.5, color: "D1D5DB" },
    autoPage: false,
  });

  addFooter(slide, slideCount);
}

// ── Slide: Team Analysis ─────────────────────────────────────────────
function addTeamAnalysisSlide() {
  slideCount++;
  const slide = pres.addSlide();
  slide.background = { color: C.white };

  addNavyHeader(slide, null, null);
  slide.addText("Team Analysis", {
    x: 0.5, y: 0, w: 9, h: 0.9,
    fontFace: FONT_B, fontSize: 22, color: C.white,
    bold: true, valign: "middle",
  });

  const ta = data.team_analysis || {};

  // ── Left half: 완료 vs 계획 ────────────────────────────────────────
  slide.addText("완료 vs 계획", {
    x: 0.5, y: 1.1, w: 4.2, h: 0.35,
    fontFace: FONT_B, fontSize: 14, color: C.dark, bold: true,
  });

  let progressY = 1.55;
  members.forEach((m) => {
    const completed = (m.last_week || []).filter((t) => t.status === "Completed").length;
    const totalTasks = (m.last_week || []).length;
    const pct = totalTasks > 0 ? Math.min(1, completed / totalTasks) : 0;
    const barFillW = Math.max(0.05, 2.8 * pct);
    const barColor = pct >= 1 ? C.green : C.amber;

    slide.addText(m.name, {
      x: 0.5, y: progressY, w: 0.9, h: 0.35,
      fontFace: FONT_B, fontSize: 12, color: C.dark,
      bold: true, valign: "middle",
    });
    // Progress bar background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.5, y: progressY + 0.08, w: 2.8, h: 0.2,
      fill: { color: C.progBg }, line: { color: C.progBg },
    });
    // Progress bar fill
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.5, y: progressY + 0.08, w: barFillW, h: 0.2,
      fill: { color: barColor }, line: { color: barColor },
    });
    // Score
    slide.addText(`${completed}/${totalTasks}`, {
      x: 4.4, y: progressY, w: 0.5, h: 0.35,
      fontFace: FONT_B, fontSize: 11, color: C.body, valign: "middle",
    });
    progressY += 0.45;
  });

  // 업무량 분포 (다음 주)
  const workloadY = progressY + 0.2;
  slide.addText("업무량 분포 (다음 주)", {
    x: 0.5, y: workloadY, w: 4.2, h: 0.35,
    fontFace: FONT_B, fontSize: 14, color: C.dark, bold: true,
  });
  let wY = workloadY + 0.4;
  members.forEach((m) => {
    const tasks = (m.this_week || []).length;
    const days = (m.this_week || []).reduce((sum, t) => sum + (t.est_days || 0), 0);
    slide.addText(`${m.name}`, {
      x: 0.5, y: wY, w: 1.0, h: 0.3,
      fontFace: FONT_B, fontSize: 11, color: C.dark, bold: true, valign: "middle",
    });
    slide.addText(`${tasks} tasks / ${days} days`, {
      x: 1.6, y: wY, w: 3.1, h: 0.3,
      fontFace: FONT_B, fontSize: 11, color: C.body, valign: "middle",
    });
    wY += 0.35;
  });

  // ── Right half: 리스크 & 플래그 ───────────────────────────────────
  let rightY = 1.1;

  slide.addText("리스크 항목", {
    x: 5.3, y: rightY, w: 4.2, h: 0.35,
    fontFace: FONT_B, fontSize: 14, color: C.dark, bold: true,
  });
  rightY += 0.4;

  const risks = ta.risks || [];
  if (risks.length) {
    risks.forEach((r) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 5.3, y: rightY, w: 4.2, h: 0.45,
        fill: { color: C.amberBg }, line: { color: C.amberBg },
      });
      slide.addText(`${r.member}: ${r.description}`, {
        x: 5.4, y: rightY, w: 4.0, h: 0.45,
        fontFace: FONT_B, fontSize: 11, color: C.amberTxt, valign: "middle",
      });
      rightY += 0.5;
    });
  } else {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: rightY, w: 4.2, h: 0.4,
      fill: { color: C.greenBg }, line: { color: C.greenBg },
    });
    slide.addText("리스크 없음", {
      x: 5.3, y: rightY, w: 4.2, h: 0.4,
      fontFace: FONT_B, fontSize: 11, color: C.greenTxt,
      align: "center", valign: "middle",
    });
    rightY += 0.5;
  }

  rightY += 0.2;

  slide.addText("플래그", {
    x: 5.3, y: rightY, w: 4.2, h: 0.35,
    fontFace: FONT_B, fontSize: 14, color: C.dark, bold: true,
  });
  rightY += 0.4;

  const flags = ta.flags || [];
  if (flags.length) {
    flags.forEach((f) => {
      let text = f.member;
      if (f.task) text += ` / ${f.task}`;
      if (f.type === "long_task") text += ` — ${f.est_days}일`;
      if (f.note) text += ` (${f.note})`;
      const iconColor = f.type === "long_task" ? C.amber : C.blue;

      // Icon circle
      slide.addShape(pres.shapes.OVAL, {
        x: 5.3, y: rightY, w: 0.28, h: 0.28,
        fill: { color: iconColor }, line: { color: iconColor },
      });
      slide.addText(f.type === "long_task" ? "!" : "i", {
        x: 5.3, y: rightY, w: 0.28, h: 0.28,
        fontFace: FONT_B, fontSize: 10, color: C.white,
        bold: true, align: "center", valign: "middle",
      });
      // Flag text
      slide.addText(text, {
        x: 5.7, y: rightY, w: 3.8, h: 0.3,
        fontFace: FONT_B, fontSize: 11, color: C.body, valign: "middle",
      });
      rightY += 0.38;
    });
  } else {
    slide.addText("플래그 없음", {
      x: 5.3, y: rightY, w: 4.2, h: 0.35,
      fontFace: FONT_B, fontSize: 11, color: C.muted,
    });
  }

  addFooter(slide, slideCount);
}

// ── Build presentation ───────────────────────────────────────────────
addCoverSlide();
addExecutiveSummarySlide();
for (const member of members) {
  addLastWeekSlide(member);
  addOutputSummarySlide(member);
  for (const link of (member.output_summary || []).filter((l) => !!l.image_data)) {
    addImageSlide(member, link);
  }
  addThisWeekSlide(member);
}
addTeamAnalysisSlide();

// ── Write output ─────────────────────────────────────────────────────
pres.writeFile({ fileName: outputPath })
  .then(() => { console.log(outputPath); })
  .catch((err) => { console.error("Failed:", err.message); process.exit(1); });
