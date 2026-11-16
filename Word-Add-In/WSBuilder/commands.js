/* WS Builder - commands.js  (PROTOTYPE: MCQ Short Ans group + SQ Q1 group)
 *
 * Rules built in:
 *  - Page width is read from Word on EVERY step of EVERY click (nothing cached).
 *  - Arial 12 + spacing is applied to the selection on EVERY step.
 *  - No silent fallback width: if the width cannot be read, nothing is inserted
 *    and an error dialog is shown.
 */
"use strict";

// Set to false once testing is finished. While true, a dialog reports the
// width that was read after every successful click.
const WSB_DEBUG = false;

const CM_TWIPS = 566.929; // used by the MCQ files (twips per cm)
const CM_PT = 28.3465;    // used by the SQ files  (points per cm)

/* ------------------------------------------------------------------ */
/* 1. Page geometry - read fresh every time                            */
/* ------------------------------------------------------------------ */

const PS_FIELDS = ["pageSetup/pageWidth", "pageSetup/leftMargin", "pageSetup/rightMargin"];

function packGeometry(section, route) {
  const ps = section.pageSetup;
  const widthPt = ps.pageWidth - ps.leftMargin - ps.rightMargin;
  if (!isFinite(widthPt) || widthPt < 100) {
    throw new Error("Page width could not be read correctly (got " + widthPt + " pt).");
  }
  return { pageWidth: ps.pageWidth, left: ps.leftMargin, right: ps.rightMargin, widthPt, route };
}

/** Returns { widthPt, pageWidth, left, right, route } for the section at the cursor. */
async function readGeometry() {
  try {
    // Preferred: the section the cursor is in.
    return await Word.run(async (ctx) => {
      const section = ctx.document.getSelection().parentBody.parentSection;
      section.load(PS_FIELDS);
      await ctx.sync();
      return packGeometry(section, "section at cursor");
    });
  } catch (cursorErr) {
    // Fallback is allowed ONLY when the document has exactly one section.
    return await Word.run(async (ctx) => {
      const secs = ctx.document.sections;
      secs.load("items");
      await ctx.sync();
      if (secs.items.length !== 1) {
        throw new Error(
          "Could not find the section at the cursor, and the document has " +
          secs.items.length + " sections. (" + cursorErr.message + ")"
        );
      }
      const section = secs.items[0];
      section.load(PS_FIELDS);
      await ctx.sync();
      return packGeometry(section, "only section (cursor lookup failed: " + cursorErr.message + ")");
    });
  }
}

function describe(stepName, g) {
  return stepName + ": " + g.route + " | page " + g.pageWidth.toFixed(1) + " pt, margins " +
    g.left.toFixed(1) + " / " + g.right.toFixed(1) + " pt -> table " +
    g.widthPt.toFixed(1) + " pt (" + (g.widthPt / CM_PT).toFixed(2) + " cm)";
}

/* ------------------------------------------------------------------ */
/* 2. Default format - applied to the selection on every step          */
/* ------------------------------------------------------------------ */

async function formatSelection(ctx, kind, widthPt) {
  const sel = ctx.document.getSelection();
  sel.font.name = "Arial"; // covers a collapsed cursor
  sel.font.size = 12;

  const paras = sel.paragraphs;
  paras.load("items");
  await ctx.sync();

  paras.items.forEach((p) => {
    p.font.name = "Arial";
    p.font.size = 12;
    p.alignment = Word.Alignment.justified;
    p.spaceBefore = 6;
    if (kind === "mcq") {
      p.spaceAfter = 6;
      p.lineSpacing = 12;
    }
  });

  if (kind === "sq") {
    // Same as SQ_Default_Format.js: tables in the selection -> borderless, full width.
    const tables = sel.tables;
    tables.load("items");
    await ctx.sync();
    tables.items.forEach((t) => {
      t.preferredWidth = widthPt;
      [
        Word.BorderType.top, Word.BorderType.bottom, Word.BorderType.left,
        Word.BorderType.right, Word.BorderType.insideHorizontal, Word.BorderType.insideVertical,
      ].forEach((b) => { t.getBorder(b).type = Word.BorderType.none; });
    });
  }
  await ctx.sync();
}

/* ------------------------------------------------------------------ */
/* 3. MCQ OOXML builders (same markup as the original MCQ_*.js files)  */
/* ------------------------------------------------------------------ */

const MCQ_PPR = `
      <w:pPr>
        <w:jc w:val="both"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="exact"/>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="24"/>
          <w:szCs w:val="24"/>
        </w:rPr>
      </w:pPr>
    `;
const MCQ_RPR_BOLD = `<w:rPr><w:b/><w:bCs/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>`;

const mcqRun = (inner) => `<w:r>${MCQ_RPR_BOLD}${inner}</w:r>`;
const mcqBoldText = (t) => mcqRun(`<w:t>${t}</w:t>`);
const mcqAutonum = () =>
  mcqRun(`<w:fldChar w:fldCharType="begin"/>`) +
  mcqRun(`<w:instrText xml:space="preserve"> AUTONUMLGL \\e </w:instrText>`) +
  mcqRun(`<w:fldChar w:fldCharType="separate"/>`) +
  mcqRun(`<w:fldChar w:fldCharType="end"/>`);

const mcqCell = (w, inner, center) =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p>${center ? MCQ_PPR_CENTER : MCQ_PPR}${inner || ""}</w:p></w:tc>`;

/** A bold option letter / numeral cell: the text, centred horizontally in its cell. */
const mcqLetter = (t) => ({ xml: mcqBoldText(t), center: true });

/** rowsXml: array of rows, each an array of finished <w:tc> strings. trailingPara: blank paragraph after the table. */
function mcqTableRaw(totalTwips, widths, rowsXml, trailingPara) {
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const trs = rowsXml.map((cells) => `<w:tr>${cells.join("")}</w:tr>`).join("");
  return `
      <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblW w:w="${totalTwips}" w:type="dxa"/>
              <w:tblBorders>
                <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>
                <w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
              </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>${grid}</w:tblGrid>
            ${trs}
          </w:tbl>
          ${trailingPara ? `<w:p>${MCQ_PPR}</w:p>` : ""}
        </w:body>
      </w:wordDocument>
    `;
}

/** rows: array of rows, each an array of per-cell inner XML (plain, borderless cells). */
function mcqTable(totalTwips, widths, rows, trailingPara) {
  return mcqTableRaw(
    totalTwips, widths,
    rows.map((inner) =>
      widths.map((w, i) => {
        const c = inner[i];
        return c && typeof c === "object" ? mcqCell(w, c.xml, c.center) : mcqCell(w, c);
      })
    ),
    trailingPara
  );
}

function mcqQnNoOoxml(total) {
  const c1 = Math.round(0.95 * CM_TWIPS);
  return mcqTable(total, [c1, total - c1], [[mcqAutonum(), ""]], true);
}

function mcqQnPicOoxml(total) {
  const c1 = Math.round(0.95 * CM_TWIPS);
  return mcqTable(total, [c1, total - c1], [["", ""]], true);
}

function mcqShortOptionOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const lab = Math.round(0.7 * CM_TWIPS);
  const con = Math.round((total - c0 - 4 * lab) / 4);
  return mcqTable(
    total,
    [c0, lab, con, lab, con, lab, con, lab, con],
    [["", mcqLetter("A"), "", mcqLetter("B"), "", mcqLetter("C"), "", mcqLetter("D"), ""]],
    true
  );
}

/** Medium: 2 rows x 5 columns (A B / C D). No blank paragraph after the table. */
function mcqMediumOptionOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const lab = Math.round(0.7 * CM_TWIPS);
  const con = Math.round((total - c0 - 2 * lab) / 2);
  return mcqTable(
    total,
    [c0, lab, con, lab, con],
    [
      ["", mcqLetter("A"), "", mcqLetter("B"), ""],
      ["", mcqLetter("C"), "", mcqLetter("D"), ""],
    ],
    false
  );
}

/** Long: 4 rows x 3 columns (A / B / C / D). No blank paragraph after the table. */
function mcqLongOptionOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const lab = Math.round(0.7 * CM_TWIPS);
  const con = Math.round(total - c0 - lab);
  return mcqTable(
    total,
    [c0, lab, con],
    [
      ["", mcqLetter("A"), ""],
      ["", mcqLetter("B"), ""],
      ["", mcqLetter("C"), ""],
      ["", mcqLetter("D"), ""],
    ],
    false
  );
}

/* ---- Table Two / Table Three (bordered grids, copied from MCQ_Table_Two/Three_Option.js) ---- */

const MCQ_PPR_CENTER = MCQ_PPR.replace('<w:jc w:val="both"/>', '<w:jc w:val="center"/>');

const mcqBdr = (top, left, bottom, right) => {
  const e = (side, on) =>
    on ? `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>` : `<w:${side} w:val="none"/>`;
  return e("top", top) + e("left", left) + e("bottom", bottom) + e("right", right);
};
const BDR_ALL = mcqBdr(1, 1, 1, 1);
const BDR_BOTTOM = mcqBdr(0, 0, 1, 0);
const BDR_VERT = mcqBdr(0, 1, 0, 1);
const BDR_BOTTOM_VERT = mcqBdr(0, 1, 1, 1);
const BDR_LEFT = mcqBdr(0, 1, 0, 0);

function mcqBorderedCell(width, center, borders, text) {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${
    borders ? `<w:tcBorders>${borders}</w:tcBorders>` : ""
  }</w:tcPr><w:p>${center || text ? MCQ_PPR_CENTER : MCQ_PPR}${text ? mcqBoldText(text) : ""}</w:p></w:tc>`;
}

/** Table Two: 5 rows x 5 columns (0.95 cm, 0.7 cm, 3 equal columns). A-D in column 2, rows 2-5. */
function mcqTableTwoOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const c1 = Math.round(0.7 * CM_TWIPS);
  const d = Math.round((total - c0 - c1) / 3);
  const widths = [c0, c1, d, d, d];
  const border = (r, c) => {
    if (r === 0 && c >= 1 && c <= 3) return BDR_ALL;
    if (r === 4 && c >= 1 && c <= 3) return c === 1 || c === 3 ? BDR_BOTTOM_VERT : BDR_BOTTOM;
    if (c === 1 || c === 3) return BDR_VERT;
    return "";
  };
  const rows = [0, 1, 2, 3, 4].map((r) =>
    widths.map((w, c) => mcqBorderedCell(w, c === 2 || c === 3, border(r, c), c === 1 && r >= 1 ? "ABCD"[r - 1] : ""))
  );
  return mcqTableRaw(total, widths, rows, false);
}

/** Table Three: 5 rows x 6 columns (0.95 cm, 0.7 cm, 4 equal columns). A-D in column 2, rows 2-5. */
function mcqTableThreeOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const c1 = Math.round(0.7 * CM_TWIPS);
  const d = Math.round((total - c0 - c1) / 4);
  const widths = [c0, c1, d, d, d, d];
  const border = (r, c) => {
    if (r === 0 && c >= 1 && c <= 4) return BDR_ALL;
    if (r === 4 && c >= 1 && c <= 4) return BDR_BOTTOM_VERT;
    if (c >= 1 && c <= 5) return BDR_LEFT;
    return "";
  };
  const rows = [0, 1, 2, 3, 4].map((r) =>
    widths.map((w, c) => mcqBorderedCell(w, c >= 2 && c <= 4, border(r, c), c === 1 && r >= 1 ? "ABCD"[r - 1] : ""))
  );
  return mcqTableRaw(total, widths, rows, false);
}

const BDR_NONE = mcqBdr(0, 0, 0, 0);

/** Multi Option: 4 rows x 4 columns (0.95, 0.7, 0.8 cm, rest). I-IV in column 3. (MCQ_Multi_Option.js) */
function mcqMultiOptionOoxml(total) {
  const c0 = Math.round(0.95 * CM_TWIPS);
  const c1 = Math.round(0.7 * CM_TWIPS);
  const c2 = Math.round(0.8 * CM_TWIPS);
  const c3 = total - c0 - c1 - c2;
  const widths = [c0, c1, c2, c3];
  const rows = ["I", "II", "III", "IV"].map((roman) =>
    widths.map((w, c) => mcqBorderedCell(w, false, BDR_NONE, c === 2 ? roman : ""))
  );
  return mcqTableRaw(total, widths, rows, false);
}

/** MCQ insert: replace selection, reuse an existing blank paragraph after the table. */
async function insertMcq(ctx, ooxml) {
  const insertedRange = ctx.document.getSelection().insertOoxml(ooxml, Word.InsertLocation.replace);
  await ctx.sync();

  const afterParas = insertedRange.getRange(Word.RangeLocation.after).paragraphs;
  afterParas.load("items/text");
  await ctx.sync();

  let target;
  if (afterParas.items.length > 0 && afterParas.items[0].text.trim() === "") {
    target = afterParas.items[0];
  } else {
    target = insertedRange.insertParagraph("", Word.InsertLocation.after);
    await ctx.sync();
  }
  target.select(Word.SelectionMode.start);
  await ctx.sync();
}

async function mcqStep(name, builder, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));
  const total = Math.round(g.widthPt * 20);             // points -> twips
  await Word.run(async (ctx) => {
    await formatSelection(ctx, "mcq");                  // Arial 12 + spacing, every step
    await insertMcq(ctx, builder(total));
  });
}

/* ------------------------------------------------------------------ */
/* 4. SQ OOXML builders (same markup as the original SQ_Qn.js)         */
/* ------------------------------------------------------------------ */

function sqCell(widthDxa, opts) {
  const bold = !!(opts && opts.bold);
  const justify = (opts && opts.justify) || "left";
  return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                    <w:jc w:val="${justify}"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${bold ? "<w:b/>" : ""}
                        <w:sz w:val="24"/>
                    </w:rPr>
                </w:pPr>
            </w:p>
        </w:tc>`;
}

/** Wraps a finished table (grid + cells) in the OOXML package that Word expects. */
const SQ_BORDERS_NONE = `
                                <w:top w:val="none"/>
                                <w:left w:val="none"/>
                                <w:bottom w:val="none"/>
                                <w:right w:val="none"/>
                                <w:insideH w:val="none"/>
                                <w:insideV w:val="none"/>
                            `;

function sqPackage(totalDxa, gridXml, cellsXml, extraTblPr, bordersXml) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
        </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                    <w:tbl>
                        <w:tblPr>
                            <w:tblW w:w="${totalDxa}" w:type="dxa"/>
                            <w:tblBorders>${bordersXml || SQ_BORDERS_NONE}</w:tblBorders>${extraTblPr || ""}
                        </w:tblPr>
                        <w:tblGrid>${gridXml}</w:tblGrid>
                        <w:tr>${cellsXml}</w:tr>
                    </w:tbl>
                </w:body>
            </w:document>
        </pkg:xmlData>
    </pkg:part>
</pkg:package>`;
}

function sqTableOoxml(colDxa, cellOpts) {
  const totalDxa = colDxa.reduce((s, w) => s + w, 0);
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const cells = colDxa.map((w, i) => sqCell(w, cellOpts[i])).join("");
  return sqPackage(totalDxa, grid, cells);
}

/** Lined-answer cell (SQ_Lined_Ans.js): the last cell is the answer cell - 16 pt, not bold, dotted bottom border. */
function sqLinedCell(widthDxa, isAnswerCell, jc) {
  const sz = isAnswerCell ? 32 : 24;
  const boldTag = isAnswerCell ? "" : "<w:b/>";
  const borders = isAnswerCell
    ? `<w:tcBorders><w:bottom w:val="dotted" w:sz="18" w:space="0" w:color="auto"/></w:tcBorders>`
    : "";
  return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
                ${borders}
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                    <w:jc w:val="${jc}"/>
                    <w:snapToGrid w:val="0"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${boldTag}
                        <w:sz w:val="${sz}"/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${boldTag}
                        <w:sz w:val="${sz}"/>
                    </w:rPr>
                    <w:t></w:t>
                </w:r>
            </w:p>
        </w:tc>`;
}

function sqLinedTableOoxml(colDxa, labelJc, answerJc) {
  const totalDxa = colDxa.reduce((s, w) => s + w, 0);
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const last = colDxa.length - 1;
  const cells = colDxa.map((w, i) => sqLinedCell(w, i === last, i === last ? answerJc : labelJc)).join("");
  return sqPackage(totalDxa, grid, cells);
}

/** SQ insert: insert after selection, make sure two blank lines follow, cursor on the last. */
async function insertSq(ctx, ooxml) {
  const insertedRange = ctx.document.getSelection().insertOoxml(ooxml, Word.InsertLocation.after);
  await ctx.sync();

  const paras = insertedRange.getRange(Word.RangeLocation.after).paragraphs;
  paras.load("items/text");
  await ctx.sync();

  let emptyCount = 0;
  for (const p of paras.items) {
    if (p.text.trim() === "") {
      emptyCount++;
      if (emptyCount >= 2) break;
    } else {
      break;
    }
  }

  let target;
  if (emptyCount >= 2) {
    target = paras.items[1];
  } else if (emptyCount === 1) {
    target = paras.items[0].insertParagraph("", Word.InsertLocation.after);
  } else {
    const first = insertedRange.insertParagraph("", Word.InsertLocation.after);
    target = first.insertParagraph("", Word.InsertLocation.after);
  }
  await ctx.sync();
  target.select(Word.SelectionMode.start);
  await ctx.sync();
}

/** Column widths in twips: the fixed columns (cm) plus a last column that takes the remaining width. */
function sqColumnWidths(fixedCm, widthPt) {
  const dxa = fixedCm.map((cm) => Math.round(cm * CM_PT * 20));
  const usedPt = fixedCm.reduce((s, cm) => s + cm * CM_PT, 0);
  const lastDxa = Math.round((widthPt - usedPt) * 20);
  if (lastDxa < 200) {
    throw new Error("Text area is too narrow for this table (last column would be " + lastDxa + " twips).");
  }
  dxa.push(lastDxa);
  return dxa;
}

/** fixedCm: widths (cm) of the fixed columns. The last column takes the remaining width. */
async function sqStep(name, fixedCm, cellOpts, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));

  await Word.run((ctx) => formatSelection(ctx, "sq", g.widthPt)); // Arial 12 + spacing, every step

  const dxa = sqColumnWidths(fixedCm, g.widthPt);
  await Word.run((ctx) => insertSq(ctx, sqTableOoxml(dxa, cellOpts)));
}

/** Lined-answer table: same flow as sqStep, different cell markup (SQ_Lined_Ans.js). */
async function sqLinedStep(name, fixedCm, labelJc, answerJc, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));

  await Word.run((ctx) => formatSelection(ctx, "sq", g.widthPt)); // Arial 12 + spacing, every step

  const dxa = sqColumnWidths(fixedCm, g.widthPt);
  await Word.run((ctx) => insertSq(ctx, sqLinedTableOoxml(dxa, labelJc, answerJc)));
}

/* ---- Marks tables (SQ_Lined_Ans_Marks.js and SQ_Qn_Marks.js) ----
 * Layout: N label columns | one flexible column (takes the remaining width) | one fixed marks column at the right. */

const SQ_SP = `<w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>`;
const SQ_FONT = `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`;
const sqTc = (w, borders, inner) =>
  `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${w}" w:type="dxa"/>${borders || ""}
            </w:tcPr>
            <w:p>${inner}</w:p>
        </w:tc>`;
const SQ_DOTTED = `<w:tcBorders><w:bottom w:val="dotted" w:sz="18" w:space="0" w:color="auto"/></w:tcBorders>`;

// Lined answer + marks: no runs inside the cells
const sqLamLabelCell = (w) =>
  sqTc(w, "", `<w:pPr>${SQ_SP}<w:jc w:val="left"/><w:rPr>${SQ_FONT}<w:b/><w:sz w:val="24"/></w:rPr></w:pPr>`);
const sqLamLineCell = (w) =>
  sqTc(w, SQ_DOTTED, `<w:pPr>${SQ_SP}<w:jc w:val="both"/><w:rPr>${SQ_FONT}<w:sz w:val="32"/></w:rPr></w:pPr>`);
const sqLamFinalCell = (w) =>
  sqTc(w, "", `<w:pPr>${SQ_SP}<w:jc w:val="right"/><w:rPr>${SQ_FONT}<w:sz w:val="24"/></w:rPr></w:pPr>`);

// Qn + marks: jc comes first, snapToGrid, and an empty run in every cell
const sqQmRun = (rpr) => `<w:r><w:rPr>${rpr}</w:rPr><w:t></w:t></w:r>`;
const sqQmCell = (w, jc, rpr) =>
  sqTc(w, "", `<w:pPr><w:jc w:val="${jc}"/>${SQ_SP}<w:snapToGrid w:val="0"/><w:rPr>${rpr}</w:rPr></w:pPr>${sqQmRun(rpr)}`);
const sqQmLabelCell = (w) => sqQmCell(w, "left", `${SQ_FONT}<w:sz w:val="24"/><w:b/>`);
const sqQmLineCell = (w) => sqQmCell(w, "both", `${SQ_FONT}<w:sz w:val="24"/>`);
const sqQmFinalCell = (w) => sqQmCell(w, "right", `${SQ_FONT}<w:sz w:val="24"/>`);

const SQ_ZERO_MARGINS = `
                            <w:tblCellMar>
                                <w:top w:w="0" w:type="dxa"/>
                                <w:left w:w="0" w:type="dxa"/>
                                <w:bottom w:w="0" w:type="dxa"/>
                                <w:right w:w="0" w:type="dxa"/>
                            </w:tblCellMar>`;

const SQ_MARKS_KINDS = {
  linedMarks: { label: sqLamLabelCell, line: sqLamLineCell, final: sqLamFinalCell, extraTblPr: "" },
  qnMarks: { label: sqQmLabelCell, line: sqQmLineCell, final: sqQmFinalCell, extraTblPr: SQ_ZERO_MARGINS },
};

function sqMarksTableOoxml(colDxa, kind) {
  const k = SQ_MARKS_KINDS[kind];
  const n = colDxa.length;
  const totalDxa = colDxa.reduce((s, w) => s + w, 0);
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const cells = colDxa
    .map((w, i) => (i === n - 1 ? k.final(w) : i === n - 2 ? k.line(w) : k.label(w)))
    .join("");
  return sqPackage(totalDxa, grid, cells, k.extraTblPr);
}

/** Columns in twips: label columns (cm), a flexible column (remaining width), a fixed marks column (cm) at the right. */
function sqMarksColumnWidths(labelCm, finalCm, widthPt) {
  const labels = labelCm.map((cm) => Math.round(cm * CM_PT * 20));
  const finalDxa = Math.round(finalCm * CM_PT * 20);
  const usedPt = labelCm.reduce((s, cm) => s + cm * CM_PT, 0) + finalCm * CM_PT;
  const flexDxa = Math.round((widthPt - usedPt) * 20);
  if (flexDxa < 200) {
    throw new Error("Text area is too narrow for this table (flexible column would be " + flexDxa + " twips).");
  }
  return [...labels, flexDxa, finalDxa];
}

async function sqMarksStep(name, labelCm, finalCm, kind, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));

  await Word.run((ctx) => formatSelection(ctx, "sq", g.widthPt)); // Arial 12 + spacing, every step

  const dxa = sqMarksColumnWidths(labelCm, finalCm, g.widthPt);
  await Word.run((ctx) => insertSq(ctx, sqMarksTableOoxml(dxa, kind)));
}

/* ---- Cal Ans / Cal Ans Unit (SQ_Cal_Ans.js) ----
 * Layout: one flexible column on the left (takes the remaining width), then fixed columns (cm). */

// Cal Ans: jc, spacing, snapToGrid and an empty run in every cell; zero cell margins
const sqCalCell = (w, jc, sz, borders) => {
  const rpr = `${SQ_FONT}<w:sz w:val="${sz}"/>`;
  return sqTc(w, borders, `<w:pPr><w:jc w:val="${jc}"/>${SQ_SP}<w:snapToGrid w:val="0"/><w:rPr>${rpr}</w:rPr></w:pPr>${sqQmRun(rpr)}`);
};
function sqCalAnsOoxml(colDxa) {
  const [c1, c2, c3] = colDxa;
  const cells = sqCalCell(c1, "right", 24, "") + sqCalCell(c2, "both", 32, SQ_DOTTED) + sqCalCell(c3, "right", 24, "");
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  return sqPackage(c1 + c2 + c3, grid, cells, SQ_ZERO_MARGINS);
}

// Cal Ans Unit: spacing then jc, no runs, no cell margins, only four table borders
const sqCalUnitCell = (w, jc, sz, borders) =>
  sqTc(w, borders, `<w:pPr>${SQ_SP}<w:jc w:val="${jc}"/><w:rPr>${SQ_FONT}<w:sz w:val="${sz}"/></w:rPr></w:pPr>`);
const SQ_BORDERS_FOUR = `
                                            <w:top w:val="none"/>
                                            <w:bottom w:val="none"/>
                                            <w:right w:val="none"/>
                                            <w:left w:val="none"/>
                                        `;
function sqCalAnsUnitOoxml(colDxa) {
  const [c1, c2, c3, c4] = colDxa;
  const cells =
    sqCalUnitCell(c1, "right", 24, "") + sqCalUnitCell(c2, "both", 32, SQ_DOTTED) +
    sqCalUnitCell(c3, "left", 24, "") + sqCalUnitCell(c4, "right", 24, "");
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  return sqPackage(c1 + c2 + c3 + c4, grid, cells, "", SQ_BORDERS_FOUR);
}

/** Columns in twips: a flexible first column (remaining width, subtracted one by one as in the original), then fixed columns (cm). */
function sqCalColumnWidths(fixedCm, widthPt) {
  let rest = widthPt;
  fixedCm.forEach((cm) => { rest -= cm * CM_PT; });
  const first = Math.round(rest * 20);
  if (first < 200) {
    throw new Error("Text area is too narrow for this table (first column would be " + first + " twips).");
  }
  return [first, ...fixedCm.map((cm) => Math.round(cm * CM_PT * 20))];
}

async function sqCalStep(name, fixedCm, builder, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));

  await Word.run((ctx) => formatSelection(ctx, "sq", g.widthPt)); // Arial 12 + spacing, every step

  const dxa = sqCalColumnWidths(fixedCm, g.widthPt);
  await Word.run((ctx) => insertSq(ctx, builder(dxa)));
}

/* ---- Join Tables (MergeAllTable.js) ---- */

/** Merges adjacent tables that are separated only by an empty paragraph. Works on the whole document. */
async function joinTables(log) {
  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    if (tables.items.length < 2) {
      log.push("Fewer than 2 tables found in document.");
      return;
    }

    let mergedCount = 0;
    // Loop backwards so indices stay stable as paragraphs are deleted
    for (let i = tables.items.length - 1; i > 0; i--) {
      const prevTable = tables.items[i - 1];
      const gapParagraph = prevTable.getRange(Word.RangeLocation.after).paragraphs.getFirst();
      gapParagraph.load("text");
      await context.sync();

      // Only merge when the paragraph between the tables is genuinely empty
      if (gapParagraph.text.trim() === "") {
        gapParagraph.delete();
        await context.sync();
        mergedCount++;
      }
    }

    log.push(
      mergedCount > 0
        ? "Successfully merged " + mergedCount + " table(s)!"
        : "No adjacent tables (separated only by an empty paragraph) found to merge."
    );
  });
}

/* ---- Blanks (SQ_Fill_In_The_Blanks.js) ---- */

/** Inserts a run of dots lowered by 3 pt, then puts the cursor after a space that is back at the ORIGINAL font position,
 *  so whatever is typed next sits on the normal baseline again. */
async function insertDots(dotCount, log) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();

    // Remember the font position the text had before the blanks (0 = normal baseline).
    // If the selection mixes several positions Word reports no single number, so fall back to 0.
    selection.font.load("position");
    await context.sync();
    const rawPos = selection.font.position;
    const originalPos = typeof rawPos === "number" && isFinite(rawPos) ? rawPos : 0;
    if (log) log.push("Blanks: font position before = " + originalPos + " pt");

    const paragraph = selection.paragraphs.getFirst();
    paragraph.spaceBefore = 6;
    paragraph.spaceAfter = 0;
    paragraph.lineSpacing = 12;
    paragraph.alignment = Word.Alignment.justified;

    const dotsRange = selection.insertText(".".repeat(dotCount), Word.InsertLocation.replace);
    dotsRange.font.name = "Arial";
    dotsRange.font.size = 12;
    dotsRange.font.position = -3; // dots lowered by 3 pt

    // A space right after the dots, set back to the original position (raised by 3 pt relative to the dots)
    const anchorRange = dotsRange.insertText(" ", Word.InsertLocation.after);
    anchorRange.font.name = "Arial";
    anchorRange.font.size = 12;
    anchorRange.font.position = originalPos;
    await context.sync();

    // Cursor goes after that space, so new typing starts at the original position
    anchorRange.select(Word.SelectionMode.end);
    await context.sync();

    // Word keeps a separate "pending" font for the next typed character, and it can still be the lowered dots.
    // Set the original position explicitly on the live cursor (a collapsed selection) as well.
    const cursor = context.document.getSelection();
    cursor.font.name = "Arial";
    cursor.font.size = 12;
    cursor.font.position = originalPos;
    await context.sync();

    cursor.font.load("position");
    await context.sync();
    if (log) log.push("Blanks: font position at the cursor now = " + cursor.font.position + " pt");
  });
}

/* ---- Total Marks (Sum_Bracketed_Numbers.js, extended) ----
 * Total 1: cells whose whole content is [n]            e.g. [2]
 * Total 2: cells whose whole content is [n marks]      e.g. [2 marks]  (lowercase, plural only)
 * No [n marks] cells  -> Total 1 only.   Totals equal -> that total.   Totals differ -> both are shown. */

const SUM_PLAIN = /^\[(\d{1,2})\]$/;
const SUM_MARKS = /^\[(\d{1,2})\s+marks\]$/;

/** Pure helper, also used by the tests. cellTexts: flat array of cell strings. */
function computeMarksText(cellTexts) {
  let total1 = 0;
  let total2 = 0;
  let found2 = 0;
  for (const raw of cellTexts) {
    const text = String(raw).trim();
    const m1 = SUM_PLAIN.exec(text);
    if (m1) total1 += parseInt(m1[1], 10);
    const m2 = SUM_MARKS.exec(text);
    if (m2) {
      total2 += parseInt(m2[1], 10);
      found2++;
    }
  }
  if (found2 === 0 || total1 === total2) return String(total1);
  return "[n] = " + total1 + ", [n marks] = " + total2;
}

async function sumMarks(log) {
  await Word.run(async (context) => {
    // Capture the cursor first, so the result goes exactly where the user clicked
    const selection = context.document.getSelection();

    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    tables.items.forEach((t) => t.load("values"));
    await context.sync();

    const cells = [];
    tables.items.forEach((t) => t.values.forEach((row) => row.forEach((c) => cells.push(c))));

    const result = computeMarksText(cells);
    selection.insertText(result, Word.InsertLocation.replace);
    await context.sync();
    log.push("Total marks: " + result);
  });
}

/* ------------------------------------------------------------------ */
/* 5. Button actions                                                    */
/* ------------------------------------------------------------------ */

/** Builds mcqQn<Size>Ans, mcqQnPic<Size>Ans, mcqQnPicQn<Size>Ans.
 *  steps: the table(s) that sit between the Qn No / Qn Pic tables and the closing Qn Pic, as [name, builder]. */
function mcqSequences(size, steps) {
  const qnNo = (log) => mcqStep("Qn No", mcqQnNoOoxml, log);
  const qnPic = (log) => mcqStep("Qn Pic", mcqQnPicOoxml, log);
  const mid = async (log) => {
    for (const [name, builder] of steps) await mcqStep(name, builder, log);
  };
  return {
    async ["mcqQn" + size + "Ans"](log) {
      await qnNo(log); await mid(log); await qnPic(log);
    },
    async ["mcqQnPic" + size + "Ans"](log) {
      await qnNo(log); await qnPic(log); await mid(log); await qnPic(log);
    },
    async ["mcqQnPicQn" + size + "Ans"](log) {
      await qnNo(log); await qnPic(log); await qnPic(log); await mid(log); await qnPic(log);
    },
  };
}

const actions = {
  // LAYOUT TEST ONLY: placeholder button, does nothing. Remove with the button.
  async testPlaceholder(log) {},

  // Group 1: Short / Medium / Long  (sequences copied from MCQ_*_Option_Tbl.js)
  //   Qn        : Qn No, Option, Qn Pic
  //   Qn Pic    : Qn No, Qn Pic, Option, Qn Pic
  //   Qn Pic Qn : Qn No, Qn Pic, Qn Pic, Option, Qn Pic
  ...mcqSequences("Short", [["Short options", mcqShortOptionOoxml]]),
  ...mcqSequences("Medium", [["Medium options", mcqMediumOptionOoxml]]),
  ...mcqSequences("Long", [["Long options", mcqLongOptionOoxml]]),

  // Group 2 "Table", 7th button: Qn Pic  (sequence copied from MCQ_Pic_Option.js: Qn No, Qn Pic, Qn Pic)
  async mcqQnPicAns(log) {
    await mcqStep("Qn No", mcqQnNoOoxml, log);
    await mcqStep("Qn Pic", mcqQnPicOoxml, log);
    await mcqStep("Qn Pic", mcqQnPicOoxml, log);
  },
  // 8th button of group 2: Qn Pic Qn  (MCQ_Pic_Option.js: Qn No, Qn Pic, Qn Pic, Qn Pic)
  async mcqQnPicQnAns(log) {
    await mcqStep("Qn No", mcqQnNoOoxml, log);
    await mcqStep("Qn Pic", mcqQnPicOoxml, log);
    await mcqStep("Qn Pic", mcqQnPicOoxml, log);
    await mcqStep("Qn Pic", mcqQnPicOoxml, log);
  },

  // Group 2 "Table": Table Two / Table Three  (sequences copied from MCQ_Table_*_Option_Tbl.js)
  ...mcqSequences("TableTwo", [["Table two options", mcqTableTwoOoxml]]),
  ...mcqSequences("TableThree", [["Table three options", mcqTableThreeOoxml]]),

  // Group 3 "Options": Multi Long / Multi Short  (sequences copied from MCQ_Multi_*_Option_Tbl.js)
  //   Qn        : Qn No, Multi option, Long/Short option, Qn Pic
  //   Qn Pic    : Qn No, Qn Pic, Multi option, Long/Short option, Qn Pic
  //   Qn Pic Qn : Qn No, Qn Pic, Qn Pic, Multi option, Long/Short option, Qn Pic
  ...mcqSequences("MultiLong", [["Multi options", mcqMultiOptionOoxml], ["Long options", mcqLongOptionOoxml]]),
  ...mcqSequences("MultiShort", [["Multi options", mcqMultiOptionOoxml], ["Short options", mcqShortOptionOoxml]]),

  // Group: SQ Q1  (widths/bold flags copied from SQ_Qn.js)
  async sqQ1(log) {
    await sqStep("Q1", [0.95], [{ bold: true }, {}], log);
  },
  async sqQ1a(log) {
    await sqStep("Q1a", [0.95, 1], [{ bold: true }, { bold: true }, {}], log);
  },
  async sqQ1ai(log) {
    await sqStep("Q1ai", [0.95, 1, 1.1],
      [{ bold: true }, { bold: true }, { bold: true }, { justify: "both" }], log);
  },
  async sqQ1ai1(log) {
    await sqStep("Q1ai1", [0.95, 1, 1.1, 0.85],
      [{ bold: true }, { bold: true }, { bold: true }, { bold: true }, {}], log);
  },

  // Group 5 "lined ans"  (widths and alignments copied from SQ_Lined_Ans.js)
  async sqAns1(log) {
    await sqLinedStep("Lined Ans 1", [0.95], "both", "both", log);
  },
  async sqAns1a(log) {
    await sqLinedStep("Lined Ans 1a", [0.95, 1], "left", "left", log);
  },
  async sqAns1ai(log) {
    // answer column is "both" here, as in the original file (every other variant uses "left")
    await sqLinedStep("Lined Ans 1ai", [0.95, 1, 1.1], "left", "both", log);
  },
  async sqAns1ai1(log) {
    await sqLinedStep("Lined Ans 1ai1", [0.95, 1, 1.1, 0.9], "left", "left", log);
  },

  // Group 6 "marks": lined answer + marks  (widths copied from SQ_Lined_Ans_Marks.js)
  async sqLinedMarks1(log) {
    await sqMarksStep("Lined Ans Marks 1", [0.95], 0.85, "linedMarks", log);
  },
  async sqLinedMarks1a(log) {
    await sqMarksStep("Lined Ans Marks 1a", [0.95, 1], 0.85, "linedMarks", log);
  },
  async sqLinedMarks1ai(log) {
    await sqMarksStep("Lined Ans Marks 1ai", [0.95, 1, 1.1], 0.85, "linedMarks", log);
  },
  async sqLinedMarks1ai1(log) {
    await sqMarksStep("Lined Ans Marks 1ai1", [0.95, 1, 1.1, 0.85], 0.85, "linedMarks", log);
  },

  // Group 7 "qn n mk": question + marks  (widths copied from SQ_Qn_Marks.js)
  async sqQnMarks1(log) {
    await sqMarksStep("Q Marks 1", [0.95], 0.85, "qnMarks", log);
  },
  async sqQnMarks1a(log) {
    await sqMarksStep("Q Marks 1a", [0.95, 1], 0.85, "qnMarks", log);
  },
  async sqQnMarks1ai(log) {
    await sqMarksStep("Q Marks 1ai", [0.95, 1, 1.1], 0.85, "qnMarks", log);
  },
  async sqQnMarks1ai1(log) {
    await sqMarksStep("Q Marks 1ai1", [0.95, 1, 1.1, 0.85], 0.85, "qnMarks", log);
  },

  // Group 8 "Cal n Mks"
  async sqCalAns(log) {
    await sqCalStep("Cal Ans", [3.2, 0.85], sqCalAnsOoxml, log);
  },
  async sqCalAnsUnit(log) {
    await sqCalStep("Cal Ans Unit", [3.2, 1.5, 0.85], sqCalAnsUnitOoxml, log);
  },
  async joinTables(log) {
    await joinTables(log);
  },
  async sumMarks(log) {
    await sumMarks(log);
  },

  // Group 9 "Blanks": 8 / 12 / 16 / 20 dots
  async blanksShort(log) { await insertDots(8, log); },
  async blanksMedium(log) { await insertDots(12, log); },
  async blanksLong(log) { await insertDots(16, log); },
  async blanksVeryLong(log) { await insertDots(20, log); },
};

/* ------------------------------------------------------------------ */
/* 6. Error / debug dialog and ribbon registration                      */
/* ------------------------------------------------------------------ */

function showDialog(text) {
  return new Promise((resolve) => {
    try {
      const url = new URL("msg.html", window.location.href).href + "?m=" + encodeURIComponent(text);
      Office.context.ui.displayDialogAsync(url, { height: 35, width: 35 }, () => resolve());
    } catch (e) {
      console.error("Dialog failed:", e, text);
      resolve();
    }
  });
}

function wrap(name, fn) {
  return async function (event) {
    const log = [];
    try {
      await fn(log);
      if (WSB_DEBUG && log.length) await showDialog(log.join("\n"));
    } catch (e) {
      console.error(name, e);
      await showDialog(name + " failed:\n" + (e && e.message ? e.message : String(e)));
    } finally {
      event.completed();
    }
  };
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(() => {
    Object.keys(actions).forEach((name) => Office.actions.associate(name, wrap(name, actions[name])));
  });
}

// Lets the OOXML builders be tested outside Word (node).
if (typeof module !== "undefined" && module.exports) {
  module.exports = { mcqQnNoOoxml, mcqQnPicOoxml, mcqShortOptionOoxml, mcqMediumOptionOoxml, mcqLongOptionOoxml, mcqTableTwoOoxml, mcqTableThreeOoxml, mcqMultiOptionOoxml, sqTableOoxml, sqLinedTableOoxml, sqMarksTableOoxml, sqCalAnsOoxml, sqCalAnsUnitOoxml, sqCalColumnWidths, computeMarksText, CM_PT };
}
