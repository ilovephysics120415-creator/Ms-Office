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

const mcqCell = (w, inner) =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr><w:p>${MCQ_PPR}${inner || ""}</w:p></w:tc>`;

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
    rows.map((inner) => widths.map((w, i) => mcqCell(w, inner[i]))),
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
    [["", mcqBoldText("A"), "", mcqBoldText("B"), "", mcqBoldText("C"), "", mcqBoldText("D"), ""]],
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
      ["", mcqBoldText("A"), "", mcqBoldText("B"), ""],
      ["", mcqBoldText("C"), "", mcqBoldText("D"), ""],
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
      ["", mcqBoldText("A"), ""],
      ["", mcqBoldText("B"), ""],
      ["", mcqBoldText("C"), ""],
      ["", mcqBoldText("D"), ""],
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
  }</w:tcPr><w:p>${center ? MCQ_PPR_CENTER : MCQ_PPR}${text ? mcqBoldText(text) : ""}</w:p></w:tc>`;
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

function sqTableOoxml(colDxa, cellOpts) {
  const totalDxa = colDxa.reduce((s, w) => s + w, 0);
  const grid = colDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  const cells = colDxa.map((w, i) => sqCell(w, cellOpts[i])).join("");
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
                            <w:tblBorders>
                                <w:top w:val="none"/>
                                <w:left w:val="none"/>
                                <w:bottom w:val="none"/>
                                <w:right w:val="none"/>
                                <w:insideH w:val="none"/>
                                <w:insideV w:val="none"/>
                            </w:tblBorders>
                        </w:tblPr>
                        <w:tblGrid>${grid}</w:tblGrid>
                        <w:tr>${cells}</w:tr>
                    </w:tbl>
                </w:body>
            </w:document>
        </pkg:xmlData>
    </pkg:part>
</pkg:package>`;
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

/** fixedCm: widths (cm) of the fixed columns. The last column takes the remaining width. */
async function sqStep(name, fixedCm, cellOpts, log) {
  const g = await readGeometry();                       // fresh width, every step
  log.push(describe(name, g));

  await Word.run((ctx) => formatSelection(ctx, "sq", g.widthPt)); // Arial 12 + spacing, every step

  const dxa = fixedCm.map((cm) => Math.round(cm * CM_PT * 20));
  const usedPt = fixedCm.reduce((s, cm) => s + cm * CM_PT, 0);
  const lastDxa = Math.round((g.widthPt - usedPt) * 20);
  if (lastDxa < 200) {
    throw new Error("Text area is too narrow for this table (last column would be " + lastDxa + " twips).");
  }
  dxa.push(lastDxa);

  await Word.run((ctx) => insertSq(ctx, sqTableOoxml(dxa, cellOpts)));
}

/* ------------------------------------------------------------------ */
/* 5. Button actions                                                    */
/* ------------------------------------------------------------------ */

/** Builds mcqQn<Size>Ans, mcqQnPic<Size>Ans, mcqQnPicQn<Size>Ans. */
function mcqSequences(size, optName, optBuilder) {
  const qnNo = (log) => mcqStep("Qn No", mcqQnNoOoxml, log);
  const qnPic = (log) => mcqStep("Qn Pic", mcqQnPicOoxml, log);
  const opt = (log) => mcqStep(optName, optBuilder, log);
  return {
    async ["mcqQn" + size + "Ans"](log) {
      await qnNo(log); await opt(log); await qnPic(log);
    },
    async ["mcqQnPic" + size + "Ans"](log) {
      await qnNo(log); await qnPic(log); await opt(log); await qnPic(log);
    },
    async ["mcqQnPicQn" + size + "Ans"](log) {
      await qnNo(log); await qnPic(log); await qnPic(log); await opt(log); await qnPic(log);
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
  ...mcqSequences("Short", "Short options", mcqShortOptionOoxml),
  ...mcqSequences("Medium", "Medium options", mcqMediumOptionOoxml),
  ...mcqSequences("Long", "Long options", mcqLongOptionOoxml),

  // Group 2 "Table": Table Two / Table Three  (sequences copied from MCQ_Table_*_Option_Tbl.js)
  ...mcqSequences("TableTwo", "Table two options", mcqTableTwoOoxml),
  ...mcqSequences("TableThree", "Table three options", mcqTableThreeOoxml),

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
  module.exports = { mcqQnNoOoxml, mcqQnPicOoxml, mcqShortOptionOoxml, mcqMediumOptionOoxml, mcqLongOptionOoxml, mcqTableTwoOoxml, mcqTableThreeOoxml, sqTableOoxml, CM_PT };
}
