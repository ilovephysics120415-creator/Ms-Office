/**
 * MCQ_Default_Format.js
 * Applies default paragraph/font formatting and calculates the global table width (in twips).
 */

// Shared global variables accessible across all scripts
window.tableWidthTwips = 0;
window.tableWidthPoints = 0;

async function applyMcqDefaultFormat(context) {
  const formatLogic = async (ctx) => {
    // 1. Calculate section table width (Page Width - Right Margin - Left Margin)
    const section = ctx.document.sections.getFirst();
    section.load(["pageSetup/pageWidth", "pageSetup/leftMargin", "pageSetup/rightMargin"]);
    await ctx.sync();

    const pageSetup = section.pageSetup;
    const printableWidthPoints = pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin;

    // Store globally on window object (1 pt = 20 twips for OOXML)
    window.tableWidthPoints = printableWidthPoints;
    window.tableWidthTwips = Math.round(printableWidthPoints * 20);

    // 2. Set default paragraph and font properties on current selection
    const selection = ctx.document.getSelection();
    selection.font.name = "Arial";
    selection.font.size = 12;

    // Corrected: Set paragraph properties directly on selection (not paragraphFormat)
    selection.spaceBefore = 6;
    selection.spaceAfter = 6;
    selection.lineSpacing = 12; // 12pt equals Single line spacing for 12pt font
    selection.alignment = Word.Alignment.justified;

    await ctx.sync();
    updateStatus("Applied MCQ Default Format and updated table width successfully!");
  };

  // If a context is passed from another script, use it; otherwise run its own Word.run block
  if (context) {
    await formatLogic(context);
  } else {
    try {
      await Word.run(async (ctx) => {
        await formatLogic(ctx);
      });
    } catch (error) {
      updateStatus("Error applying default format: " + error.message);
      console.error("Format Error: ", error);
    }
  }
}

// Utility function to update the status pane UI
function updateStatus(msg) {
  const statusEl = document.getElementById("status");
  if (statusEl) {
    statusEl.textContent = msg;
  }
  console.log(msg);
}

// Register handler when Office JS is initialized
Office.onReady(() => {
  const btn = document.getElementById("btn-mcq-default-format");
  if (btn) {
    btn.addEventListener("click", () => applyMcqDefaultFormat());
  }
});