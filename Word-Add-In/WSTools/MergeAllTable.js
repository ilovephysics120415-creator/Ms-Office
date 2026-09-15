// Function to display status messages in the UI task pane
function updateStatus(msg) {
  const statusEl = document.getElementById("status");
  if (statusEl) {
    statusEl.textContent = msg;
  }
  console.log(msg);
}

// Merges all adjacent tables in the Word document
async function mergeAllTables() {
  updateStatus("Scanning document to merge adjacent tables...");

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const tables = body.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length < 2) {
        updateStatus("Fewer than 2 tables found in document.");
        return;
      }

      let mergedCount = 0;

      // Loop backwards through tables so indices remain stable as we go
      for (let i = tables.items.length - 1; i > 0; i--) {
        const currentTable = tables.items[i];
        const prevTable = tables.items[i - 1];

        // Get the paragraph sitting immediately after the previous table.
        // Note: Word.RangeLocation.before is only valid for ContentControl,
        // not for Table - calling currentTable.getRange(before) throws
        // InvalidArgument. So instead of spanning between the two tables,
        // we just grab the single paragraph right after prevTable.
        const gapRange = prevTable.getRange(Word.RangeLocation.after);
        const gapParagraph = gapRange.paragraphs.getFirst();
        gapParagraph.load("text");
        await context.sync();

        // Only merge if that paragraph is genuinely empty. If there's real
        // content between the tables (a heading, notes, etc.), leave that pair alone.
        if (gapParagraph.text.trim() === "") {
          // Deleting the empty paragraph is exactly what happens when you
          // manually forward-delete between two tables in Word - Word's own
          // engine merges them automatically once nothing separates them.
          gapParagraph.delete();
          await context.sync();
          mergedCount++;
        }
      }

      if (mergedCount > 0) {
        updateStatus(`Successfully merged ${mergedCount} table(s)!`);
      } else {
        updateStatus("No adjacent tables (separated only by an empty paragraph) found to merge.");
      }
    });
  } catch (error) {
    updateStatus("Error merging tables: " + error.message);
    console.error(error);
  }
}

// Attach event listener when Office JS initialized
Office.onReady(() => {
  const mergeBtn = document.getElementById("btn-merge-tables");
  if (mergeBtn) {
    mergeBtn.addEventListener("click", mergeAllTables);
  }
});
