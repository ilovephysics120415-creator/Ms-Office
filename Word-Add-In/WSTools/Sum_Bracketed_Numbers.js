/* global Office, Word, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initSumBracketedNumbers);
    } else {
      initSumBracketedNumbers();
    }
  }
});

function initSumBracketedNumbers() {
  const btn = document.getElementById("btn-sum-bracket");
  if (btn) {
    btn.addEventListener("click", sumBracketedNumbers);
  }
}

async function sumBracketedNumbers() {
  try {
    await Word.run(async (context) => {
      // 1. Capture the current cursor/selection position FIRST, before any
      // reading happens, so we insert the result exactly where the user's
      // cursor was when they clicked the button.
      const selection = context.document.getSelection();

      // 2. Get every table in the document body.
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      let total = 0;
      // Anchored with ^ and $: the ENTIRE cell content (after trimming
      // surrounding whitespace) must be exactly "[" + 1-2 digits + "]" --
      // nothing before the "[" and nothing after the "]". A cell like
      // "See [12] above" no longer counts; only a cell that is just "[12]" does.
      const numberPattern = /^\[(\d{1,2})\]$/;

      // 3. For each table, load its full 2D cell-text array in one call.
      for (const table of tables.items) {
        table.load("values");
      }
      await context.sync();

      // 4. Walk every row, every cell -- a cell contributes to the total
      // only if its whole (trimmed) content matches the pattern.
      for (const table of tables.items) {
        for (const row of table.values) {
          for (const cellText of row) {
            const match = numberPattern.exec(cellText.trim());
            if (match) {
              total += parseInt(match[1], 10);
            }
          }
        }
      }

      // 5. Insert the total at the captured cursor position.
      selection.insertText(String(total), Word.InsertLocation.replace);
      await context.sync();

      updateStatus("Summed bracketed numbers: " + total);
    });
  } catch (error) {
    updateStatus("Error summing bracketed numbers: " + error.message);
    console.error("Error summing bracketed numbers: ", error);
  }
}
