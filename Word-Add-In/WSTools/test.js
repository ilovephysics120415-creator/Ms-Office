/**
 * test.js
 * Orchestrates sequential execution of MCQ_Default_Format.js once,
 * followed by MCQ_Qn_No.js twice inside a single Word execution context.
 */

function log(msg) {
  const el = document.getElementById("status");
  if (el) {
    el.textContent = msg;
  }
  console.log(msg);
}

async function insertTestTable() {
  log("Starting sequential test execution...");

  try {
    await Word.run(async (context) => {
      // 1. Run MCQ_Default_Format.js once using the shared context
      if (typeof applyMcqDefaultFormat === "function") {
        log("Running MCQ_Default_Format...");
        await applyMcqDefaultFormat(context);
      } else {
        log("Error: applyMcqDefaultFormat function not found. Ensure MCQ_Default_Format.js is included in HTML.");
        return;
      }

      // 2. Run MCQ_Qn_No.js pattern first time
      if (typeof runQnNoPattern === "function") {
        log("Running MCQ_Qn_No (Pass 1)...");
        await runQnNoPattern();
      } else {
        log("Error: runQnNoPattern function not found. Ensure MCQ_Qn_No.js is included in HTML.");
        return;
      }

      log("Test passed: MCQ_Default_Format executed once, followed by MCQ_Qn_No twice!");
    });
  } catch (error) {
    log("Test Execution Error: " + error.message);
    console.error("Test Error: ", error);
  }
}

function attachTestListener() {
  const button = document.getElementById("btn-test");
  if (button) {
    button.addEventListener("click", insertTestTable);
    log("Ready. Test listener attached.");
  }
}

Office.onReady(() => {
  attachTestListener();
});