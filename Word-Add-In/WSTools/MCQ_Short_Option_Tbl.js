/* global Office, Word, createQnNoTable, createShortOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initShortOptionTbl);
    } else {
      initShortOptionTbl();
    }
  }
});

function initShortOptionTbl() {
  const btnqnshort = document.getElementById("btn-mcq-qn-short-option");
  if (btnqnshort) {
    btnqnshort.addEventListener("click", createQnShortOption);
  }

  const btnqnpicshort = document.getElementById("btn-mcq-qn-pic-short-option");
  if (btnqnpicshort) {
    btnqnpicshort.addEventListener("click", createQnPicShortOption);
  }

  const btnqnpicqnshortoption = document.getElementById("btn-mcq-qn-pic-qn-short-option");
  if (btnqnpicqnshortoption) {
    btnqnpicqnshortoption.addEventListener("click", createQnPicQnShortOption);
  }
}

async function createQnShortOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Short_Option.js
    await createShortOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Short Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicShortOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Qn_Pic.js
    await createQnPicTable();

    await createShortOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnShortOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Qn_Pic.js
    await createQnPicTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    // 4. MCQ_Short_Option.js
    await createShortOptionTable();

    // 5. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Qn Short Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Qn Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

