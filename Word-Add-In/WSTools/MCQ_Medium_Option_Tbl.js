/* global Office, Word, createQnNoTable, createMediumOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnMediumOptionTbl);
    } else {
      initQnMediumOptionTbl();
    }
  }
});

function initQnMediumOptionTbl() {
  const btnqnmedium = document.getElementById("btn-mcq-qn-medium-option");
  if (btnqnmedium) {
    btnqnmedium.addEventListener("click", createQnMediumOption);
  }

  const btnqnpicmedium = document.getElementById("btn-mcq-qn-pic-medium-option");
  if (btnqnpicmedium) {
    btnqnpicmedium.addEventListener("click", createQnPicMediumOption);
  }

  const btnqnpicqnmedium = document.getElementById("btn-mcq-qn-pic-qn-medium-option");
  if (btnqnpicqnmedium) {
    btnqnpicqnmedium.addEventListener("click", createQnPicQnMediumOption);
  }
}

async function createQnMediumOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Medium_Option.js
    await createMediumOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Medium Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicMediumOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    // 2. MCQ_Medium_Option.js
    await createMediumOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Medium Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnMediumOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Medium_Option.js
    await createMediumOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Medium Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}