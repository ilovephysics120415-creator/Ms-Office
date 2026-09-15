/* global Office, Word, createQnNoTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnPicOptionTbl);
    } else {
      initQnPicOptionTbl();
    }
  }
});

function initQnPicOptionTbl() {
  const btnqnpic = document.getElementById("btn-mcq-qn-pic-option");
  if (btnqnpic) {
    btnqnpic.addEventListener("click", createQnPicOption);
  }

  const btnqnpicqn = document.getElementById("btn-mcq-qn-pic-qn-option");
  if (btnqnpicqn) {
    btnqnpicqn.addEventListener("click", createQnPicQnOption);
  }
}

async function createQnPicOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Qn_Pic.js
    await createQnPicTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Qn_Pic.js
    await createQnPicTable();

    await createQnPicTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}