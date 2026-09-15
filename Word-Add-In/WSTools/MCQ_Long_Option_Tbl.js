/* global Office, Word, createQnNoTable, createLongOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnLongOptionTbl);
    } else {
      initQnLongOptionTbl();
    }
  }
});

function initQnLongOptionTbl() {
  const btnqnlong = document.getElementById("btn-mcq-qn-long-option");
  if (btnqnlong) {
    btnqnlong.addEventListener("click", createQnLongOption);
  }

  const btnqnpiclong = document.getElementById("btn-mcq-qn-pic-long-option");
  if (btnqnpiclong) {
    btnqnpiclong.addEventListener("click", createQnPicLongOption);
  }

  const btnqnpicqnlong = document.getElementById("btn-mcq-qn-pic-qn-long-option");
  if (btnqnpicqnlong) {
    btnqnpicqnlong.addEventListener("click", createQnPicQnLongOption);
  }

}

async function createQnLongOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Long_Option.js
    await createLongOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicLongOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    // 2. MCQ_Long_Option.js
    await createLongOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnLongOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Long_Option.js
    await createLongOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Pic Qn Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Pic Qn Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}