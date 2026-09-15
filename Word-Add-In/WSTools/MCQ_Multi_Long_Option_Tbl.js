/* global Office, Word, createQnNoTable, createMultiOptionTable, createLongOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnMultiLongOptionTbl);
    } else {
      initQnMultiLongOptionTbl();
    }
  }
});

function initQnMultiLongOptionTbl() {
  const btnqnmultilong = document.getElementById("btn-mcq-qn-multi-long-option");
  if (btnqnmultilong) {
    btnqnmultilong.addEventListener("click", createQnMultiLong);
  }

  const btnqnpicmultilong = document.getElementById("btn-mcq-qn-pic-multi-long-option");
  if (btnqnpicmultilong) {
    btnqnpicmultilong.addEventListener("click", createQnPicMultiLong);
  }

  const btnqnpicqnmultilong = document.getElementById("btn-mcq-qn-pic-qn-multi-long-option");
  if (btnqnpicqnmultilong) {
    btnqnpicqnmultilong.addEventListener("click", createQnPicQnMultiLong);
  }
}

async function createQnMultiLong() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Long_Option.js
    await createLongOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicMultiLong() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Long_Option.js
    await createLongOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnMultiLong() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Long_Option.js
    await createLongOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Long Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Long Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}