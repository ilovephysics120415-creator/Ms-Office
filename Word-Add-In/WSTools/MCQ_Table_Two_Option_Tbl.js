/* global Office, Word, createQnNoTable, createTableTwoOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnTableTwoOptionTbl);
    } else {
      initQnTableTwoOptionTbl();
    }
  }
});

function initQnTableTwoOptionTbl() {
  const btnqntabletwo = document.getElementById("btn-mcq-qn-table-two-option");
  if (btnqntabletwo) {
    btnqntabletwo.addEventListener("click", createQnTableTwoOption);
  }

  const btnqnpictabletwo = document.getElementById("btn-mcq-qn-pic-table-two-option");
  if (btnqnpictabletwo) {
    btnqnpictabletwo.addEventListener("click", createQnPicTableTwoOption);
  }

  const btnqnpicqntabletwo = document.getElementById("btn-mcq-qn-pic-qn-table-two-option");
  if (btnqnpicqntabletwo) {
    btnqnpicqntabletwo.addEventListener("click", createQnPicQnTableTwoOption);
  }
}

async function createQnTableTwoOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Table_Two_Option.js
    await createTableTwoOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Two Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Two Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicTableTwoOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    // 2. MCQ_Table_Two_Option.js
    await createTableTwoOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Two Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Two Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnTableTwoOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Table_Two_Option.js
    await createTableTwoOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Two Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Two Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}