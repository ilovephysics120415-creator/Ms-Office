/* global Office, Word, createQnNoTable, createTableThreeOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnTableThreeOptionTbl);
    } else {
      initQnTableThreeOptionTbl();
    }
  }
});

function initQnTableThreeOptionTbl() {
  const btnqntablethree = document.getElementById("btn-mcq-qn-table-three-option");
  if (btnqntablethree) {
    btnqntablethree.addEventListener("click", createQnTableThreeOption);
  }

  const btnqnpictablethree = document.getElementById("btn-mcq-qn-pic-table-three-option");
  if (btnqnpictablethree) {
    btnqnpictablethree.addEventListener("click", createQnPicTableThreeOption);
  }

  const btnqnpicqntablethree = document.getElementById("btn-mcq-qn-pic-qn-table-three-option");
  if (btnqnpicqntablethree) {
    btnqnpicqntablethree.addEventListener("click", createQnPicQnTableThreeOption);
  }
}

async function createQnTableThreeOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Table_Three_Option.js
    await createTableThreeOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Three Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Three Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicTableThreeOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    // 2. MCQ_Table_Three_Option.js
    await createTableThreeOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Three Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Three Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnTableThreeOption() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Table_Three_Option.js
    await createTableThreeOptionTable();

    // 3. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Table Three Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Table Three Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}