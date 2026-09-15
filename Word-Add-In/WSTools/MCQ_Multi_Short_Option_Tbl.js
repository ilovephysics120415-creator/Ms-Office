/* global Office, Word, createQnNoTable, createMultiOptionTable, createShortOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnMultiShortOptionTbl);
    } else {
      initQnMultiShortOptionTbl();
    }
  }
});

function initQnMultiShortOptionTbl() {
  const btnqnmultishort = document.getElementById("btn-mcq-qn-multi-short-option");
  if (btnqnmultishort) {
    btnqnmultishort.addEventListener("click", createQnMultiShort);
  }

  const btnqnpicmultishort = document.getElementById("btn-mcq-qn-pic-multi-short-option");
  if (btnqnpicmultishort) {
    btnqnpicmultishort.addEventListener("click", createQnPicMultiShort);
  }

  const btnqnpicqnmultishort = document.getElementById("btn-mcq-qn-pic-qn-multi-short-option");
  if (btnqnpicqnmultishort) {
    btnqnpicqnmultishort.addEventListener("click", createQnPicQnMultiShort);
  }
}

async function createQnMultiShort() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createShortOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Short Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}
  async function createQnPicMultiShort() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();
  
    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createShortOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Short Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnMultiShort() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createShortOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Short Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}