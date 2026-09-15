/* global Office, Word, createQnNoTable, createMultiOptionTable, createShortOptionTable, createQnPicTable, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnMultiMediumOptionTbl);
    } else {
      initQnMultiMediumOptionTbl();
    }
  }
});

function initQnMultiMediumOptionTbl() {
  const btnqnmultimedium = document.getElementById("btn-mcq-qn-multi-medium-option");
  if (btnqnmultimedium) {
    btnqnmultimedium.addEventListener("click", createQnMultiMedium);
  }

  const btnqnpicmultimedium = document.getElementById("btn-mcq-qn-pic-multi-medium-option");
  if (btnqnpicmultimedium) {
    btnqnpicmultimedium.addEventListener("click", createQnPicMultiMedium);
  }

  const btnqnpicqnmultimedium = document.getElementById("btn-mcq-qn-pic-qn-multi-medium-option");
  if (btnqnpicqnmultimedium) {
    btnqnpicqnmultimedium.addEventListener("click", createQnPicQnMultiMedium);
  }
}

async function createQnMultiMedium() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createMediumOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}
  async function createQnPicMultiMedium() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();
  
    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createMediumOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}

async function createQnPicQnMultiMedium() {
  try {
    // 1. MCQ_Qn_No.js
    await createQnNoTable();

    await createQnPicTable();

    await createQnPicTable();

    // 2. MCQ_Multi_Option.js
    await createMultiOptionTable();

    // 3. MCQ_Short_Option.js
    await createMediumOptionTable();

    // 4. MCQ_Qn_Pic.js
    await createQnPicTable();

    updateStatus("Executed Qn Multi Medium Option sequence successfully!");
  } catch (error) {
    updateStatus("Error executing Qn Multi Short Option sequence: " + error.message);
    console.error("Error executing sequence: ", error);
  }
}