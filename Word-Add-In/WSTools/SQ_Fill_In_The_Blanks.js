
Office.onReady((info)=>{
    if(info.host === Office.HostType.Word){
        if(document.readyState === "loading") {
            document.addEventListener("DOMContentLoaded", initDots);
        } else {
            initDots();
        }
    }
});

function initDots(){

    const dotBtn = document.querySelectorAll("[data-dots]");

    for (const btn of dotBtn){
        const count = parseInt(btn.dataset.dots, 10);
        
        if(!isNaN(count)){
            btn.addEventListener("click", () => insertDots(count));
        }
    }
}

async function insertDots(dotCount) {
    await Word.run(async (context) => {
        //const selection = context.document.getSelection();

        console.log("context:", context);
        console.log("context.document:", context.document);
        const selection = context.document.getSelection();
        console.log("selection:", selection);   // <-- see if this logs undefined
        const paragraph = selection.paragraphs.getFirst();

        // 1. Set paragraph formatting
        paragraph.spaceBefore = 6;
        paragraph.spaceAfter = 0;
        paragraph.lineSpacing = 12;
        paragraph.alignment = Word.Alignment.justified;
        //paragraph.paragraphFormat.spaceBefore = 6;
        //paragraph.paragraphFormat.spaceAfter = 0;
        //paragraph.paragraphFormat.lineSpacing = 12;
        //paragraph.paragraphFormat.alignment = Word.Alignment.justified;

        // 2. Generate specified number of dots
        const dotString = ".".repeat(dotCount);
        const dotsRange = selection.insertText(dotString, Word.InsertLocation.replace);
        dotsRange.font.name = "Arial";
        dotsRange.font.size = 12;
        dotsRange.font.position = -3; // Lower offset by 3pt

        // 3. Reset formatting immediately after dots
        dotsRange.select(Word.SelectionMode.End);
        const cursorPos = dotsRange.getRange(Word.RangeLocation.end);
        //cursorPos.font.name = "Arial";
        //cursorPos.font.size = 12;
        //cursorPos.font.position = 0;
        cursorPos.select();
        const anchorRange = dotsRange.insertText(" ", Word.InsertLocation.after);
        anchorRange.font.name = "Arial";
        anchorRange.font.size = 12;
        anchorRange.font.position = 0;
        //await context.sync();

        //anchorRange.delete();
        //const normalRange = dotsRange.insertText("", Word.InsertLocation.after);
        //normalRange.font.name = "Arial";
        //normalRange.font.size = 12;
        //normalRange.font.position = 0; // Reset back to normal baseline

        //normalRange.select(Word.SelectionMode.end);

        await context.sync();
    });
}