// SQ_Default_Format.js

var SQ_TABLE_WIDTH = null;

async function applySQDefaultFormat() {
    await Word.run(async (context) => {
        // 1. Calculate Page Width minus Margins
        const section = context.document.sections.getFirst();
        const pageSetup = section.pageSetup;
        
        pageSetup.load(["pageWidth", "leftMargin", "rightMargin"]);
        await context.sync();

        // Calculate maximum table width
        SQ_TABLE_WIDTH = pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin;
        window.SQ_TABLE_WIDTH = SQ_TABLE_WIDTH;

        // 2. Apply formatting to Selection / Paragraphs
        const selection = context.document.getSelection();
        const paragraphs = selection.paragraphs;
        
        paragraphs.load("items");
        await context.sync();

        paragraphs.items.forEach((paragraph) => {
            paragraph.font.name = "Arial";
            paragraph.font.size = 12;
            paragraph.alignment = Word.Alignment.justified;
            paragraph.spaceBefore = 6;
        });

        // 3. Format Tables in Selection to be Borderless & Match Width
        const tables = selection.tables;
        tables.load("items");
        await context.sync();

        tables.items.forEach((table) => {
            table.preferredWidth = SQ_TABLE_WIDTH;
            
            const borderTypes = [
                Word.BorderType.top,
                Word.BorderType.bottom,
                Word.BorderType.left,
                Word.BorderType.right,
                Word.BorderType.insideHorizontal,
                Word.BorderType.insideVertical
            ];

            borderTypes.forEach((borderType) => {
                const border = table.getBorder(borderType);
                border.type = Word.BorderType.none;
            });
        });

        await context.sync();
    });
}

// Calculate width automatically when Office initializes
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        Word.run(async (context) => {
            const section = context.document.sections.getFirst();
            const pageSetup = section.pageSetup;
            pageSetup.load(["pageWidth", "leftMargin", "rightMargin"]);
            await context.sync();
            
            SQ_TABLE_WIDTH = pageSetup.pageWidth - pageSetup.leftMargin - pageSetup.rightMargin;
            window.SQ_TABLE_WIDTH = SQ_TABLE_WIDTH;
        }).catch((err) => {
            console.error("Error calculating initial page width:", err);
        });
    }
});