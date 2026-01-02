/* global Word, console */

async function debugDocumentStructure() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("text, tableNestingLevel");

        const tables = body.tables;
        tables.load("items");

        await context.sync();

        console.log("--- Document Structure Debug ---");
        console.log(`Total Paragraphs: ${paragraphs.items.length}`);

        paragraphs.items.forEach((p, i) => {
            console.log(`[P${i + 1}] (Nesting: ${p.tableNestingLevel}): "${p.text.substring(0, 50)}..."`);
        });

        console.log(`Total Tables: ${tables.items.length}`);
    });
}
