/* global Office, PowerPoint */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("PowerPoint Add-in loaded");

        // UI initialisieren
        document.getElementById("status").textContent = "Add-in bereit";
    }
});


async function runStandardization() {
    const status = document.getElementById("status");
    status.innerText = "Scanning document...";

    try {
        await PowerPoint.run(async (context) => {
            const targetFont = document.getElementById("font-select").value;
            const hMin = parseFloat(document.getElementById("h-min").value);
            const sMin = parseFloat(document.getElementById("s-min").value);

            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();

            let totalUpdated = 0;

            for (let slide of slides.items) {
                const shapes = slide.shapes;
                shapes.load("items");
                await context.sync();

                // Process shapes using a recursive helper to handle Groups
                totalUpdated += await processShapesCollection(shapes.items, targetFont, hMin, sMin, context);
            }

            await context.sync();
            status.innerText = `Done! Processed ${totalUpdated} text blocks.`;
        });
    } catch (error) {
        status.innerText = "Error: " + error.message;
    }
}

async function processShapesCollection(shapes, targetFont, hMin, sMin, context) {
    let count = 0;
    for (let shape of shapes) {
        shape.load("hasTextFrame, type");
        await context.sync();

        if (shape.type === "Group") {
            // Drill down into groups
            const groupShapes = shape.groupItems;
            groupShapes.load("items");
            await context.sync();
            count += await processShapesCollection(groupShapes.items, targetFont, hMin, sMin, context);
        } else if (shape.hasTextFrame) {
            const textRange = shape.textFrame.textRange;
            textRange.load("font/size");
            await context.sync();

            // Apply font change to everything
            // Note: You can add specific 'textRange.font.size = X' logic here if needed
            textRange.font.name = targetFont;
            count++;
        }
    }
    return count;
}
