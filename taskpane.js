/* global Office, PowerPoint */

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("apply-btn").onclick = runStandardization;
    }
});

async function runStandardization() {
    const status = document.getElementById("status");
    status.innerText = "Processing slides...";

    try {
        await PowerPoint.run(async (context) => {
            const targetFont = document.getElementById("font-select").value;
            const hMin = parseFloat(document.getElementById("h-min").value);
            const sMin = parseFloat(document.getElementById("s-min").value);

            const presentation = context.presentation;
            const slides = presentation.slides;
            slides.load("items");
            await context.sync();

            let totalUpdated = 0;

            for (let i = 0; i < slides.items.length; i++) {
                const slide = slides.items[i];
                const shapes = slide.shapes;
                shapes.load("items");
                await context.sync();

                for (let j = 0; j < shapes.items.length; j++) {
                    const shape = shapes.items[j];
                    
                    // We load hasTextFrame to verify if we can manipulate text
                    shape.load("hasTextFrame");
                    await context.sync();

                    if (shape.hasTextFrame) {
                        const textRange = shape.textFrame.textRange;
                        textRange.load("font/size, font/name");
                        await context.sync();

                        // Logic: Compare current size to UI thresholds
                        const currentSize = textRange.font.size;

                        if (currentSize >= hMin || (currentSize >= sMin && currentSize < hMin) || currentSize < sMin) {
                            // In this implementation, we apply the chosen font to all 
                            // but you can add specific size overrides here if desired.
                            textRange.font.name = targetFont;
                            totalUpdated++;
                        }
                    }
                }
            }

            await context.sync();
            status.innerText = `Success! Updated ${totalUpdated} text blocks.`;
        });
    } catch (error) {
        status.innerText = "Error: " + error.message;
        console.error(error);
    }
}
