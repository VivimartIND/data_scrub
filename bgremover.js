const fs = require('fs');
const path = require('path');
const removeBg = require('remove-bg'); // Ensure you have the remove-bg package installed

const inputPath = 'C:\\Users\\surya\\Pictures\\Screenshots'; // Input directory
const outputPath = 'C:\\Users\\surya\\Pictures\\Screenshots\\output'; // Output directory

async function removeBackgroundFromImages() {
    try {
        // Ensure the output directory exists
        if (!fs.existsSync(outputPath)) {
            fs.mkdirSync(outputPath, { recursive: true });
        }

        const files = await fs.promises.readdir(inputPath);
        for (const file of files) {
            const ext = path.extname(file).toLowerCase();
            if (['.jpg', '.jpeg', '.png'].includes(ext)) {
                const { name } = path.parse(file);
                const outputFile = path.join(outputPath, `${name}-no-bg.png`);
                try {
                    await removeBg.removeBackgroundFromImageFile({
                        path: path.join(inputPath, file),
                        apiKey: 'YOUR_API_KEY', // Replace with your actual API key
                        outputFile: outputFile,
                    });
                    console.log(`Processed: ${file} -> ${outputFile}`);
                } catch (error) {
                    console.error(`Error processing ${file}:`, error);
                }
            } else {
                console.log(`Skipped non-image file: ${file}`);
            }
        }
    } catch (error) {
        console.error('Error reading input directory:', error);
    }
}

removeBackgroundFromImages(); // Call the function to start processing