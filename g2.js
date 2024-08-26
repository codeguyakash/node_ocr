const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const vision = require('@google-cloud/vision');
const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    AlignmentType,
    PageBreak,
    BorderStyle,
} = require('docx');

require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;
const time = new Date();

// Google Cloud Vision Client setup
const client = new vision.ImageAnnotatorClient({
    keyFilename: path.join(__dirname, 'google_cloud_keyfile.json'),
});

// Set EJS as the templating engine
app.set('view engine', 'ejs');

// Function to create formatted paragraphs from detected text
function createFormattedParagraphs(text) {
    const lines = text.split('\n');
    const children = [];

    const headingRegex = /^\s*(\d+\.\s*|\w+\s*)\s*$/;
    const listItemRegex = /^\s*\(\w\)\s*$/;

    lines.forEach(line => {
        const trimmedLine = line.trim();

        if (headingRegex.test(trimmedLine)) {
            const headingLevel = HeadingLevel.HEADING_1; // You can adjust heading level logic if necessary

            children.push(
                new Paragraph({
                    children: [new TextRun({ text: trimmedLine, bold: true })],
                    heading: headingLevel,
                })
            );
        } else if (listItemRegex.test(trimmedLine)) {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: trimmedLine })],
                    bullet: { level: 0 },
                })
            );
        } else {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: trimmedLine })],
                    alignment: AlignmentType.LEFT, // You can change alignment if necessary
                })
            );
        }
    });

    return children;
}

// Route to render the upload form
app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

// Route to handle file upload and text extraction
app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const outputFileName = `File-${time}.docx`;

    try {
        const [result] = await client.textDetection(imagePath);
        const detections = result.textAnnotations;
        const text = detections[0].description;

        const paragraphs = createFormattedParagraphs(text);

        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: paragraphs,
                },
            ],
        });

        const buffer = await Packer.toBuffer(doc);
        const outputDir = path.join(__dirname, 'output');
        const docxPath = path.join(outputDir, outputFileName);

        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }

        fs.writeFileSync(docxPath, buffer);
        fs.unlinkSync(imagePath);
        console.log("Extraction Done");
        res.render('index', { docxFile: outputFileName });
    } catch (error) {
        console.error('Error extracting text:', error);
        res.status(500).send('Error processing image.');
    }
});

// Route to handle document download
app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, 'output', filename);
    res.download(filePath);
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});


//############################################################################################ =============================================================================================>


// const express = require('express');
// const multer = require('multer');
// const path = require('path');
// const fs = require('fs');
// const vision = require('@google-cloud/vision');
// const { Document, Packer, Paragraph, TextRun } = require('docx');

// require('dotenv').config();

// const app = express();
// const upload = multer({ dest: 'uploads/' });
// const PORT = process.env.PORT || 3000;

// // Google Cloud Vision Client setup
// const client = new vision.ImageAnnotatorClient({
//     keyFilename: path.join(__dirname, 'google_cloud_keyfile.json'),
// });

// // Set EJS as the templating engine
// app.set('view engine', 'ejs');

// // Function to create paragraphs from detected text
// function createParagraphsFromText(text) {
//     const lines = text.split('\n');
//     return lines.map(line => {
//         const trimmedLine = line.trim();
//         const isBullet = trimmedLine.startsWith('•') || trimmedLine.startsWith('-') || trimmedLine.startsWith('*') || /^\d+\./.test(trimmedLine);
//         const isBold = trimmedLine.toUpperCase() === trimmedLine;
//         const isItalic = trimmedLine.startsWith('_') && trimmedLine.endsWith('_');

//         const textRun = new TextRun({
//             text: trimmedLine.replace(/^[-•*]\s*/, ''),
//             bold: isBold,
//             italics: isItalic,
//             size: 24,
//         });

//         return new Paragraph({
//             children: [textRun],
//             bullet: isBullet ? { level: 0 } : undefined,
//         });
//     });
// }

// // Route to render the upload form
// app.get('/', (req, res) => {
//     res.render('index', { docxFile: null });
// });

// // Route to handle file upload and text extraction
// app.post('/upload', upload.single('image'), async (req, res) => {
//     const imagePath = req.file.path;
//     const outputFileName = `File_${Date.now()}.docx`;

//     try {
//         const [result] = await client.textDetection(imagePath);
//         const detections = result.textAnnotations;
//         const text = detections[0].description;

//         const paragraphs = createParagraphsFromText(text);

//         const doc = new Document({
//             sections: [
//                 {
//                     properties: {},
//                     children: paragraphs,
//                 },
//             ],
//         });

//         const buffer = await Packer.toBuffer(doc);
//         const outputDir = path.join(__dirname, 'output');
//         const docxPath = path.join(outputDir, outputFileName);

//         if (!fs.existsSync(outputDir)) {
//             fs.mkdirSync(outputDir);
//         }

//         fs.writeFileSync(docxPath, buffer);
//         fs.unlinkSync(imagePath);

//         res.render('index', { docxFile: outputFileName });
//     } catch (error) {
//         console.error('Error extracting text:', error);
//         res.status(500).send('Error processing image.');
//     }
// });

// // Route to handle document download
// app.get('/download/:filename', (req, res) => {
//     const filename = req.params.filename;
//     const filePath = path.join(__dirname, 'output', filename);
//     res.download(filePath);
// });

// // Start the server
// app.listen(PORT, () => {
//     console.log(`Server started on http://localhost:${PORT}`);
// });
