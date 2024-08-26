const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const vision = require('@google-cloud/vision');
const { Document, Packer, Paragraph, TextRun } = require('docx');
require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

const client = new vision.ImageAnnotatorClient({
    keyFilename: path.join(__dirname, 'google_cloud_keyfile.json'),
});

app.set('view engine', 'ejs');

// Format the extracted text
function formatText(inputText) {
    return inputText.replace(/(\d+\.|\(\w+\))\s*\n\s*/g, '$1 ')
        .replace(/\n(?=[a-zA-Z])/g, ' ')
        .replace(/([.!?])\s*\n\s*/g, '$1\n\n');
}

// Create paragraphs for the Word document
function createParagraphsFromText(text) {
    const lines = text.split('\n');
    return lines.map(line => {
        const trimmedLine = line.trim();

        const isBold = trimmedLine.startsWith('7.') || trimmedLine.startsWith('ALL') || trimmedLine === trimmedLine.toUpperCase();
        const isBullet = /^\d+\./.test(trimmedLine) || /^\(\w+\)/.test(trimmedLine);
        const isItalic = trimmedLine.startsWith('_') && trimmedLine.endsWith('_');

        const textRun = new TextRun({
            text: trimmedLine.replace(/^[-•*]\s*/, ''),
            bold: isBold,
            italics: isItalic,
            size: 24,
        });

        return new Paragraph({
            children: [textRun],
            bullet: isBullet ? { level: 0 } : undefined,
        });
    });
}

// Route to render the upload form
app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

// Route to handle file upload and text extraction
app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const randomNumber = Math.floor(Math.random() * 1000000);
    const outputFileName = `File_${randomNumber}.docx`;

    try {
        const [result] = await client.textDetection(imagePath);
        const detections = result.textAnnotations;
        let text = detections[0].description;

        // Format the text
        text = formatText(text);

        // Create document paragraphs
        const paragraphs = createParagraphsFromText(text);

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

        console.log(`Extracting Done: ${outputFileName}`);
        res.render('index', { docxFile: outputFileName });
    } catch (error) {
        console.error('Error extracting text:', error);
        res.status(500).send('Error processing image.');
    }
});

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, 'output', filename);
    res.download(filePath);
});

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});
