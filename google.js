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
    keyFilename: path.join(__dirname, 'your-google-cloud-keyfile.json'),
});

app.set('view engine', 'ejs');

function createParagraphsFromText(text) {
    const lines = text.split('\n');
    return lines.map(line => {
        const trimmedLine = line.trim();
        const isBullet = trimmedLine.startsWith('•') || trimmedLine.startsWith('-') || trimmedLine.startsWith('*') || /^\d+\./.test(trimmedLine);
        const isBold = trimmedLine.toUpperCase() === trimmedLine;  // Simple bold detection
        const isIndented = line.startsWith('    ') || line.startsWith('\t');
        const isItalic = trimmedLine.startsWith('_') && trimmedLine.endsWith('_');  // Simple italic detection

        const textRun = new TextRun({
            text: trimmedLine.replace(/^[-•*]\s*/, ''),  // Remove bullet symbols
            bold: isBold,
            italics: isItalic,
            size: 24,
        });

        return new Paragraph({
            children: [textRun],
            bullet: isBullet ? { level: 0 } : undefined,
            indent: isIndented ? { left: 720 } : undefined,
        });
    });
}

app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const outputFileName = `File_${Date.now()}.docx`;

    try {
        const [result] = await client.textDetection(imagePath);
        const detections = result.textAnnotations;
        const text = detections[0].description;

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

        // Delete the uploaded image after processing
        fs.unlinkSync(imagePath);

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
