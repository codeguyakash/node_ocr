const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { TextractClient, DetectDocumentTextCommand } = require('@aws-sdk/client-textract');
const { Document, Packer, Paragraph, TextRun } = require('docx');
require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

const textractClient = new TextractClient({
    region: process.env.AWS_REGION,
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
});

app.set('view engine', 'ejs');

function formatText(inputText) {
    return inputText.replace(/(\d+\.|\(\w+\))\s*\n\s*/g, '$1 ')
        .replace(/\n(?=[a-zA-Z])/g, ' ')
        .replace(/([.!?])\s*\n\s*/g, '$1\n\n');
}
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

app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const randomNumber = Math.floor(Math.random() * 1000000);
    const outputFileName = `File:${randomNumber}.docx`;

    try {
        const imageBytes = fs.readFileSync(imagePath);

        const command = new DetectDocumentTextCommand({
            Document: {
                Bytes: imageBytes,
            },
        });

        const textractResult = await textractClient.send(command);

        let text = textractResult.Blocks
            .filter(block => block.BlockType === 'LINE')
            .map(block => block.Text)
            .join('\n');

        text = formatText(text);

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
    console.log(`Server with AWS http://localhost:${PORT}`);
});
