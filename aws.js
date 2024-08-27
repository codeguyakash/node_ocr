const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { TextractClient, DetectDocumentTextCommand } = require('@aws-sdk/client-textract');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat, HeadingLevel, convertInchesToTwip } = require('docx');
require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

// Initialize AWS Textract client using v3 SDK
const textractClient = new TextractClient({
    region: process.env.AWS_REGION,
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
});

app.set('view engine', 'ejs');

// Function to detect if the text is in two-column format
function isTwoColumnFormat(text) {
    const lines = text.split('\n');
    const leftColumnLines = lines.filter(line => /^\d+\.\s|^[a-z]\)/i.test(line.trim())).length;
    const rightColumnLines = lines.length - leftColumnLines;
    return Math.abs(leftColumnLines - rightColumnLines) < lines.length * 0.3;
}

// Function to process and reorder text from a two-column layout
function processTwoColumnText(text) {
    const lines = text.split('\n');
    const leftColumn = [];
    const rightColumn = [];
    let midpoint = Math.floor(lines.length / 2);

    lines.forEach((line, index) => {
        if (index < midpoint) {
            leftColumn.push(line);
        } else {
            rightColumn.push(line);
        }
    });

    // Interleave left and right columns
    const mergedText = [];
    const maxLength = Math.max(leftColumn.length, rightColumn.length);

    for (let i = 0; i < maxLength; i++) {
        if (i < leftColumn.length) mergedText.push(leftColumn[i]);
        if (i < rightColumn.length) mergedText.push(rightColumn[i]);
    }

    return mergedText.join('\n');
}

// Function to format the text for proper structure
function formatText(inputText) {
    return inputText.replace(/(\d+\.|\(\w+\))\s*\n\s*/g, '$1 ')
        .replace(/\n(?=[a-zA-Z])/g, ' ')
        .replace(/([.!?])\s*\n\s*/g, '$1\n\n');
}

// Function to create paragraphs from formatted text
function createParagraphsFromText(text) {
    const lines = text.split('\n');
    return lines.map(line => {
        const trimmedLine = line.trim();

        const isBold = /^[\d]+\./.test(trimmedLine) || trimmedLine.toUpperCase() === trimmedLine;
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

        // Determine if the text is in a two-column format
        if (isTwoColumnFormat(text)) {
            text = processTwoColumnText(text); // Process text for a two-column layout
        }

        text = formatText(text); // Further format the text

        const paragraphs = createParagraphsFromText(text);

        const doc = new Document({
            numbering: {
                config: [
                    {
                        reference: "custom-numbering",
                        levels: [
                            {
                                level: 0,
                                format: LevelFormat.UPPER_ROMAN,
                                text: "%1.",
                                alignment: AlignmentType.START,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                            {
                                level: 1,
                                format: LevelFormat.DECIMAL,
                                text: "%2.",
                                alignment: AlignmentType.START,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(1), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                            {
                                level: 2,
                                format: LevelFormat.LOWER_LETTER,
                                text: "%3)",
                                alignment: AlignmentType.START,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(1.5), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                        ],
                    },
                    {
                        reference: "custom-bullets",
                        levels: [
                            {
                                level: 0,
                                format: LevelFormat.BULLET,
                                text: "•",
                                alignment: AlignmentType.LEFT,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                            {
                                level: 1,
                                format: LevelFormat.BULLET,
                                text: "◦",
                                alignment: AlignmentType.LEFT,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(1), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                            {
                                level: 2,
                                format: LevelFormat.BULLET,
                                text: "▪",
                                alignment: AlignmentType.LEFT,
                                style: {
                                    paragraph: {
                                        indent: { left: convertInchesToTwip(1.5), hanging: convertInchesToTwip(0.25) },
                                    },
                                },
                            },
                        ],
                    },
                ],
            },
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
    console.log(`Server with AWS running on http://localhost:${PORT}`);
});
