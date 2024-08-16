const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { TextractClient, DetectDocumentTextCommand } = require('@aws-sdk/client-textract');
const { Document, Packer, Paragraph, TextRun } = require('docx');

require('dotenv').config();
const PORT = process.env.PORT || 3000;
const app = express();
const upload = multer({ dest: 'uploads/' });

const client = new TextractClient({
    region: process.env.AWS_REGION || 'us-west-2',
    credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
    },
});

app.set('view engine', 'ejs');
app.use(express.static('uploads'));

app.get('/', (req, res) => {
    console.log('GET / - Rendering index page');
    res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
    console.log('POST /upload - Received image upload');
    const imagePath = req.file.path;
    const outputFileName = `extracted_${Date.now()}.docx`;
    const imageBytes = fs.readFileSync(imagePath);

    try {
        console.log('POST /upload - Sending image to Textract for processing');
        const command = new DetectDocumentTextCommand({ Document: { Bytes: imageBytes } });
        const data = await client.send(command);

        console.log('POST /upload - Textract response received');
        const text = data.Blocks.filter(block => block.BlockType === 'LINE')
            .map(block => block.Text)
            .join('\n');

        console.log('POST /upload - Extracted text:', text);

        const lines = text.split('\n').map(line => {
            return new Paragraph({
                children: [
                    new TextRun({
                        text: line,
                        font: 'Arial',
                        size: 24,
                    }),
                ],
            });
        });

        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: lines,
                },
            ],
        });

        const buffer = await Packer.toBuffer(doc);
        const docxPath = path.join('uploads', outputFileName);
        fs.writeFileSync(docxPath, buffer);

        console.log(`POST /upload - DOCX file created: ${outputFileName}`);
        res.render('index', { docxFile: outputFileName });
    } catch (error) {
        console.error('POST /upload - Error processing image:', error);
        res.status(500).send('Error processing image.');
    }
});

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, 'uploads', filename);
    console.log(`GET /download/${filename} - Downloading file`);
    res.download(filePath);
});

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});
