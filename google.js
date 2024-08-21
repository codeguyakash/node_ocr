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
const file = path.join(__dirname, 'codeguyakash-dev-56489cba98da.json');

const client = new vision.ImageAnnotatorClient({
    keyFilename: file,
});

app.set('view engine', 'ejs');

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

        const lines = text.split('\n').map(line => {
            const isBullet = line.trim().startsWith('•') || line.trim().startsWith('-') || line.trim().startsWith('*');
            const isIndented = line.startsWith('    ') || line.startsWith('\t');

            return new Paragraph({
                children: [
                    new TextRun({
                        text: line.trim(),
                        bold: isBullet,
                        size: 24,
                    }),
                ],
                bullet: isBullet ? { level: 0 } : undefined,
                indent: isIndented ? { left: 720 } : undefined,
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

        res.render('index', { docxFile: outputFileName });
    } catch (error) {
        console.error('Error extracting text:', error);
        res.status(500).send('Error processing image.');
    }
});

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, 'uploads', filename);
    res.download(filePath);
});

app.listen(PORT, () => {
    console.log(`Server started with google on http://localhost:${PORT}`);
});
