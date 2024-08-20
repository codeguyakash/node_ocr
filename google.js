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
const file = path.join(__dirname, 'credentials.json');
console.log(file)

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
        // Perform text detection using Google Cloud Vision
        const [result] = await client.textDetection(imagePath);
        const detections = result.textAnnotations;
        const text = detections[0].description;

        // Prepare the extracted text for the DOCX file
        const lines = text.split('\n').map(line => {
            return new Paragraph({
                children: [new TextRun({ text: line, size: 24 })],
            });
        });

        // Create and save the DOCX file
        const doc = new Document({ sections: [{ properties: {}, children: lines }] });
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
    console.log(`Server started on http://localhost:${PORT}`);
});
