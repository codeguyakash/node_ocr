const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const { ComputerVisionClient } = require('@azure/cognitiveservices-computervision');
const { CognitiveServicesCredentials } = require('@azure/ms-rest-js');

require('dotenv').config();

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

const credentials = new CognitiveServicesCredentials(process.env.AZURE_COMPUTER_VISION_KEY);
const client = new ComputerVisionClient(credentials, process.env.AZURE_COMPUTER_VISION_ENDPOINT);

app.set('view engine', 'ejs');

app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const outputFileName = `rent_${Date.now()}.docx`;

    try {
        const readStream = fs.createReadStream(imagePath);
        const result = await client.recognizePrintedTextInStream(true, readStream);

        const lines = result.regions.flatMap(region => region.lines).map(line => {
            const text = line.words.map(word => word.text).join(' ');
            return new Paragraph({
                children: [new TextRun({ text, size: 24 })],
            });
        });

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
