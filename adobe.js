
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const { Document, Packer, Paragraph, TextRun } = require('docx');


const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');

app.get('/', (req, res) => {
    res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
    const imagePath = req.file.path;
    const outputFileName = `rent_${Date.now()}.docx`;


    const text = 'Extracted text from Adobe PDF Services';
    const lines = text.split('\n').map(line => {
        return new Paragraph({
            children: [new TextRun({ text: line, size: 24 })],
        });
    });

    const doc = new Document({ sections: [{ properties: {}, children: lines }] });
    const buffer = await Packer.toBuffer(doc);
    const docxPath = path.join('uploads', outputFileName);
    fs.writeFileSync(docxPath, buffer);

    res.render('index', { docxFile: outputFileName });
});

app.get('/download/:filename', (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, 'uploads', filename);
    res.download(filePath);
});

app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});
