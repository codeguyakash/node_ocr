const express = require('express');
const multer = require('multer');
const path = require('path');
const Tesseract = require('tesseract.js');
const sharp = require('sharp');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun } = require('docx');

require('dotenv').config()

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = process.env.PORT || 3000;

const targetFont = 'Roboto';

app.set('view engine', 'ejs');

app.use(express.static('uploads'));
app.get('/', (req, res) => {
  res.render('index', { docxFile: null });
});

app.post('/upload', upload.single('image'), async (req, res) => {
  const imagePath = req.file.path;
  const outputFileName = `rent_${Date.now()}.docx`;

  const processedImagePath = `${imagePath}_processed.png`;
  await sharp(imagePath)
    .grayscale()
    .normalize()
    .toFile(processedImagePath);

  Tesseract.recognize(
    processedImagePath,
    'eng',
    {
      tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789,.;!?-()[]{}\'\"@#$%^&*<>_=+\\|/`~ ',
      preserve_interword_spaces: 1,
      logger: m => console.log(m),
    }
  ).then(({ data: { text } }) => {
    const lines = text.split('\n').map(line => {
      return new Paragraph({
        children: [
          new TextRun({
            text: line,
            font: targetFont,
            size: 24,
          })
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

    Packer.toBuffer(doc).then((buffer) => {
      const docxPath = path.join('uploads', outputFileName);
      fs.writeFileSync(docxPath, buffer);
      console.log(`Text extraction complete. Check ${outputFileName} for the result.`);
      res.render('index', { docxFile: outputFileName });
    });
  }).catch(error => {
    console.error('Error extracting text:', error);
    res.status(500).send('Error processing image.');
  });
});

app.get('/download/:filename', (req, res) => {
  const filename = req.params.filename;
  const filePath = path.join(__dirname, 'uploads', filename);
  res.download(filePath);
});

app.listen(PORT, () => {
  console.log(`Server started on http://localhost:${PORT}`);
});
