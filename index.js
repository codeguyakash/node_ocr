const Tesseract = require('tesseract.js');
const sharp = require('sharp');
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun } = require('docx');

const imagePath = './image.png';
const outputFileName = 'rent.docx';
const targetFont = 'Roboto';

async function extractText() {

  const processedImagePath = 'image_sample.png';
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
      fs.writeFileSync(outputFileName, buffer);
      console.log(`Text Extraction complete. Check ${outputFileName} for the result.`);
    });
  }).catch(error => {
    console.error('Error extracting text:', error);
  });
}

extractText();