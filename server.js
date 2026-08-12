const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun } = require('docx');

const app = express();
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Use memory storage so files don't need to be written to disk
const storage = multer.memoryStorage();
const upload = multer({ storage });

function addJsonAsParagraphs(children, obj, prefix = '') {
  if (obj === null) {
    children.push(new Paragraph(`${prefix}: null`));
    return;
  }
  if (typeof obj === 'object' && !Array.isArray(obj)) {
    for (const [k, v] of Object.entries(obj)) {
      addJsonAsParagraphs(children, v, prefix ? `${prefix}.${k}` : k);
    }
    return;
  }
  if (Array.isArray(obj)) {
    obj.forEach((item, idx) => {
      addJsonAsParagraphs(children, item, `${prefix}[${idx}]`);
    });
    return;
  }
  // primitive
  children.push(new Paragraph({
    children: [new TextRun({ text: `${prefix}: `, bold: true }), new TextRun(String(obj))],
  }));
}

app.post('/generate', upload.fields([{ name: 'json' }, { name: 'logo' }]), async (req, res) => {
  try {
    const jsonFile = req.files && req.files['json'] && req.files['json'][0];
    if (!jsonFile) return res.status(400).send('Missing JSON file');

    const jsonText = jsonFile.buffer.toString('utf8');
    let data = {};
    try {
      data = JSON.parse(jsonText);
    } catch (e) {
      return res.status(400).send('Invalid JSON uploaded');
    }

    const doc = new Document();
    const children = [];

    // If a logo was uploaded, add it as the first element
    const logoFile = req.files && req.files['logo'] && req.files['logo'][0];
    if (logoFile) {
      const img = new ImageRun({
        data: logoFile.buffer,
        transformation: { width: 300, height: Math.round(300 * 0.25) },
      });
      children.push(new Paragraph({children: [img]}));
    }

    // Add title if present in the JSON, otherwise generic
    const titleText = (data && data.title) ? String(data.title) : 'Helseopplysninger';
    children.push(new Paragraph({ text: titleText, heading: HeadingLevel.HEADING1 }));

    // Add a short meta paragraph if available
    if (data && data.meta && typeof data.meta === 'string') {
      children.push(new Paragraph(data.meta));
    }

    // Convert the JSON body to paragraphs
    addJsonAsParagraphs(children, data);

    doc.addSection({ children });

    const buffer = await Packer.toBuffer(doc);

    const filename = (data && data.filename) ? `${data.filename}.docx` : 'helseopplysninger.docx';
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).send('Server error');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
