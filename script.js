// script.js - klient-side .docx generator
// Avhengigheter: docx (via CDN) og FileSaver.js

const jsonFileInput = document.getElementById("jsonFile");
const logoFileInput = document.getElementById("logoFile");
const generateBtn = document.getElementById("generateBtn");
const statusSpan = document.getElementById("status");
const jsonPreview = document.getElementById("jsonPreview");

let loadedJson = null;
let logoDataUrl = null;

jsonFileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  statusSpan.textContent = "Laster JSON...";
  try {
    const text = await f.text();
    loadedJson = JSON.parse(text);
    jsonPreview.textContent = JSON.stringify(loadedJson, null, 2);
    statusSpan.textContent = "JSON lastet.";
  } catch (err) {
    statusSpan.textContent = "Feil ved lesing/parsing av JSON";
    console.error(err);
  }
});

logoFileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) {
    logoDataUrl = null; return;
  }
  const reader = new FileReader();
  reader.onload = () => { logoDataUrl = reader.result; statusSpan.textContent = "Logo lastet."; };
  reader.readAsDataURL(f);
});

function dataUrlToUint8Array(dataUrl){
  const base64 = dataUrl.split(',')[1];
  const binary = atob(base64);
  const len = binary.length;
  const u8 = new Uint8Array(len);
  for (let i=0;i<len;i++) u8[i] = binary.charCodeAt(i);
  return u8;
}

// Detect "long text" fields by name or length
function isLongField(key, value){
  if (typeof value !== 'string') return false;
  const keyLower = key.toLowerCase();
  if (keyLower.includes('fritekst') || keyLower.includes('begrunn') || keyLower.includes('beskrivelse') || keyLower.includes('innhold') || keyLower.includes('begrunnelse')) return true;
  if (value.length > 300) return true;
  return false;
}

async function buildDocument(json){
  const Docx = window.docx;
  const { Document, Packer, Paragraph, TextRun, Header, Footer, ImageRun, Table, TableRow, TableCell, WidthType, PageNumber, NumberOfTotalPages } = Docx;

  const doc = new Document({
    sections: []
  });

  // header with logo (if provided)
  let header = null;
  if (logoDataUrl){
    const arr = dataUrlToUint8Array(logoDataUrl);
    // approximate size - we'll set width 120
    const img = new ImageRun({ data: arr, transformation: { width: 120, height: 40 } });
    header = new Header({ children: [ new Paragraph({ children: [ img ] }) ] });
  } else {
    header = new Header({ children: [ new Paragraph('Helse-skjema') ] });
  }

  // footer with page number
  const footer = new Footer({ children: [ new Paragraph({ children: [ new TextRun('Side '), new PageNumber(), new TextRun(' av '), new NumberOfTotalPages() ], alignment: Docx.AlignmentType.CENTER }) ] });

  // build content paragraphs from JSON
  const children = [];
  // top title
  children.push(new Paragraph({ children: [ new TextRun({ text: json.title || 'Helseopplysninger', bold: true, size: 32 }) ], spacing: { after: 200 } }));

  // If JSON contains structured sections, try to respect that
  function addKeyValue(key, value){
    if (isLongField(key, value)){
      // heading + a large paragraph (preserve newlines)
      children.push(new Paragraph({ children: [ new TextRun({ text: key+':', bold:true }) ], spacing: { before: 200 } }));
      const lines = String(value||'').split(/\r?\n/);
      for (const ln of lines){
        children.push(new Paragraph(ln));
      }
      // add empty line
      children.push(new Paragraph(''));
    } else {
      const label = new TextRun({ text: key+': ', bold:true });
      const content = new TextRun(String(value||''));
      children.push(new Paragraph({ children: [ label, content ] }));
    }
  }

  // If JSON is array of fields
  if (Array.isArray(json.fields)){
    for (const f of json.fields){
      const k = f.label || f.key || 'felt';
      const v = f.value || f.default || '';
      addKeyValue(k, v);
    }
  } else {
    // iterate keys
    for (const [k,v] of Object.entries(json)){
      if (k === 'title' || k === 'fields') continue;
      addKeyValue(k, v);
    }
  }

  doc.addSection({ headers: { default: header }, footers: { default: footer }, children });
  return { doc, Packer };
}

generateBtn.addEventListener('click', async () => {
  if (!loadedJson){ statusSpan.textContent = 'Ingen JSON lastet.'; return; }
  statusSpan.textContent = 'Genererer...';
  try{
    const { doc, Packer } = await buildDocument(loadedJson);
    const blob = await Packer.toBlob(doc);
    const filename = (loadedJson.filename || 'helse-skjema') + '.docx';
    saveAs(blob, filename);
    statusSpan.textContent = 'Ferdig — fil lastes ned (se nedlastingsmappe).';
  }catch(err){
    console.error(err);
    statusSpan.textContent = 'Feil ved generering: se konsoll.';
  }
});

// Quick demo: hvis ingen JSON lastet opp, bruk sample
// (fjerner denne linjen hvis du ikke ønsker demo-autofill)
fetch('sample.json').then(r=>r.json()).then(j=>{ if(!loadedJson){ loadedJson = j; jsonPreview.textContent = JSON.stringify(j,null,2);} }).catch(()=>{});
