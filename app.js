// Enkel klient-side PDF-generator som leser JSON + valgfri logo og lager en PDF
// Bruker html2pdf (html2canvas + jsPDF)

const jsonFileInput = document.getElementById('jsonFile');
const logoFileInput = document.getElementById('logoFile');
const generateBtn = document.getElementById('generateBtn');
const printArea = document.getElementById('printArea');
const exampleJsonPre = document.getElementById('exampleJson');
const logoPreview = document.getElementById('logoPreview');

let currentData = null;
let logoDataUrl = null;

const EXAMPLE = {
  "tittel": "Skjema om helseopplysninger",
  "pasient": {
    "navn": "Ola Nordmann",
    "fodselsnummer": "01019012345",
    "adresse": "Eksempelveien 1, 0001 Oslo"
  },
  "henvisning": "Beskrivelse av hvorfor informasjonen etterspørres...",
  "journal": {
    "avdeling": "Klinikk for indremedisin",
    "dato": "2026-08-12",
    "fritekst": "Stor fritekst. Dette skal vises som et stort tekstfelt i PDF'en (flere linjer)."
  },
  "signatur": "Navn på rekvirent / kontaktinfo"
};
exampleJsonPre.textContent = JSON.stringify(EXAMPLE, null, 2);

jsonFileInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(reader.result);
      currentData = parsed;
      exampleJsonPre.textContent = JSON.stringify(parsed, null, 2);
      buildPreview(parsed);
    } catch (err) {
      alert('Feil ved lesing av JSON: ' + err.message);
    }
  };
  reader.readAsText(f, 'utf-8');
});

logoFileInput.addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    logoDataUrl = reader.result;
    logoPreview.innerHTML = `<img src="${logoDataUrl}" alt="logo" style="max-height:60px">`;
    if (currentData) buildPreview(currentData);
  };
  reader.readAsDataURL(f);
});

function buildPreview(data){
  printArea.classList.remove('hidden');
  printArea.innerHTML = '';

  const header = document.createElement('div');
  header.className = 'doc-header';
  const left = document.createElement('div');
  const hTitle = document.createElement('h2');
  hTitle.textContent = data.tittel || 'Skjema om helseopplysninger';
  left.appendChild(hTitle);

  const right = document.createElement('div');
  if (logoDataUrl) {
    const img = document.createElement('img');
    img.src = logoDataUrl;
    img.style.maxHeight = '60px';
    right.appendChild(img);
  }
  header.appendChild(left);
  header.appendChild(right);
  printArea.appendChild(header);

  // Pasient-info
  if (data.pasient){
    const sec = document.createElement('div');
    sec.className = 'section';
    const lbl = document.createElement('div'); lbl.className='field-label'; lbl.textContent='Pasient';
    sec.appendChild(lbl);
    for (const k of Object.keys(data.pasient)){
      const v = document.createElement('div'); v.className='field-value';
      v.innerHTML = `<strong>${k}:</strong> ${escapeHtml(String(data.pasient[k]||''))}`;
      sec.appendChild(v);
    }
    printArea.appendChild(sec);
  }

  // Henvisning / bakgrunn
  if (data.henvisning){
    const sec = document.createElement('div');
    sec.className = 'section';
    const lbl = document.createElement('div'); lbl.className='field-label'; lbl.textContent='Forespørsel / bakgrunn';
    const v = document.createElement('div'); v.className='field-value';
    v.textContent = data.henvisning;
    sec.appendChild(lbl); sec.appendChild(v);
    printArea.appendChild(sec);
  }

  // Journal/fritekst (stor tekstfelt)
  if (data.journal){
    const sec = document.createElement('div');
    sec.className = 'section';
    const lbl = document.createElement('div'); lbl.className='field-label'; lbl.textContent='Journal / fritekst';
    const v = document.createElement('div'); v.className='field-value';
    // Hvis fritekst finnes, vis som stort felt
    if (data.journal.fritekst){
      const p = document.createElement('div');
      p.style.minHeight = '180px'; p.style.whiteSpace='pre-wrap'; p.textContent = data.journal.fritekst;
      v.appendChild(p);
    }
    for (const k of Object.keys(data.journal)){
      if (k === 'fritekst') continue;
      const info = document.createElement('div'); info.style.marginTop='6px';
      info.innerHTML = `<strong>${k}:</strong> ${escapeHtml(String(data.journal[k]||''))}`;
      v.appendChild(info);
    }
    sec.appendChild(lbl); sec.appendChild(v); printArea.appendChild(sec);
  }

  // signatur
  if (data.signatur){
    const sec = document.createElement('div');
    sec.className = 'section';
    const lbl = document.createElement('div'); lbl.className='field-label'; lbl.textContent='Signatur / kontakt';
    const v = document.createElement('div'); v.className='field-value'; v.textContent = data.signatur;
    sec.appendChild(lbl); sec.appendChild(v);
    printArea.appendChild(sec);
  }

  // Footer note
  const foot = document.createElement('div'); foot.style.marginTop='22px'; foot.style.fontSize='12px'; foot.style.color='#666';
  foot.textContent = 'Dokument generert lokalt fra JSON. Sidetall legges i bunn ved eksport til PDF.';
  printArea.appendChild(foot);
}

function escapeHtml(s){ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

// Generer PDF med page numbers
async function generatePdf(){
  if (!currentData){ alert('Du må laste opp en JSON-fil først.'); return; }

  // Ensure preview built
  buildPreview(currentData);

  const opt = {
    margin: 12,
    filename: (currentData.tittel||'skjema') + '.pdf',
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2 },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  // toPdf().get('pdf') gir oss jsPDF-objektet slik at vi kan legge på sidetall
  const worker = html2pdf().from(printArea).set(opt);
  await worker.toPdf().get('pdf').then(function (pdf) {
    const totalPages = pdf.internal.getNumberOfPages();
    const pageWidth = pdf.internal.pageSize.getWidth();
    // legg sidetall nederst midt på hver side
    for (let i = 1; i <= totalPages; i++){
      pdf.setPage(i);
      pdf.setFontSize(10);
      pdf.setTextColor(120);
      pdf.text(`${i} / ${totalPages}`, pageWidth/2, pdf.internal.pageSize.getHeight() - 8, { align: 'center' });
    }
  }).then(function(){
    // Lagre filen (nedlasting i nettleser)
    worker.save();
  }).catch(err => {
    console.error(err); alert('Feil ved generering av PDF: ' + (err && err.message));
  });
}

generateBtn.addEventListener('click', generatePdf);

// Hvis brukeren ikke laster opp, vis eksempel-JSON klar til bruk
buildPreview(EXAMPLE);
currentData = EXAMPLE;

