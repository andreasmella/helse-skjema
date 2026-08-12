/* script.js

Client-side generator that:
- leser et JSON-objekt fra bruker
- valgfritt: leser en logo (PNG/JPG)
- lager en .docx med header (logo), sidetall i footer, og innhold fra JSON

Design: vi bygger et enkelt, men fleksibelt dokument. Tilpass etter behov.
*/

(async function(){
  const jsonInput = document.getElementById('jsonFile');
  const logoInput = document.getElementById('logoFile');
  const generateBtn = document.getElementById('generateBtn');

  // hent docx API fra global (UMD)
  const docx = window.docx;
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Footer, Header, ImageRun } = docx;

  function fileToText(file){
    return new Promise((res, rej)=>{
      const r = new FileReader();
      r.onload = ()=>res(r.result);
      r.onerror = rej;
      r.readAsText(file, 'utf-8');
    });
  }
  function fileToDataUrl(file){
    return new Promise((res, rej)=>{
      const r = new FileReader();
      r.onload = ()=>res(r.result);
      r.onerror = rej;
      r.readAsDataURL(file);
    });
  }
  function dataUrlToArrayBuffer(dataUrl){
    const base64 = dataUrl.split(',')[1];
    const binary = atob(base64);
    const len = binary.length;
    const buffer = new ArrayBuffer(len);
    const view = new Uint8Array(buffer);
    for(let i=0;i<len;i++) view[i]=binary.charCodeAt(i);
    return buffer;
  }

  function addParagraphsFromObject(obj){
    // Forenklet rendering: går gjennom top-level keys og produserer overskrift + verdi
    const paragraphs = [];
    for(const key of Object.keys(obj)){
      const value = obj[key];
      // Hvis verdi er tekst eller tall
      if(typeof value === 'string' || typeof value === 'number'){
        paragraphs.push(new Paragraph({
          children:[ new TextRun({text: key.toString()+':',bold:true}) ]
        }));
        // for lange tekster, behold linjeskift som egne avsnitt
        const lines = String(value).split(/\r?\n/);
        for(const ln of lines){
          paragraphs.push(new Paragraph(ln));
        }
      } else if(Array.isArray(value)){
        paragraphs.push(new Paragraph({children:[ new TextRun({text:key+':',bold:true}) ]}));
        value.forEach(item=>{
          paragraphs.push(new Paragraph('- '+String(item)));
        });
      } else if(typeof value === 'object' && value !== null){
        paragraphs.push(new Paragraph({children:[ new TextRun({text:key+':',bold:true}) ]}));
        // rekursivt inntil ett nivå for lesbarhet
        for(const subKey of Object.keys(value)){
          const v = value[subKey];
          paragraphs.push(new Paragraph({children:[ new TextRun({text: '  '+subKey+': ',bold:true}), new TextRun(String(v))]}));
        }
      }
    }
    return paragraphs;
  }

  async function buildDocx(jsonObject, logoDataUrl){
    // header (logo) – hvis ikke logo, header kan være tom
    let header = undefined;
    if(logoDataUrl){
      const imgBuffer = dataUrlToArrayBuffer(logoDataUrl);
      const image = new ImageRun({data: imgBuffer, transformation:{width:160,height:48}});
      header = new Header({children:[ new Paragraph({children:[image],alignment:AlignmentType.LEFT}) ]});
    }

    // footer med sidetall (sentrert)
    const footer = new Footer({children:[ new Paragraph({children:[ new TextRun({text:'Side '}), new TextRun({children:["{PAGE}"]}), new TextRun({text:' av '}), new TextRun({children:["{NUMPAGES}"]})],alignment:AlignmentType.CENTER}) ]});

    // body: dokumentet kan bestå av flere seksjoner – her en seksjon
    const content = [];

    // Hvis JSON har 'title' eller 'overskrift', bruk som heading
    if(jsonObject.title || jsonObject.overskrift){
      const titleText = jsonObject.title || jsonObject.overskrift || '';
      content.push(new Paragraph({children:[ new TextRun({text:titleText, bold:true}) ], heading: HeadingLevel.HEADING_1}));
    }

    // Hvis JSON har en hoved-samling, for eksempel 'fields' eller 'sections'
    if(jsonObject.sections && Array.isArray(jsonObject.sections)){
      for(const sec of jsonObject.sections){
        if(sec.title) content.push(new Paragraph({children:[ new TextRun({text:sec.title, bold:true}) ], heading: HeadingLevel.HEADING_2}));
        if(sec.body) {
          const lines = String(sec.body).split(/\r?\n/);
          lines.forEach(l=>content.push(new Paragraph(l)));
        }
        if(sec.fields) content.push(...addParagraphsFromObject(sec.fields));
      }
    } else {
      // fallback: skriv ut top-level nøkler
      content.push(...addParagraphsFromObject(jsonObject));
    }

    const doc = new Document({ sections:[{
      headers: header ? {default: header} : {},
      footers: {default: footer},
      children: content
    }]});

    return doc;
  }

  generateBtn.addEventListener('click', async ()=>{
    try{
      if(!jsonInput.files || !jsonInput.files[0]){ alert('Velg en JSON-fil først'); return; }
      const txt = await fileToText(jsonInput.files[0]);
      let data;
      try{ data = JSON.parse(txt); }catch(e){ alert('Feil ved parsing av JSON: '+e.message); return; }
      let logoDataUrl = null;
      if(logoInput.files && logoInput.files[0]){
        logoDataUrl = await fileToDataUrl(logoInput.files[0]);
      }

      // bygg dokumentet
      const doc = await buildDocx(data, logoDataUrl);

      // pakk og last ned
      const blob = await Packer.toBlob(doc);
      const fileName = (data.filename && data.filename.endsWith('.docx')) ? data.filename : (data.title ? data.title.replace(/[^a-z0-9_-]/gi,'_')+'.docx' : 'helse-skjema.docx');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = fileName; document.body.appendChild(a); a.click();
      setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 5000);

    }catch(err){
      console.error(err); alert('Feil under generering: '+err.message);
    }
  });
})();
