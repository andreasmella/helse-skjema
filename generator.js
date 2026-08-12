// generator.js
// Bruker docx (dolanmiu/docx) i nettleser + FileSaver for å lagre .docx

(function(){
  const jsonFile = document.getElementById('jsonFile');
  const logoFile = document.getElementById('logoFile');
  const generateBtn = document.getElementById('generateBtn');
  const longFieldsInput = document.getElementById('longFields');
  const log = document.getElementById('log');

  function writeLog(msg){ log.textContent = msg; }

  function readFileAsText(file){
    return new Promise((resolve,reject)=>{
      const r = new FileReader();
      r.onload = ()=>resolve(r.result);
      r.onerror = ()=>reject(r.error);
      r.readAsText(file);
    });
  }

  function readFileAsDataURL(file){
    return new Promise((resolve,reject)=>{
      const r = new FileReader();
      r.onload = ()=>resolve(r.result);
      r.onerror = ()=>reject(r.error);
      r.readAsDataURL(file);
    });
  }

  function heuristicLongFieldNames(){
    // Vanlige ord som indikerer lange fritekstfelt
    return ['forespørsel','forespørsel om','begrunnelse','merknad','kommentar','tilleggsopplysninger','opplysninger','grunn','begrunnelse','beskrivelse','journal','utredning','vilkår','saksopplysninger'];
  }

  function fieldIsLong(key,explicitList){
    if(!key) return false;
    const k = key.toString().toLowerCase();
    if(explicitList && explicitList.length){
      for(const e of explicitList){ if(e && k.includes(e.trim().toLowerCase())) return true; }
    }
    // heuristisk
    for(const h of heuristicLongFieldNames()) if(k.includes(h)) return true;
    return false;
  }

  async function buildDocxFromJson(jsonObj,logoDataURL,longFieldNames){
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun, Footer, Header } = window.docx;

    const doc = new Document({ sections: [] });

    // Prepare header (first page) if logo
    let headerForFirst = null;
    if(logoDataURL){
      // convert dataURL to base64 (strip prefix)
      const base64 = logoDataURL.split(',')[1];
      headerForFirst = new Header({ children: [
        new Paragraph({ children: [ new ImageRun({ data: base64, transformation: { width: 240, height: 50 } }) ] })
      ]});
    }

    // Footer with page numbers
    const footer = new Footer({ children: [ new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun("Side "), window.docx.PageNumber.CURRENT, new TextRun(" av "), window.docx.PageNumber.TOTAL_PAGES ]) }) });

    // Build content: title and then fields
    const children = [];

    children.push(new Paragraph({ text: jsonObj && jsonObj.title ? jsonObj.title : 'Skjema om helseopplysninger', heading: HeadingLevel.HEADING_1 }));

    // If jsonObj has metadata / patient info, put them first (single-line fields)
    const shortFields = ['navn','fornavn','etternavn','adresse','postnummer','poststed','telefon','epost','fodselsnummer','fødselsdato','fødselsdato','kommunenummer'];

    if(jsonObj){
      // Try to find a top-level patient object or fields
      const patientLike = jsonObj.patient || jsonObj.pasient || jsonObj.person || jsonObj.personalia || jsonObj.personopplysninger || {};
      // add short fields from patientLike
      let anyShort=false;
      for(const sf of shortFields){
        if(patientLike && (patientLike[sf] || patientLike[sf.toLowerCase()])){
          const v = patientLike[sf] || patientLike[sf.toLowerCase()];
          children.push(new Paragraph({ children: [ new TextRun({ text: sf + ': ', bold: true }), new TextRun(String(v)) ] }));
          anyShort = true;
        }
      }
      // fallback: top-level short fields
      for(const sf of shortFields){
        if(jsonObj[sf]){ children.push(new Paragraph({ children: [ new TextRun({ text: sf + ': ', bold: true }), new TextRun(String(jsonObj[sf])) ] })); anyShort=true; }
      }
      if(anyShort) children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

      // Now iterate over keys and output
      function pushField(key,value){
        if(value === null || value === undefined) return;
        if(typeof value === 'object'){
          try{ value = JSON.stringify(value, null, 2); }catch(e){ value = String(value); }
        }
        const isLong = fieldIsLong(key,longFieldNames);
        // Add heading for field
        children.push(new Paragraph({ children: [ new TextRun({ text: prettifyKey(key) + ':', bold: true }) ] }));
        if(isLong){
          // split on newlines into multiple paragraphs
          const parts = String(value).split(/\r?\n/).filter(Boolean);
          if(parts.length===0) parts.push(String(value));
          for(const p of parts){ children.push(new Paragraph({ children: [ new TextRun(p) ] })); }
        } else {
          // single line
          children.push(new Paragraph({ children: [ new TextRun(String(value)) ] }));
        }
        children.push(new Paragraph({ text: '' }));
      }

      // If jsonObj has an ordered fields array, prefer that
      if(Array.isArray(jsonObj.fields) && jsonObj.fields.length){
        for(const f of jsonObj.fields){
          if(f && typeof f === 'object'){
            const key = f.name || f.id || f.key || 'felt';
            const val = f.value !== undefined ? f.value : (f.default || '');
            pushField(key,val);
          }
        }
      } else {
        // iterate top-level keys sorted
        const topKeys = Object.keys(jsonObj).filter(k=>!['title','patient','pasient','person','personalia','fields'].includes(k));
        topKeys.sort();
        for(const k of topKeys){ pushField(k, jsonObj[k]); }
      }
    }

    // Create a single section with headers/footers
    const section = {
      properties: {},
      headers: headerForFirst ? { first: headerForFirst } : {},
      footers: { default: footer },
      children
    };

    doc.addSection(section);

    return Packer.toBlob(doc);
  }

  function prettifyKey(k){
    if(!k) return '';
    // replace _ and camelCase -> separate words
    const spaced = k.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[_\-]+/g,' ');
    return spaced.charAt(0).toUpperCase() + spaced.slice(1);
  }

  generateBtn.addEventListener('click', async ()=>{
    try{
      log.textContent = '';
      if(!jsonFile.files || jsonFile.files.length===0) return writeLog('Velg en JSON-fil først.');
      writeLog('Laster JSON...');
      const txt = await readFileAsText(jsonFile.files[0]);
      let jsonObj;
      try{ jsonObj = JSON.parse(txt); }catch(e){ return writeLog('JSON-feil: ' + e.message); }

      let logoDataURL = null;
      if(logoFile.files && logoFile.files.length>0){
        writeLog('Laster logo...');
        logoDataURL = await readFileAsDataURL(logoFile.files[0]);
      }

      const explicit = (longFieldsInput.value || '').split(',').map(s=>s.trim()).filter(Boolean);
      writeLog('Genererer dokument...');

      const blob = await buildDocxFromJson(jsonObj,logoDataURL,explicit);
      saveAs(blob, (jsonObj && jsonObj.filename) ? jsonObj.filename.replace(/\.[^.]+$/, '') + '.docx' : 'helse-skjema.docx');
      writeLog('Ferdig — fil lastes ned.');
    }catch(err){
      console.error(err);
      writeLog('Feil: ' + (err && err.message ? err.message : String(err)));
    }
  });
})();
