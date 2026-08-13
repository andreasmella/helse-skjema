// generator.js - client-side DOCX generator (browser)
// Requires docx UMD (window.docx) and optionally FileSaver.js (saveAs).
(function(){
  // DOM refs
  const jsonFile = document.getElementById('jsonFile');
  const logoFile = document.getElementById('logoFile');
  const generateBtn = document.getElementById('generateBtn');
  const longFieldsInput = document.getElementById('longFields');
  const logEl = document.getElementById('log');
  const logoPreview = document.getElementById('logoPreview');
  console.log("VERSION 101");
  function writeLog(msg){
    if(logEl) logEl.textContent = msg;
    console.log('[helse-skjema] ' + msg);
  }

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

  function dataURLToUint8Array(dataURL){
    // convert base64 dataURL to Uint8Array
    const idx = dataURL.indexOf(',');
    const base64 = idx >= 0 ? dataURL.slice(idx+1) : dataURL;
    const binary_string = atob(base64);
    const len = binary_string.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++){
      bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes;
  }

  function heuristicLongFieldNames(){
    return ['forespørsel','forespørsel om','begrunnelse','merknad','kommentar','tilleggsopplysninger','opplysninger','grunn','beskrivelse','journal','utredning','vilkår','saksopplysninger'];
  }
  function fieldIsLong(key, explicitList){
    if(!key) return false;
    const k = key.toString().toLowerCase();
    if(explicitList && explicitList.length){
      for(const e of explicitList){ if(e && k.includes(e.trim().toLowerCase())) return true; }
    }
    for(const h of heuristicLongFieldNames()) if(k.includes(h)) return true;
    return false;
  }

  // --- START: FHIR Questionnaire renderer ---
  function renderFHIRQuestionnaire(q, Paragraph, TextRun, HeadingLevel) {
    console.log("FHIR renderer kjører");
    console.log(q);
    console.log("Items:", q.item);
    const paragraphs = [];

    function pushQuestion(text, level = 0, isBold = true) {

  if (!text) return;

  paragraphs.push(
    new Paragraph({
      spacing: {
        before: 120,
        after: 80
      },
      children: [
        new TextRun({
          text: text,
          bold: true,
          size: 22
        })
      ]
    })
  );
}
function pushSection(title) {

  paragraphs.push(
    new Paragraph({
      spacing: {
        before: 300,
        after: 150
      },
      children: [
        new TextRun({
          text: title.toUpperCase(),
          bold: true,
          color: "1F4E79",
          size: 28
        })
      ]
    })
  );
}
function addShortField() {

  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun(
          "____________________________________________"
        )
      ]
    })
  );
}
function addLongField() {

  for(let i = 0; i < 5; i++) {

    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun(
            "_______________________________________________________________"
          )
        ]
      })
    );

  }
}

function renderOptions(answerOptions) {

  if (!Array.isArray(answerOptions))
    return;

  for (const opt of answerOptions) {

    let label = null;

    if (opt.valueString)
      label = opt.valueString;

    else if (
      opt.valueCoding &&
      opt.valueCoding.display
    )
      label = opt.valueCoding.display;

    if (!label)
      continue;

    paragraphs.push(
      new Paragraph({
        children:[
          new TextRun(
            "☐ " + label
          )
        ]
      })
    );
  }
}



    function pushText(text) {
      if (text === null || text === undefined) return;
      const parts = String(text).split(/\r?\n/);
      for (const p of parts) {
        paragraphs.push(new Paragraph({ children: [ new TextRun(String(p)) ] }));
      }
    }

    function renderOptions(answerOptions) {
      if (!Array.isArray(answerOptions) || answerOptions.length === 0) return;
      for (const opt of answerOptions) {
        let label = null;
        if (opt.valueString) label = opt.valueString;
        else if (opt.valueCoding && opt.valueCoding.display) label = opt.valueCoding.display;
        else if (opt.valueInteger !== undefined) label = String(opt.valueInteger);
        else if (typeof opt === 'string') label = opt;
        if (label) {
          paragraphs.push(new Paragraph({ children: [ new TextRun({ text: '  • ' + label }) ] }));
        }
      }
    }

function getEnableWhenText(it, questionMap) {

    if (!it.enableWhen || !it.enableWhen.length)
        return null;

    const rule = it.enableWhen[0];

    const parentQuestion =
        questionMap[rule.question];

    if (!parentQuestion)
        return "Besvares kun dersom betingelsen er oppfylt.";

    return (
        "Besvares kun dersom du svarte på spørsmålet: " +
        parentQuestion
    );
}

const questionMap = {};



    function walkItems(items, level = 0) {
      if (!Array.isArray(items)) return;
      for (const it of items) {
        console.log("ITEM",it.linkId,it.type,it.text);
        const text = it.text || "";
if (it.linkId && text) {
    questionMap[it.linkId] = text;
}

const conditionalText =
    getEnableWhenText(
        it,
        questionMap
    );

if (conditionalText) {

    paragraphs.push(
        new Paragraph({
            children: [
                new TextRun({
                    text: conditionalText,
                    italics: true
                })
            ]
        })
    );

}


	if (
  text &&
  text.includes("(Skjult)")
) {
  continue;
}
if (
    text &&
    (
        text.toLowerCase().includes("query") ||
        text.toLowerCase().includes("reshid")
    )
) {
    continue;
}
        if (it.type === 'group') {

  pushSection(text);

  if (it.item)
    walkItems(
      it.item,
      level + 1
    );

} 
else if (it.type === 'display') {
          pushText(text);
          if (it.item) walkItems(it.item, level + 1);
        } else {

if (
    it.type === "string" ||
    it.type === "integer" ||
    it.type === "date" ||
    it.type === "quantity"
) {

    pushQuestion(text);

    addShortField();

    continue;
}
if (
    it.type === "text"
) {

    pushQuestion(text);

    addLongField();

    continue;
}
          pushQuestion(text, level, true);
          if (it.answerOption) renderOptions(it.answerOption);
          if (it.initial && it.initial.length) {
            for (const init of it.initial) {
              const v = init.valueString || init.valueBoolean || (init.valueCoding && init.valueCoding.display) || init.valueDate || init.valueInteger;
              if (v !== undefined) paragraphs.push(new Paragraph({ children: [ new TextRun(String('  (Default: ' + v + ')')) ] }));
            }
          }
          if (it.item) walkItems(it.item, level + 1);
        }
      }
    }

    if (q.title) paragraphs.push(new Paragraph({ text: q.title, heading: HeadingLevel.HEADING_1 }));
    if (q.description) pushText(q.description);
    walkItems(
    q.item || q.items || [],
    0
);

paragraphs.push(
    new Paragraph({
        text: ""
    })
);

paragraphs.push(
    new Paragraph({
        text:
            "Sted og dato:"
    })
);

paragraphs.push(
    new Paragraph({
        text:
            "____________________________________________"
    })
);

paragraphs.push(
    new Paragraph({
        text:
            "Underskrift:"
    })
);

paragraphs.push(
    new Paragraph({
        text:
            "____________________________________________"
    })
);

return paragraphs;
  }
  // --- END: FHIR renderer ---

  async function buildDocxFromJson(jsonObj, logoDataURL, longFieldNames){
    if(!window.docx) throw new Error('docx library (window.docx) ikke funnet. Sjekk at du har lastet scriptet via CDN.');
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, Footer, Header, AlignmentType } = window.docx;

    const doc = new Document({ sections: [] });

    // Header with logo if provided (robust handling)
    let headerChildren = [];
    if(logoDataURL){
      try {
        // convert to binary for ImageRun
        const bytes = dataURLToUint8Array(logoDataURL);
        headerChildren.push(new Paragraph({
          children: [ new ImageRun({ data: bytes, transformation: { width: 200, height: 40 } }) ]
        }));
      } catch (e) {
        console.warn('[helse-skjema] Kunne ikke bruke logo, hopper over. Error:', e);
      }
    }

    // create a Header instance if we have header children
    const headerObj = headerChildren.length ? new Header({ children: headerChildren }) : undefined;

    // Footer - simple text
    const footer = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [ new TextRun({ text: "Helse sør-øst — generert skjema", italics: true }) ]
        })
      ]
    });

    const children = [];

    // If JSON looks like a FHIR Questionnaire, render nicely
    //if (jsonObj && (jsonObj.resourceType === 'Questionnaire' || jsonObj.resourceType === 'QuestionnaireResponse')) {
    const questionnaire =
    jsonObj?.resourceType === "Questionnaire"
        ? jsonObj
        : jsonObj?.entry?.find(
              e => e.resource?.resourceType === "Questionnaire"
          )?.resource;

    if (questionnaire) {

        const fhirChildren =
            renderFHIRQuestionnaire(
                questionnaire,
                Paragraph,
                TextRun,
                HeadingLevel
            );

        for (const p of fhirChildren) {
            children.push(p);
        }

    }
    else {
      // default heading
      children.push(new Paragraph({ text: jsonObj && jsonObj.title ? jsonObj.title : 'Skjema om helseopplysninger', heading: HeadingLevel.HEADING_1 }));

      const shortFields = ['navn','fornavn','etternavn','adresse','postnummer','poststed','telefon','epost','fodselsnummer','fødselsdato','kommunenummer'];

      if(jsonObj){
        const patientLike = jsonObj.patient || jsonObj.pasient || jsonObj.person || jsonObj.personalia || jsonObj.personopplysninger || {};
        let anyShort = false;
        for(const sf of shortFields){
          if(patientLike && (patientLike[sf] || patientLike[sf.toLowerCase()])){
            const v = patientLike[sf] || patientLike[sf.toLowerCase()];
            children.push(new Paragraph({ children: [ new TextRun({ text: sf + ': ', bold: true }), new TextRun(String(v)) ] }));
            anyShort = true;
          }
        }
        for(const sf of shortFields){
          if(jsonObj[sf]){ children.push(new Paragraph({ children: [ new TextRun({ text: sf + ': ', bold: true }), new TextRun(String(jsonObj[sf])) ] })); anyShort=true; }
        }
        if(anyShort) children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

        function pushField(key,value){
          if(value === null || value === undefined) return;
          if(typeof value === 'object'){
            try{ value = JSON.stringify(value, null, 2); }catch(e){ value = String(value); }
          }
          const isLong = fieldIsLong(key,longFieldNames);
          children.push(new Paragraph({ children: [ new TextRun({ text: prettifyKey(key) + ':', bold: true }) ] }));
          if(isLong){
            const parts = String(value).split(/\r?\n/);
            if(parts.length===0) parts.push(String(value));
            for(const p of parts){ children.push(new Paragraph({ children: [ new TextRun(String(p)) ] })); }
          } else {
            children.push(new Paragraph({ children: [ new TextRun(String(value)) ] }));
          }
          children.push(new Paragraph({ text: '' }));
        }

        if(Array.isArray(jsonObj.fields) && jsonObj.fields.length){
          for(const f of jsonObj.fields){
            if(f && typeof f === 'object'){
              const key = f.name || f.id || f.key || 'felt';
              const val = f.value !== undefined ? f.value : (f.default || '');
              pushField(key,val);
            }
          }
        } else {
          const topKeys = Object.keys(jsonObj).filter(k=>!['title','patient','pasient','person','personalia','fields'].includes(k));
          topKeys.sort();
          for(const k of topKeys) pushField(k, jsonObj[k]);
        }
      }
    }

    const section = {
      properties: {},
      headers: headerObj ? { default: headerObj } : {},
      footers: { default: footer },
      children
    };

    doc.addSection(section);
    return Packer.toBlob(doc);
  }

  function prettifyKey(k){
    if(!k) return '';
    const spaced = k.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[_\-]+/g,' ');
    return spaced.charAt(0).toUpperCase() + spaced.slice(1);
  }

  generateBtn.addEventListener('click', async ()=>{
    try{
      writeLog('Starter generering...');
      if(!jsonFile.files || jsonFile.files.length===0) return writeLog('Velg en JSON-fil først.');

      // library checks
      if(!window.docx){
        writeLog('Feil: docx-bibliotek ikke funnet. Sjekk at script tag for docx er med i index.html.');
        return;
      }

      // read JSON
      writeLog('Laster JSON-fil...');
      const txt = await readFileAsText(jsonFile.files[0]);
      let jsonObj;
      try{ jsonObj = JSON.parse(txt); }catch(e){ return writeLog('JSON-feil: ' + e.message); }

      // optional logo
      let logoDataURL = null;
      if(logoFile.files && logoFile.files.length>0){
        writeLog('Laster logo...');
        try {
          logoDataURL = await readFileAsDataURL(logoFile.files[0]);
          if(logoPreview) logoPreview.innerHTML = '<img src="'+logoDataURL+'" alt="logo" style="max-height:40px;">';
        } catch(e) {
          writeLog('Kunne ikke lese logo-fil: ' + (e && e.message));
          logoDataURL = null;
        }
      }

      const explicit = longFieldsInput && longFieldsInput.value ? longFieldsInput.value.split(',').map(s=>s.trim()).filter(Boolean) : [];
      writeLog('Bygger DOCX — vent litt...');

      const blob = await buildDocxFromJson(jsonObj, logoDataURL, explicit);
      const filename = (jsonObj && jsonObj.filename) ? jsonObj.filename.replace(/\.[^.]+$/, '') + '.docx' : 'helse-skjema.docx';

      // Try FileSaver.saveAs, fallback to anchor download
      if(typeof saveAs === 'function'){
        saveAs(blob, filename);
      } else {
        // fallback
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 5000);
      }
      writeLog('Ferdig — fil forsøkes lastet ned: ' + filename + '. Hvis ingenting skjer, se konsoll (F12).');
    }catch(err){
      console.error(err);
      writeLog('Feil under generering: ' + (err && err.message ? err.message : String(err)));
    }
  });
})();
