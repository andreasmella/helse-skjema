// generator.js - client-side DOCX generator (browser)
// Requires docx UMD (window.docx) and optionally FileSaver.js (saveAs).
(function() {
const config =
    window.HSO_CONFIG || {};
    // DOM refs
    const jsonFile = document.getElementById('jsonFile');
    const logoFile = document.getElementById('logoFile');
    const generateBtn = document.getElementById('generateBtn');
    const logEl = document.getElementById('log');
    const logoPreview = document.getElementById('logoPreview');
    const languageSelect = document.getElementById('languageSelect');
    console.log("VERSION 136");

    function writeLog(msg) {
        if (logEl) logEl.textContent = msg;
        console.log('[helse-skjema] ' + msg);
    }

    function readFileAsText(file) {
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = () => resolve(r.result);
            r.onerror = () => reject(r.error);
            r.readAsText(file);
        });
    }

    function readFileAsDataURL(file) {
        return new Promise((resolve, reject) => {
            const r = new FileReader();
            r.onload = () => resolve(r.result);
            r.onerror = () => reject(r.error);
            r.readAsDataURL(file);
        });
    }

    function dataURLToUint8Array(dataURL) {
        // convert base64 dataURL to Uint8Array
        const idx = dataURL.indexOf(',');
        const base64 = idx >= 0 ? dataURL.slice(idx + 1) : dataURL;
        const binary_string = atob(base64);
        const len = binary_string.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binary_string.charCodeAt(i);
        }
        return bytes;
    }

    function heuristicLongFieldNames() {
        return ['forespørsel', 'forespørsel om', 'begrunnelse', 'merknad', 'kommentar', 'tilleggsopplysninger', 'opplysninger', 'grunn', 'beskrivelse', 'journal', 'utredning', 'vilkår', 'saksopplysninger'];
    }

    function fieldIsLong(key, explicitList) {
        if (!key) return false;
        const k = key.toString().toLowerCase();
        if (explicitList && explicitList.length) {
            for (const e of explicitList) {
                if (e && k.includes(e.trim().toLowerCase())) return true;
            }
        }
        for (const h of heuristicLongFieldNames())
            if (k.includes(h)) return true;
        return false;
    }

    // --- START: FHIR Questionnaire renderer ---
    function renderFHIRQuestionnaire(q, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType) {
        console.log("FHIR renderer kjører");
        console.log(q);
        console.log("Items:", q.item);
const language = q.language || "nb-NO";
const labels = {
    "nb-NO": {
        page: "Side",
        of: "av",
        signature: "Underskrift",
        placeDate: "Sted og dato"
    },

    "nn-NO": {
        page: "Side",
        of: "av",
        signature: "Underskrift",
        placeDate: "Stad og dato"
    },

    "en": {
        page: "Page",
        of: "of",
        signature: "Signature",
        placeDate: "Place and date"
    }
};

const L =
    labels[language] ||
    labels["nb-NO"];
        const paragraphs = [];
        const personFields = [
            "Navn",
            "Fødselsnummer (11 siffer)",
            "Mobiltelefonnummer"
        ];

        function addPersonTable() {

    const fields = [

        "Navn",

        "Fødselsnummer (11 siffer)",

        "Mobiltelefon",

        "Alternativ telefon"

    ];

    for (const field of fields) {

        paragraphs.push(

            new Paragraph({

                spacing: {
                    before: 120,
                    after: 80
                },

                children: [

                    new TextRun({

                        text: field,

                        bold: true

                    })

                ]

            })

        );

        paragraphs.push(

            new Paragraph({

                spacing: {
                    after: 200
                },

                children: [

                    new TextRun(

                        "____________________________________________________________"

                    )

                ]

            })

        );

    }

}


        function isSimpleDataGroup(group) {

            if (!group.item)
                return false;

            const fields = group.item.filter(
                x =>
                x.type === "string" ||
                x.type === "date" ||
                x.type === "integer"
            );

            return (
                fields.length >= 2 &&
                fields.length <= 6
            );

        }

        function renderSimpleGroupTable(group) {

            const fields = group.item.filter(
                x =>
                x.type === "string" ||
                x.type === "date" ||
                x.type === "integer"
            );

            const rows = [];

            for (let i = 0; i < fields.length; i += 2) {

                const left = fields[i];
                const right = fields[i + 1];

                rows.push(

                    new TableRow({

                        children: [

                            new TableCell({
                                children: [
                                    new Paragraph(
                                        left?.text || ""
                                    )
                                ]
                            }),

                            new TableCell({
                                children: [
                                    new Paragraph(
                                        right?.text || ""
                                    )
                                ]
                            })

                        ]

                    })

                );

                rows.push(

                    new TableRow({

                        children: [

                            new TableCell({
                                children: [
                                    new Paragraph(
                                        "________________"
                                    )
                                ]
                            }),

                            new TableCell({
                                children: [
                                    new Paragraph(
                                        "________________"
                                    )
                                ]
                            })

                        ]

                    })

                );

            }

            paragraphs.push(

                new Table({

                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE
                    },

                    rows

                })

            );

        }


        function pushQuestion(text, level = 0, isBold = true) {

            if (!text) return;

            paragraphs.push(
                new Paragraph({

                    spacing: {
                        before: 80,
                        after: 120
                    },
                    children: [
                        new TextRun({

    text: text,

    bold: false,

    size: 24

})
                    ]
                })
            );
        }

        function pushSection(title) {
    paragraphs.push(
        new Paragraph({
            spacing: {
                before: 400,
                after: 200
            },
            children: [
                new TextRun({
                    text: title.toUpperCase(),
                    bold: true,
                    color: "006E8A",
                    size: 34
                })
            ]
        })
    );

    }

        function pushSubSection(title) {

    paragraphs.push(

        new Paragraph({

            spacing: {
                before: 400,
                after: 180
            },

            children: [

                new TextRun({

                    text: title.toUpperCase(),

                    bold: true,

                    color: "4A4A4A",

                    size: 30

                })

            ]

        })

    );

}

        function addShortField() {

            paragraphs.push(
    new Paragraph({

        spacing: {
            after: 300
        },

        children: [

            new TextRun(
                "____________________________________________________________"
            )

        ]

    })
);}

        function addLongField() {

    paragraphs.push(
        new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE
            },
            rows: [
                new TableRow({
                    height: {
                        value: 2500
                    },
                    children: [
                        new TableCell({
                            children: [
                                new Paragraph("")
                            ]
                        })
                    ]
                })
            ]
        })
    );

}

        function renderOptions(answerOptions) {

    if (!Array.isArray(answerOptions))
        return;

    const labels = [];

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

        labels.push(label);

    }

    const numericScale =

        labels.length >= 3 &&
        labels.every(x => !isNaN(Number(x)));

    if (numericScale) {

        const scaleText =
            labels
                .map(x => "☐ " + x)
                .join("      ")

        paragraphs.push(
    new Paragraph({

        spacing: {
            after: 220
        },

        children: [
            new TextRun({
    text: scaleText,
    size: 24
})
        ]

    })
);

        return;

    }

    for (const label of labels) {

    paragraphs.push(

        new Paragraph({

            children: [

                new TextRun({
    text: "☐ " + label,
    size: 24
})

            ]

        })

    );

}

}



function pushText(text) {

    if (text === null || text === undefined)
        return;

    const parts =
        String(text).split(/\r?\n/);

    for (const p of parts) {

        paragraphs.push(

            new Paragraph({

                spacing: {
                    after: 100
                },

                children: [

                    new TextRun({

                        text: String(p),

                        size: 24

                    })

                ]

            })

        );

    }

}

        function getEnableWhenText(it, questionMap) {

            if (!it.enableWhen || !it.enableWhen.length)
                return null;

            const rule = it.enableWhen[0];

            const parentQuestion =
                questionMap[rule.question];

            if (!parentQuestion)
                return null;

            const ignoredQuestions = [
                "Alder",
                "Navn",
                "Høyde",
                "Vekt",
                "Kroppsmasseindeks"
            ];

            if (
                ignoredQuestions.includes(
                    parentQuestion
                )
            ) {
                return null;
            }

            const usefulQuestions = [

                "Har du barn",

                "Er du gravid",

                "Har du en eller flere arbeidsgivere",

                "Ønsker du å registere flere pårørende",

                "Bruker du medisiner",

                "Er du allergisk mot noen medisiner"

            ];

            const shouldShow =
                usefulQuestions.some(
                    q => parentQuestion.includes(q)
                );

            if (!shouldShow)
                return null;

            return (
                "Hvis JA på: " +
                parentQuestion
            );
        }

        function getDependencyLevel(it) {

            if (
                !it.enableWhen ||
                !it.enableWhen.length
            ) {
                return 0;
            }

            return 1;

        }


        const questionMap = {};
        const dependencyMap = {};


        function walkItems(items, level = 0) {
            if (!Array.isArray(items)) return;
            for (const it of items) {
                console.log("ITEM", it.linkId, it.type, it.text);

const text = it.text || "";

if (
    text &&
    text.toLowerCase().includes("symptom")
) {
    console.log("SYMPTOM TREFF:", it);
}
if (
    text &&
    text.toLowerCase().includes(
        "symptomsett kronisk syke"
    )
) {
    continue;
}
                if (it.linkId && text) {

                    questionMap[it.linkId] = text;

                    if (
                        it.enableWhen &&
                        it.enableWhen.length
                    ) {

                        const parentId =
                            it.enableWhen[0].question;

                        if (!dependencyMap[parentId]) {
                            dependencyMap[parentId] = [];
                        }

                        dependencyMap[parentId].push(
                            it.linkId
                        );

                    }

                }

                const conditionalText =
                    getEnableWhenText(
                        it,
                        questionMap
                    );

                if (false && conditionalText) {

                    paragraphs.push(
                        new Paragraph({

                            spacing: {
                                before: 120,
                                after: 60
                            },

                            indent: {
                                left: 240
                            },

                            children: [

                                new TextRun({

                                    text: conditionalText,

                                    bold: true,

                                    italics: true,

                                    color: "7F6000",

                                    size: 20

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
        text.toLowerCase().includes("reshid") ||
        text.toLowerCase().includes("symptomsett")
    )
) {
    continue;
}
                if (it.type === "group") {

                    const hiddenGroups = [
    "MOTTAKER",
    "INNSENDING",
    "OPPSUMMERING AV HELSERELATERTE SPØRSMÅL",
    "PERSONOPPLYSNINGER",
    "OPPSUMMERING"
    
];

                    if (
                        hiddenGroups.includes(
                            text.toUpperCase()
                        )
                    ) {

                        continue;

                    }
                    /*if (
                        isSimpleDataGroup(it)
                    ) {

                        if (level === 0)
                            pushSection(text);
                        else
                            pushSubSection(text);

                        renderSimpleGroupTable(it);

                        continue;
                    }*/

                    if (level === 0) {

                        pushSection(text);

                    } else {

                        pushSubSection(text);

                    }

                    if (it.item)
                        walkItems(it.item, level + 1);

                } else if (it.type === 'display') {

    if (
        text &&
        (
            text.toLowerCase().includes("send inn") ||
            text.toLowerCase().includes("lagret i din journal") ||
            text.toLowerCase().includes("lagret i din journal")
        )
    ) {
console.log("FILTERT BORT", text);
        continue;
    }

    pushText(text);

    if (it.item)
        walkItems(it.item, level + 1);
} else {
                    const hiddenPersonFields = [

                        "Navn",

                        "Fødselsnummer (11 siffer)",

                        "Mobiltelefonnummer",

                        "Eventuelt annet telefonnummer vi kan nå deg på"

                    ];

                    if (
                        hiddenPersonFields.includes(text)
                    ) {
                        continue;
                    }

                    if (
                        it.type === "string" ||
                        it.type === "integer" ||
                        it.type === "date" ||
                        it.type === "quantity"
                    ) {

                        pushQuestion(
                            text,
                            getDependencyLevel(it)
                        );

                        addShortField();

                        continue;
                    }
                    const skipTextGroups = [

                        "HOVEDPÅRØRENDE",

                        "PÅRØRENDE",

                        "BARN I FAMILIEN"

                    ];

                    if (
                        skipTextGroups.includes(
                            text.toUpperCase()
                        )
                    ) {

                        continue;

                    }
                    if (
                        it.type === "text" &&
                        !it.item &&
                        text.length > 0
                    ) {

                        pushQuestion(text);

                        addLongField();

                        continue;
                    }
                    pushQuestion(
                        text,
                        level,
                        true
                    );

                    if (it.answerOption) renderOptions(it.answerOption);
                    if (it.initial && it.initial.length) {
                        for (const init of it.initial) {
                            const v = init.valueString || init.valueBoolean || (init.valueCoding && init.valueCoding.display) || init.valueDate || init.valueInteger;
                            if (v !== undefined) paragraphs.push(new Paragraph({
                                children: [new TextRun(String('  (Default: ' + v + ')'))]
                            }));
                        }
                    }
                    if (it.item) walkItems(it.item, level + 1);
                }
            }
        }

        if (q.title) {
    paragraphs.push(
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: {
                after: 500
            },
            children: [
                new TextRun({
                    text: q.title.toUpperCase(),
                    bold: true,
                    size: 48
                })
            ]
        })
    );
}

        paragraphs.push(
            new Paragraph({
                text: ""
            })
        );

        pushSubSection(
            "Personopplysninger"
        );

        addPersonTable();
	console.log(
    "QUESTIONNAIRE DESCRIPTION:",
    q.description
);

if (
    q.description &&
    !q.description.toLowerCase().includes(
        "symptomsett"
    )
) {
    paragraphs.push(
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: {
                after: 400
            },
            children: [
                new TextRun({
                    text: q.description,
                    italics: true,
                    color: "666666",
                    size: 22
                })
            ]
        })
    );
}
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
        spacing: {
            before: 400,
after: 180
        },
        children: [
            new TextRun({
                text: L.placeDate + ":",
                bold: true,
                size: 24
            })
        ]
    })
);

paragraphs.push(
    new Paragraph({
        spacing: {
            after: 250
        },
        children: [
            new TextRun(
                "________________________________________________________________________________"
            )
        ]
    })
);

paragraphs.push(
    new Paragraph({
        spacing: {
            before: 200,
after: 180
        },
        children: [
            new TextRun({
                text: L.signature + ":",
                bold: true,
                size: 24
            })
        ]
    })
);

paragraphs.push(
    new Paragraph({
        spacing: {
            after: 400
        },
        children: [
            new TextRun(
                "________________________________________________________________________________"
            )
        ]
    })
);

        return paragraphs;
    }
    // --- END: FHIR renderer ---

    async function buildDocxFromJson(jsonObj, logoDataURL, longFieldNames) {
        if (!window.docx) throw new Error('docx library (window.docx) ikke funnet. Sjekk at du har lastet scriptet via CDN.');
        const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    HeadingLevel,
    ImageRun,
    Footer,
    Header,
    AlignmentType,
    Table,
    TableRow,
    TableCell,
    WidthType,
    PageNumber
} = window.docx;

        const doc = new Document({

    styles: {

        default: {

            document: {

                run: {

                    font: "Arial"

                }

            }

        }

    },

    sections: []

});

        // Header with logo if provided (robust handling)
        let headerChildren = [];
        if (logoDataURL) {
            try {
                // convert to binary for ImageRun
                const bytes = dataURLToUint8Array(logoDataURL);
                headerChildren.push(new Paragraph({
                    children: [new ImageRun({
                        data: bytes,
                        transformation: {
                            width: 200,
                            height: 40
                        }
                    })]
                }));
            } catch (e) {
                console.warn('[helse-skjema] Kunne ikke bruke logo, hopper over. Error:', e);
            }
        }

        // create a Header instance if we have header children
        const headerObj = headerChildren.length ? new Header({
            children: headerChildren
        }) : undefined;

        // Footer - simple text
        const footer = new Footer({
    children: [
        new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [
                new TextRun("Side "),
                PageNumber.CURRENT,
                new TextRun(" av "),
                PageNumber.TOTAL_PAGES
            ]
        })
    ]
});

        const children = [];

        // If JSON looks like a FHIR Questionnaire, render nicely
        //if (jsonObj && (jsonObj.resourceType === 'Questionnaire' || jsonObj.resourceType === 'QuestionnaireResponse')) {
        const questionnaire =
            jsonObj?.resourceType === "Questionnaire" ?
            jsonObj :
            jsonObj?.entry?.find(
                e => e.resource?.resourceType === "Questionnaire"
            )?.resource;
if (
    questionnaire &&
    languageSelect &&
    languageSelect.value
) {
    questionnaire.language =
        languageSelect.value;
}


        if (questionnaire) {

            const fhirChildren =
                renderFHIRQuestionnaire(
                    questionnaire,
                    Paragraph,
                    TextRun,
                    HeadingLevel,
                    Table,
                    TableRow,
                    TableCell,
                    WidthType,
                    AlignmentType
                );

            for (const p of fhirChildren) {
                children.push(p);
            }

        } else {
            // default heading
            children.push(new Paragraph({
                text: jsonObj && jsonObj.title ? jsonObj.title : 'Skjema om helseopplysninger',
                heading: HeadingLevel.HEADING_1
            }));

            const shortFields = ['navn', 'fornavn', 'etternavn', 'adresse', 'postnummer', 'poststed', 'telefon', 'epost', 'fodselsnummer', 'fødselsdato', 'kommunenummer'];

            if (jsonObj) {
                const patientLike = jsonObj.patient || jsonObj.pasient || jsonObj.person || jsonObj.personalia || jsonObj.personopplysninger || {};
                let anyShort = false;
                for (const sf of shortFields) {
                    if (patientLike && (patientLike[sf] || patientLike[sf.toLowerCase()])) {
                        const v = patientLike[sf] || patientLike[sf.toLowerCase()];
                        children.push(new Paragraph({
                            children: [new TextRun({
                                text: sf + ': ',
                                bold: true
                            }), new TextRun(String(v))]
                        }));
                        anyShort = true;
                    }
                }
                for (const sf of shortFields) {
                    if (jsonObj[sf]) {
                        children.push(new Paragraph({
                            children: [new TextRun({
                                text: sf + ': ',
                                bold: true
                            }), new TextRun(String(jsonObj[sf]))]
                        }));
                        anyShort = true;
                    }
                }
                if (anyShort) children.push(new Paragraph({
                    text: '',
                    spacing: {
                        after: 200
                    }
                }));

                function pushField(key, value) {
                    if (value === null || value === undefined) return;
                    if (typeof value === 'object') {
                        try {
                            value = JSON.stringify(value, null, 2);
                        } catch (e) {
                            value = String(value);
                        }
                    }
                    const isLong = fieldIsLong(key, longFieldNames);
                    children.push(new Paragraph({
                        children: [new TextRun({
                            text: prettifyKey(key) + ':',
                            bold: true
                        })]
                    }));
                    if (isLong) {
                        const parts = String(value).split(/\r?\n/);
                        if (parts.length === 0) parts.push(String(value));
                        for (const p of parts) {
                            children.push(new Paragraph({
                                children: [new TextRun(String(p))]
                            }));
                        }
                    } else {
                        children.push(new Paragraph({
                            children: [new TextRun(String(value))]
                        }));
                    }
                    children.push(new Paragraph({
                        text: ''
                    }));
                }

                if (Array.isArray(jsonObj.fields) && jsonObj.fields.length) {
                    for (const f of jsonObj.fields) {
                        if (f && typeof f === 'object') {
                            const key = f.name || f.id || f.key || 'felt';
                            const val = f.value !== undefined ? f.value : (f.default || '');
                            pushField(key, val);
                        }
                    }
                } else {
                    const topKeys = Object.keys(jsonObj).filter(k => !['title', 'patient', 'pasient', 'person', 'personalia', 'fields'].includes(k));
                    topKeys.sort();
                    for (const k of topKeys) pushField(k, jsonObj[k]);
                }
            }
        }

        const section = {
            properties: {},
            headers: headerObj ? {
                default: headerObj
            } : {},
            footers: {
                default: footer
            },
            children
        };

        doc.addSection(section);
        return Packer.toBlob(doc);
    }

    function prettifyKey(k) {
        if (!k) return '';
        const spaced = k.replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[_\-]+/g, ' ');
        return spaced.charAt(0).toUpperCase() + spaced.slice(1);
    }

    generateBtn.addEventListener('click', async () => {
        try {
            writeLog('Starter generering...');
            if (!jsonFile.files || jsonFile.files.length === 0) return writeLog('Velg en JSON-fil først.');

            // library checks
            if (!window.docx) {
                writeLog('Feil: docx-bibliotek ikke funnet. Sjekk at script tag for docx er med i index.html.');
                return;
            }

            // read JSON
            writeLog('Laster JSON-fil...');
            const txt = await readFileAsText(jsonFile.files[0]);
            let jsonObj;
            try {
                jsonObj = JSON.parse(txt);
if (
    jsonObj &&
    languageSelect &&
    languageSelect.value
) {
    jsonObj.language =
        languageSelect.value;
}
let selectedLanguage =
    languageSelect?.value || "auto";
            } catch (e) {
                return writeLog('JSON-feil: ' + e.message);
            }

            // optional logo
            let logoDataURL = null;
            if (logoFile.files && logoFile.files.length > 0) {
                writeLog('Laster logo...');
                try {
                    logoDataURL = await readFileAsDataURL(logoFile.files[0]);
                    if (logoPreview) logoPreview.innerHTML = '<img src="' + logoDataURL + '" alt="logo" style="max-height:40px;">';
                } catch (e) {
                    writeLog('Kunne ikke lese logo-fil: ' + (e && e.message));
                    logoDataURL = null;
                }
            }
            writeLog('Bygger DOCX — vent litt...');

            const blob = await buildDocxFromJson(jsonObj, logoDataURL);
            const filename = (jsonObj && jsonObj.filename) ? jsonObj.filename.replace(/\.[^.]+$/, '') + '.docx' : 'helse-skjema.docx';

            // Try FileSaver.saveAs, fallback to anchor download
            if (typeof saveAs === 'function') {
                saveAs(blob, filename);
            } else {
                // fallback
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                setTimeout(() => {
                    URL.revokeObjectURL(url);
                    a.remove();
                }, 5000);
            }
            writeLog('Ferdig — fil forsøkes lastet ned: ' + filename + '. Hvis ingenting skjer, se konsoll (F12).');
        } catch (err) {
            console.error(err);
            writeLog('Feil under generering: ' + (err && err.message ? err.message : String(err)));
        }
    });
})();