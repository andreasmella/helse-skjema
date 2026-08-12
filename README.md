# helse-skjema — klient-side generator

Dette repoet inneholder en enkel klient-side .docx-generator som kjører i nettleseren. Du kan åpne `index.html` lokalt (dobbeltklikk eller åpne i nettleser) og gjøre alt derfra: laste opp JSON og valgfritt logo-bilde, så lastes en .docx-fil ned.

Hva jeg har lagt til:
- index.html — enkel UI for å velge JSON + logo og generere .docx
- generator.js — logikk som bruker `docx` (dolanmiu/docx) og FileSaver i nettleseren
- styles.css — litt styling

Hvordan bruke:
1. Åpne `index.html` i nettleseren.
2. Velg JSON-filen (skjemaet). Hvis du har en stor fil (f.eks. ~700kb) går det fint — alt skjer i nettleseren.
3. Valgfritt: last opp en PNG/JPG-logo (den settes som header på første side).
4. Valgfritt: skriv inn navn på felter som skal behandles som lange fritekstfelt (komma-separert). Hvis du lar feltet stå tomt brukes heuristiske regler (sjekker ord som "forespørsel", "merknad", "begrunnelse" osv.).
5. Klikk "Generer .docx". Filen lastes ned automatisk.

Design og layout:
- Jeg følger et enkelt hierarki: tittel (hvis finnes i JSON), så pasientinfo (hvis finnes), så feltene fra JSON.
- Store fritekstfelt deles i avsnitt ved newline, korte felter settes som én linje.
- Logo settes i header for første side; sidetall i footer ("Side X av Y").

Hvis du ønsker nøyaktig identisk layout med PDF-en (eksakt skrift, linjehøyder, kolonner, posisjonering osv.) anbefaler jeg å bruke en server-side løsning som genererer PDF direkte (f.eks. Puppeteer/HTML/CSS eller LibreOffice/docx-til-PDF), siden Word/docx har begrensninger i styling og html->docx-konvertering vanligvis kan variere.

Neste steg jeg kan gjøre for deg (hvis du ønsker):
- Finjustere layout for å matche PDFen nøyaktig (må få eksempler og en liste over hvilke JSON-felter som skal være lange).
- Legge til en server-side generator (Node) som også kan eksportere ferdig PDF.


