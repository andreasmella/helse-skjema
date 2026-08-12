# helse-skjema — klient-side generator

Dette er et enkelt klient-side verktøy for å generere en .docx-fil fra et JSON-objekt i nettleseren.

Hva som er lagt til:

- index.html — skjema for å laste opp JSON og logo og generere .docx
- script.js — bygger .docx med docx.js og laster ned med FileSaver.js
- style.css — enkel styling
- sample.json — lite eksempel du kan laste ned / bruke

Hvordan bruke lokalt:

1. Last ned repoet eller klon.
2. Åpne `index.html` i nettleseren (dobbelklikk). Moderne nettlesere støtter File API og nedlasting uten server.
3. Last opp JSON-filen din og eventuelt logo (PNG/JPG). Trykk 'Generer .docx'.
4. Filen lastes ned automatisk til din nettlesers nedlastingsmappe (vanligvis "Nedlastinger" / "Downloads").

Kommentarer og videre arbeid:

- For å få nøyaktig layout som PDF-en kan vi tilpasse oppbygning, fonter og marger. Det er også mulig å bygge en mer avansert template (f.eks. med tables og absolute positioned elements) dersom du ønsker.
- Hvis du foretrekker å generere PDF i stedet for .docx kan vi legge til jsPDF-eksport som alternativ.

Sikkerhet:

All generering skjer i klientens nettleser — ingen filer sendes til server.

