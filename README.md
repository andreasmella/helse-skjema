# helse-skjema

Dette er en enkel klient-side generator for skjema om helseopplysninger. Målet er at du kan laste ned repoet, åpne `index.html` i nettleseren (lokalt) og laste opp din JSON-fil + logo, og generere en PDF som lastes ned i nettleseren.

Hva jeg la til:

- index.html: enkel brukerflate for å laste opp JSON og logo, og knapp for å generere PDF.
- style.css: grunnleggende styling.
- app.js: funksjonaliteten — leser JSON, bygger forhåndsvisning og genererer PDF (html2pdf/jsPDF). Sidetall legges i bunn.
- README.md: denne filen.

Hvordan bruke lokalt:

1. Last ned repo (git clone eller last det ned zip).
2. Åpne `index.html` i en moderne nettleser (Chrome/Edge/Firefox). Ingen server kreves (alt kjører klient-side).
3. Klikk "Velg JSON" og velg din JSON-fil (den du nevnte på ~725kb). Filen må være gyldig JSON.
4. (Valgfritt) Last opp logo-bildet.
5. Klikk "Generer PDF". PDF-en lastes ned til nettleserens nedlastingsmappe.

Merknader og videre utvikling:

- Løsningen genererer PDF. Hvis du trenger DOCX i tillegg kan vi legge inn en klient-side DOCX-generator (for eksempel ved å pakke inn docx-templater eller bruke server-side generering). DOCX-klientløninger kan kreve tunge bundler eller server-komponent.
- Designet er ment å være nært original PDF som du la ved — vi kan viderejustere layout (marg, fonter, sidetopp/bunn, farger) når du bekrefter nøyaktig reglene for hvordan felter skal plasseres.
- Hvis du ønsker at jeg skal legge inn ferdig Helse Sør-Øst-logo i repoet, kan du laste opp logo-filen her eller gi meg beskjed så legger jeg den inn (jeg har brukt en plassholder som kunne lastes opp i nettleseren).

Hvis dette ser greit ut, så kan jeg:
- legge til flere templates (f.eks. side 2: "Forespørsel om endring i pasientjournal" med stort fritekstfelt formatert som i eksempelfilen), eller
- implementere DOCX-eksport (tar litt mer arbeid/avklaringer).

Gi beskjed hva du vil at jeg prioriterer videre (DOCX, nøyaktig layout-tilpasning, ferdig logo sjonglering osv.).
