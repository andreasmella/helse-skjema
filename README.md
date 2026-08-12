# Node-based generator for helse-skjema DOCX

Dette repoet inneholder en enkel web-app (Node + Express) som lar deg laste opp en JSON-fil og en valgfri logo, og så genererer en .docx-fil basert på JSON-innholdet.

Hurtigstart (lokalt):

1. Klon repoet:
   git clone https://github.com/andreasmella/helse-skjema.git
   cd helse-skjema

2. Installer avhengigheter:
   npm install

3. Start serveren:
   npm start

4. Åpne nettleseren på:
   http://localhost:3000

Bruk:
- Velg JSON-filen din i skjemaet og trykk "Generer DOCX". Hvis du vil kan du laste opp en logo som da blir satt inn øverst i dokumentet.
- Dokumentet genereres og lastes ned i nettleseren.

Teknisk:
- Serveren bruker multer (memory storage) for opplasting.
- docx (npm) brukes for å generere .docx.
- JSON-objektet konverteres til en rekke avsnitt: nøkler/verdier. Dette er en enkel baseline — dersom du trenger formatering (tabeller, egne oppsett for bestemte felter), kan jeg legge til tilpassede maler.

Neste steg jeg kan gjøre for deg om du ønsker:
- Lage en mer avansert template (med tabeller og bestemte felt på bestemte plasser) som matcher PDF-designet ditt.
- Forhåndvise DOCX i nettleser som PDF (krever ekstra verktøy/konvertering).
- Legge til Dockerfile og GitHub Actions for å kjøre appen i en container eller på GitHub Pages/Heroku-like tjenester.

Viktig: Om du har en PDF-design (du har lastet opp den), si fra hvilke konkrete felter som skal plasseres hvor, så kan jeg oppdatere generatoren til å fylle et riktig layoutet dokument (tabeller/overskrifter osv.).
