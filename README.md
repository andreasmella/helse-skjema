# helse-skjema — klient-side generator

Dette repoet inneholder en enkel, klient-side generator som tar et JSON-objekt og genererer en .docx-fil i nettleseren.

Hvordan bruke

1. Last ned repoet eller klon: `git clone https://github.com/andreasmella/helse-skjema.git`
2. Åpne `index.html` i nettleseren (dobbelklikk eller høyreklikk -> Åpne i nettleser).
3. Last opp JSON-filen (skjema-data) og eventuelt logo (PNG/JPG).
4. Klikk "Generer .docx" — dokumentet lastes ned i nettleserens nedlastingsmappe.

Merk:
- All prosessering skjer i klienten (nettleseren); ingenting lastes opp til en server.
- Layouten i dette eksempelet er grunnleggende. Jeg kan videreutvikle malen for å matche den vedlagte PDF-designen mer nøyaktig (kolonner, tabeller, spacing, skriftstørrelser osv.).

Neste steg jeg trenger fra deg hvis du ønsker videre utvikling:

- Et eksempel-JSON (du nevnte at den er 725 kB) — du kan laste den inn lokalt i nettleseren når du tester, eller jeg kan legge til et eksempel i repoet hvis du ønsker.
- Eksakt krav til hvilke felter som skal presenteres, rekkefølge, og hvordan lange fritekst-felt skal se ut (font, størrelse, begrenset høyde med linjer?).
- Eventuelle grafiske elementer (logo: du kan laste opp `logo-helse-sorost.png` i nettleser UI eller jeg kan committe en logo-fil hvis du vil at den skal ligge i repoet).

Jeg har lagt til: `index.html`, `script.js`, `README.md`.

Hvis du vil at jeg skal committe flere endringer (for eksempel en mer komplett mal som følger PDF-en nærmere), svar "Ja, videre med X" og beskriv ønsket (tabell-layout, sidemarg, font osv.).
