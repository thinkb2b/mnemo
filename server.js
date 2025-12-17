const express = require('express');
const path = require('path');
const app = express();

// Render setzt automatisch die PORT Umgebungsvariable (meist 10000)
const port = process.env.PORT || 3000;

// CORS Header erlauben (Wichtig für Office Add-ins, falls sie in iframes geladen werden)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// Statische Dateien (index.html, index.tsx, assets) aus dem aktuellen Ordner ausliefern
app.use(express.static(__dirname));

// Fallback: Alle unbekannten Routen zur index.html leiten
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});