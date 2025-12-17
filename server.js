const express = require('express');
const path = require('path');
const app = express();

const port = process.env.PORT || 3000;

// CORS Header erlauben
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

// WICHTIG: Explizit den Content-Type für .tsx Dateien setzen,
// damit Babel im Browser sie korrekt laden kann.
app.use(express.static(__dirname, {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.tsx')) {
      res.set('Content-Type', 'text/plain');
    }
  }
}));

// Fallback zur index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});