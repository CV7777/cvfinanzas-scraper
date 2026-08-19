const express = require('express');
const fs = require('fs');
const path = require('path');

function loadLocalEnv() {
  const envPath = path.join(__dirname, '../.env');

  if (!fs.existsSync(envPath)) {
    return;
  }

  const lines = fs.readFileSync(envPath, 'utf8').split(/\r?\n/);

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#') || !trimmed.includes('=')) {
      return;
    }

    const [key, ...valueParts] = trimmed.split('=');
    if (!process.env[key]) {
      process.env[key] = valueParts.join('=').replace(/^['"]|['"]$/g, '');
    }
  });
}

try {
  require('dotenv').config();
} catch (error) {
  loadLocalEnv();
}

const app = express();
const PORT = process.env.PORT || 3000;
const { requireAuth } = require('./middleware/auth');

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get('/login', (req, res) => {
  res.sendFile(path.join(__dirname, '../public/login.html'));
});

app.get('/login.html', (req, res) => {
  res.redirect('/login');
});

app.get('/dashboard', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '../public/dashboard.html'));
});

app.get('/dashboard.html', requireAuth, (req, res) => {
  res.redirect('/dashboard');
});

app.get('/search', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '../public/search.html'));
});

app.get('/search.html', requireAuth, (req, res) => {
  res.redirect('/search');
});

app.get('/gastos', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '../public/gastos.html'));
});

app.get('/gastos.html', requireAuth, (req, res) => {
  res.redirect('/gastos');
});

app.get('/tipo-cambio', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, '../public/tipo-cambio.html'));
});

app.get('/tipo-cambio.html', requireAuth, (req, res) => {
  res.redirect('/tipo-cambio');
});

app.use(express.static(path.join(__dirname, '../public')));
app.use('/styles', express.static(path.join(__dirname, 'styles')));

// View engine setup
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Routes
const indexRoutes = require('./routes/index');
app.use('/', indexRoutes);

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).send('Error al procesar la solicitud');
});

// Start server
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
  });
}

module.exports = app;
