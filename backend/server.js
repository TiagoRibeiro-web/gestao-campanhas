require('dotenv').config();
const express = require('express');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3002;

// ============ CORS ============
app.use(cors());
app.use(express.json());

// ============ USUÁRIOS ============
const VALID_USERS = {
  'admin': { password: 'admin123', name: 'Administrador', role: 'admin' },
  'sicoob': { password: 'cocred2026', name: 'Gestor Sicoob', role: 'gestor' },
  'cristini.cordesco@ideatoreamericas.com': { password: 'cocred2026', name: 'Cristini Cordesco', role: 'admin' }
};

// ============ ROTAS ============

// Rota raiz
app.get('/', (req, res) => {
  res.json({
    message: '🚀 SICOOB COCRED - Backend API',
    version: '2.1.0',
    status: 'online',
    timestamp: new Date().toISOString(),
    endpoints: {
      health: '/api/health',
      login: '/api/login',
      usuarios: '/api/usuarios'
    }
  });
});

// Health Check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'online',
    message: 'Backend funcionando!',
    timestamp: new Date().toISOString(),
    users: Object.keys(VALID_USERS).length
  });
});

// Login
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  
  if (!username || !password) {
    return res.status(400).json({ success: false, error: 'Usuário e senha são obrigatórios' });
  }

  const user = VALID_USERS[username];
  if (!user || user.password !== password) {
    return res.status(401).json({ success: false, error: 'Usuário ou senha inválidos' });
  }

  const token = Buffer.from(`${username}:${Date.now()}`).toString('base64');
  
  res.json({
    success: true,
    message: 'Login realizado com sucesso!',
    user: { username, name: user.name, role: user.role },
    token
  });
});

// Listar usuários
app.get('/api/usuarios', (req, res) => {
  const users = Object.keys(VALID_USERS).map(username => ({
    username: username,
    name: VALID_USERS[username].name,
    role: VALID_USERS[username].role
  }));
  
  res.json({ 
    success: true, 
    users: users 
  });
});

// Buscar planilha
app.post('/api/buscar-planilha', (req, res) => {
  res.json({
    success: true,
    message: 'Funcionalidade em desenvolvimento',
    data: [],
    totalRows: 0
  });
});

// Salvar campanhas
app.post('/api/salvar-campanhas', (req, res) => {
  res.json({
    success: true,
    message: 'Funcionalidade em desenvolvimento',
    totalRows: 0
  });
});

// ============ INICIAR ============
app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════════════╗
║  🚀 SICOOB COCRED - BACKEND COMPLETO                                ║
║  📡 Servidor rodando em http://localhost:${PORT}                      ║
║  👥 Usuários: ${Object.keys(VALID_USERS).length}                     ║
║  ✅ Rotas: /, /api/health, /api/login, /api/usuarios                ║
╚═══════════════════════════════════════════════════════════════════════╝
  `);
});