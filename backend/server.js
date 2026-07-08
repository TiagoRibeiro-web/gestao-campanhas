require('dotenv').config();
const express = require('express');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3002;

// Middleware
app.use(cors());
app.use(express.json());

// ============ ROTAS ============

// Rota raiz
app.get('/', (req, res) => {
  res.json({
    message: '🚀 SICOOB COCRED - Backend Online!',
    status: 'online',
    version: '1.0.0',
    timestamp: new Date().toISOString()
  });
});

// Health Check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'online',
    message: 'Backend funcionando perfeitamente!',
    timestamp: new Date().toISOString(),
    uptime: process.uptime()
  });
});

// Login (simples)
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  
  if (!username || !password) {
    return res.status(400).json({ 
      success: false, 
      error: 'Usuário e senha são obrigatórios' 
    });
  }

  // Validação simples
  if (username === 'admin' && password === 'admin123') {
    return res.json({
      success: true,
      message: 'Login realizado com sucesso!',
      user: {
        username: 'admin',
        name: 'Administrador',
        role: 'admin'
      },
      token: Buffer.from('admin:123456').toString('base64')
    });
  }

  return res.status(401).json({
    success: false,
    error: 'Usuário ou senha inválidos'
  });
});

// Listar usuários (simples)
app.get('/api/usuarios', (req, res) => {
  res.json({
    success: true,
    users: [
      { username: 'admin', name: 'Administrador', role: 'admin' },
      { username: 'sicoob', name: 'Gestor Sicoob', role: 'gestor' }
    ]
  });
});

// Buscar planilha (mock)
app.post('/api/buscar-planilha', (req, res) => {
  res.json({
    success: true,
    message: 'Funcionalidade em desenvolvimento',
    data: [],
    totalRows: 0
  });
});

// ============ INICIAR SERVIDOR ============

app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════════════╗
║  🚀 SICOOB COCRED - BACKEND (VERSÃO MÍNIMA)                         ║
║  📡 Servidor rodando em http://localhost:${PORT}                      ║
║  🌍 Ambiente: ${process.env.RAILWAY_ENVIRONMENT ? 'RAILWAY' : 'LOCAL'}
║  🔗 URL: https://gestao-campanhas-production.up.railway.app          ║
╚═══════════════════════════════════════════════════════════════════════╝
  `);
  console.log(`\n📋 Endpoints disponíveis:`);
  console.log(`   GET  /                - Rota raiz`);
  console.log(`   GET  /api/health      - Health Check`);
  console.log(`   POST /api/login       - Login (admin/admin123)`);
  console.log(`   GET  /api/usuarios    - Listar usuários`);
  console.log(`   POST /api/buscar-planilha - Buscar planilha (mock)`);
});