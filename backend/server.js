require('dotenv').config();
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3002;

app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With'],
  credentials: true
}));
app.options('*', cors());

app.use(express.json({ limit: '50mb' }));

// ============ CARREGAR USUÁRIOS ============

const USERS_PATH = path.join(__dirname, 'users.json');

function loadUsers() {
  try {
    // 1. Tenta carregar do arquivo users.json (local)
    if (fs.existsSync(USERS_PATH)) {
      const data = fs.readFileSync(USERS_PATH, 'utf8');
      const users = JSON.parse(data);
      console.log(`✅ ${Object.keys(users).length} usuários carregados do users.json`);
      return users;
    }
    
    // 2. Tenta carregar da variável de ambiente USERS_JSON (Railway)
    if (process.env.USERS_JSON) {
      const users = JSON.parse(process.env.USERS_JSON);
      console.log(`✅ ${Object.keys(users).length} usuários carregados da variável USERS_JSON`);
      return users;
    }
    
    // 3. Fallback: usuários padrão
    console.log('⚠️ Nenhum usuário encontrado. Usando usuários padrão.');
    return {
      'admin': { password: 'admin123', name: 'Administrador', role: 'admin' },
      'sicoob': { password: 'cocred2026', name: 'Gestor Sicoob', role: 'gestor' },
      'gestor': { password: 'gestor456', name: 'Gerente', role: 'gerente' }
    };
  } catch (error) {
    console.error('❌ Erro ao carregar usuários:', error.message);
    return {
      'admin': { password: 'admin123', name: 'Administrador', role: 'admin' }
    };
  }
}

let VALID_USERS = loadUsers();

// ============ CONFIGURAÇÃO MICROSOFT GRAPH (COM FALLBACK) ============

let graphClient = null;
let credential = null;

// Verifica se as variáveis do SharePoint estão configuradas
const hasSharePointConfig = !!(process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET);

if (hasSharePointConfig) {
  try {
    const { Client } = require('@microsoft/microsoft-graph-client');
    const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
    const { ClientSecretCredential } = require('@azure/identity');

    credential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });

    graphClient = Client.initWithMiddleware({ authProvider });
    console.log('✅ Microsoft Graph configurado com sucesso!');
  } catch (error) {
    console.error('❌ Erro ao configurar Microsoft Graph:', error.message);
    graphClient = null;
  }
} else {
  console.log('⚠️ Variáveis do SharePoint não configuradas. Funcionalidades do SharePoint indisponíveis.');
  console.log('   Configure TENANT_ID, CLIENT_ID e CLIENT_SECRET no Railway.');
}

// IDs do SharePoint (com fallback para undefined)
const FILE_ID = process.env.SHAREPOINT_FILE_ID || '';
const USER_EMAIL = process.env.SHAREPOINT_USERNAME || '';
const SOURCEDOC = process.env.SHAREPOINT_SOURCEDOC || '';
const SHEET_NAME = process.env.SHEET_NAME || 'VALOR POR PROJETO';

// ============ FUNÇÕES AUXILIARES ============

// Função para extrair o ID do arquivo da URL
function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/sourcedoc=\{([A-F0-9-]+)\}/i);
  if (match) return match[1];
  const match2 = url.match(/sourcedoc=%7B([A-F0-9-]+)%7D/i);
  if (match2) return match2[1];
  return null;
}

// Função para obter token de acesso
async function getAccessToken() {
  if (!credential) {
    throw new Error('Microsoft Graph não configurado. Verifique as variáveis de ambiente.');
  }
  try {
    const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
    return tokenResponse.token;
  } catch (error) {
    console.error('❌ Erro ao obter token:', error.message);
    throw new Error('Falha na autenticação com Microsoft Graph');
  }
}

// Função para baixar arquivo usando axios
async function downloadFileWithAxios(driveId, itemId, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const response = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: 'arraybuffer',
    timeout: 30000
  });
  return Buffer.from(response.data);
}

// Função para fazer upload de arquivo usando axios
async function uploadFileWithAxios(driveId, itemId, buffer, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const response = await axios.put(url, buffer, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    timeout: 60000
  });
  return response.data;
}

// ============ FUNÇÕES DO SHAREPOINT (COM VERIFICAÇÃO) ============

async function getExcelFromSharePoint(fileUrl) {
  if (!graphClient) {
    return { success: false, error: 'Microsoft Graph não configurado. Verifique as variáveis de ambiente.' };
  }
  
  try {
    console.log('🔍 Iniciando busca da planilha...');
    
    if (!USER_EMAIL || !FILE_ID) {
      throw new Error('Configurações do SharePoint incompletas.');
    }
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) fileId = extractedId;
    }
    
    console.log('👤 Usuário:', USER_EMAIL);
    console.log('📁 File ID:', fileId);
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    console.log('✅ Drive encontrado');
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo encontrado:', fileItem.name);
    } catch (error) {
      console.error('❌ Arquivo não encontrado pelo ID:', error.message);
      const files = await graphClient.api(`/drives/${drive.id}/root/children`).get();
      const excelFile = files.value.find(f => 
        f.name && f.name.toLowerCase().includes('cocred') && f.name.toLowerCase().includes('.xlsx')
      );
      if (excelFile) {
        fileItem = excelFile;
        console.log('📄 Arquivo encontrado pelo nome:', fileItem.name);
      } else {
        throw new Error('Arquivo não encontrado.');
      }
    }
    
    console.log('📥 Baixando conteúdo...');
    const accessToken = await getAccessToken();
    const buffer = await downloadFileWithAxios(drive.id, fileItem.id, accessToken);
    
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetNames = workbook.SheetNames;
    console.log('📑 Abas encontradas:', sheetNames.join(', '));
    
    let sheetName = sheetNames[0];
    const targetSheet = SHEET_NAME;
    const exactMatch = sheetNames.find(s => s.toUpperCase() === targetSheet.toUpperCase());
    if (exactMatch) {
      sheetName = exactMatch;
    } else {
      const partialMatch = sheetNames.find(s => 
        s.toUpperCase().includes('VALOR') && s.toUpperCase().includes('PROJETO')
      );
      if (partialMatch) sheetName = partialMatch;
    }
    
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    console.log(`✅ ${data.length} registros carregados da aba "${sheetName}"`);
    
    return { 
      success: true, 
      data, 
      sheetName, 
      totalRows: data.length,
      driveId: drive.id,
      fileId: fileItem.id,
      fileName: fileItem.name
    };
  } catch (error) {
    console.error('❌ Erro detalhado:', error.message);
    return { success: false, error: error.message };
  }
}

async function writeExcelToSharePoint(campanhas, fileUrl) {
  if (!graphClient) {
    return { success: false, error: 'Microsoft Graph não configurado.' };
  }
  
  try {
    console.log('💾 Iniciando escrita da planilha...');
    
    if (!USER_EMAIL || !FILE_ID) {
      throw new Error('Configurações do SharePoint incompletas.');
    }
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) fileId = extractedId;
    }
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo encontrado:', fileItem.name);
    } catch (error) {
      throw new Error('Arquivo não encontrado.');
    }
    
    const headers = [
      'Projeto/Campanha',
      'Valor Planejado Mídia/Projeto',
      'Valor Realizado Mídia/Projeto',
      'Saldo Mídia/Projeto',
      'Valor Planejado Prod',
      'Valor Realizado Prod',
      'Saldo Prod',
      'Bolsa',
      'Período'
    ];
    
    const values = [headers];
    campanhas.forEach(camp => {
      values.push([
        camp.nome || '',
        Number(camp.planejadoMidia || 0),
        Number(camp.realizadoMidia || 0),
        Number(camp.saldoMidia || 0),
        Number(camp.planejadoProd || 0),
        Number(camp.realizadoProd || 0),
        Number(camp.saldoProd || 0),
        camp.bolsa || 'Avulsa',
        camp.periodo || 'Q1'
      ]);
    });
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(values);
    XLSX.utils.book_append_sheet(wb, ws, 'VALOR POR PROJETO');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    let tentativas = 0;
    let maxTentativas = 3;
    let sucesso = false;
    let ultimoErro = null;
    
    while (!sucesso && tentativas < maxTentativas) {
      tentativas++;
      try {
        console.log(`📤 Tentativa ${tentativas}/${maxTentativas}...`);
        const accessToken = await getAccessToken();
        await uploadFileWithAxios(drive.id, fileItem.id, buffer, accessToken);
        sucesso = true;
        console.log(`✅ ${campanhas.length} registros salvos (tentativa ${tentativas})`);
      } catch (error) {
        ultimoErro = error;
        console.log(`❌ Erro na tentativa ${tentativas}: ${error.message}`);
        if (tentativas < maxTentativas) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
    }
    
    if (!sucesso) {
      throw new Error(`Falha após ${maxTentativas} tentativas.`);
    }
    
    return { success: true, totalRows: campanhas.length };
  } catch (error) {
    console.error('❌ Erro detalhado:', error.message);
    return { success: false, error: error.message };
  }
}

// ============ ENDPOINTS ============

// Rota raiz
app.get('/', (req, res) => {
  res.json({
    message: '🚀 SICOOB COCRED - Backend API',
    version: '2.1.0',
    status: 'online',
    sharepoint: hasSharePointConfig ? '✅ Configurado' : '⚠️ Não configurado',
    endpoints: {
      health: '/api/health',
      login: '/api/login',
      usuarios: '/api/usuarios',
      buscar: '/api/buscar-planilha',
      salvar: '/api/salvar-campanhas'
    }
  });
});

// Health Check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'online',
    message: 'Backend funcionando!',
    version: '2.1.0',
    timestamp: new Date().toISOString(),
    environment: process.env.RAILWAY_ENVIRONMENT ? 'railway' : 'local',
    sharepoint: {
      configured: hasSharePointConfig,
      user: USER_EMAIL || '❌',
      fileId: FILE_ID ? '✅' : '❌'
    },
    users: Object.keys(VALID_USERS).length
  });
});

// ============ ENDPOINTS DE AUTENTICAÇÃO ============

app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    console.log(`🔐 Tentativa de login: ${username}`);

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        error: 'Usuário e senha são obrigatórios'
      });
    }

    const user = VALID_USERS[username];
    
    if (!user || user.password !== password) {
      console.log(`❌ Falha no login: ${username}`);
      return res.status(401).json({
        success: false,
        error: 'Usuário ou senha inválidos'
      });
    }

    console.log(`✅ Login bem-sucedido: ${username}`);
    const token = Buffer.from(`${username}:${Date.now()}`).toString('base64');
    
    res.json({
      success: true,
      message: 'Login realizado com sucesso!',
      user: {
        username: username,
        name: user.name,
        role: user.role
      },
      token: token
    });
  } catch (error) {
    console.error('❌ Erro no login:', error);
    res.status(500).json({
      success: false,
      error: 'Erro interno do servidor'
    });
  }
});

app.post('/api/verificar-auth', async (req, res) => {
  try {
    const { token } = req.body;
    if (!token) {
      return res.status(401).json({ success: false, error: 'Token não fornecido' });
    }
    try {
      const decoded = Buffer.from(token, 'base64').toString('utf-8');
      const [username] = decoded.split(':');
      if (VALID_USERS[username]) {
        return res.json({
          success: true,
          user: {
            username: username,
            name: VALID_USERS[username].name,
            role: VALID_USERS[username].role
          }
        });
      }
    } catch (e) {}
    return res.status(401).json({ success: false, error: 'Token inválido' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ============ ENDPOINTS DE USUÁRIOS ============

app.get('/api/usuarios', (req, res) => {
  try {
    const users = Object.keys(VALID_USERS).map(username => ({
      username: username,
      name: VALID_USERS[username].name,
      role: VALID_USERS[username].role,
      hasPassword: !!VALID_USERS[username].password
    }));
    res.json({ success: true, users });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/usuarios', (req, res) => {
  try {
    const { username, password, name, role } = req.body;
    if (!username || !password || !name) {
      return res.status(400).json({ success: false, error: 'Campos obrigatórios' });
    }
    if (VALID_USERS[username]) {
      return res.status(409).json({ success: false, error: 'Usuário já existe' });
    }
    VALID_USERS[username] = { password, name, role: role || 'usuario' };
    res.json({ success: true, message: `Usuário ${username} criado!` });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ============ ENDPOINTS DO SHAREPOINT ============

app.post('/api/buscar-planilha', async (req, res) => {
  console.log('📥 Requisição em /api/buscar-planilha');
  
  if (!hasSharePointConfig) {
    return res.status(400).json({
      success: false,
      error: 'SharePoint não configurado. Configure as variáveis de ambiente no Railway.'
    });
  }
  
  const result = await getExcelFromSharePoint(req.body.fileUrl);
  
  if (result.success) {
    const mappedData = result.data.map((row, idx) => ({
      id: `sp_${Date.now()}_${idx}`,
      nome: row["Projeto/Campanha"] || row["Projeto"] || `Item ${idx + 1}`,
      planejadoMidia: Number(row["Valor Planejado Mídia/Projeto"] || 0),
      realizadoMidia: Number(row["Valor Realizado Mídia/Projeto"] || 0),
      saldoMidia: Number(row["Saldo Mídia/Projeto"] || 0),
      planejadoProd: Number(row["Valor Planejado Prod"] || 0),
      realizadoProd: Number(row["Valor Realizado Prod"] || 0),
      saldoProd: Number(row["Saldo Prod"] || 0),
      bolsa: row["Bolsa"] || "Avulsa",
      periodo: row["Período"] || "Q1"
    }));
    res.json({ success: true, data: mappedData, sheetName: result.sheetName, totalRows: mappedData.length });
  } else {
    res.status(500).json({ success: false, error: result.error });
  }
});

app.post('/api/salvar-campanhas', async (req, res) => {
  console.log('💾 Requisição em /api/salvar-campanhas');
  
  if (!hasSharePointConfig) {
    return res.status(400).json({
      success: false,
      error: 'SharePoint não configurado. Configure as variáveis de ambiente no Railway.'
    });
  }
  
  const { campanhas, fileUrl } = req.body;
  if (!campanhas || !Array.isArray(campanhas)) {
    return res.status(400).json({ success: false, error: 'Dados inválidos.' });
  }
  
  const result = await writeExcelToSharePoint(campanhas, fileUrl);
  if (result.success) {
    res.json({ success: true, message: `${result.totalRows} campanhas salvas!`, totalRows: result.totalRows });
  } else {
    res.status(500).json({ success: false, error: result.error });
  }
});

// ============ INICIAR SERVIDOR ============

app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════════════╗
║  🚀 SICOOB COCRED - BACKEND v2.1.0                                  ║
║  📡 Servidor rodando em http://localhost:${PORT}                      ║
║  🌍 Ambiente: ${process.env.RAILWAY_ENVIRONMENT ? 'RAILWAY' : 'LOCAL'}
║  📊 SharePoint: ${hasSharePointConfig ? '✅ CONFIGURADO' : '⚠️ NÃO CONFIGURADO'}
║  👥 Usuários: ${Object.keys(VALID_USERS).length}                     ║
╚═══════════════════════════════════════════════════════════════════════╝
  `);
  
  if (!hasSharePointConfig) {
    console.log(`
⚠️  ATENÇÃO: SharePoint não configurado!
   Para usar o SharePoint, configure no Railway:
   - TENANT_ID
   - CLIENT_ID
   - CLIENT_SECRET
   - SHAREPOINT_USERNAME
   - SHAREPOINT_FILE_ID
    `);
  }
});