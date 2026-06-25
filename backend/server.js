require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { ClientSecretCredential } = require('@azure/identity');
const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3002;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// ============ CONFIGURAÇÃO MICROSOFT GRAPH ============

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const graphClient = Client.initWithMiddleware({ authProvider });

// IDs fixos do arquivo
const FILE_ID = process.env.SHAREPOINT_FILE_ID || 'CAAE07F7-6116-48A3-94E1-C75827ADD126';
const USER_EMAIL = process.env.SHAREPOINT_USERNAME || 'cristini.cordesco@ideatoreamericas.com';
const SOURCEDOC = process.env.SHAREPOINT_SOURCEDOC || 'CAAE07F7-6116-48A3-94E1-C75827ADD126';
const SHEET_NAME = process.env.SHEET_NAME || 'VALOR POR PROJETO';

// ============ FUNÇÕES AUXILIARES ============

function extractFileIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/sourcedoc=\{([A-F0-9-]+)\}/i);
  if (match) return match[1];
  const match2 = url.match(/sourcedoc=%7B([A-F0-9-]+)%7D/i);
  if (match2) return match2[1];
  return null;
}

async function getAccessToken() {
  const tokenResponse = await credential.getToken(['https://graph.microsoft.com/.default']);
  return tokenResponse.token;
}

async function downloadFileWithAxios(driveId, itemId, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const response = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: 'arraybuffer'
  });
  return Buffer.from(response.data);
}

async function uploadFileWithAxios(driveId, itemId, buffer, accessToken) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const response = await axios.put(url, buffer, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
  });
  return response.data;
}

// ============ FUNÇÃO DE AUTENTICAÇÃO ============

// Função para ler usuários do arquivo JSON
function getUsers() {
  try {
    const usersPath = path.join(__dirname, 'users.json');
    const data = fs.readFileSync(usersPath, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    console.error('❌ Erro ao ler users.json:', error.message);
    // Se o arquivo não existir, retorna um array vazio
    return { users: [] };
  }
}

// ============ FUNÇÕES PARA DESBLOQUEAR ARQUIVO ============

async function unlockFileIfNeeded(driveId, itemId) {
  try {
    console.log('🔓 Verificando bloqueio...');
    const file = await graphClient.api(`/drives/${driveId}/items/${itemId}`).get();
    
    if (file.ctag && file.ctag.includes('checkout')) {
      console.log('⚠️ Checkout ativo. Forçando check-in...');
      try {
        await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
          checkInAs: 'Major',
          comment: 'Check-in automático'
        });
        console.log('✅ Check-in realizado!');
      } catch (e) {
        try {
          await graphClient.api(`/drives/${driveId}/items/${itemId}/checkout`).post({});
          await new Promise(r => setTimeout(r, 1000));
          await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
            checkInAs: 'Major',
            comment: 'Check-in forçado'
          });
          console.log('✅ Checkout/Checkin forçado!');
        } catch (e2) {
          console.log('⚠️ Não foi possível forçar');
        }
      }
      await new Promise(r => setTimeout(r, 2000));
    }
    return true;
  } catch (error) {
    return true;
  }
}

async function forceUnlockFile(driveId, itemId) {
  try {
    console.log('🔓 Desbloqueando...');
    await graphClient.api(`/drives/${driveId}/items/${itemId}/checkout`).post({});
    await new Promise(r => setTimeout(r, 1500));
    await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
      checkInAs: 'Major',
      comment: 'Desbloqueio automático'
    });
    await new Promise(r => setTimeout(r, 2000));
    return true;
  } catch (error) {
    return false;
  }
}

// ============ FUNÇÃO DE LEITURA ============

async function getExcelFromSharePoint(fileUrl) {
  try {
    console.log('🔍 Buscando planilha...');
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) fileId = extractedId;
      console.log('📁 ID extraído da URL:', fileId);
    }
    
    console.log('👤 Usuário:', USER_EMAIL);
    console.log('📁 File ID:', fileId);
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    console.log('✅ Drive encontrado');
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo:', fileItem.name);
    } catch (error) {
      console.log('🔍 Buscando pelo nome...');
      const files = await graphClient.api(`/drives/${drive.id}/root/children`).get();
      const excelFile = files.value.find(f => 
        f.name && f.name.toLowerCase().includes('cocred') && f.name.includes('.xlsx')
      );
      if (excelFile) {
        fileItem = excelFile;
        console.log('📄 Arquivo:', fileItem.name);
      } else {
        throw new Error('Arquivo não encontrado');
      }
    }
    
    const accessToken = await getAccessToken();
    const buffer = await downloadFileWithAxios(drive.id, fileItem.id, accessToken);
    
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetNames = workbook.SheetNames;
    console.log('📑 Abas:', sheetNames.join(', '));
    
    let sheetName = sheetNames[0];
    const targetSheet = SHEET_NAME;
    const exactMatch = sheetNames.find(s => s.toUpperCase() === targetSheet.toUpperCase());
    if (exactMatch) {
      sheetName = exactMatch;
      console.log(`✅ Usando aba "${sheetName}"`);
    }
    
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);
    console.log(`📊 ${data.length} registros encontrados`);
    
    // 🔥 FILTRAR LINHA DE TOTAL
    const dadosFiltrados = data.filter(row => {
      const nome = row["Projeto/Campanha"] || row["Projeto"] || row["Campanha"] || '';
      const isTotal = nome.toLowerCase().includes('total') || 
                      nome.toLowerCase().includes('totais') ||
                      nome.toLowerCase().includes('todos os projetos') ||
                      nome.toLowerCase().includes('geral');
      const isEmpty = !nome || nome.trim() === '';
      return !isTotal && !isEmpty;
    });
    
    console.log(`✅ ${dadosFiltrados.length} registros carregados (${data.length - dadosFiltrados.length} filtrados)`);
    
    return { 
      success: true, 
      data: dadosFiltrados, 
      sheetName, 
      totalRows: dadosFiltrados.length,
      originalRows: data.length,
      filteredRows: data.length - dadosFiltrados.length,
      driveId: drive.id,
      fileId: fileItem.id,
      fileName: fileItem.name
    };
  } catch (error) {
    console.error('❌ Erro:', error.message);
    return { success: false, error: error.message };
  }
}

// ============ FUNÇÃO DE ESCRITA ============

async function writeExcelToSharePoint(campanhas, fileUrl) {
  try {
    console.log('💾 Atualizando planilha...');
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) fileId = extractedId;
    }
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    console.log('✅ Drive encontrado');
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo:', fileItem.name);
    } catch (error) {
      throw new Error('Arquivo não encontrado');
    }
    
    await unlockFileIfNeeded(drive.id, fileItem.id);
    
    console.log('📥 Baixando arquivo existente...');
    const accessToken = await getAccessToken();
    const existingFile = await downloadFileWithAxios(drive.id, fileItem.id, accessToken);
    
    const workbook = XLSX.read(existingFile, { type: 'buffer' });
    console.log('📑 Abas existentes:', workbook.SheetNames.join(', '));
    
    const targetSheet = SHEET_NAME;
    
    if (!workbook.SheetNames.includes(targetSheet)) {
      console.log(`⚠️ Aba "${targetSheet}" não encontrada. Criando...`);
      const ws = XLSX.utils.json_to_sheet([]);
      XLSX.utils.book_append_sheet(workbook, ws, targetSheet);
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
    
    const ws = XLSX.utils.aoa_to_sheet(values);
    workbook.Sheets[targetSheet] = ws;
    
    const buffer = XLSX.write(workbook, { 
      type: 'buffer', 
      bookType: 'xlsx'
    });
    
    let tentativas = 0;
    let maxTentativas = 5;
    let sucesso = false;
    let ultimoErro = null;
    
    while (!sucesso && tentativas < maxTentativas) {
      tentativas++;
      try {
        console.log(`📤 Tentativa ${tentativas}/${maxTentativas}...`);
        if (tentativas > 1) await forceUnlockFile(drive.id, fileItem.id);
        const token = await getAccessToken();
        await uploadFileWithAxios(drive.id, fileItem.id, buffer, token);
        sucesso = true;
        console.log(`✅ ${campanhas.length} registros salvos! (tentativa ${tentativas})`);
      } catch (error) {
        ultimoErro = error;
        const statusCode = error.response?.status;
        console.log(`❌ Erro ${statusCode}: ${error.message}`);
        if (statusCode === 423 || statusCode === 429) {
          const tempo = statusCode === 423 ? 3000 : 5000;
          console.log(`⏳ Aguardando ${tempo/1000}s...`);
          await new Promise(r => setTimeout(r, tempo));
        } else {
          throw error;
        }
      }
    }
    
    if (!sucesso) {
      throw new Error(`Falha após ${maxTentativas} tentativas: ${ultimoErro?.message}`);
    }
    
    return { success: true, totalRows: campanhas.length };
  } catch (error) {
    console.error('❌ Erro:', error.message);
    return { success: false, error: error.message };
  }
}

// ============ ENDPOINTS DA API ============

// ============ ENDPOINT DE LOGIN ============
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;
  
  console.log('🔐 Tentativa de login:', username);
  
  if (!username || !password) {
    return res.status(400).json({
      success: false,
      error: 'Usuário e senha são obrigatórios'
    });
  }
  
  try {
    const { users } = getUsers();
    
    const found = users.find(u => 
      u.username.toLowerCase() === username.toLowerCase() && 
      u.password === password
    );
    
    if (found) {
      console.log('✅ Login bem-sucedido:', username);
      res.json({ 
        success: true, 
        user: username,
        message: 'Login realizado com sucesso!'
      });
    } else {
      console.log('❌ Falha no login:', username);
      res.status(401).json({ 
        success: false, 
        error: 'Credenciais inválidas. Verifique seu usuário e senha.' 
      });
    }
  } catch (error) {
    console.error('❌ Erro ao validar usuário:', error.message);
    res.status(500).json({ 
      success: false, 
      error: 'Erro interno ao validar credenciais' 
    });
  }
});

// Endpoint para listar usuários (apenas para debug)
app.get('/api/usuarios', async (req, res) => {
  try {
    const { users } = getUsers();
    // Retorna apenas os nomes de usuário (sem as senhas)
    const usuarios = users.map(u => ({ username: u.username }));
    res.json({ success: true, usuarios });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ============ ENDPOINTS DO SHAREPOINT ============

app.post('/api/buscar-planilha', async (req, res) => {
  console.log('📥 /buscar-planilha');
  const result = await getExcelFromSharePoint(req.body.fileUrl);
  
  if (result.success) {
    if (result.data.length > 0) {
      console.log('\n📋 COLUNAS ENCONTRADAS:');
      const keys = Object.keys(result.data[0]);
      keys.forEach((key, index) => {
        console.log(`   ${index}: "${key}"`);
      });
    }
    
    const mappedData = result.data.map((row, idx) => {
      const nome = row["Projeto/Campanha"] || row["Projeto"] || row["Campanha"] || `Item ${idx + 1}`;
      
      const planejadoMidia = Number(row["Valor Planejado Mídia/Projeto"] || row["Planejado Mídia"] || row["Planejado Mídia/Projeto"] || 0);
      const realizadoMidia = Number(row["Valor Realizado Mídia/Projeto"] || row["Realizado Mídia"] || row["Realizado Mídia/Projeto"] || 0);
      
      const planejadoProd = Number(
        row["Valor Planejado Prod"] || 
        row["Valor Planejado  Prod"] || 
        row["Planejado Prod"] || 
        row["Planejado Produção"] ||
        row["Valor Planejado Produção"] ||
        0
      );
      
      const realizadoProd = Number(
        row["Valor Realizado Prod"] || 
        row["Valor Realizado  Prod"] ||
        row["Realizado Prod"] || 
        row["Realizado Produção"] ||
        row["Valor Realizado Produção"] ||
        0
      );
      
      const saldoMidia = Number(row["Saldo Mídia/Projeto"] || row["Saldo Mídia"] || (planejadoMidia - realizadoMidia));
      const saldoProd = Number(row["Saldo Prod"] || row["Saldo Produção"] || (planejadoProd - realizadoProd));
      
      const bolsa = row["Bolsa"] || row["Tipo"] || "Avulsa";
      const periodo = row["Período"] || row["Periodo"] || "Q1";
      
      return {
        id: `sp_${Date.now()}_${idx}`,
        nome,
        planejadoMidia,
        realizadoMidia,
        saldoMidia,
        planejadoProd,
        realizadoProd,
        saldoProd,
        bolsa,
        periodo
      };
    });
    
    console.log(`✅ ${mappedData.length} campanhas convertidas`);
    
    res.json({ 
      success: true, 
      data: mappedData, 
      totalRows: mappedData.length,
      filteredRows: result.filteredRows || 0
    });
  } else {
    console.log('❌ Erro:', result.error);
    res.status(500).json({ success: false, error: result.error });
  }
});

app.post('/api/salvar-campanhas', async (req, res) => {
  console.log('💾 /salvar-campanhas');
  const { campanhas, fileUrl } = req.body;
  
  if (!campanhas || !Array.isArray(campanhas) || campanhas.length === 0) {
    return res.status(400).json({
      success: false,
      error: 'Dados inválidos. Envie um array de campanhas.'
    });
  }
  
  console.log(`📊 ${campanhas.length} campanhas para salvar`);
  
  const result = await writeExcelToSharePoint(campanhas, fileUrl);
  
  if (result.success) {
    res.json({ 
      success: true, 
      message: `${result.totalRows} campanhas salvas!`,
      totalRows: result.totalRows 
    });
  } else {
    res.status(500).json({ success: false, error: result.error });
  }
});

app.post('/api/desbloquear-arquivo', async (req, res) => {
  console.log('🔓 /desbloquear-arquivo');
  try {
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    const result = await forceUnlockFile(drive.id, FILE_ID);
    res.json({
      success: result,
      message: result ? 'Arquivo desbloqueado!' : 'Falha ao desbloquear'
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/listar-arquivos', async (req, res) => {
  try {
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    const files = await graphClient.api(`/drives/${drive.id}/root/children`).get();
    const excelFiles = files.value
      .filter(f => f.name && f.name.includes('.xlsx'))
      .map(f => ({ name: f.name, id: f.id, size: f.size }));
    res.json({ success: true, files: excelFiles });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ 
    status: 'online',
    version: '2.0.5',
    endpoints: {
      login: 'POST /api/login',
      buscar: 'POST /api/buscar-planilha',
      salvar: 'POST /api/salvar-campanhas',
      desbloquear: 'POST /api/desbloquear-arquivo',
      listar: 'GET /api/listar-arquivos',
      health: 'GET /api/health'
    }
  });
});

app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════════════╗
║  🚀 SICOOB COCRED - BACKEND v2.0.5                                  ║
║  📡 Servidor: http://localhost:${PORT}                               ║
║  🔐 POST /api/login                                                 ║
║  📊 POST /api/buscar-planilha                                       ║
║  ✏️ POST /api/salvar-campanhas                                      ║
║  🔓 POST /api/desbloquear-arquivo                                   ║
║  📁 GET  /api/listar-arquivos                                       ║
║  🏥 GET  /api/health                                                ║
╚═══════════════════════════════════════════════════════════════════════╝
  `);
  console.log(`📋 Configurações:`);
  console.log(`   👤 Usuário: ${USER_EMAIL}`);
  console.log(`   📁 File ID: ${FILE_ID}`);
  console.log(`   📑 Aba alvo: ${SHEET_NAME}`);
  console.log(`   🔑 Client ID: ${process.env.CLIENT_ID?.substring(0, 10)}...`);
  console.log(`\n🔐 Autenticação via backend/users.json`);
});