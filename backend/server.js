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

// ============ CARREGAR USUÁRIOS (COMPATÍVEL COM RAILWAY) ============

// Caminho do arquivo users.json (para ambiente local)
const USERS_PATH = path.join(__dirname, 'users.json');

// Função para carregar usuários
function loadUsers() {
  try {
    // 1. Tenta carregar do arquivo users.json (local)
    if (fs.existsSync(USERS_PATH)) {
      const data = fs.readFileSync(USERS_PATH, 'utf8');
      const users = JSON.parse(data);
      console.log(`✅ ${Object.keys(users).length} usuários carregados do users.json`);
      return users;
    }
    
    // 2. Tenta carregar da variável de ambiente USERS_JSON (Railway/Produção)
    if (process.env.USERS_JSON) {
      const users = JSON.parse(process.env.USERS_JSON);
      console.log(`✅ ${Object.keys(users).length} usuários carregados da variável USERS_JSON`);
      return users;
    }
    
    // 3. Fallback: usuário admin padrão (apenas para desenvolvimento)
    console.log('⚠️ Nenhum usuário encontrado. Usando admin padrão.');
    return {
      'admin': {
        password: 'admin123',
        name: 'Administrador',
        role: 'admin'
      }
    };
  } catch (error) {
    console.error('❌ Erro ao carregar usuários:', error.message);
    return {
      'admin': {
        password: 'admin123',
        name: 'Administrador',
        role: 'admin'
      }
    };
  }
}

// Função para salvar usuários (apenas ambiente local)
function saveUsers(users) {
  try {
    // Só salva em ambiente local (não no Railway)
    if (!process.env.RAILWAY_ENVIRONMENT) {
      fs.writeFileSync(USERS_PATH, JSON.stringify(users, null, 2), 'utf8');
      console.log('✅ Usuários salvos no users.json');
    } else {
      console.log('ℹ️ Ambiente Railway: usuários salvos apenas na memória');
    }
    return true;
  } catch (error) {
    console.error('❌ Erro ao salvar users.json:', error.message);
    return false;
  }
}

// Carrega os usuários
let VALID_USERS = loadUsers();

// ============ CONFIGURAÇÃO MICROSOFT GRAPH ============

// Verifica se as variáveis de ambiente estão configuradas
const requiredEnvVars = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET'];
const missingEnvVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingEnvVars.length > 0) {
  console.error(`❌ Variáveis de ambiente faltando: ${missingEnvVars.join(', ')}`);
  console.error('   Configure as variáveis no Railway ou no arquivo .env');
}

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const graphClient = Client.initWithMiddleware({ authProvider });

// IDs fixos do arquivo (com fallback para Railway)
const FILE_ID = process.env.SHAREPOINT_FILE_ID || '';
const USER_EMAIL = process.env.SHAREPOINT_USERNAME || '';
const SOURCEDOC = process.env.SHAREPOINT_SOURCEDOC || '';
const SHEET_NAME = process.env.SHEET_NAME || 'VALOR POR PROJETO';

// Verifica se as configurações do SharePoint estão completas
if (!FILE_ID || !USER_EMAIL) {
  console.warn('⚠️ Configurações do SharePoint incompletas:');
  if (!FILE_ID) console.warn('   - SHAREPOINT_FILE_ID não configurado');
  if (!USER_EMAIL) console.warn('   - SHAREPOINT_USERNAME não configurado');
  console.warn('   A funcionalidade de SharePoint pode não funcionar corretamente.');
}

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
    timeout: 30000 // 30 segundos de timeout
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
    timeout: 60000 // 60 segundos de timeout
  });
  return response.data;
}

// ============ FUNÇÕES PARA DESBLOQUEAR ARQUIVO ============

// Função para verificar e desbloquear arquivo
async function unlockFileIfNeeded(driveId, itemId) {
  try {
    console.log('🔓 Verificando se o arquivo está bloqueado...');
    
    const file = await graphClient.api(`/drives/${driveId}/items/${itemId}`).get();
    
    if (file.ctag && file.ctag.includes('checkout')) {
      console.log('⚠️ Arquivo está com checkout ativo. Forçando check-in...');
      
      try {
        await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
          checkInAs: 'Major',
          comment: 'Check-in automático para salvar dados'
        });
        console.log('✅ Check-in realizado com sucesso!');
      } catch (checkinError) {
        console.log('⚠️ Erro no check-in, tentando checkout/checkin forçado...');
        try {
          await graphClient.api(`/drives/${driveId}/items/${itemId}/checkout`).post({});
          await new Promise(resolve => setTimeout(resolve, 1000));
          await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
            checkInAs: 'Major',
            comment: 'Check-in automático para salvar dados'
          });
          console.log('✅ Checkout/Check-in forçado realizado com sucesso!');
        } catch (forceError) {
          console.log('⚠️ Não foi possível forçar checkout/checkin');
        }
      }
      
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
    
    return true;
  } catch (error) {
    console.log('ℹ️ Arquivo não está bloqueado ou já está disponível');
    return true;
  }
}

// Função para desbloquear forçadamente
async function forceUnlockFile(driveId, itemId) {
  try {
    console.log('🔓 Tentando desbloquear arquivo...');
    
    await graphClient.api(`/drives/${driveId}/items/${itemId}/checkout`).post({});
    console.log('✅ Checkout realizado');
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    await graphClient.api(`/drives/${driveId}/items/${itemId}/checkin`).post({
      checkInAs: 'Major',
      comment: 'Desbloqueio automático'
    });
    console.log('✅ Check-in realizado, arquivo desbloqueado');
    
    await new Promise(resolve => setTimeout(resolve, 2000));
    return true;
  } catch (error) {
    console.error('❌ Erro ao desbloquear:', error.message);
    return false;
  }
}

// ============ FUNÇÃO DE LEITURA ============

async function getExcelFromSharePoint(fileUrl) {
  try {
    console.log('🔍 Iniciando busca da planilha...');
    
    if (!USER_EMAIL || !FILE_ID) {
      throw new Error('Configurações do SharePoint não estão completas. Verifique as variáveis de ambiente.');
    }
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) {
        fileId = extractedId;
        console.log('📁 ID extraído da URL:', fileId);
      }
    }
    
    console.log('👤 Usuário:', USER_EMAIL);
    console.log('📁 File ID:', fileId);
    
    let drive;
    try {
      drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
      console.log('✅ Drive encontrado');
    } catch (error) {
      console.error('❌ Erro ao acessar drive:', error.message);
      throw new Error(`Não foi possível acessar o drive de ${USER_EMAIL}. Verifique se o usuário existe e tem acesso.`);
    }
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo encontrado:', fileItem.name);
      console.log('📏 Tamanho:', fileItem.size, 'bytes');
    } catch (error) {
      console.error('❌ Arquivo não encontrado pelo ID:', error.message);
      
      console.log('🔍 Tentando buscar pelo nome...');
      const files = await graphClient.api(`/drives/${drive.id}/root/children`).get();
      const excelFile = files.value.find(f => 
        f.name && f.name.toLowerCase().includes('cocred') && f.name.toLowerCase().includes('.xlsx')
      );
      
      if (excelFile) {
        fileItem = excelFile;
        console.log('📄 Arquivo encontrado pelo nome:', fileItem.name);
      } else {
        throw new Error('Arquivo não encontrado. Verifique se a planilha está no drive.');
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
      console.log(`✅ Usando aba "${sheetName}" (exata)`);
    } else {
      const partialMatch = sheetNames.find(s => 
        s.toUpperCase().includes('VALOR') && s.toUpperCase().includes('PROJETO')
      );
      if (partialMatch) {
        sheetName = partialMatch;
        console.log(`✅ Usando aba "${sheetName}" (parcial)`);
      } else {
        console.log(`⚠️ Usando primeira aba disponível: ${sheetName}`);
      }
    }
    
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    console.log(`✅ ${data.length} registros carregados da aba "${sheetName}"`);
    
    if (data.length > 0) {
      const columns = Object.keys(data[0]);
      console.log('📋 Colunas encontradas:', columns.slice(0, 8).join(', '));
    }
    
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

// ============ FUNÇÃO DE ESCRITA ============

async function writeExcelToSharePoint(campanhas, fileUrl) {
  try {
    console.log('💾 Iniciando escrita da planilha...');
    
    if (!USER_EMAIL || !FILE_ID) {
      throw new Error('Configurações do SharePoint não estão completas. Verifique as variáveis de ambiente.');
    }
    
    let fileId = FILE_ID;
    if (fileUrl) {
      const extractedId = extractFileIdFromUrl(fileUrl);
      if (extractedId) {
        fileId = extractedId;
        console.log('📁 ID extraído da URL:', fileId);
      }
    }
    
    console.log('👤 Usuário:', USER_EMAIL);
    console.log('📁 File ID:', fileId);
    
    let drive;
    try {
      drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
      console.log('✅ Drive encontrado');
    } catch (error) {
      console.error('❌ Erro ao acessar drive:', error.message);
      throw new Error(`Não foi possível acessar o drive de ${USER_EMAIL}.`);
    }
    
    let fileItem;
    try {
      fileItem = await graphClient.api(`/drives/${drive.id}/items/${fileId}`).get();
      console.log('📄 Arquivo encontrado:', fileItem.name);
    } catch (error) {
      console.error('❌ Arquivo não encontrado pelo ID:', error.message);
      throw new Error('Arquivo não encontrado. Verifique se o File ID está correto.');
    }
    
    await unlockFileIfNeeded(drive.id, fileItem.id);
    
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
    let maxTentativas = 5;
    let sucesso = false;
    let ultimoErro = null;
    
    while (!sucesso && tentativas < maxTentativas) {
      tentativas++;
      try {
        console.log(`📤 Tentativa ${tentativas}/${maxTentativas} de upload...`);
        
        if (tentativas > 1) {
          console.log('🔄 Tentando desbloquear novamente...');
          await forceUnlockFile(drive.id, fileItem.id);
        }
        
        const accessToken = await getAccessToken();
        await uploadFileWithAxios(drive.id, fileItem.id, buffer, accessToken);
        
        sucesso = true;
        console.log(`✅ ${campanhas.length} registros salvos no SharePoint (tentativa ${tentativas})`);
      } catch (error) {
        ultimoErro = error;
        const statusCode = error.response?.status;
        console.log(`❌ Erro na tentativa ${tentativas}: Status ${statusCode} - ${error.message}`);
        
        if (statusCode === 423) {
          console.log(`⚠️ Arquivo bloqueado. Aguardando ${tentativas * 2} segundos...`);
          await new Promise(resolve => setTimeout(resolve, tentativas * 2000));
        } else if (statusCode === 429) {
          console.log(`⚠️ Limite de requisições excedido. Aguardando ${tentativas * 3} segundos...`);
          await new Promise(resolve => setTimeout(resolve, tentativas * 3000));
        } else if (statusCode === 401 || statusCode === 403) {
          console.error('❌ Erro de autenticação. Verifique as credenciais.');
          throw error;
        } else {
          throw error;
        }
      }
    }
    
    if (!sucesso) {
      const mensagem = ultimoErro?.response?.data?.error?.message || ultimoErro?.message || 'Erro desconhecido';
      throw new Error(`Não foi possível salvar após ${maxTentativas} tentativas. Último erro: ${mensagem}`);
    }
    
    return { success: true, totalRows: campanhas.length };
  } catch (error) {
    console.error('❌ Erro detalhado:', error.message);
    if (error.response) {
      console.error('📋 Detalhes da resposta:', error.response.data);
    }
    return { success: false, error: error.message };
  }
}

// ============ ENDPOINTS DE AUTENTICAÇÃO ============

// Endpoint de login
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

// Endpoint para verificar autenticação
app.post('/api/verificar-auth', async (req, res) => {
  try {
    const { token } = req.body;
    
    if (!token) {
      return res.status(401).json({
        success: false,
        error: 'Token não fornecido'
      });
    }

    try {
      const decoded = Buffer.from(token, 'base64').toString('utf-8');
      const [username, timestamp] = decoded.split(':');
      
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
    } catch (e) {
      // Token inválido
    }

    return res.status(401).json({
      success: false,
      error: 'Token inválido ou expirado'
    });

  } catch (error) {
    console.error('❌ Erro na verificação:', error);
    res.status(500).json({
      success: false,
      error: 'Erro interno do servidor'
    });
  }
});

// ============ ENDPOINTS PARA GERENCIAR USUÁRIOS ============

// Listar todos os usuários (com senha escondida)
app.get('/api/usuarios', async (req, res) => {
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

// Adicionar novo usuário
app.post('/api/usuarios', async (req, res) => {
  try {
    const { username, password, name, role } = req.body;
    
    if (!username || !password || !name) {
      return res.status(400).json({
        success: false,
        error: 'Username, password e name são obrigatórios'
      });
    }

    if (VALID_USERS[username]) {
      return res.status(409).json({
        success: false,
        error: 'Usuário já existe'
      });
    }

    VALID_USERS[username] = {
      password: password,
      name: name,
      role: role || 'usuario'
    };

    saveUsers(VALID_USERS);

    res.json({
      success: true,
      message: `Usuário ${username} criado com sucesso!`,
      user: {
        username: username,
        name: name,
        role: role || 'usuario'
      }
    });
  } catch (error) {
    console.error('❌ Erro ao criar usuário:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Atualizar usuário
app.put('/api/usuarios/:username', async (req, res) => {
  try {
    const { username } = req.params;
    const { name, role, password } = req.body;
    
    if (!VALID_USERS[username]) {
      return res.status(404).json({
        success: false,
        error: 'Usuário não encontrado'
      });
    }

    if (name) VALID_USERS[username].name = name;
    if (role) VALID_USERS[username].role = role;
    if (password) VALID_USERS[username].password = password;

    saveUsers(VALID_USERS);

    res.json({
      success: true,
      message: `Usuário ${username} atualizado com sucesso!`,
      user: {
        username: username,
        name: VALID_USERS[username].name,
        role: VALID_USERS[username].role
      }
    });
  } catch (error) {
    console.error('❌ Erro ao atualizar usuário:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Deletar usuário
app.delete('/api/usuarios/:username', async (req, res) => {
  try {
    const { username } = req.params;
    
    if (username === 'admin') {
      return res.status(403).json({
        success: false,
        error: 'Não é possível deletar o usuário admin'
      });
    }

    if (!VALID_USERS[username]) {
      return res.status(404).json({
        success: false,
        error: 'Usuário não encontrado'
      });
    }

    delete VALID_USERS[username];
    
    saveUsers(VALID_USERS);

    res.json({
      success: true,
      message: `Usuário ${username} deletado com sucesso!`
    });
  } catch (error) {
    console.error('❌ Erro ao deletar usuário:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Endpoint para alterar senha (mantido para compatibilidade)
app.post('/api/alterar-senha', async (req, res) => {
  try {
    const { username, oldPassword, newPassword } = req.body;
    
    if (!username || !oldPassword || !newPassword) {
      return res.status(400).json({
        success: false,
        error: 'Todos os campos são obrigatórios'
      });
    }

    if (!VALID_USERS[username]) {
      return res.status(404).json({
        success: false,
        error: 'Usuário não encontrado'
      });
    }

    if (VALID_USERS[username].password !== oldPassword) {
      return res.status(401).json({
        success: false,
        error: 'Senha atual incorreta'
      });
    }

    VALID_USERS[username].password = newPassword;
    saveUsers(VALID_USERS);
    
    console.log(`✅ Senha alterada para: ${username}`);

    res.json({
      success: true,
      message: 'Senha alterada com sucesso!'
    });

  } catch (error) {
    console.error('❌ Erro ao alterar senha:', error);
    res.status(500).json({
      success: false,
      error: 'Erro interno do servidor'
    });
  }
});

// ============ ENDPOINTS DO SHAREPOINT ============

// Endpoint para BUSCAR dados da planilha
app.post('/api/buscar-planilha', async (req, res) => {
  const { fileUrl } = req.body;
  console.log('📥 Requisição recebida em /api/buscar-planilha');
  
  const result = await getExcelFromSharePoint(fileUrl);
  
  if (result.success) {
    const mappedData = result.data.map((row, idx) => ({
      id: `sp_${Date.now()}_${idx}`,
      nome: row["Projeto/Campanha"] || row["Projeto"] || row["Campanha"] || `Item ${idx + 1}`,
      planejadoMidia: Number(row["Valor Planejado Mídia/Projeto"] || row["Planejado Mídia"] || 0),
      realizadoMidia: Number(row["Valor Realizado Mídia/Projeto"] || row["Realizado Mídia"] || 0),
      saldoMidia: Number(row["Saldo Mídia/Projeto"] || (Number(row["Valor Planejado Mídia/Projeto"] || 0) - Number(row["Valor Realizado Mídia/Projeto"] || 0))),
      planejadoProd: Number(row["Valor Planejado Prod"] || row["Planejado Prod"] || 0),
      realizadoProd: Number(row["Valor Realizado Prod"] || row["Realizado Prod"] || 0),
      saldoProd: Number(row["Saldo Prod"] || (Number(row["Valor Planejado Prod"] || 0) - Number(row["Valor Realizado Prod"] || 0))),
      bolsa: row["Bolsa"] || row["Tipo"] || "Avulsa",
      periodo: row["Período"] || row["Periodo"] || "Q1"
    }));
    
    console.log(`✅ ${mappedData.length} campanhas convertidas com sucesso`);
    
    res.json({ 
      success: true, 
      data: mappedData,
      sheetName: result.sheetName,
      totalRows: mappedData.length,
      originalRows: result.totalRows
    });
  } else {
    console.log('❌ Erro retornado ao cliente:', result.error);
    res.status(500).json({ success: false, error: result.error });
  }
});

// Endpoint para SALVAR dados no SharePoint
app.post('/api/salvar-campanhas', async (req, res) => {
  console.log('💾 Requisição recebida em /api/salvar-campanhas');
  const { campanhas, fileUrl } = req.body;
  
  if (!campanhas || !Array.isArray(campanhas)) {
    return res.status(400).json({
      success: false,
      error: 'Dados inválidos. Envie um array de campanhas.'
    });
  }
  
  console.log(`📊 ${campanhas.length} campanhas para salvar`);
  
  try {
    const result = await writeExcelToSharePoint(campanhas, fileUrl);
    
    if (result.success) {
      res.json({
        success: true,
        message: `${result.totalRows} campanhas salvas com sucesso no SharePoint!`,
        totalRows: result.totalRows
      });
    } else {
      res.status(500).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('❌ Erro ao salvar:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint para DESBLOQUEAR arquivo
app.post('/api/desbloquear-arquivo', async (req, res) => {
  try {
    console.log('🔓 Requisição para desbloquear arquivo...');
    const { fileId } = req.body;
    
    const id = fileId || FILE_ID;
    
    if (!id) {
      return res.status(400).json({
        success: false,
        error: 'File ID não fornecido'
      });
    }
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    
    const result = await forceUnlockFile(drive.id, id);
    
    res.json({
      success: result,
      message: result ? 'Arquivo desbloqueado com sucesso!' : 'Falha ao desbloquear arquivo'
    });
  } catch (error) {
    console.error('❌ Erro ao desbloquear:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint para CRIAR novo arquivo Excel
app.post('/api/criar-novo-arquivo', async (req, res) => {
  try {
    console.log('📄 Criando novo arquivo Excel no SharePoint...');
    
    const { nomeArquivo } = req.body;
    const fileName = nomeArquivo || 'campanhas_cocred.xlsx';
    
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    console.log('✅ Drive encontrado:', drive.id);
    
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
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    XLSX.utils.book_append_sheet(wb, ws, 'VALOR POR PROJETO');
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    
    const accessToken = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${fileName}:/content`;
    
    await axios.put(url, buffer, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    });
    
    const fileResponse = await graphClient.api(`/drives/${drive.id}/root:/${fileName}`).get();
    
    console.log('✅ Arquivo criado com sucesso!');
    console.log(`📁 Nome: ${fileResponse.name}`);
    console.log(`📄 ID: ${fileResponse.id}`);
    
    res.json({
      success: true,
      message: `Arquivo "${fileName}" criado com sucesso!`,
      fileId: fileResponse.id,
      fileName: fileResponse.name,
      driveId: drive.id,
      envUpdate: `
# Atualize seu .env com:
SHAREPOINT_FILE_ID=${fileResponse.id}
      `.trim()
    });
  } catch (error) {
    console.error('❌ Erro ao criar arquivo:', error.message);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Endpoint para listar arquivos do drive
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

// Endpoint para buscar informações do drive
app.get('/api/drive-info', async (req, res) => {
  try {
    const drive = await graphClient.api(`/users/${USER_EMAIL}/drive`).get();
    res.json({ 
      success: true, 
      drive: {
        id: drive.id,
        name: drive.name,
        owner: drive.owner,
        quota: drive.quota
      }
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// Endpoint de health check
app.get('/api/health', (req, res) => {
  const hasSharePointConfig = !!(FILE_ID && USER_EMAIL);
  
  res.json({ 
    status: 'online', 
    message: 'Backend funcionando!',
    version: '2.1.0',
    environment: process.env.RAILWAY_ENVIRONMENT ? 'railway' : 'local',
    usersCount: Object.keys(VALID_USERS).length,
    sharepoint: {
      configured: hasSharePointConfig,
      user: USER_EMAIL ? '✓' : '✗',
      fileId: FILE_ID ? '✓' : '✗'
    },
    endpoints: {
      auth: 'POST /api/login',
      verify: 'POST /api/verificar-auth',
      users: 'GET /api/usuarios',
      createUser: 'POST /api/usuarios',
      updateUser: 'PUT /api/usuarios/:username',
      deleteUser: 'DELETE /api/usuarios/:username',
      changePassword: 'POST /api/alterar-senha',
      buscar: 'POST /api/buscar-planilha',
      salvar: 'POST /api/salvar-campanhas',
      desbloquear: 'POST /api/desbloquear-arquivo',
      criarArquivo: 'POST /api/criar-novo-arquivo',
      listar: 'GET /api/listar-arquivos',
      drive: 'GET /api/drive-info',
      health: 'GET /api/health'
    }
  });
});

// ============ INICIAR SERVIDOR ============

app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════════════════════════════════╗
║  🚀 SICOOB COCRED - BACKEND v2.1.0                                  ║
║  📡 Servidor rodando em http://localhost:${PORT}                      ║
║  🌍 Ambiente: ${process.env.RAILWAY_ENVIRONMENT ? 'RAILWAY' : 'LOCAL'}     ║
║  🔐 Endpoint: POST /api/login                                       ║
║  👥 Endpoint: GET  /api/usuarios (GERENCIAMENTO DE USUÁRIOS)       ║
║  📊 Endpoint: POST /api/buscar-planilha                             ║
║  ✏️ Endpoint: POST /api/salvar-campanhas (COM RETRY AUTOMÁTICO!)    ║
║  🔓 Endpoint: POST /api/desbloquear-arquivo                         ║
║  📄 Endpoint: POST /api/criar-novo-arquivo                         ║
║  📁 Endpoint: GET  /api/listar-arquivos                             ║
║  🏥 Health check: GET /api/health                                   ║
╚═══════════════════════════════════════════════════════════════════════╝
  `);
  
  console.log(`📋 Configurações:`);
  console.log(`   🌍 Ambiente: ${process.env.RAILWAY_ENVIRONMENT ? 'Railway' : 'Local'}`);
  console.log(`   👤 Usuário SharePoint: ${USER_EMAIL || '⚠️ NÃO CONFIGURADO'}`);
  console.log(`   📁 File ID: ${FILE_ID ? '✅ Configurado' : '⚠️ NÃO CONFIGURADO'}`);
  console.log(`   📑 Aba alvo: ${SHEET_NAME}`);
  console.log(`   🔑 Client ID: ${process.env.CLIENT_ID?.substring(0, 10)}...`);
  
  console.log(`\n👥 Usuários cadastrados: ${Object.keys(VALID_USERS).length}`);
  
  if (process.env.RAILWAY_ENVIRONMENT) {
    console.log(`\n🌍 Railway URL: https://${process.env.RAILWAY_PUBLIC_DOMAIN || 'sua-url'}`);
  } else {
    console.log(`   📁 Arquivo: ${USERS_PATH}`);
  }
  
  console.log(`\n🔐 Credenciais de teste:`);
  console.log(`   admin / admin123`);
  
  // Lista alguns usuários (sem mostrar senhas completas)
  const userList = Object.keys(VALID_USERS).slice(0, 3);
  if (userList.length > 0) {
    console.log(`   ${userList.map(u => `${u} / ****`).join('\n   ')}`);
  }
  
  console.log(`\n💡 Use o botão "Salvar no Excel" no frontend para escrever dados!`);
  console.log(`🔓 Se o arquivo estiver bloqueado, use POST /api/desbloquear-arquivo`);
});