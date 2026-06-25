// Função para extrair informações do SharePoint a partir do link
function extractSharePointInfo(shareUrl) {
  console.log('🔍 Extraindo informações do link:', shareUrl);
  
  // Padrão 1: Link de compartilhamento padrão
  const pattern1 = /https:\/\/([^\/]+)\.sharepoint\.com\/:x:\/g\/personal\/([^\/]+)\/([A-Za-z0-9_-]+)/;
  const match1 = shareUrl.match(pattern1);
  
  if (match1) {
    const domain = match1[1];
    let userEmailRaw = match1[2];
    const fileId = match1[3];
    
    // Corrigir formato do email corretamente
    // Exemplo: cristini_cordesco_ideatoreamericas_com -> cristini.cordesco@ideatoreamericas.com
    let userEmail = userEmailRaw.replace(/_/g, '.');
    // Remover domínio duplicado se existir
    userEmail = userEmail.replace(/\.ideatoreamericas\.com$/, '');
    // Garantir formato correto
    if (!userEmail.includes('@')) {
      userEmail = userEmail + '@ideatoreamericas.com';
    }
    
    console.log('✅ Extraído via padrão 1:', { domain, userEmail, fileId });
    return { success: true, domain, userEmail, fileId, type: 'share-link' };
  }
  
  // Padrão 2: Link com sourcedoc
  const pattern2 = /sourcedoc=\{([A-F0-9-]+)\}/i;
  const match2 = shareUrl.match(pattern2);
  if (match2) {
    return { success: true, fileId: match2[1], type: 'resource-link' };
  }
  
  // Padrão 3: Apenas o ID no caminho
  const pattern3 = /\/([A-Za-z0-9_-]{30,})/;
  const match3 = shareUrl.match(pattern3);
  if (match3) {
    return { success: true, fileId: match3[1], type: 'path-id' };
  }
  
  console.error('❌ Nenhum padrão reconhecido');
  return { success: false, error: 'Link não reconhecido' };
}

function extractUserEmail(shareUrl) {
  const match = shareUrl.match(/\/personal\/([^\/]+)\//);
  if (match) {
    let email = match[1].replace(/_/g, '.');
    email = email.replace(/\.ideatoreamericas\.com$/, '');
    if (!email.includes('@')) {
      email = email + '@ideatoreamericas.com';
    }
    return email;
  }
  return null;
}

module.exports = {
  extractSharePointInfo,
  extractUserEmail
};