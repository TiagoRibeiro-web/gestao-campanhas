require('dotenv').config();
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { ClientSecretCredential } = require('@azure/identity');

async function testFile() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });

  const graphClient = Client.initWithMiddleware({ authProvider });

  const fileId = '269E2FE1-B5C1-4E0A-995F-4131F5149C8E';
  const userEmail = 'cristini.cordesco@ideatoreamericas.com';

  console.log('🔍 Testando acesso ao arquivo...\n');

  try {
    // Acessar drive da usuária
    const drive = await graphClient.api(`/users/${userEmail}/drive`).get();
    console.log('✅ Drive encontrado!');
    
    // Tentar acessar o arquivo pelo ID
    try {
      const file = await graphClient.api(`/users/${userEmail}/drive/items/${fileId}`).get();
      console.log('✅ Arquivo encontrado!');
      console.log(`   Nome: ${file.name}`);
      console.log(`   ID: ${file.id}`);
      console.log(`   Tamanho: ${file.size} bytes`);
      console.log(`   URL: ${file.webUrl}`);
    } catch (error) {
      console.log('❌ Erro ao acessar arquivo pelo ID:', error.message);
      console.log('\n💡 Isso significa que o aplicativo não tem permissão!');
      console.log('   A Cristini precisa compartilhar com:');
      console.log(`   ${process.env.CLIENT_ID}@agenciaideatore.onmicrosoft.com`);
    }
    
  } catch (error) {
    console.error('❌ Erro:', error.message);
  }
}

testFile();