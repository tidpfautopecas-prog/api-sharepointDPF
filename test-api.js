import fetch from 'node-fetch';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Configurações da API
const API_BASE = 'http://localhost:3001';

// Cores para logs
const colors = {
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  reset: '\x1b[0m',
  bold: '\x1b[1m'
};

function log(color, icon, message) {
  console.log(`${color}${icon} ${message}${colors.reset}`);
}

// Função para criar um PDF de teste em Base64
function createTestPDF() {
  // PDF mínimo válido em Base64
  const pdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj

4 0 obj
<<
/Length 44
>>
stream
BT
/F1 12 Tf
100 700 Td
(Teste PDF API SharePoint) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000206 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
300
%%EOF`;

  return Buffer.from(pdfContent).toString('base64');
}

// Teste 1: Verificar status da API
async function testStatus() {
  try {
    log(colors.blue, '🧪', 'Testando status da API...');
    
    const response = await fetch(`${API_BASE}/status`);
    const data = await response.json();
    
    if (response.ok) {
      log(colors.green, '✅', 'API está online');
      console.log('   📋 Configurações:', JSON.stringify(data.config, null, 2));
      return true;
    } else {
      log(colors.red, '❌', 'API não está respondendo');
      return false;
    }
  } catch (error) {
    log(colors.red, '❌', `Erro ao conectar com a API: ${error.message}`);
    log(colors.yellow, '⚠️', 'Certifique-se de que a API está rodando: npm start');
    return false;
  }
}

// Teste 2: Testar conexão com SharePoint
async function testConnection() {
  try {
    log(colors.blue, '🧪', 'Testando conexão com SharePoint...');
    
    const response = await fetch(`${API_BASE}/test-connection`);
    const data = await response.json();
    
    if (response.ok && data.success) {
      log(colors.green, '✅', 'Conexão com SharePoint funcionando');
      console.log('   🏢 Site:', data.siteInfo.name);
      console.log('   🔗 URL:', data.siteInfo.url);
      console.log('   📁 Status da pasta:', data.folderStatus);
      return true;
    } else {
      log(colors.red, '❌', `Erro na conexão: ${data.error}`);
      console.log('   💡 Detalhes:', data.details);
      return false;
    }
  } catch (error) {
    log(colors.red, '❌', `Erro no teste de conexão: ${error.message}`);
    return false;
  }
}

// Teste 3: Criar pasta Laudos se necessário
async function testCreateFolder() {
  try {
    log(colors.blue, '🧪', 'Verificando/criando pasta Laudos...');
    
    const response = await fetch(`${API_BASE}/create-folder`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      log(colors.green, '✅', 'Pasta Laudos criada/verificada com sucesso');
      console.log('   📁 Nome:', data.folder.name);
      console.log('   🔗 URL:', data.folder.url);
      return true;
    } else {
      // Se a pasta já existe, ainda é um sucesso
      if (data.details && data.details.includes('already exists')) {
        log(colors.green, '✅', 'Pasta Laudos já existe');
        return true;
      }
      log(colors.yellow, '⚠️', `Aviso na criação da pasta: ${data.error}`);
      return false;
    }
  } catch (error) {
    log(colors.red, '❌', `Erro ao criar pasta: ${error.message}`);
    return false;
  }
}

// Teste 4: Upload de PDF de teste
async function testUploadPDF() {
  try {
    log(colors.blue, '🧪', 'Testando upload de PDF...');
    
    const testPDFBase64 = createTestPDF();
    const fileName = `Teste_API_${new Date().toISOString().slice(0, 19).replace(/[:-]/g, '')}.pdf`;
    
    const response = await fetch(`${API_BASE}/upload-pdf`, {
      method: 'POST',
      headers: { 
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        fileName,
        fileBase64: testPDFBase64,
        ticketNumber: '#TESTE-001',
        ticketTitle: 'Teste de Upload da API',
        isReport: false
      })
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      log(colors.green, '✅', 'Upload de PDF realizado com sucesso!');
      console.log('   📄 Arquivo:', data.fileName);
      console.log('   📍 Local:', data.location);
      console.log('   ⏱️ Tempo:', data.uploadTime);
      console.log('   📊 Tamanho:', data.fileSize);
      console.log('   🔗 URL SharePoint:', data.sharePointUrl);
      return true;
    } else {
      log(colors.red, '❌', `Erro no upload: ${data.error}`);
      console.log('   💡 Detalhes:', data.details);
      if (data.troubleshooting) {
        console.log('   🔧 Soluções:');
        Object.entries(data.troubleshooting).forEach(([key, value]) => {
          console.log(`      - ${value}`);
        });
      }
      return false;
    }
  } catch (error) {
    log(colors.red, '❌', `Erro no teste de upload: ${error.message}`);
    return false;
  }
}

// Teste 5: Validar integração com o frontend
async function testFrontendIntegration() {
  try {
    log(colors.blue, '🧪', 'Testando integração com frontend...');
    
    // Simular dados que viriam do frontend
    const mockTicket = {
      numero: '#LDO-999',
      titulo: 'Teste de Integração Frontend',
      responsavel: 'Sistema de Testes',
      itens: [
        { numeroItem: 1, quantidade: 2, motivo: 'Teste', observacao: 'Integração' }
      ]
    };
    
    const testPDFBase64 = createTestPDF();
    const dataAtual = new Date();
    const dataFormatada = dataAtual.toLocaleDateString('pt-BR').replace(/\//g, '-');
    const horaFormatada = dataAtual.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }).replace(':', 'h') + 'min';
    const fileName = `Laudo_${mockTicket.numero.replace('#', '')}_${dataFormatada}_${horaFormatada}.pdf`;
    
    const response = await fetch(`${API_BASE}/upload-pdf`, {
      method: 'POST',
      headers: { 
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        fileName,
        fileBase64: testPDFBase64,
        ticketNumber: mockTicket.numero,
        ticketTitle: mockTicket.titulo,
        isReport: false
      })
    });
    
    const data = await response.json();
    
    if (response.ok && data.success) {
      log(colors.green, '✅', 'Integração com frontend funcionando!');
      console.log('   🎫 Laudo:', mockTicket.numero);
      console.log('   📄 Arquivo:', fileName);
      console.log('   📍 Salvo em:', data.location);
      return true;
    } else {
      log(colors.red, '❌', `Erro na integração: ${data.error}`);
      return false;
    }
  } catch (error) {
    log(colors.red, '❌', `Erro no teste de integração: ${error.message}`);
    return false;
  }
}

// Executar todos os testes
async function runAllTests() {
  console.log(`${colors.bold}${colors.blue}🚀 INICIANDO TESTES DA API SHAREPOINT${colors.reset}\n`);
  
  const tests = [
    { name: 'Status da API', fn: testStatus },
    { name: 'Conexão SharePoint', fn: testConnection },
    { name: 'Criar Pasta Laudos', fn: testCreateFolder },
    { name: 'Upload de PDF', fn: testUploadPDF },
    { name: 'Integração Frontend', fn: testFrontendIntegration }
  ];
  
  let passed = 0;
  let failed = 0;
  
  for (const test of tests) {
    console.log(`\n${colors.yellow}📋 Executando: ${test.name}${colors.reset}`);
    const result = await test.fn();
    
    if (result) {
      passed++;
    } else {
      failed++;
    }
    
    // Pequena pausa entre testes
    await new Promise(resolve => setTimeout(resolve, 1000));
  }
  
  // Resultado final
  console.log(`\n${colors.bold}📊 RESULTADO DOS TESTES${colors.reset}`);
  console.log(`${colors.green}✅ Passou: ${passed}${colors.reset}`);
  console.log(`${colors.red}❌ Falhou: ${failed}${colors.reset}`);
  
  if (failed === 0) {
    console.log(`\n${colors.bold}${colors.green}🎉 TODOS OS TESTES PASSARAM!${colors.reset}`);
    console.log(`${colors.green}✅ A API SharePoint está funcionando perfeitamente${colors.reset}`);
    console.log(`${colors.green}✅ Integração com o frontend está pronta${colors.reset}`);
    console.log(`${colors.green}✅ PDFs serão salvos automaticamente na pasta Laudos${colors.reset}`);
  } else {
    console.log(`\n${colors.bold}${colors.red}⚠️ ALGUNS TESTES FALHARAM${colors.reset}`);
    console.log(`${colors.yellow}💡 Verifique os erros acima e corrija antes de usar em produção${colors.reset}`);
  }
  
  console.log(`\n${colors.blue}📋 Para usar no sistema:${colors.reset}`);
  console.log(`   1. Certifique-se de que a API está rodando: ${colors.bold}npm start${colors.reset}`);
  console.log(`   2. Acesse o sistema de laudos no navegador`);
  console.log(`   3. Crie um novo laudo e gere o PDF`);
  console.log(`   4. O PDF será salvo automaticamente no SharePoint`);
}

// Executar testes
runAllTests().catch(error => {
  console.error(`${colors.red}💥 Erro crítico nos testes:${colors.reset}`, error);
  process.exit(1);
});