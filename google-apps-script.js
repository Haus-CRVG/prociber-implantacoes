// ============================================================
// GOOGLE APPS SCRIPT — Controle de Implantação de Backup
// ============================================================
// COMO INSTALAR:
//  1. Abra seu Google Sheets > Extensões > Apps Script
//  2. Apague o código padrão e cole este arquivo inteiro
//  3. Clique em "Implantar" > "Nova implantação"
//  4. Tipo: Aplicativo da Web
//     Executar como: Eu mesmo
//     Quem pode acessar: Qualquer pessoa
//  5. Copie a URL gerada e cole no formulario-implantacao.html (variável SCRIPT_URL)
// ============================================================

const SHEET_NAME = 'Implantações';
const ANALISTAS = ['', 'João Silva', 'Maria Souza', 'Carlos Lima']; // Edite com seus analistas

// ── Recebe dados do formulário (POST) ─────────────────────
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  // Cria a aba se não existir e monta o cabeçalho
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupSheet(sheet);
  }

  const now = new Date();
  const row = [
    now,                     // A – Data de entrada
    data.empresa  || '',     // B – Cliente (Empresa)
    data.contato  || '',     // C – Contato
    data.telefone || '',     // D – Telefone
    data.representante || '',// E – Representante
    data.status   || 'Aguardando aceite', // F – Status
    '',                      // G – Analista responsável (preenchido internamente)
    '',                      // H – Data de conclusão (automática via onEdit)
    data.observacoes || '',  // I – Observações
  ];

  sheet.appendRow(row);

  // Formata a linha recém-adicionada
  const lastRow = sheet.getLastRow();
  formatNewRow(sheet, lastRow);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Detecta mudança de Status para preencher data de conclusão ──
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Coluna F (6) = Status
  if (col === 6 && row > 1) {
    const newStatus = e.range.getValue();
    const conclusaoCell = sheet.getRange(row, 8); // Coluna H = Data conclusão

    if (newStatus === 'Concluído') {
      if (!conclusaoCell.getValue()) {
        conclusaoCell.setValue(new Date());
        conclusaoCell.setNumberFormat('dd/MM/yyyy HH:mm');
      }
    } else {
      // Se voltar de Concluído, limpa a data
      conclusaoCell.clearContent();
    }

    // Aplica cor de fundo conforme status
    colorizeStatusRow(sheet, row, newStatus);
  }
}

// ── Instala o trigger onEdit como trigger instalável (rode 1x manualmente) ──
function instalarTrigger() {
  // Apaga triggers duplicados
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('✅ Trigger instalado com sucesso!');
}

// ── Configura cabeçalho e formatação inicial da planilha ──
function setupSheet(sheet) {
  const headers = [
    'Data de Entrada',
    'Cliente (Empresa)',
    'Contato',
    'Telefone',
    'Representante',
    'Status',
    'Analista Responsável',
    'Data de Conclusão',
    'Observações',
  ];

  sheet.appendRow(headers);

  // Estilo do cabeçalho
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a237e')
             .setFontColor('#ffffff')
             .setFontWeight('bold')
             .setFontSize(11)
             .setHorizontalAlignment('center');

  // Larguras das colunas
  sheet.setColumnWidth(1, 160); // Data entrada
  sheet.setColumnWidth(2, 220); // Empresa
  sheet.setColumnWidth(3, 180); // Contato
  sheet.setColumnWidth(4, 140); // Telefone
  sheet.setColumnWidth(5, 180); // Representante
  sheet.setColumnWidth(6, 160); // Status
  sheet.setColumnWidth(7, 180); // Analista
  sheet.setColumnWidth(8, 160); // Data conclusão
  sheet.setColumnWidth(9, 300); // Observações

  // Congela o cabeçalho
  sheet.setFrozenRows(1);

  // Validação de Status na coluna F
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Aguardando aceite', 'Em implantação', 'Concluído'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(statusRule);

  // Validação de Analista na coluna G
  if (ANALISTAS.length > 1) {
    const analistaRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(ANALISTAS, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('G2:G1000').setDataValidation(analistaRule);
  }
}

// ── Formata linha nova ──
function formatNewRow(sheet, row) {
  const range = sheet.getRange(row, 1, 1, 9);
  range.setFontSize(10);

  // Data de entrada com formato
  sheet.getRange(row, 1).setNumberFormat('dd/MM/yyyy HH:mm');

  // Cor alternada
  const bg = (row % 2 === 0) ? '#f8f9ff' : '#ffffff';
  range.setBackground(bg);

  // Aplica cor de status
  const status = sheet.getRange(row, 6).getValue();
  colorizeStatusRow(sheet, row, status);
}

// ── Coloriza a linha conforme o status ──
function colorizeStatusRow(sheet, row, status) {
  const statusCell = sheet.getRange(row, 6);
  switch (status) {
    case 'Aguardando aceite':
      statusCell.setBackground('#fff3cd').setFontColor('#856404');
      break;
    case 'Em implantação':
      statusCell.setBackground('#cfe2ff').setFontColor('#084298');
      break;
    case 'Concluído':
      statusCell.setBackground('#d1e7dd').setFontColor('#0a3622');
      break;
    default:
      statusCell.setBackground('#f8f9fa').setFontColor('#333333');
  }
  statusCell.setFontWeight('bold').setHorizontalAlignment('center');
}

// ── Cria menu personalizado na planilha ──
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Implantações')
    .addItem('📋 Configurar planilha', 'setupPlanilhaManual')
    .addItem('🔧 Instalar trigger de automação', 'instalarTrigger')
    .addToUi();
}

function setupPlanilhaManual() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    setupSheet(sheet);
    SpreadsheetApp.getUi().alert('✅ Aba "Implantações" criada e configurada!');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ A aba "Implantações" já existe.');
  }
}
