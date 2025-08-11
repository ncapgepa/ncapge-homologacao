const SHEET_ID = '1Cnb-tqz1b5uvaW4rK3rlGjlYW3QJGEaz9sKPXCzEcxY';
const REQUESTS_SHEET_NAME = 'Pedidos Prescrição';
const ACCESS_SHEET_NAME = 'Acessos';
const EMAIL_QUEUE_SHEET_NAME = 'EmailQueue';

/**
 * Função principal que serve o painel do atendente.
 */
function doGet(e) {
  const accessInfo = checkUserAccess();
  // Mostra sempre o email detectado, mesmo se não tiver acesso
  if (accessInfo.hasAccess) {
    const template = HtmlService.createTemplateFromFile('painel');
    template.userName = accessInfo.nome;
    template.userEmail = accessInfo.email;
    template.userRole = accessInfo.role;
    return template.evaluate().setTitle('Painel do Atendente');
  } else {
    return HtmlService.createHtmlOutput(
      '<h1>Acesso Negado</h1><p>O seu email (<strong>' + 
      (accessInfo.email || 'Não identificado') + 
      '</strong>) não tem permissão para aceder a esta página. Por favor, contacte o administrador do sistema.</p>'
    );
  }
}

/**
 * Verifica se o utilizador atual tem acesso ao sistema.
 */
function checkUserAccess() {
  const userEmail = Session.getEffectiveUser().getEmail();
  if (!userEmail) return { hasAccess: false, nome: null, email: null, role: null };
  try {
    const accessSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ACCESS_SHEET_NAME);
    const data = accessSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].trim().toLowerCase() === userEmail.trim().toLowerCase()) {
        return { hasAccess: true, nome: data[i][0], email: userEmail, role: data[i][2].trim() };
      }
    }
    return { hasAccess: false, nome: null, email: userEmail, role: null };
  } catch (e) {
    Logger.log('Erro ao verificar acesso para ' + userEmail + ': ' + e.message);
    return { hasAccess: false, nome: null, email: userEmail, role: null };
  }
}

/**
 * Retorna todos os pedidos para o painel.
 */
function getRequests() {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess) throw new Error('Acesso não autorizado.');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove cabeçalho
  return data.map(row => ({
    protocolo: row[0],
    data: row[1].toLocaleString(),
    nome: row[2],
    status: row[8]
  }));
}

/**
 * Consulta TODOS os detalhes de um protocolo para o atendente.
 */
function consultarProtocoloCompleto(protocolo) {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess) throw new Error('Acesso não autorizado.');
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === protocolo) {
      return {
        protocolo: data[i][0], data: data[i][1].toLocaleString(), nome: data[i][2],
        email: data[i][3], telefone: data[i][4], tipo: data[i][5], cdas: data[i][6],
        linkDocumentos: data[i][7], status: data[i][8], atendente: data[i][9], historico: data[i][10],
        attusSaj: data[i][12] // NOVO: Lê da coluna M (índice 12)
      };
    }
  }
  return { erro: 'Protocolo não encontrado.' };
}

/**
 * Atualiza o status de um pedido.
 */
function updateStatus(protocolo, status, historico, attusSaj) { 
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess) throw new Error('Acesso negado para esta operação.');
  const atendente = accessInfo.nome;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === protocolo) {
      const row = i + 1;
      const nomeContribuinte = data[i][2];
      const emailContribuinte = data[i][3];
      const statusAntigo = data[i][8];
      sheet.getRange(row, 9).setValue(status);
      sheet.getRange(row, 10).setValue(atendente);
      const oldHistorico = sheet.getRange(row, 11).getValue();
      const newHistoricoEntry = `\n${new Date().toLocaleString()} - ${atendente}: ${historico}`;
      sheet.getRange(row, 11).setValue(oldHistorico + newHistoricoEntry);
      sheet.getRange(row, 13).setValue(attusSaj); 
      if (status === 'Deferido' || status === 'Indeferido') {
        sheet.getRange(row, 12).setValue(new Date());
      }
      if (status !== statusAntigo) {
        const emailQueueSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_QUEUE_SHEET_NAME);
        emailQueueSheet.appendRow([
          new Date(), protocolo, nomeContribuinte, emailContribuinte, status, historico
        ]);
        return { success: true, needsRedirect: true };
      }
      return { success: true, needsRedirect: false };
    }
  }
  return { success: false };
}

function prepareEmailAndCreateTrigger(protocolo, nomeContribuinte, emailContribuinte, novoStatus, observacao) {
  try {
    const emailQueueSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_QUEUE_SHEET_NAME);
    emailQueueSheet.appendRow([
      new Date(), protocolo, nomeContribuinte, emailContribuinte, novoStatus, observacao
    ]);
    ScriptApp.newTrigger('processEmailQueue')
      .timeBased()
      .after(1)
      .create();
  } catch(e) {
    Logger.log("Erro ao preparar o email para a fila: " + e.message);
  }
}

function processEmailQueue(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (e) {
    Logger.log('Não foi possível obter o bloqueio. Outro processo pode estar em execução.');
    return;
  }
  const emailQueueSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_QUEUE_SHEET_NAME);
  if (emailQueueSheet.getLastRow() < 2) {
    lock.releaseLock();
    return;
  }
  const dataRange = emailQueueSheet.getRange("A2:F" + emailQueueSheet.getLastRow());
  const data = dataRange.getValues();
  if (data.length > 0 && data[0][0] !== "") {
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const protocolo = row[1], nomeContribuinte = row[2], emailContribuinte = row[3], novoStatus = row[4], observacao = row[5];
      const assunto = `Atualização do seu Protocolo: ${protocolo}`;
      const corpo = `<p>Prezado(a) ${nomeContribuinte},</p><p>Houve uma atualização no seu pedido de Análise de Prescrição (protocolo <strong>${protocolo}</strong>).</p><p><strong>Novo Status:</strong> ${novoStatus}</p><p><strong>Observação do Atendente:</strong><br/><i>${observacao}</i></p><p>Você pode consultar o seu pedido a qualquer momento.</p><p>Atenciosamente,<br>Equipe de Atendimento</p>`;
      try {
        MailApp.sendEmail({ to: emailContribuinte, subject: assunto, htmlBody: corpo, name: "PGE - Atendimento" });
        Logger.log(`Email enviado para ${emailContribuinte} (Protocolo: ${protocolo})`);
      } catch (err) {
        Logger.log(`Falha ao enviar email para ${emailContribuinte}. Erro: ${err.message}`);
      }
    }
    dataRange.clearContent();
  }
  if (e && e.triggerUid) {
    const allTriggers = ScriptApp.getProjectTriggers();
    for (const trigger of allTriggers) {
      if (trigger.getUniqueId() === e.triggerUid) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
  }
  lock.releaseLock();
}

/**
 * Retorna a lista de utilizadores.
 */
function getUsers() {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess || accessInfo.role !== 'admin') {
    throw new Error('Acesso negado.');
  }
  const accessSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ACCESS_SHEET_NAME);
  const data = accessSheet.getDataRange().getValues();
  data.shift();
  return data.map(row => ({ nome: row[0], email: row[1], role: row[2] }));
}

/**
 * Adiciona ou atualiza um utilizador.
 */
function addOrUpdateUser(nome, email, role) {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess || accessInfo.role !== 'admin') {
    throw new Error('Acesso negado.');
  }
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ACCESS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 1).setValue(nome);
      sheet.getRange(i + 1, 3).setValue(role);
      return { status: 'success', message: 'Utilizador atualizado.' };
    }
  }
  sheet.appendRow([nome, email, role]);
  return { status: 'success', message: 'Utilizador adicionado.' };
}

/**
 * Remove um utilizador.
 */
function removeUser(email) {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess || accessInfo.role !== 'admin') {
    throw new Error('Acesso negado.');
  }
  if (email.toLowerCase() === accessInfo.email.toLowerCase()) {
    throw new Error('Não pode remover-se a si próprio.');
  }
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ACCESS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][1].toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { status: 'success', message: 'Utilizador removido.' };
    }
  }
  throw new Error('Utilizador não encontrado.');
}

/**
 * NOVA FUNÇÃO: Atualiza todos os dados de um pedido, incluindo os dados do requerente.
 */
function updateRequestData(dataObject) {
  const accessInfo = checkUserAccess();
  if (!accessInfo.hasAccess) throw new Error('Acesso negado.');

  const atendente = accessInfo.nome;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Pega os cabeçalhos para encontrar os índices das colunas

  const protocolIndex = headers.indexOf('Protocolo');
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][protocolIndex] === dataObject.protocolo) {
      rowIndex = i + 2; // +1 pelo índice 0, +1 pelo cabeçalho removido
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Protocolo não encontrado.');

  const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  let historyUpdates = [];
  let needsRedirect = false;

  // Função auxiliar para verificar e registar alterações
  function checkAndUpdate(fieldName, newValue, columnIndex) {
    const oldValue = rowValues[columnIndex];
    if (String(oldValue).trim() !== String(newValue).trim()) {
      sheet.getRange(rowIndex, columnIndex + 1).setValue(newValue);
      historyUpdates.push(`${fieldName} alterado de "${oldValue}" para "${newValue}".`);
    }
  }

  // Verifica cada campo editável
  checkAndUpdate('NomeSolicitante', dataObject.nome, headers.indexOf('NomeSolicitante'));
  checkAndUpdate('Email', dataObject.email, headers.indexOf('Email'));
  checkAndUpdate('Telefone', dataObject.telefone, headers.indexOf('Telefone'));
  checkAndUpdate('TipoPessoa', dataObject.tipo, headers.indexOf('TipoPessoa'));
  checkAndUpdate('CDAs', dataObject.cdas, headers.indexOf('CDAs'));
  checkAndUpdate('ATTUS/SAJ', dataObject.attusSaj, headers.indexOf('ATTUS/SAJ'));

  const statusIndex = headers.indexOf('Status');
  const oldStatus = rowValues[statusIndex];
  if (oldStatus !== dataObject.status) {
    sheet.getRange(rowIndex, statusIndex + 1).setValue(dataObject.status);
    historyUpdates.push(`Status alterado de "${oldStatus}" para "${dataObject.status}".`);
    // Adiciona à fila de envio de email se o status mudou
    const emailQueueSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_QUEUE_SHEET_NAME);
    emailQueueSheet.appendRow([
      new Date(), dataObject.protocolo, dataObject.nome, dataObject.email, dataObject.status, dataObject.observacao
    ]);
    needsRedirect = true;
  }

  // Atualiza Histórico e Atendente se houver qualquer alteração
  if (historyUpdates.length > 0 || dataObject.observacao) {
    const historyIndex = headers.indexOf('Historico');
    const atendenteIndex = headers.indexOf('AtendenteResp');
    const oldHistorico = rowValues[historyIndex] || '';
    let newHistoricoEntry = `\n--- ATUALIZAÇÃO: ${new Date().toLocaleString()} - ${atendente} ---\n`;
    if(dataObject.observacao) {
      newHistoricoEntry += `Observação: ${dataObject.observacao}\n`;
    }
    if (historyUpdates.length > 0) {
      newHistoricoEntry += `Alterações de Dados: \n- ${historyUpdates.join('\n- ')}`;
    }
    sheet.getRange(rowIndex, historyIndex + 1).setValue(oldHistorico + newHistoricoEntry);
    sheet.getRange(rowIndex, atendenteIndex + 1).setValue(atendente);
    if (dataObject.status === 'Deferido' || dataObject.status === 'Indeferido') {
      const dataEncerramentoIndex = headers.indexOf('DataEncerramento');
      sheet.getRange(rowIndex, dataEncerramentoIndex + 1).setValue(new Date());
    }
  }
  return { success: true, needsRedirect: needsRedirect };
}

/**
 * NOVA FUNÇÃO: Gera um PDF com os detalhes do protocolo.
 */
function generateProtocolPdf(protocolo) {
  const dados = consultarProtocoloCompleto(protocolo);
  if (dados.erro) {
    throw new Error('Não foi possível gerar o PDF: ' + dados.erro);
  }
  const htmlContent = `
    <html>
      <head>
        <style>
          body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 11px; }
          h1 { color: #333; border-bottom: 2px solid #ccc; padding-bottom: 5px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background-color: #f2f2f2; }
          pre { background-color: #f8f8f8; padding: 10px; border: 1px solid #eee; white-space: pre-wrap; word-wrap: break-word; }
        </style>
      </head>
      <body>
        <h1>Relatório do Protocolo: ${dados.protocolo}</h1>
        <h3>Dados do Requerente</h3>
        <table>
          <tr><th>Nome</th><td>${dados.nome}</td></tr>
          <tr><th>E-mail</th><td>${dados.email}</td></tr>
          <tr><th>Telefone</th><td>${dados.telefone}</td></tr>
          <tr><th>Tipo de Requerente</th><td>${dados.tipo}</td></tr>
        </table>
        <h3>Dados do Pedido</h3>
        <table>
          <tr><th>Data da Solicitação</th><td>${dados.data}</td></tr>
          <tr><th>Status Atual</th><td>${dados.status}</td></tr>
          <tr><th>Nº Processo ATTUS/SAJ</th><td>${dados.attusSaj || 'Não informado'}</td></tr>
          <tr><th>CDAs</th><td>${dados.cdas}</td></tr>
        </table>
        <h3>Histórico Completo</h3>
        <pre>${dados.historico || 'Nenhum histórico registado.'}</pre>
        <p style="text-align:center; color:#777; font-size:9px; margin-top: 30px;">
          Documento gerado pelo SisNCA em ${new Date().toLocaleString()}
        </p>
      </body>
    </html>
  `;
  const blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(`Protocolo_${dados.protocolo}.pdf`);
  return {
    fileName: blob.getName(),
    contentType: blob.getContentType(),
    fileContent: Utilities.base64Encode(blob.getBytes())
  };
}