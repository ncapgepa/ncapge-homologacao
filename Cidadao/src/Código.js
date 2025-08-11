// ID da sua Planilha Google - VERIFIQUE SE ESTÁ CORRETO
const SHEET_ID = '1k0ytrIaumadc4Dfp29i5KSdqG93RR2GXMMwBd96jXdQ'; 
const REQUESTS_SHEET_NAME = 'Pedidos Prescrição';
const DRIVE_FOLDER_NAME = 'Documentos prescricao (homologacao)';

// URL base para consulta do protocolo (ajuste conforme o ambiente)
const CONSULTA_URL_BASE = 'https://script.google.com/macros/s/AKfycbzVUDExZbAyLYVQ-8CAbAg3JjKA3cQ1NVv60P3-c9F2HC8Gtvkr4wb1uxjgrf65NZF7';

/**
 * Função principal que serve as páginas públicas.
 */
function doGet(e) {
  var page = e.parameter && e.parameter.page ? e.parameter.page : 'cidadao';
  
  if (page === 'consulta') {
    return HtmlService.createTemplateFromFile('consulta').evaluate().setTitle('Consulta de Protocolo');
  } else {
    // A página padrão é a do cidadão
    return HtmlService.createTemplateFromFile('cidadao').evaluate().setTitle('Análise de Prescrição de Dívida Ativa');
  }
}

/**
 * ATUALIZADO: Processa o formulário, agora com protocolo dinâmico.
 */
function submitForm(formObject) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`A planilha com o nome "${REQUESTS_SHEET_NAME}" não foi encontrada.`);
    }

    const submittedCDAs = Array.isArray(formObject['cda[]']) ? formObject['cda[]'] : [formObject['cda[]']];
    const duplicateCheck = findDuplicateCDAs(sheet, submittedCDAs);
    if (duplicateCheck.isDuplicate) {
      return { erro: `A CDA nº ${duplicateCheck.cda} já existe em uma solicitação com status "${duplicateCheck.status}". Não é possível criar um novo pedido.` };
    }

    // --- INÍCIO DA ALTERAÇÃO ---
    const lastRow = sheet.getLastRow();
    const nextNumber = lastRow;
    const currentYear = new Date().getFullYear(); // Pega o ano atual
    const protocolo = `PGE-PRESC-${currentYear}-${String(nextNumber).padStart(4, '0')}`;
    // --- FIM DA ALTERAÇÃO ---

    const nomeSolicitante = formObject.nome;

    let driveFolder = getOrCreateFolder(DRIVE_FOLDER_NAME);
    let submissionFolder = driveFolder.createFolder(`${protocolo} - ${nomeSolicitante}`);
    let outrosDocumentosLista = [];
    
    // Processa todos os ficheiros enviados, incluindo os dinâmicos
    for (let key in formObject) {
      if (key.startsWith('doc_') && formObject[key] && typeof formObject[key].getName === 'function') {
        let fileBlob = formObject[key];
        let file = submissionFolder.createFile(fileBlob);
        
        // Se for um documento "outro", associa a sua descrição
        if (key.startsWith('doc_outro_')) {
          const fileIndex = key.split('_').pop();
          const descKey = 'desc_outro_' + fileIndex;
          const description = formObject[descKey] || 'Documento sem descrição';
          outrosDocumentosLista.push({ descricao: description, nomeArquivo: file.getName(), url: file.getUrl() });
        }
      }
    }
    const folderUrl = submissionFolder.getUrl();
    const cdasString = submittedCDAs.join(', ');

    // Prepara os novos campos para a planilha
    let nomeRepresentado = formObject.nomeRepresentado || '';
    let cpfCnpjRepresentado = formObject.cpfCnpjRepresentado || '';
    const tipoRepresentante = formObject.tipoRepresentante || '';
    const tipoDocumentoRepresentante = formObject.tipoDocumentoRepresentante || '';
    const numeroDocumentoRepresentante = formObject.numeroDocumentoRepresentante || '';

    // Se não houver representante, o nomeRepresentado recebe o nome do solicitante e o CPF/CNPJ do titular vai para a coluna O
    if (!tipoRepresentante && !tipoDocumentoRepresentante && !numeroDocumentoRepresentante) {
      nomeRepresentado = nomeSolicitante;
      cpfCnpjRepresentado = formObject.cpfCnpjTitular || '';
    }
    // Se houver representante, grava o campo do representado normalmente, mas também grava o cpfCnpjTitular na coluna O
    else if (formObject.cpfCnpjTitular) {
      cpfCnpjRepresentado = formObject.cpfCnpjRepresentado || formObject.cpfCnpjTitular;
    }
    
    const newRow = [
      protocolo,                      // A
      new Date(),                     // B
      nomeSolicitante,                // C (Sempre quem preenche, seja o titular ou representante)
      formObject.email,               // D
      formObject.telefone,            // E
      formObject.tipo,                // F (Refere-se ao titular/representado)
      cdasString,                     // G
      folderUrl,                      // H
      'Novo',                         // I
      '',                             // J (AtendenteResp)
      `Pedido criado em ${new Date().toLocaleString()}`, // K (Historico)
      '',                             // L (DataEncerramento)
      '',                             // M (ATTUS/SAJ) - Coluna vazia por enquanto
      nomeRepresentado,               // N
      cpfCnpjRepresentado,            // O (Agora sempre recebe o CPF/CNPJ do titular, mesmo sem representante)
      tipoRepresentante,              // P (NOVO)
      tipoDocumentoRepresentante,     // Q (NOVO)
      numeroDocumentoRepresentante,   // R (NOVO)
      JSON.stringify(outrosDocumentosLista) // S (Lista de outros documentos)
    ];
    
    sheet.appendRow(newRow);

    sendConfirmationEmail(protocolo, formObject.email, nomeSolicitante);
    return { protocolo: protocolo };

  } catch (error) {
    Logger.log(error.toString());
    return { erro: error.toString() };
  }
}

/**
 * NOVA FUNÇÃO: Procura por CDAs duplicadas na planilha.
 * @param {Sheet} sheet O objeto da planilha de pedidos.
 * @param {Array<string>} cdasToCheck A lista de CDAs enviadas pelo utilizador.
 * @returns {object} Um objeto indicando se há duplicata e qual a CDA.
 */
function findDuplicateCDAs(sheet, cdasToCheck) {
  const cdaColumnIndex = 7; // G = 7
  const statusColumnIndex = 9; // I = 9
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // Não há dados além do cabeçalho, então não há duplicatas
    return { isDuplicate: false };
  }
  const range = sheet.getRange(2, cdaColumnIndex, lastRow - 1, 3); // Lê as colunas G, H, I
  const data = range.getValues();

  // Cria um mapa de CDAs existentes para o seu status
  const existingCDAs = new Map();
  data.forEach(row => {
    const cdasInSheet = row[0].split(',').map(cda => cda.trim());
    const status = row[statusColumnIndex - cdaColumnIndex]; // Índice relativo ao range lido
    
    // Considera apenas pedidos que não foram indeferidos.
    if (status.toLowerCase() !== 'indeferido') {
      cdasInSheet.forEach(cda => {
        if (cda) existingCDAs.set(cda, status);
      });
    }
  });

  // Verifica cada CDA enviada contra o mapa de existentes
  for (const cda of cdasToCheck) {
    if (existingCDAs.has(cda.trim())) {
      return { 
        isDuplicate: true, 
        cda: cda.trim(),
        status: existingCDAs.get(cda.trim())
      };
    }
  }

  return { isDuplicate: false };
}

/**
 * Envia email de confirmação para o cidadão.
 */
function sendConfirmationEmail(protocolo, destinatario, nome) {
  const assunto = `Confirmação de Recebimento - Protocolo ${protocolo}`;
  // Adiciona o protocolo como parâmetro na URL
  const consultaUrl = `${CONSULTA_URL_BASE}/exec?page=consulta&protocolo=${encodeURIComponent(protocolo)}`;
  const corpo = `
    <p>Prezado(a) ${nome},</p>
    <p>A sua solicitação de Análise de Prescrição de Dívida Ativa foi recebida com sucesso.</p>
    <p>O seu número de protocolo é: <strong>${protocolo}</strong></p>
    <p>Guarde este número para futuras consultas sobre o andamento do seu pedido.</p>
    <p>
      <a href="${consultaUrl}" style="display:inline-block;padding:12px 24px;background:#004d40;color:#fff;text-decoration:none;border-radius:4px;font-weight:bold;">Consultar andamento do pedido</a>
    </p>
    <p style="color:#888;font-size:0.95em;">Por favor, não responda a este e-mail. Esta caixa não é monitorada.</p>
    <p>Atenciosamente,<br>Procuradoria-Geral do Estado do Pará</p>
  `;
  try {
    MailApp.sendEmail({ to: destinatario, subject: assunto, htmlBody: corpo });
  } catch (e) {
    Logger.log(`Falha ao enviar email para ${destinatario}. Erro: ${e.message}`);
  }
}

/**
 * Encontra ou cria a pasta no Google Drive.
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * Consulta um protocolo para o cidadão.
 */
function consultarProtocolo(protocolo) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(REQUESTS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === protocolo) {
        return {
          protocolo: data[i][0],
          data: data[i][1].toLocaleString(),
          status: data[i][8] // Apenas informações públicas
        };
      }
    }
    return { erro: 'Protocolo não encontrado.' };
  } catch (e) {
    return { erro: 'Ocorreu um erro ao consultar.' };
  }
}