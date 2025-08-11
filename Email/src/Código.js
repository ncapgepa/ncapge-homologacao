const SHEET_ID = '1k0ytrIaumadc4Dfp29i5KSdqG93RR2GXMMwBd96jXdQ'; // O MESMO ID da sua planilha
const EMAIL_QUEUE_SHEET_NAME = 'EmailQueue';

/**
 * Esta é a única função deste projeto. Ela é executada assim que a página é aberta.
 */
function doGet(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000); 

    const emailQueueSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAIL_QUEUE_SHEET_NAME);
    
    // Verifica se há linhas para processar
    if (emailQueueSheet.getLastRow() < 2) {
      lock.releaseLock();
      return HtmlService.createHtmlOutput('<p>Nenhum email na fila para ser enviado. Pode fechar esta janela.</p>');
    }

    const dataRange = emailQueueSheet.getRange("A2:F" + emailQueueSheet.getLastRow());
    const data = dataRange.getValues();
    let emailsSent = 0;

    if (data.length > 0 && data[0][0] !== "") {
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const protocolo = row[1], nomeContribuinte = row[2], emailContribuinte = row[3], novoStatus = row[4], observacao = row[5];
        
        const assunto = `Atualização do seu Protocolo: ${protocolo}`;
        const corpo = `<p>Prezado(a) ${nomeContribuinte},</p><p>Houve uma atualização no seu pedido de Análise de Prescrição (protocolo <strong>${protocolo}</strong>).</p><p><strong>Novo Status:</strong> ${novoStatus}</p><p><strong>Observação do Atendente:</strong><br/><i>${observacao}</i></p><p>Você pode consultar o seu pedido a qualquer momento.</p><p>Atenciosamente,<br>Equipe de Atendimento</p>`;
        
        // Como este script é executado como "USER_DEPLOYING", o email sairá da conta do proprietário.
        MailApp.sendEmail({ to: emailContribuinte, subject: assunto, htmlBody: corpo, name: "PGE - Atendimento" });
        emailsSent++;
      }
      dataRange.clearContent();
    }

    lock.releaseLock();
    return HtmlService.createHtmlOutput(`<p>${emailsSent} email(s) enviados com sucesso! Pode fechar esta janela.</p><script>setTimeout(function(){ window.close(); }, 3000);</script>`);

  } catch (err) {
    Logger.log("Erro no processo de envio de email: " + err.message);
    return HtmlService.createHtmlOutput("<p>Ocorreu um erro ao processar a fila de emails.</p>");
  }
}
