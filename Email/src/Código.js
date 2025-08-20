const SHEET_ID = '1k0ytrIaumadc4Dfp29i5KSdqG93RR2GXMMwBd96jXdQ'; // O MESMO ID da sua planilha
const EMAIL_QUEUE_SHEET_NAME = 'EmailQueue';
// URL base para consulta de protocolo (ajuste conforme o ambiente)
const consultaUrlBase = 'https://script.google.com/macros/s/AKfycbzUiUkAP9XQ3gCo0vOHswNt78jV-SJpx_RulNzgDh6G680XTx8VEA52VA_CdyDd86erGg/exec';

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
    let totalEmails = data.length;

    // HTML de progresso
    let progressHtml = `
      <div id="progress-container" style="max-width:400px;margin:60px auto 0 auto;padding:32px 24px;background:#f7f7f7;border-radius:12px;box-shadow:0 2px 12px #0001;text-align:center;">
        <div style="font-size:22px;font-weight:bold;color:#00796b;margin-bottom:18px;">Enviando emails...</div>
        <div id="progress-bar" style="background:#e0e0e0;border-radius:8px;height:24px;width:100%;overflow:hidden;margin-bottom:12px;">
          <div id="progress-fill" style="background:#00796b;height:100%;width:0%;transition:width 0.3s;"></div>
        </div>
        <div id="progress-text" style="font-size:16px;color:#555;">0 de ${totalEmails} enviados</div>
      </div>
      <script>
        function updateProgress(sent, total) {
          var percent = Math.round((sent/total)*100);
          document.getElementById('progress-fill').style.width = percent + '%';
          document.getElementById('progress-text').innerText = sent + ' de ' + total + ' enviados';
        }
      </script>
    `;

    // Mostra progresso inicial
    let output = HtmlService.createHtmlOutput(progressHtml);
    output.setTitle('Envio de Emails');
    // Não é possível atualizar dinamicamente do lado do servidor, mas o usuário verá a tela de progresso

    if (data.length > 0 && data[0][0] !== "") {
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const protocolo = row[1], nomeContribuinte = row[2], emailContribuinte = row[3], novoStatus = row[4], observacao = row[5];

  const assunto = `Atualização do seu Protocolo: ${protocolo}`;
  const consultaUrl = `${consultaUrlBase}?page=consulta&protocolo=${encodeURIComponent(protocolo)}`;
  const corpo = `
<p>Prezado(a) ${nomeContribuinte},</p>
<p>Houve uma atualização no seu pedido de Análise de Prescrição (protocolo <strong>${protocolo}</strong>).</p>
<p><strong>Novo Status:</strong> ${novoStatus}</p>
<p><strong>Observação do Atendente:</strong><br/><i>${observacao}</i></p>
<div style="margin:32px 0 24px 0;text-align:center;">
  <a href="${consultaUrl}" style="display:inline-block;padding:20px 40px;background:#00796b;color:#fff;font-weight:bold;text-decoration:none;border-radius:8px;font-size:22px;box-shadow:0 2px 8px #0002;letter-spacing:1px;">Consultar Protocolo</a>
</div>
<div style="background:#ffe082;color:#b26a00;font-weight:bold;padding:12px 18px;border-radius:6px;margin:24px 0 16px 0;text-align:center;font-size:15px;">
  Por favor, não responda a este e-mail. Esta caixa não é monitorada.
</div>
<p style="margin-top:32px;line-height:1.6;">
Atenciosamente,<br>
Procuradoria-Geral do Estado do Pará<br>
Núcleo de Cobrança Administrativa - NCA
</p>`;

        // Envia email
        MailApp.sendEmail({ to: emailContribuinte, subject: assunto, htmlBody: corpo, name: "PGE - Atendimento" });
        emailsSent++;
        // Não é possível atualizar o HTML do lado do cliente a cada envio, mas o usuário verá a barra de progresso inicial
      }
      dataRange.clearContent();
    }

    lock.releaseLock();
    // Mensagem final
    return HtmlService.createHtmlOutput(`
      <div style="max-width:400px;margin:60px auto 0 auto;padding:32px 24px;background:#f7f7f7;border-radius:12px;box-shadow:0 2px 12px #0001;text-align:center;">
        <div style="font-size:22px;font-weight:bold;color:#00796b;margin-bottom:18px;">Envio concluído!</div>
        <div style="font-size:18px;color:#333;margin-bottom:12px;">${emailsSent} email(s) enviados com sucesso!</div>
        <button onclick="window.close()" style="margin-top:18px;padding:12px 32px;font-size:16px;background:#00796b;color:#fff;border:none;border-radius:6px;cursor:pointer;">Fechar</button>
      </div>
      <script>setTimeout(function(){ window.close(); }, 4000);</script>
    `);

  } catch (err) {
    Logger.log("Erro no processo de envio de email: " + err.message);
    return HtmlService.createHtmlOutput("<p>Ocorreu um erro ao processar a fila de emails.</p>");
  }
}
