function verificarVolumetria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Respostas ao formulário 1");
  const data = sheet.getDataRange().getValues();

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const diaSemana = hoje.getDay();
  // Ignora finais de semana
  if (diaSemana === 0 || diaSemana === 6) return;

  // Grupos de e-mails (exemplo)
  const ITAPEVA_EMAILS = [
    "gerencia.itapeva@empresa.com"
  ];

  const EXTREMA_EMAILS = [
    "gerencia.extrema@empresa.com"
  ];

  const LIDERES_EMAILS = [
    "lideranca.logistica@empresa.com"
  ];

  let itapeva = null;
  let extrema = null;

  // Busca a ÚLTIMA resposta DO DIA para cada CD
  for (let i = data.length - 1; i > 0; i--) {
    const timestamp = new Date(data[i][0]);
    timestamp.setHours(0, 0, 0, 0);

    if (timestamp.getTime() !== hoje.getTime()) continue;

    const cd = data[i][1];
    const responsavel = data[i][3];
    const volume = Number(data[i][4]); // volume 0 é válido
    const fotos = data[i][5] || "";
    const obs = data[i][6] || "";

    if (cd === "Itapeva" && !itapeva) {
      itapeva = { volume, responsavel, fotos, obs };
    }

    if (cd === "Extrema" && !extrema) {
      extrema = { volume, responsavel, fotos, obs };
    }
  }

  // ALERTAS — apenas se NÃO houver resposta no dia
  if (!itapeva) {
    MailApp.sendEmail({
      to: ITAPEVA_EMAILS.join(","),
      cc: LIDERES_EMAILS.join(","),
      subject: "⚠️ Alerta — Itapeva não enviou a volumetria",
      htmlBody: "<p>O CD de Itapeva ainda não preencheu a volumetria do dia.</p>"
    });
  }

  if (!extrema) {
    MailApp.sendEmail({
      to: EXTREMA_EMAILS.join(","),
      cc: LIDERES_EMAILS.join(","),
      subject: "⚠️ Alerta — Extrema não enviou a volumetria",
      htmlBody: "<p>O CD de Extrema ainda não preencheu a volumetria do dia.</p>"
    });
  }

  // CONSOLIDADO — enviado mesmo quando volume é 0
  if (itapeva && extrema) {
    const html = `
      <p><b>📦 Volumetria Consolidada do Dia</b></p>

      <p><b>Itapeva</b><br>
      Volumes: ${itapeva.volume}<br>
      Responsável: ${itapeva.responsavel}<br>
      <a href="${itapeva.fotos}">Fotos</a><br>
      Observações: ${itapeva.obs}</p>

      <p><b>Extrema</b><br>
      Volumes: ${extrema.volume}<br>
      Responsável: ${extrema.responsavel}<br>
      <a href="${extrema.fotos}">Fotos</a><br>
      Observações: ${extrema.obs}</p>

      <p>Atenciosamente,<br>Equipe Logística</p>
    `;

    MailApp.sendEmail({
      to: LIDERES_EMAILS.join(","),
      subject: "📦 Volumetria do dia — Consolidada",
      htmlBody: html
    });
  }
}