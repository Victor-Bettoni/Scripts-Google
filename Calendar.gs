function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 Calendar')
    .addItem('Abrir Criador', 'abrirSidebar')
    .addToUi();
}

function abrirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarUI')
    .setTitle('Criador de Eventos')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getCalendarios() {
  return CalendarApp.getAllOwnedCalendars().map(c => c.getName());
}

function getSelecaoAtual() {
  const range = SpreadsheetApp.getActiveRange();
  return range.getA1Notation();
}

function criarEventos(config) {
  try {
    if (!config) throw new Error("Configuração não fornecida");
    
    const { 
      calendario = "",
      titulos = "",
      descricao = null,
      datas = "",
      inicio = "",
      fim = null,
      notificacoes = null
    } = config;

    if (!calendario || !titulos || !datas || !inicio) {
      throw new Error("Preencha todos os campos obrigatórios (Títulos, Datas e Início)");
    }

    const calendar = CalendarApp.getCalendarsByName(calendario)[0];
    if (!calendar) throw new Error(`Calendário "${calendario}" não encontrado`);

    const ss = SpreadsheetApp.getActive();
    const valores = {
      titulos: ss.getRange(titulos).getValues(),
      descricao: descricao ? ss.getRange(descricao).getValues() : null,
      datas: ss.getRange(datas).getValues(),
      inicio: ss.getRange(inicio).getValues(),
      fim: fim ? ss.getRange(fim).getValues() : null,
      notificacoes: notificacoes ? ss.getRange(notificacoes).getValues() : null
    };

    const numEventos = valores.titulos.length;
    if (valores.datas.length !== numEventos || valores.inicio.length !== numEventos) {
      throw new Error("Todos os intervalos devem ter o mesmo número de linhas");
    }

    let criados = 0;
    const resultados = [];

    for (let i = 0; i < numEventos; i++) {
      try {
        if (!valores.titulos[i][0] || !valores.datas[i][0] || !valores.inicio[i][0]) {
          throw new Error(`Dados incompletos na linha ${i+1}`);
        }

        const startTime = new Date(valores.datas[i][0]);
        const horaInicio = new Date(valores.inicio[i][0]);
        startTime.setHours(horaInicio.getHours(), horaInicio.getMinutes());

        const endTime = valores.fim?.[i]?.[0] 
          ? new Date(
              new Date(valores.datas[i][0])
                .setHours(
                  new Date(valores.fim[i][0]).getHours(),
                  new Date(valores.fim[i][0]).getMinutes()
                )
            )
          : new Date(startTime.getTime() + 3600000);

        const evento = calendar.createEvent(
          valores.titulos[i][0].toString(),
          startTime,
          endTime
        );

        if (valores.descricao?.[i]?.[0]) {
          evento.setDescription(valores.descricao[i][0].toString());
        }

        // Nova lógica para múltiplas notificações
        if (valores.notificacoes?.[i]?.[0]) {
          const textoNotificacoes = valores.notificacoes[i][0].toString();
          const partes = textoNotificacoes.split(',').map(p => p.trim());
          
          partes.forEach(parte => {
            try {
              const minutos = converterParaMinutos(parte);
              if (minutos > 0) {
                evento.addPopupReminder(minutos);
                resultados.push(`🔔 Notificação: ${formatarTempo(minutos)} antes`);
              }
            } catch (e) {
              resultados.push(`⚠️ Notificação inválida: "${parte}"`);
            }
          });
        }

        criados++;
        resultados.push(`✅ ${valores.titulos[i][0]} (${startTime.toLocaleString()})`);

      } catch (e) {
        resultados.push(`❌ Linha ${i+1}: ${e.message}`);
      }
    }

    return {
      sucesso: true,
      total: numEventos,
      criados: criados,
      resultados: resultados
    };

  } catch (e) {
    return {
      sucesso: false,
      error: e.message
    };
  }
}

function converterParaMinutos(texto) {
  const numero = parseFloat(texto.replace(',', '.').match(/-?\d+([.,]\d+)?/)?.[0]?.replace(',', '.') || 0);
  
  if (/(hora|hr|h)/i.test(texto)) return Math.round(numero * 60);
  if (/(dia|d)/i.test(texto)) return Math.round(numero * 1440);
  return Math.round(numero); // Padrão: minutos
}

function formatarTempo(minutos) {
  if (minutos >= 1440) {
    const dias = minutos / 1440;
    return `${dias.toFixed(1)} dia${dias !== 1 ? 's' : ''}`;
  }
  if (minutos >= 60) {
    const horas = minutos / 60;
    return `${horas.toFixed(1)} hora${horas !== 1 ? 's' : ''}`;
  }
  return `${minutos} minuto${minutos !== 1 ? 's' : ''}`;
}
