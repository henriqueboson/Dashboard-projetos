/*
  > histórico de revisões

      - 20250424 - R01
        - autor: Henrique
        - observações:
          - correção na contabilização de atividades
*/

function f_recuperaProjetos() {
    var idPastaTodosProjetos = 'CÓDIGO DA PASTA DOS PROJETOS'; // Substitua pelo ID da pasta que contém os projetos
    var idMasterPlan = 'CÓDIGO DO MASTERPLAN'; // Substitua pelo ID do Plano Mensal
    var pastaTodosProjetos = DriveApp.getFolderById(idPastaTodosProjetos);
    var listaProjetos = pastaTodosProjetos.getFilesByType(MimeType.GOOGLE_SHEETS);
    var MasterPlan = SpreadsheetApp.openById(idMasterPlan);
  
    // Busca a aba do mês atual no MasterPlan
    var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    var mesAtual = meses[new Date().getMonth()];
    var abaMasterPlan = MasterPlan.getSheetByName(mesAtual);
  
    if (!abaMasterPlan) {
        Logger.log("Erro: Aba do mês '" + mesAtual + "' não encontrada no MasterPlan.");
        return;
    }
  
    // Busca as abas na planilha ativa para plotar dados
    var planilhaListaProjetos = SpreadsheetApp.getActiveSpreadsheet();
    var planilhaAux = planilhaListaProjetos.getSheetByName('_aux');
    var abaIndicadoresMacro = planilhaListaProjetos.getSheetByName('INDICADORES MACRO');
  
    if (!planilhaAux || !abaIndicadoresMacro) {
        Logger.log("Erro: Aba '_aux' ou 'INDICADORES MACRO' não encontrada.");
        return;
    }
  
    // Limpa dados anteriores
    planilhaAux.getRange("U3:AH").clearContent();
  
    // Recupera dados do MasterPlan
    var dadosMasterPlanH = abaMasterPlan.getRange('H:H').getValues().flat().map(String);
    var datasD = abaMasterPlan.getRange('D:D').getValues().flat();
    var datasF = abaMasterPlan.getRange('F:F').getValues().flat();
  
    // Plota data atual na celula H7
    var dataAtual = new Date();
    var dataAtualFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    abaIndicadoresMacro.getRange("H7").setValue(dataAtualFormatada);
  
    var linhaAtual = 3;  // Define a linha inicial para escrita
  
    // Busca dados nos arquivos da pasta de projetos
    while (listaProjetos.hasNext()) {
        var arquivoDaVez = listaProjetos.next();
        var nomeArquivoDaVez = arquivoDaVez.getName();
  
        // Tira planilhas gerais e planilha de modelo
        if (!nomeArquivoDaVez.includes('___modelo') && !nomeArquivoDaVez.includes('GERAL')) {
            var planilhaDaVez = SpreadsheetApp.open(arquivoDaVez);
            var abaVisaoGeral = planilhaDaVez.getSheetByName('VISÃO GERAL');
            var abaAux = planilhaDaVez.getSheetByName('aux');
  
            // Recupera dados das Planilhas
            if (abaVisaoGeral && abaAux) {
                var porcentagemDiasConsumidos = Math.round(abaVisaoGeral.getRange('G15').getValue() * 100) / 100;
                var porcentagemHorasApropriadas = Math.round(abaVisaoGeral.getRange('G21').getValue() * 100) / 100;
                var fimPrevisto = abaVisaoGeral.getRange('G8').getValue();
                var coordenador = abaVisaoGeral.getRange('E9').getValue(); 
  
                var atividadesConcluidasProjeto = 0;
                var atividadesAtrasadasProjeto = 0;
                var atividadesTotaisProjeto = 0;
                var atividadesFuturasProjeto = 0;
  
                // Contabiliza atividades
                for (var i = 0; i < dadosMasterPlanH.length; i++) {
                  var nomeAtividadeH = dadosMasterPlanH[i];
                  if (!nomeAtividadeH || typeof nomeAtividadeH !== 'string') continue;
                  if (!nomeAtividadeH.trim().endsWith(nomeArquivoDaVez)) continue;
              
                  var dataD = datasD[i]; // Data Final Planejada
                  var dataF = datasF[i]; // Data Final Real
              
                  if (!(dataD instanceof Date) || isNaN(dataD)) continue; // Ignora atividades sem data planejada válida
              
                  if (dataD > dataAtual){
                      atividadesFuturasProjeto++;  // Contabiliza atividades planejadas para o futuro 
                  }
              
                  atividadesTotaisProjeto++; // Contabiliza atividade válida
              
                  if (dataF instanceof Date && !isNaN(dataF)) {
                      atividadesConcluidasProjeto++; // Concluída se dataF preenchida corretamente
                  } else if (dataAtual > dataD) {
                      atividadesAtrasadasProjeto++;  // Atrasada se data final planejada passou e não foi concluída
                  }
              }
              
                if (atividadesTotaisProjeto > 0) {
                    var atividadesPlanejadasHoje = (atividadesTotaisProjeto - atividadesFuturasProjeto);
                    var percentualTotalProjeto = (atividadesConcluidasProjeto / atividadesPlanejadasHoje) * 100;
                    var percentualTotalProjetoFormatado = percentualTotalProjeto.toFixed(2).replace('.', ',');
  
                    var percentualAtrasadasProjeto = (atividadesAtrasadasProjeto / atividadesTotaisProjeto) * 100;
                    var percentualAtrasadasProjetoFormatado = percentualAtrasadasProjeto.toFixed(2).replace('.', ',');
  
                    var percentualAtividadesAtuais = atividadesConcluidasProjeto / atividadesTotaisProjeto * 100;
                    var percentualAtividadesAtuaisFormatado = percentualAtividadesAtuais.toFixed(2).replace('.', ',');    
  
                    // Escreve os dados linha a linha
                    planilhaAux.getRange(linhaAtual, 21).setValue(nomeArquivoDaVez);
                    planilhaAux.getRange(linhaAtual, 22).setValue(porcentagemDiasConsumidos);
                    planilhaAux.getRange(linhaAtual, 23).setValue(porcentagemHorasApropriadas);
                    planilhaAux.getRange(linhaAtual, 24).setValue(fimPrevisto);
                    planilhaAux.getRange(linhaAtual, 27).setValue(new Date());
                    planilhaAux.getRange(linhaAtual, 28).setValue(atividadesConcluidasProjeto);
                    planilhaAux.getRange(linhaAtual, 29).setValue(atividadesAtrasadasProjeto);
                    planilhaAux.getRange(linhaAtual, 30).setValue(atividadesTotaisProjeto);
                    planilhaAux.getRange(linhaAtual, 31).setValue(percentualTotalProjetoFormatado);
                    planilhaAux.getRange(linhaAtual, 32).setValue(percentualAtrasadasProjetoFormatado);
                    planilhaAux.getRange(linhaAtual, 33).setValue(coordenador);
                    planilhaAux.getRange(linhaAtual, 34).setFormula(`=HYPERLINK("${arquivoDaVez.getUrl()}";"LINK")`);
                    planilhaAux.getRange(linhaAtual, 35).setValue(percentualAtividadesAtuaisFormatado);
  
                    linhaAtual++; // Avança para próxima linha
                }
            }
        }
    }
  
    Logger.log("Dados atualizados na aba '_aux' com sucesso!");
  }
  