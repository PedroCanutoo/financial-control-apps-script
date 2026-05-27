// ===============================
// Financial Control System (Apps Script)
// ===============================

var planilha = SpreadsheetApp.getActiveSpreadsheet();

var cadastro = planilha.getSheetByName("Cadastro");                       // acessa aba cadastro
var auxiliar = planilha.getSheetByName("Auxiliar");                       // acessa aba auxiliar
var movimentacoes = planilha.getSheetByName("Movimentações");             // acessa aba movimentações
var relatorio = planilha.getSheetByName("Relatório");                     // acessa aba relatorio
var gerador = planilha.getSheetByName("Gerador de relatórios");           // acessa aba gerador


// ===============================
// CADASTRO DE GASTOS / ENTRADA
// ===============================
function cadastrar() {
  var data = cadastro.getRange("C3:G3").getValue();                       //
  var tipo = cadastro.getRange("C5").getValue();                          //
  var categoria = cadastro.getRange("F5:G5").getValue();                  //    pega os valores das células selecionadas
  var descricao = cadastro.getRange("C7:G7").getValue();                  //
  var valor = cadastro.getRange("C9:G9").getValue();                      //

  if (!data || !tipo || !categoria || !valor) {
    SpreadsheetApp.getUi().alert("Preencha todos os campos!");            //   valida os campos, se tiver erro manda alerta
    return;
  }

  if (tipo !== 'Entrada') {
    valor *= -1;                                                          //    caso seja saída, vira negativo
  }

  var ultimaLinha = auxiliar.getLastRow() + 1;                            //    pega ultima linha

  // salva no banco auxiliar
  auxiliar.getRange(ultimaLinha, 1, 1, 8).setValues([[                    //
    data,                                                                 //
    '',                                                                   //
    '',                                                                   //    coloca os dados na aba auxiliar
    '',                                                                   //  
    tipo,                                                                 //  
    categoria,                                                            //
    descricao,                                                            //
    valor                                                                 //
  ]]);

  auxiliar.getRange(ultimaLinha, 2).setFormula('=SPLIT(A' + ultimaLinha + ';"/")');         // formata data

  if (ultimaLinha === 2) {                                                                  //
    movimentacoes.getRange(ultimaLinha, 9).setValue(valor);                                 //
  } else {                                                                                  // calcula saldo na aba movimentações
    var saldoAnterior = movimentacoes.getRange(ultimaLinha - 1, 9).getValue();              // 
    movimentacoes.getRange(ultimaLinha, 9).setValue(saldoAnterior + valor);                 //
  }

  limparCadastro();
}


// ===============================
// GERAR RELATÓRIO
// ===============================
function gerar() {

  relatorio.getRange("F2:F").clearContent();                                              // limpa conteúdo da coluna F

  var ultimaLinha = relatorio.getLastRow();                                               // pega última linha
  if (ultimaLinha < 2) return;                              

  var saldo = 0;

  for (var i = 2; i <= ultimaLinha; i++) {                                                // preenche o relatório                                   
    var valor = relatorio.getRange(i, 5).getValue();                                      //    com os saldos
    saldo += valor;                                                                       //
                                                                                          //
    relatorio.getRange(i, 6).setValue(saldo);                                             //
  }

  SpreadsheetApp.setActiveSheet(relatorio);                                               // manda pra pag de relatorio
}


// ===============================
// ENVIAR RELATÓRIO POR EMAIL
// ===============================
function enviar() {

  try {
    var destinatario = gerador.getRange("K4").getValue();                                 // pega os valores dentro do campo email
    var mensagem = gerador.getRange("I4").getValue();                                     // pega os valores dentro do campo de mensagem

    if (!destinatario) {
      SpreadsheetApp.getUi().alert("Informe um e-mail!");                                 // alerta de erro de campo
      return;
    }

    var email = {                                                                         // obj com informações para enviar email
      to: destinatario,
      subject: "Relatório Financeiro",
      body: mensagem,
      name: "Sistema Financeiro",
      attachments: [
        planilha.getAs(MimeType.PDF).setName("Relatorio_Financeiro.pdf")
      ]
    };

    // oculta abas
    cadastro.hideSheet();
    auxiliar.hideSheet();
    gerador.hideSheet();

    MailApp.sendEmail(email);                                                             // envia email 

    SpreadsheetApp.getUi().alert("Relatório enviado com sucesso!");                       // alerta falando que deu certo

  } catch (erro) {
    Logger.log(erro);
    SpreadsheetApp.getUi().alert("Erro ao enviar: " + erro.message);                      // alerta falando q deu erro

  } finally {
    // garante que as abas voltam
    cadastro.showSheet();
    auxiliar.showSheet();
    gerador.showSheet();
  }

  limparGerador();
  limparMail();
}


// ===============================
// LIMPEZA DE CAMPOS
// ===============================
function limparCadastro() {
  cadastro.getRange("C3:G3").clearContent();
  cadastro.getRange("C5").clearContent();
  cadastro.getRange("F5:G5").clearContent();
  cadastro.getRange("C7:G7").clearContent();
  cadastro.getRange("C9:G9").clearContent();
}

function limparGerador() {
  gerador.getRange("C3").clearContent();
  gerador.getRange("F3").clearContent();
  gerador.getRange("C5:F5").clearContent();
}

function limparMail() {
  gerador.getRange("I4:I6").clearContent();
  gerador.getRange("K4:K5").clearContent();
}
