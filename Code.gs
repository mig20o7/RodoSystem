const SHEET_VIAGENS = "Viagens";
const SHEET_PASSAGENS = "Passagens";
function myFunction() {
  inicializarSistema();
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("RodoSystem - Gestão de Passagens")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let viagens = ss.getSheetByName(SHEET_VIAGENS);

  if (!viagens) {
    viagens = ss.insertSheet(SHEET_VIAGENS);
    viagens.appendRow([
      "ID",
      "Origem",
      "Destino",
      "Data",
      "Horario",
      "Empresa",
      "Preco",
      "TotalAssentos",
      "Status"
    ]);
  }

  const viagensPadrao = [
    ["V001", "Porto Velho", "Ariquemes", "2026-05-15", "08:30", "Expresso Norte", 85, 40, "Ativa"],
    ["V002", "Porto Velho", "Ji-Paraná", "2026-05-15", "14:00", "Real Bus", 135, 44, "Ativa"],
    ["V003", "Porto Velho", "Vilhena", "2026-05-16", "21:30", "Rota Amazônica", 210, 46, "Ativa"],
    ["V004", "Porto Velho", "Machadinho do Oeste", "2026-05-17", "20:00", "Empresa de Transportes Abel", 200, 30, "Ativa"],
    ["V005", "Porto Velho", "Rio Branco", "2026-05-19", "19:00", "Canário Branco", 240, 40, "Ativa"],
    
  ];

  const dadosExistentes = viagens.getDataRange().getValues();
  const idsExistentes = dadosExistentes.map(function(linha) {
    return String(linha[0]);
  });

  viagensPadrao.forEach(function(viagem) {
    if (!idsExistentes.includes(viagem[0])) {
      viagens.appendRow(viagem);
    }
  });

  let passagens = ss.getSheetByName(SHEET_PASSAGENS);

  if (!passagens) {
    passagens = ss.insertSheet(SHEET_PASSAGENS);
    passagens.appendRow([
      "ID",
      "IDViagem",
      "Passageiro",
      "CPF",
      "Assento",
      "FormaPagamento",
      "Preco",
      "Status",
      "DataVenda"
    ]);
  }

  return "Sistema inicializado com sucesso.";
}

function getViagens() {
  inicializarSistema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_VIAGENS);
  const dados = sheet.getDataRange().getDisplayValues();

  const viagens = [];
  const idsJaExibidos = {};

  for (let i = 1; i < dados.length; i++) {
    const id = String(dados[i][0]).trim();
    const status = String(dados[i][8]).trim();

    if (status === "Ativa" && !idsJaExibidos[id]) {
      viagens.push({
        id: id,
        origem: String(dados[i][1]),
        destino: String(dados[i][2]),
        data: String(dados[i][3]),
        horario: String(dados[i][4]),
        empresa: String(dados[i][5]),
        preco: Number(dados[i][6]),
        totalAssentos: Number(dados[i][7]),
        status: status
      });

      idsJaExibidos[id] = true;
    }
  }

  return viagens;
}


function getPassagens() {
  inicializarSistema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PASSAGENS);
  const dados = sheet.getDataRange().getDisplayValues();

  const passagens = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      passagens.push({
        id: String(dados[i][0]),
        idViagem: String(dados[i][1]),
        passageiro: String(dados[i][2]),
        cpf: String(dados[i][3]),
        assento: String(dados[i][4]),
        formaPagamento: String(dados[i][5]),
        preco: Number(String(dados[i][6]).replace(",", ".")),
        status: String(dados[i][7]),
        dataVenda: String(dados[i][8])
      });
    }
  }

  return passagens;
}

function venderPassagem(dados) {
  inicializarSistema();

  if (!dados.passageiro || !dados.cpf || !dados.idViagem || !dados.assento || !dados.formaPagamento) {
    throw new Error("Preencha todos os campos obrigatórios.");
  }

  const cpfLimpo = String(dados.cpf).replace(/\D/g, "");
  if (cpfLimpo.length !== 11) {
    throw new Error("CPF inválido. Informe 11 números.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const viagensSheet = ss.getSheetByName(SHEET_VIAGENS);
  const passagensSheet = ss.getSheetByName(SHEET_PASSAGENS);

  const viagens = viagensSheet.getDataRange().getValues();
  let viagemEncontrada = null;

  for (let i = 1; i < viagens.length; i++) {
    if (viagens[i][0] === dados.idViagem && viagens[i][8] === "Ativa") {
      viagemEncontrada = {
        id: viagens[i][0],
        preco: viagens[i][6],
        totalAssentos: viagens[i][7]
      };
      break;
    }
  }
function getDadosIniciais() {
  return {
    viagens: getViagens(),
    passagens: getPassagens()
  };
}
  if (!viagemEncontrada) {
    throw new Error("Viagem não encontrada ou inativa.");
  }

  const assento = Number(dados.assento);

  if (assento < 1 || assento > Number(viagemEncontrada.totalAssentos)) {
    throw new Error("Assento inválido para esta viagem.");
  }

  const passagens = passagensSheet.getDataRange().getValues();

  for (let i = 1; i < passagens.length; i++) {
    const mesmaViagem = passagens[i][1] === dados.idViagem;
    const mesmoAssento = Number(passagens[i][4]) === assento;
    const ativa = passagens[i][7] === "Vendida";

    if (mesmaViagem && mesmoAssento && ativa) {
      throw new Error("Este assento já está ocupado.");
    }
  }

  const idPassagem = "P" + new Date().getTime();

  passagensSheet.appendRow([
    idPassagem,
    dados.idViagem,
    dados.passageiro,
    cpfLimpo,
    assento,
    dados.formaPagamento,
    viagemEncontrada.preco,
    "Vendida",
    new Date()
  ]);

  return {
    sucesso: true,
    mensagem: "Passagem vendida com sucesso.",
    id: idPassagem
  };
}
function resetarPassagens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PASSAGENS);

  const ultimaLinha = sheet.getLastRow();

  // Mantém apenas o cabeçalho
  if (ultimaLinha > 1) {
    sheet.deleteRows(2, ultimaLinha - 1);
  }

  return {
    sucesso: true,
    mensagem: "Todas as passagens foram removidas."
  };
}
