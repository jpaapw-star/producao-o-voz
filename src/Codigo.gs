/*
 * Produção por voz — app da Web que grava a produção dos colaboradores na aba "Registros"
 * da planilha "Controle de Produção". Não altera colunas nem abas existentes: só acrescenta
 * linhas em Registros e guarda quem lançou na aba LOG_ABA (criada sozinha).
 *
 * Configuração (fica fora do repositório):
 *   PLANILHA_ID   — ID da planilha de produção
 * Quem tem o link do app pode lançar; a aba LOG_ABA guarda quem lançou cada vez.
 * Vêm do arquivo 00_Segredos.gs, gerado pelo GitHub Actions a partir dos segredos do repositório,
 * ou das Propriedades do script.
 */

var VERSAO = '1.15.0';
var LOG_ABA = 'Lançamentos por voz';
var APELIDOS_ABA = 'Apelidos';
var HORARIOS_ABA = 'Horários';
var CHAMADA_ABA = 'Chamada';
var AGENDA_ABA = 'Agenda da chamada';
var CHAMADA_CABECALHO = ['QUANDO', 'DATA', 'FAIXA', 'PESSOA', 'RESPOSTA'];
var RESPOSTAS = ['lançou', 'nada', 'não veio', 'não respondeu'];
var LOG_CABECALHO = ['ID', 'QUANDO', 'QUEM LANÇOU', 'O QUE FOI FALADO', 'LINHAS', 'DESFEITO EM'];
var LINHAS_HISTORICO = 3000;
var MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro',
  'Novembro', 'Dezembro'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Produção')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function incluir_(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function config_() {
  var segredos = (typeof SEGREDOS !== 'undefined' && SEGREDOS) ? SEGREDOS : {};
  var planilhaId = segredos.planilhaId || PropertiesService.getScriptProperties().getProperty('PLANILHA_ID') || '';
  if (!planilhaId) throw new Error('O app ainda não foi configurado (falta a planilha).');
  return { planilhaId: planilhaId };
}

function semAcento_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
}

/* ------------------------------------------------------------------ planilha */

function abrir_(cfg) {
  return SpreadsheetApp.openById(cfg.planilhaId);
}

function abaPorNome_(ss, nome) {
  var alvo = semAcento_(nome);
  var abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) if (semAcento_(abas[i].getName()) === alvo) return abas[i];
  return null;
}

// Acha a linha de cabeçalho de Registros e a posição de cada coluna pelo nome.
function colunasRegistros_(aba) {
  var topo = aba.getRange(1, 1, Math.min(5, aba.getMaxRows()), aba.getLastColumn()).getValues();
  for (var r = 0; r < topo.length; r++) {
    var nomes = topo[r].map(semAcento_);
    var col = {};
    nomes.forEach(function (n, i) {
      if (n === 'data') col.data = i;
      else if (n.indexOf('hora') === 0) col.hora = i;
      else if (n === 'pessoa' || n === 'colaborador') col.pessoa = i;
      else if (n === 'categoria') col.categoria = i;
      else if (n.indexOf('testad') >= 0) col.testados = i;
      else if (n.indexOf('finaliz') >= 0) col.finalizados = i;
      else if (n.indexOf('defeit') >= 0) col.defeitos = i;
      else if (n === 'inicio') col.inicio = i;
      else if (n === 'fim') col.fim = i;
    });
    if (col.pessoa !== undefined && col.categoria !== undefined && col.data !== undefined) {
      // Coluna do mês: a primeira sem título antes de DATA (na planilha atual, a coluna A).
      if (col.data > 0 && !nomes[col.data - 1]) col.mes = col.data - 1;
      col.linhaCabecalho = r + 1;
      col.largura = Math.max.apply(null, Object.keys(col).filter(function (k) { return k !== 'linhaCabecalho'; })
        .map(function (k) { return col[k]; })) + 1;
      return col;
    }
  }
  throw new Error('Não achei o cabeçalho da aba Registros (PESSOA, CATEGORIA, DATA).');
}

function ultimaLinhaComDados_(aba, col) {
  var ultima = aba.getLastRow();
  var passo = 500;
  while (ultima > col.linhaCabecalho) {
    var inicio = Math.max(col.linhaCabecalho + 1, ultima - passo + 1);
    var vals = aba.getRange(inicio, col.pessoa + 1, ultima - inicio + 1, 1).getValues();
    for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim()) return inicio + i;
    ultima = inicio - 1;
  }
  return col.linhaCabecalho;
}

function listaDaColuna_(aba, coluna) {
  var n = aba.getLastRow();
  if (n < 2) return [];
  var vistos = {};
  return aba.getRange(2, coluna, n - 1, 1).getValues()
    .map(function (l) { return String(l[0]).trim(); })
    .filter(function (v) {
      var k = semAcento_(v);
      if (!v || vistos[k]) return false;
      vistos[k] = true;
      return true;
    });
}

function listas_(ss) {
  var cfgAba = abaPorNome_(ss, 'Configurações');
  if (!cfgAba) throw new Error('Não achei a aba Configurações.');
  var cab = cfgAba.getRange(1, 1, 1, cfgAba.getLastColumn()).getValues()[0].map(semAcento_);
  var colPessoas = cab.indexOf('pessoas') + 1 || 1;
  var colCategorias = cab.indexOf('categorias') + 1 || 4;
  var colCargo = cab.indexOf('cargo') + 1;
  var supervisores = [];
  if (colCargo && cfgAba.getLastRow() > 1) {
    cfgAba.getRange(2, 1, cfgAba.getLastRow() - 1, Math.max(colPessoas, colCargo)).getValues().forEach(function (l) {
      var nome = String(l[colPessoas - 1]).trim();
      if (nome && /^supervis/.test(semAcento_(l[colCargo - 1]))) supervisores.push(nome);
    });
  }
  return { pessoas: listaDaColuna_(cfgAba, colPessoas), categorias: listaDaColuna_(cfgAba, colCategorias), supervisores: supervisores };
}

// Aba "Apelidos": APELIDO | NOME OU CATEGORIA. Criada vazia na primeira vez; você mesmo preenche.
function apelidos_(ss, l) {
  var aba = abaPorNome_(ss, APELIDOS_ABA);
  if (!aba) {
    aba = ss.insertSheet(APELIDOS_ABA, ss.getSheets().length);
    aba.appendRow(['APELIDO', 'NOME OU CATEGORIA']);
    aba.setFrozenRows(1);
    return [];
  }
  var oficial = {};
  l.pessoas.concat(l.categorias).forEach(function (x) { oficial[semAcento_(x)] = x; });
  var n = aba.getLastRow();
  if (n < 2) return [];
  var vals = aba.getRange(2, 1, n - 1, 3).getValues();
  var lista = [], mudou = false;
  // Coluna C diz se o app entendeu a linha, para você ver na hora um nome escrito errado.
  var situacao = vals.map(function (r) {
    var apelido = String(r[0]).trim(), alvo = oficial[semAcento_(r[1])];
    var txt = !apelido && !String(r[1]).trim() ? '' : !apelido ? 'falta o apelido' : !alvo ? 'não achei esse nome em Configurações' : 'ok';
    if (apelido && alvo) lista.push([apelido, alvo]);
    if (String(r[2]) !== txt) mudou = true;
    return [txt];
  });
  if (mudou) {
    if (!String(aba.getRange(1, 3).getValue())) aba.getRange(1, 3).setValue('SITUAÇÃO');
    aba.getRange(2, 3, situacao.length, 1).setValues(situacao);
  }
  return lista;
}

// Máximo que cada pessoa já fez em cada categoria (aba "Referência Metas"), para avisar número fora do normal.
function maximos_(ss) {
  var aba = abaPorNome_(ss, 'Referência Metas');
  if (!aba || aba.getLastRow() < 2) return {};
  var vals = aba.getRange(1, 1, aba.getLastRow(), Math.min(aba.getLastColumn(), 12)).getValues();
  var cab = vals[0].map(semAcento_);
  var cp = cab.indexOf('pessoa'), cc = cab.indexOf('categoria'), cm = -1;
  cab.forEach(function (c, i) { if (/^max/.test(c)) cm = i; });
  if (cp < 0 || cc < 0 || cm < 0) return {};
  var maximos = {};
  vals.slice(1).forEach(function (l) {
    var m = Number(String(l[cm]).replace(',', '.'));
    if (l[cp] && l[cc] && m > 0) maximos[String(l[cp]).trim() + '|' + String(l[cc]).trim()] = m;
  });
  return maximos;
}

function ehData_(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

function dataTexto_(v, tz) {
  if (ehData_(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}

// Hábitos recentes de cada pessoa: categorias usadas (mais recente primeiro) e o tipo de quantidade de costume.
function habitos_(aba, col, ultima, tz, pessoas) {
  var habitos = {}, ativos = [], hoje = [];
  var oficial = {};
  pessoas.forEach(function (p) { oficial[semAcento_(p)] = p; });
  if (ultima <= col.linhaCabecalho) return { habitos: habitos, ativos: ativos };
  var inicio = Math.max(col.linhaCabecalho + 1, ultima - LINHAS_HISTORICO + 1);
  var vals = aba.getRange(inicio, 1, ultima - inicio + 1, col.largura).getValues();
  var limite = Utilities.formatDate(new Date(Date.now() - 7 * 864e5), tz, 'yyyy-MM-dd');
  var somas = {};
  for (var i = vals.length - 1; i >= 0; i--) {
    var l = vals[i];
    var pessoa = oficial[semAcento_(l[col.pessoa])];
    if (!pessoa) continue;
    var cat = String(l[col.categoria]).trim();
    var h = habitos[pessoa] || (habitos[pessoa] = { categorias: [], tipo: null });
    if (cat && h.categorias.indexOf(cat) < 0 && h.categorias.length < 6) h.categorias.push(cat);
    if (ativos.indexOf(pessoa) < 0 && dataTexto_(l[col.data], tz) >= limite) ativos.push(pessoa);
    var s = somas[pessoa] || (somas[pessoa] = { testados: 0, finalizados: 0, defeitos: 0 });
    ['testados', 'finalizados', 'defeitos'].forEach(function (k) {
      if (col[k] !== undefined) s[k] += Number(l[col[k]]) || 0;
    });
  }
  Object.keys(somas).forEach(function (p) {
    var s = somas[p], total = s.testados + s.finalizados + s.defeitos;
    if (total >= 5) {
      if (s.finalizados / total >= 0.8) habitos[p].tipo = 'finalizados';
      else if (s.testados / total >= 0.8) habitos[p].tipo = 'testados';
    }
  });
  return { habitos: habitos, ativos: ativos };
}

/* ------------------------------------------------------------------ API do app */

function apiInicio() {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var aba = abaPorNome_(ss, 'Registros');
  if (!aba) throw new Error('Não achei a aba Registros.');
  var col = colunasRegistros_(aba);
  var l = listas_(ss);
  var ultima = ultimaLinhaComDados_(aba, col);
  var h = habitos_(aba, col, ultima, tz, l.pessoas);
  return {
    versao: VERSAO,
    pessoas: l.pessoas,
    categorias: l.categorias,
    apelidos: apelidos_(ss, l),
    comHorario: Object.keys(horarios_(ss, l)).length,
    agenda: agenda_(ss),
    supervisores: l.supervisores,
    maximos: maximos_(ss),
    habitos: h.habitos,
    ativos: h.ativos,
    hora: faixaDe_(Utilities.formatDate(new Date(), tz, 'HH:mm')),
    ia: !!(iaChave_() || ollamaChave_())
  };
}

// Faixas de meia em meia hora: a faixa "7" vai das 07:30 às 08:30. Lançamento antigo em hora cheia (08:00) conta na faixa
// em que começa a meia hora anterior (07:30 – 08:30).
function faixaDe_(hhmm) {
  var m = String(hhmm).match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return (Number(m[1]) + (Number(m[2]) >= 30 ? 0 : -1) + 24) % 24;
}

function faixa_(hora) {
  var h = Math.floor(Number(hora));
  if (!(h >= 0 && h <= 23)) throw new Error('Faixa de horário inválida.');
  var dois = function (n) { return ('0' + (n % 24)).slice(-2) + ':30'; };
  return { inicio: dois(h), fim: dois(h + 1) };
}

function apiGravar(pedido) {
  var cfg = config_();
  pedido = pedido || {};
  var quem = String(pedido.quem || '').trim();
  if (!quem) quem = 'Sem aparelho';
  var f = faixa_(pedido.hora);

  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var l = listas_(ss);
  var pessoas = {}, categorias = {};
  l.pessoas.forEach(function (p) { pessoas[semAcento_(p)] = p; });
  l.categorias.forEach(function (c) { categorias[semAcento_(c)] = c; });

  var linhas = (pedido.linhas || []).map(function (x, i) {
    var p = pessoas[semAcento_(x.pessoa)];
    var c = categorias[semAcento_(x.categoria)];
    if (!p) throw new Error('Linha ' + (i + 1) + ': pessoa fora da lista (' + x.pessoa + ').');
    if (!c) throw new Error('Linha ' + (i + 1) + ': categoria fora da lista (' + x.categoria + ').');
    var q = ['testados', 'finalizados', 'defeitos'].map(function (k) {
      var n = Number(x[k]) || 0;
      if (n < 0 || n > 999 || Math.floor(n) !== n) throw new Error('Linha ' + (i + 1) + ': quantidade inválida.');
      return n;
    });
    if (q[0] + q[1] + q[2] === 0) throw new Error('Linha ' + (i + 1) + ': sem quantidade.');
    return { pessoa: p, categoria: c, testados: q[0], finalizados: q[1], defeitos: q[2] };
  });
  if (!linhas.length) throw new Error('Nada para gravar.');
  if (linhas.length > 60) throw new Error('Muitas linhas de uma vez (máximo 60).');

  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    var aba = abaPorNome_(ss, 'Registros');
    var col = colunasRegistros_(aba);
    var ultima = ultimaLinhaComDados_(aba, col);
    var agora = new Date();
    var dataTxt = Utilities.formatDate(agora, tz, 'dd/MM/yyyy');
    var horaTxt = Utilities.formatDate(agora, tz, 'HH:mm');
    var mesNum = Number(Utilities.formatDate(agora, tz, 'M'));
    var anoTxt = Utilities.formatDate(agora, tz, 'yyyy');

    var faltam = ['testados', 'finalizados', 'defeitos', 'inicio', 'fim'].filter(function (k) { return col[k] === undefined; });
    if (faltam.length) throw new Error('Não achei em Registros a coluna de: ' + faltam.join(', ') + '. Confira o cabeçalho.');

    // Mesma pessoa e categoria já lançadas hoje nesta faixa: confirma antes (evita lançar duas vezes).
    if (!pedido.confirmarRepetidos) {
      var repetidas = repetidasNaFaixa_(aba, col, ultima, tz, dataTxt, Math.floor(Number(pedido.hora)), linhas);
      if (repetidas.length) return { ok: false, repetidas: repetidas, faixa: f.inicio + ' – ' + f.fim };
    }

    var temModelo = ultima > col.linhaCabecalho;
    var modelo = temModelo ? aba.getRange(ultima, 1, 1, col.largura).getValues()[0] : [];
    var formulas = temModelo ? aba.getRange(ultima, 1, 1, col.largura).getFormulasR1C1()[0] : [];
    var preenchidas = colunasComArrayFormula_(aba, col);

    // Data e horas vão como texto, igual a quem digita: a planilha converte no formato da coluna.
    var valores = linhas.map(function (x) {
      var linha = [];
      for (var i = 0; i < col.largura; i++) linha.push('');
      if (col.mes !== undefined) {
        linha[col.mes] = ehData_(modelo[col.mes])
          ? '01/' + ('0' + mesNum).slice(-2) + '/' + anoTxt
          : MESES[mesNum - 1] + '/' + anoTxt;
      }
      linha[col.data] = dataTxt;
      if (col.hora !== undefined) linha[col.hora] = horaTxt;
      linha[col.pessoa] = x.pessoa;
      linha[col.categoria] = x.categoria;
      if (col.testados !== undefined) linha[col.testados] = x.testados;
      if (col.finalizados !== undefined) linha[col.finalizados] = x.finalizados;
      if (col.defeitos !== undefined) linha[col.defeitos] = x.defeitos;
      if (col.inicio !== undefined) linha[col.inicio] = f.inicio;
      if (col.fim !== undefined) linha[col.fim] = f.fim;
      // Coluna preenchida por ARRAYFORMULA fica vazia, senão a fórmula quebra.
      preenchidas.forEach(function (i) { linha[i] = ''; });
      return linha;
    });

    // Outro app (ou alguém digitando) pode ter escrito embaixo entre a leitura e agora: anda até achar linhas vazias.
    var primeira = ultima + 1;
    for (var tentativa = 0; tentativa < 5; tentativa++) {
      var alvo = aba.getRange(primeira, col.pessoa + 1, valores.length, 1).getValues();
      var ocupada = -1;
      alvo.forEach(function (v, k) { if (String(v[0]).trim()) ocupada = k; });
      if (ocupada < 0) break;
      primeira += ocupada + 1;
    }
    var falta = primeira + valores.length - 1 - aba.getMaxRows();
    if (falta > 0) aba.insertRowsAfter(aba.getMaxRows(), falta);
    var destino = aba.getRange(primeira, 1, valores.length, col.largura);
    destino.setValues(valores);
    if (temModelo) {
      aba.getRange(ultima, 1, 1, col.largura).copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      // Colunas que eram fórmula na linha de cima continuam sendo fórmula.
      formulas.forEach(function (fr, i) {
        if (fr) aba.getRange(primeira, i + 1, valores.length, 1).setFormulaR1C1(fr);
      });
    }

    SpreadsheetApp.flush();
    if (col.mes !== undefined) consertarMes_(aba, col, Math.max(col.linhaCabecalho + 1, primeira - 200), primeira + valores.length - 1, tz);
    var gravadas = destino.getDisplayValues().map(function (v) { return assinatura_(v, col); });

    var id = Utilities.getUuid().slice(0, 8);
    var log = abaLog_(ss);
    log.appendRow([id, Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm:ss'), quem,
      String(pedido.texto || '').slice(0, 1000), JSON.stringify({ linhas: gravadas }), '']);
    return { ok: true, id: id, n: valores.length, faixa: f.inicio + ' – ' + f.fim };
  } finally {
    trava.releaseLock();
  }
}

// A coluna do mês tem de aparecer como nas outras linhas ("Outubro/2026"). Se aparecer número (46296, que é a data
// 01/10/2026 sem formato) ou data, troca pelo texto do mês, tirado da coluna DATA da mesma linha.
function consertarMes_(aba, col, de, ate, tz) {
  if (ate < de) return 0;
  var n = ate - de + 1;
  var mes = aba.getRange(de, col.mes + 1, n, 1).getDisplayValues();
  var datas = aba.getRange(de, col.data + 1, n, 1).getDisplayValues();
  var porNome = mes.some(function (m) { return /^[A-Za-zÀ-ú]+\/\d{4}$/.test(String(m[0]).trim()); });
  if (!porNome) return 0;
  var consertadas = 0;
  mes.forEach(function (m, i) {
    var v = String(m[0]).trim();
    if (!/^\d+([.,]\d+)?$/.test(v) && !/^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(v)) return;
    var d = dataTexto_(datas[i][0], tz);
    if (!d) return;
    var celula = aba.getRange(de + i, col.mes + 1);
    celula.setNumberFormat('@');
    celula.setValue(MESES[Number(d.slice(5, 7)) - 1] + '/' + d.slice(0, 4));
    consertadas++;
  });
  return consertadas;
}

function colunasComArrayFormula_(aba, col) {
  var formulas = aba.getRange(1, 1, col.linhaCabecalho + 1, col.largura).getFormulas();
  var cols = [];
  formulas.forEach(function (linha) {
    linha.forEach(function (fx, i) {
      if (/arrayformula|^=\s*(map|byrow|query|filter)\s*\(/i.test(fx) && cols.indexOf(i) < 0) cols.push(i);
    });
  });
  return cols;
}

function repetidasNaFaixa_(aba, col, ultima, tz, dataTxt, hora, linhas) {
  if (ultima <= col.linhaCabecalho) return [];
  var de = Math.max(col.linhaCabecalho + 1, ultima - 400);
  var vals = aba.getRange(de, 1, ultima - de + 1, col.largura).getDisplayValues();
  var hoje = dataTexto_(dataTxt, tz);
  var vistas = {};
  vals.forEach(function (l) {
    if (dataTexto_(l[col.data], tz) !== hoje) return;
    if (faixaDe_(l[col.inicio]) !== hora) return;
    vistas[semAcento_(l[col.pessoa]) + '|' + semAcento_(l[col.categoria])] = true;
  });
  return linhas.filter(function (x) { return vistas[semAcento_(x.pessoa) + '|' + semAcento_(x.categoria)]; })
    .map(function (x) { return x.pessoa + ' · ' + x.categoria; });
}

function abaLog_(ss) {
  var log = ss.getSheetByName(LOG_ABA);
  if (!log) {
    log = ss.insertSheet(LOG_ABA, ss.getSheets().length);
    log.appendRow(LOG_CABECALHO);
    log.setFrozenRows(1);
  }
  return log;
}

// Identifica uma linha gravada pelo conteúdo (como aparece na tela), para o Desfazer achar
// a linha mesmo se outras entrarem depois.
function assinatura_(linha, col) {
  return ['data', 'hora', 'pessoa', 'categoria', 'testados', 'finalizados', 'defeitos', 'inicio', 'fim']
    .map(function (k) { return col[k] === undefined ? '' : String(linha[col[k]]).trim(); }).join('|');
}

// Lançamentos de hoje feitos neste aparelho, do mais novo para o mais velho.
// Aba Horários (opcional): PESSOA, ENTRADA, SAÍDA. Pode ter mais de uma linha por pessoa (manhã e tarde).
// Devolve { PESSOA: [horas cheias em que trabalha] } e escreve na coluna D se entendeu cada linha.
function horarios_(ss, l) {
  var aba = abaPorNome_(ss, HORARIOS_ABA);
  if (!aba) {
    aba = ss.insertSheet(HORARIOS_ABA, ss.getSheets().length);
    aba.appendRow(['PESSOA', 'ENTRADA', 'SAÍDA', 'SITUAÇÃO']);
    aba.setFrozenRows(1);
  }
  var oficial = {};
  l.pessoas.forEach(function (x) { oficial[semAcento_(x)] = x; });
  var horaDe = function (v) {
    var m = String(v).match(/(\d{1,2})(?:\s*[:h]\s*(\d{2}))?/);
    return m && Number(m[1]) <= 24 ? Number(m[1]) + (Number(m[2]) || 0) / 60 : null;
  };
  var SEM = 'sem horário: chama pelo costume';
  var n = aba.getLastRow();
  var vals = n >= 2 ? aba.getRange(2, 1, n - 1, 4).getDisplayValues() : [];
  // Mantém a aba com as mesmas pessoas de Configurações (menos supervisores): quem falta entra no fim, sem horário.
  var presentes = {};
  vals.forEach(function (r) { presentes[semAcento_(r[0])] = true; });
  var chefes = l.supervisores.map(semAcento_);
  var novos = l.pessoas.filter(function (p) { return !presentes[semAcento_(p)] && chefes.indexOf(semAcento_(p)) < 0; })
    .map(function (p) { return [p, '', '', SEM]; });
  if (novos.length) aba.getRange(n + 1, 1, novos.length, 4).setValues(novos);
  var mapa = {}, periodos = {}, mudou = false;
  // Entrada e saída de cada um, para a chamada não chamar quem já foi embora ou ainda não chegou.
  Object.defineProperty(mapa, 'periodos', { value: periodos });
  var situacao = vals.map(function (r) {
    var nome = oficial[semAcento_(r[0])], de = horaDe(r[1]), ate = horaDe(r[2]);
    var vazio = !String(r[1]).trim() && !String(r[2]).trim();
    var txt = !String(r[0]).trim() && vazio ? '' : !nome ? 'não achei esse nome em Configurações'
      : vazio ? SEM : de === null || ate === null || ate <= de ? 'confira entrada e saída (ex.: 08:00 e 12:00)' : 'ok';
    if (txt === 'ok') {
      var horas = mapa[nome] || (mapa[nome] = []);
      (periodos[nome] || (periodos[nome] = [])).push([de, ate]);
      // Trabalha na faixa x (x:30 – x+1:30) se fica nela pelo menos 15 minutos.
      for (var x = 0; x < 24; x++) if (Math.min(ate, x + 1.5) - Math.max(de, x + 0.5) >= 0.25 && horas.indexOf(x) < 0) horas.push(x);
    }
    if (String(r[3]) !== txt) mudou = true;
    return [txt];
  });
  if (mudou) aba.getRange(2, 4, situacao.length, 1).setValues(situacao);
  return mapa;
}

// Agenda da chamada (aba opcional): uma hora por linha na coluna A. Vazia = 07:30, 08:30… 16:30 (menos 13:30) e 16:45.
// Devolve as horas em "HH:MM", em ordem e sem repetir.
function horaAgenda_(v) {
  var m = String(v).match(/^\s*(\d{1,2})\s*(?:[:h]\s*(\d{2}))?/);
  if (!m || Number(m[1]) > 23 || Number(m[2] || 0) > 59) return '';
  return ('0' + m[1]).slice(-2) + ':' + ('0' + Number(m[2] || 0)).slice(-2);
}
function agenda_(ss) {
  var aba = abaPorNome_(ss, AGENDA_ABA);
  if (!aba || aba.getLastRow() < 2) return [];
  var lista = [];
  aba.getRange(2, 1, aba.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) {
    var h = horaAgenda_(r[0]);
    if (h && lista.indexOf(h) < 0) lista.push(h);
  });
  return lista.sort();
}
// Faixa que a chamada de um horário cobra: a que acabou (08:30 cobra 07:30 – 08:30). Com 15 minutos ou mais de faixa nova,
// cobra a faixa em curso (16:45 cobra 16:30 – 17:30, para o fim do expediente).
function horaDaAgenda_(h) {
  var t = Number(h.slice(0, 2)) * 60 + Number(h.slice(3));
  var atual = Math.floor((t - 30) / 60);
  var x = t - (atual * 60 + 30) >= 15 ? atual : atual - 1;
  return (x + 48) % 24;
}
// Horários da chamada: os da aba Agenda da chamada; sem nenhum, 07:30, 08:30… 16:30 (menos 13:30) e 16:45 (igual ao app).
function horariosDaChamada_(ss) {
  var ag = agenda_(ss);
  if (ag.length) return ag;
  var todas = [];
  for (var x = 7; x <= 16; x++) if (x !== 13) todas.push(('0' + x).slice(-2) + ':30');
  return todas.concat(['16:45']);
}
// Momento (em horas) em que a faixa h é cobrada: o último horário da chamada que cobra essa faixa (16:45 para 16:30 – 17:30).
function momentoDaChamada_(ss, h) {
  var momento = null;
  horariosDaChamada_(ss).forEach(function (t) {
    if (horaDaAgenda_(t) === h) momento = Number(t.slice(0, 2)) + Number(t.slice(3)) / 60;
  });
  return momento === null ? h + 1.5 : momento;
}
function faixaDaAgenda_(h) {
  var f = faixa_(horaDaAgenda_(h));
  return f.inicio + ' – ' + f.fim;
}

function apiSalvarAgenda(lista) {
  var ss = abrir_(config_());
  var horas = [];
  (lista || []).forEach(function (v) {
    var h = horaAgenda_(v);
    if (!h) throw new Error('Horário inválido: ' + v + ' (use 09:50).');
    if (horas.indexOf(h) < 0) horas.push(h);
  });
  horas.sort();
  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    var aba = abaPorNome_(ss, AGENDA_ABA);
    if (!aba) {
      aba = ss.insertSheet(AGENDA_ABA, ss.getSheets().length);
      aba.getRange(1, 1, 1, 2).setValues([['HORÁRIO DA CHAMADA', 'COBRA A FAIXA']]);
      aba.setFrozenRows(1);
    }
    var antes = Math.max(0, aba.getLastRow() - 1);
    var linhas = horas.map(function (h) { return [h, faixaDaAgenda_(h)]; });
    while (linhas.length < antes) linhas.push(['', '']);
    if (linhas.length) {
      var r = aba.getRange(2, 1, linhas.length, 2);
      r.setNumberFormat('@');
      r.setValues(linhas);
    }
  } finally { trava.releaseLock(); }
  return horas;
}

// Horário pelo costume: horas em que a pessoa lançou em pelo menos 2 dias e em 40% dos dias que trabalhou (últimos 14 dias).
function horariosAprendidos_(aba, col, ultima, tz, pessoas) {
  var mapa = {};
  if (ultima <= col.linhaCabecalho || col.inicio === undefined) return mapa;
  var oficial = {};
  pessoas.forEach(function (p) { oficial[semAcento_(p)] = p; });
  var inicio = Math.max(col.linhaCabecalho + 1, ultima - LINHAS_HISTORICO + 1);
  var vals = aba.getRange(inicio, 1, ultima - inicio + 1, col.largura).getDisplayValues();
  var limite = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), tz, 'yyyy-MM-dd');
  var dias = {}, porHora = {};
  vals.forEach(function (r) {
    var p = oficial[semAcento_(r[col.pessoa])], d = dataTexto_(r[col.data], tz);
    var fx = faixaDe_(r[col.inicio]);
    if (!p || fx === null || d < limite) return;
    (dias[p] || (dias[p] = {}))[d] = true;
    var ph = porHora[p] || (porHora[p] = {});
    (ph[fx] || (ph[fx] = {}))[d] = true;
  });
  Object.keys(porHora).forEach(function (p) {
    var total = Object.keys(dias[p]).length;
    mapa[p] = Object.keys(porHora[p]).filter(function (h) {
      var n = Object.keys(porHora[p][h]).length;
      return n >= 2 && n >= 0.4 * total;
    }).map(Number);
  });
  return mapa;
}

// Chamada: quem trabalhou nos últimos 7 dias (menos supervisores) e ainda não tem linha hoje na faixa pedida.
// Quem já mandou a faixa h (h:30 – h+1:30), por qualquer app:
// - INÍCIO na faixa (este app);
// - INÍCIO – FIM cobrindo pelo menos 30 minutos da faixa (o outro app grava hora cheia: 11:00 – 12:00 vale 10:30 – 11:30 e 11:30 – 12:30);
// - registrado na meia hora depois do fim da faixa (respondendo à chamada, mesmo com a faixa errada).
function horasDe_(v) {
  var m = String(v).match(/(\d{1,2})\s*[:h]\s*(\d{2})/);
  return m ? Number(m[1]) + Number(m[2]) / 60 : null;
}
function lancouNaFaixa_(r, col, h) {
  if (faixaDe_(r[col.inicio]) === h) return true;
  var ini = horasDe_(r[col.inicio]), fim = horasDe_(r[col.fim]);
  if (ini !== null && fim !== null && fim > ini && Math.min(fim, h + 1.5) - Math.max(ini, h + 0.5) >= 0.5) return true;
  var reg = col.hora === undefined ? null : horasDe_(r[col.hora]);
  return reg !== null && reg >= h + 1.5 && reg < h + 2;
}

function apiChamada(hora, instante) {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var aba = abaPorNome_(ss, 'Registros');
  if (!aba) throw new Error('Não achei a aba Registros.');
  var col = colunasRegistros_(aba);
  var l = listas_(ss);
  var ultima = ultimaLinhaComDados_(aba, col);
  var f = faixa_(hora);
  var chefes = l.supervisores.map(semAcento_);
  var h = Math.floor(Number(hora));
  var fixos = horarios_(ss, l), aprendidos = horariosAprendidos_(aba, col, ultima, tz, l.pessoas);
  // Pela aba Horários, só chama quem está no trabalho na hora da chamada (até 10 minutos depois da saída):
  // quem sai 12:45 não é chamado às 13:30.
  var momento = momentoDaChamada_(ss, h);
  // Chamando de novo mais tarde: vale quem ainda está no trabalho agora (quem saiu para de ser chamado).
  if (instante) {
    var agoraTxt = Utilities.formatDate(new Date(Number(instante)), tz, 'HH:mm');
    momento = Math.max(momento, Number(agoraTxt.slice(0, 2)) + Number(agoraTxt.slice(3)) / 60);
  }
  var presente = function (p) {
    var per = fixos.periodos[p];
    return !per || per.some(function (x) { return x[0] <= momento && momento <= x[1] + 10 / 60; });
  };
  // Só chama quem trabalha nessa hora: pela aba Horários; quem não está lá, pelo costume das últimas 2 semanas.
  var ativos = habitos_(aba, col, ultima, tz, l.pessoas).ativos.filter(function (p) {
    if (chefes.indexOf(semAcento_(p)) >= 0) return false;
    if (!presente(p)) return false;
    var horas = fixos[p] || aprendidos[p] || [];
    return horas.indexOf(h) >= 0;
  });
  var feitos = {};
  if (ultima > col.linhaCabecalho) {
    var de = Math.max(col.linhaCabecalho + 1, ultima - 600);
    var hoje = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    aba.getRange(de, 1, ultima - de + 1, col.largura).getDisplayValues().forEach(function (r) {
      if (dataTexto_(r[col.data], tz) !== hoje) return;
      if (lancouNaFaixa_(r, col, h)) feitos[semAcento_(r[col.pessoa])] = true;
    });
  }
  // Respostas de hoje na aba Chamada: "nada" vale para a faixa; "não veio" vale para o dia todo.
  var nada = {}, foraHoje = {};
  var cham = ss.getSheetByName(CHAMADA_ABA);
  if (cham && cham.getLastRow() > 1) {
    var hojeIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var n = cham.getLastRow(), ini = Math.max(2, n - 500);
    cham.getRange(ini, 1, n - ini + 1, CHAMADA_CABECALHO.length).getDisplayValues().forEach(function (r) {
      if (dataTexto_(r[1], tz) !== hojeIso) return;
      var p = semAcento_(r[3]);
      if (r[4] === 'não veio') foraHoje[p] = true;
      // Linhas "falta" antigas (marcadas sozinhas até a 1.13.1) não contam mais.
      if (r[4] === 'nada' && String(r[2]).indexOf(f.inicio) === 0) nada[p] = true;
    });
  }
  var lista = function (teste) { return ativos.filter(function (p) { return teste(semAcento_(p)); }).sort(); };
  return {
    faixa: f.inicio + ' – ' + f.fim,
    faltam: lista(function (p) { return !feitos[p] && !nada[p] && !foraHoje[p]; }),
    feitos: lista(function (p) { return feitos[p]; }),
    nada: lista(function (p) { return !feitos[p] && nada[p]; }),
    naoVieram: lista(function (p) { return !feitos[p] && !nada[p] && foraHoje[p]; }),
    faltas: []
  };
}

// Guarda a resposta de cada nome chamado na aba Chamada, para ver depois se a chamada funciona.
function apiResposta(pedido) {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var l = listas_(ss);
  var oficial = {};
  l.pessoas.forEach(function (x) { oficial[semAcento_(x)] = x; });
  var pessoa = oficial[semAcento_(pedido && pedido.pessoa)];
  if (!pessoa) throw new Error('Pessoa não está em Configurações.');
  if (RESPOSTAS.indexOf(pedido.resposta) < 0) throw new Error('Resposta inválida.');
  var f = faixa_(pedido.hora);
  var aba = ss.getSheetByName(CHAMADA_ABA);
  if (!aba) {
    aba = ss.insertSheet(CHAMADA_ABA, ss.getSheets().length);
    aba.appendRow(CHAMADA_CABECALHO);
    aba.setFrozenRows(1);
  }
  var agora = new Date();
  var dia = Utilities.formatDate(agora, tz, 'dd/MM/yyyy');
  var trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    // Vários aparelhos chamam juntos: "não respondeu" da mesma pessoa e faixa no dia entra uma vez só.
    if (pedido.resposta === 'não respondeu') {
      var ult = aba.getLastRow(), de = Math.max(2, ult - 200);
      var repetida = ult >= 2 && aba.getRange(de, 1, ult - de + 1, CHAMADA_CABECALHO.length).getDisplayValues().some(function (r) {
        return dataTexto_(r[1], tz) === dataTexto_(dia, tz) && r[2] === f.inicio + ' – ' + f.fim && semAcento_(r[3]) === semAcento_(pessoa) && r[4] === 'não respondeu';
      });
      if (repetida) return { ok: true };
    }
    aba.appendRow([Utilities.formatDate(agora, tz, 'HH:mm'), dia, f.inicio + ' – ' + f.fim, pessoa, pedido.resposta]);
  } finally { trava.releaseLock(); }
  // Falta nunca é marcada sozinha (pedido do João Pedro): só "Não veio hoje", tocado por alguém, tira a pessoa do dia.
  return { ok: true };
}

// Fala sincronizada: o aparelho que comanda a chamada publica o texto com a hora certa de falar (relógio do servidor);
// os outros aparelhos leem e falam no mesmo instante.
function apiAnunciar(aparelho, texto) {
  var agora = Date.now();
  var anuncio = { id: aparelho + '-' + agora, de: String(aparelho), texto: String(texto).slice(0, 300), falarEm: agora + 3000 };
  CacheService.getScriptCache().put('anuncio', JSON.stringify(anuncio), 600);
  return { agora: Date.now(), anuncio: anuncio };
}

function apiAnuncio(aparelho, ouvindo) {
  if (aparelho && ouvindo) marcarOuvinte_(String(aparelho));
  var bruto = CacheService.getScriptCache().get('anuncio');
  return { agora: Date.now(), anuncio: bruto ? JSON.parse(bruto) : null, versao: VERSAO };
}

// Vários aparelhos com a chamada ligada: só o primeiro que pede a hora fala; os outros ficam em silêncio.
function apiChamador(aparelho, hora, forcar) {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var chave = 'chamador-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd') + '-' + Math.floor(Number(hora));
  var trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    var cache = CacheService.getScriptCache();
    var atual = cache.get(chave);
    // Quem aperta o botão Chamada assume a hora.
    if (atual && atual !== String(aparelho) && !forcar) return false;
    cache.put(chave, String(aparelho), 3600);
    return true;
  } finally {
    trava.releaseLock();
  }
}

function apiHoje(quem) {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var log = ss.getSheetByName(LOG_ABA);
  if (!log || log.getLastRow() < 2) return [];
  var n = log.getLastRow();
  var inicio = Math.max(2, n - 300);
  var hoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var alvo = semAcento_(quem);
  var lista = [];
  log.getRange(inicio, 1, n - inicio + 1, LOG_CABECALHO.length).getDisplayValues().forEach(function (r) {
    if (String(r[1]).indexOf(hoje) !== 0 || semAcento_(r[2]) !== alvo) return;
    var linhas = [];
    try { linhas = JSON.parse(r[4]).linhas; } catch (e) {}
    lista.push({
      id: r[0], hora: String(r[1]).slice(11, 16), desfeito: !!r[5],
      linhas: linhas.map(function (a) {
        var p = String(a).split('|');
        return { pessoa: p[2], categoria: p[3], testados: Number(p[4]) || 0, finalizados: Number(p[5]) || 0,
          defeitos: Number(p[6]) || 0, faixa: p[7] + ' – ' + p[8] };
      })
    });
  });
  return lista.reverse();
}

function apiDesfazer(id) {
  var cfg = config_();
  var ss = abrir_(cfg);
  var tz = ss.getSpreadsheetTimeZone();
  var log = ss.getSheetByName(LOG_ABA);
  if (!log) throw new Error('Lançamento não encontrado.');
  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    var n = log.getLastRow();
    var inicio = Math.max(2, n - 200);
    var vals = log.getRange(inicio, 1, n - inicio + 1, LOG_CABECALHO.length).getValues();
    var linhaLog = -1, registro = null;
    for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]) === String(id)) { linhaLog = inicio + i; registro = vals[i]; break; }
    if (!registro) throw new Error('Lançamento não encontrado.');
    if (registro[5]) throw new Error('Esse lançamento já foi desfeito.');
    var quando = ehData_(registro[1]) ? registro[1] : Utilities.parseDate(String(registro[1]), tz, 'dd/MM/yyyy HH:mm:ss');
    if (Utilities.formatDate(quando, tz, 'yyyy-MM-dd') !== Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')) {
      throw new Error('Só dá para desfazer lançamentos de hoje. Os mais antigos se corrigem na planilha.');
    }

    var faltando = JSON.parse(registro[4]).linhas.slice();
    var aba = abaPorNome_(ss, 'Registros');
    var col = colunasRegistros_(aba);
    var ultima = ultimaLinhaComDados_(aba, col);
    var de = Math.max(col.linhaCabecalho + 1, ultima - 500);
    var regs = aba.getRange(de, 1, ultima - de + 1, col.largura).getDisplayValues();
    var apagar = [];
    for (var r = regs.length - 1; r >= 0 && faltando.length; r--) {
      var k = faltando.indexOf(assinatura_(regs[r], col));
      if (k >= 0) { apagar.push(de + r); faltando.splice(k, 1); }
    }
    apagar.sort(function (a, b) { return b - a; }).forEach(function (linha) { aba.deleteRow(linha); });
    log.getRange(linhaLog, 6).setValue(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'));
    return { ok: true, apagadas: apagar.length, naoAchadas: faltando.length };
  } finally {
    trava.releaseLock();
  }
}

// Lista da vez da chamada: o primeiro computador a pedir lê a planilha e guarda; os outros recebem a mesma lista,
// para todos falarem os mesmos nomes no mesmo instante. A chave é o instante da fala (igual em todos).
function apiListaDaVez(hora, instante) {
  var r = listaDaVez_(hora, instante);
  // Alguém respondendo pelo microfone: os computadores esperam a conversa acabar antes de chamar (lido na hora, não guardado).
  r.conversaAte = Number(CacheService.getScriptCache().get('conversa')) || 0;
  return r;
}
// Um computador ouviu uma resposta e vai confirmar: ninguém começa a chamar até esse instante (relógio do servidor).
function apiConversa(ate) {
  var cache = CacheService.getScriptCache();
  var fim = Math.min(Number(ate) || 0, Date.now() + 60000);
  var atual = Number(cache.get('conversa')) || 0;
  if (fim > atual) cache.put('conversa', String(Math.round(fim)), 120);
  return { ok: true };
}
function listaDaVez_(hora, instante) {
  var cache = CacheService.getScriptCache();
  var chave = 'vez_' + Math.floor(Number(hora)) + '_' + Math.round(Number(instante) / 1000);
  var pronta = cache.get(chave);
  if (pronta) return JSON.parse(pronta);
  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    pronta = cache.get(chave);
    if (pronta) return JSON.parse(pronta);
    var r = apiChamada(hora, instante);
    cache.put(chave, JSON.stringify(r), 600);
    return r;
  } finally {
    trava.releaseLock();
  }
}

// ---------- Chamada com conversa (um de cada vez) ----------
// Quando algum computador está ouvindo pelo microfone, a chamada vira conversa: todos os computadores chamam juntos
// uma pessoa; o primeiro computador que ouve a pessoa responder "pega a vez" e conversa com ela; os outros ficam quietos.
// Terminou (ou ninguém respondeu em 2 minutos), todos chamam a próxima. O estado fica no cache do servidor.
var ESPERA_VEZ = 120000;
function vezChave_(hora) { return 'vezEstado_' + Math.floor(Number(hora)); }
function lerVez_(hora) { var t = CacheService.getScriptCache().get(vezChave_(hora)); return t ? JSON.parse(t) : null; }
function guardarVez_(hora, e) { CacheService.getScriptCache().put(vezChave_(hora), JSON.stringify(e), 21600); }
function segundoCheio_(ms) { return Math.ceil(ms / 1000) * 1000; }
function comTrava_(f) {
  var trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try { return f(); } finally { trava.releaseLock(); }
}

// Computadores com "Ouvir as respostas" ligado avisam a cada 30 s (apiAnuncio).
function marcarOuvinte_(aparelho) {
  var cache = CacheService.getScriptCache();
  var l = JSON.parse(cache.get('ouvintes') || '{}');
  l[aparelho] = Date.now();
  cache.put('ouvintes', JSON.stringify(l), 600);
}
function temOuvinte_() {
  var l = JSON.parse(CacheService.getScriptCache().get('ouvintes') || '{}');
  return Object.keys(l).some(function (k) { return Date.now() - l[k] < 150000; });
}

// Passa para a próxima pessoa: quem acabou de ser chamada e não respondeu vai para o fim da fila.
function avancarVez_(hora, e, agora) {
  var faltam = apiChamada(hora, agora).faltam;
  var antes = e.fila.filter(function (n) { return faltam.indexOf(n) >= 0; });
  var atual = e.pessoa;
  antes = antes.filter(function (n) { return n !== atual; });
  if (atual && faltam.indexOf(atual) >= 0) { antes.push(atual); e.tentativas[atual] = (e.tentativas[atual] || 0) + 1; }
  faltam.forEach(function (n) { if (antes.indexOf(n) < 0) antes.push(n); });
  e.fila = antes;
  e.pessoa = antes[0] || null;
  // Chegou no fim da lista e voltou ao começo (a próxima já foi chamada nesta volta): os computadores falam
  // a lista de quem ainda falta antes de chamar.
  e.anunciar = false;
  if (e.pessoa && (e.tentativas[e.pessoa] || 0) > (e.volta || 0)) { e.volta = (e.volta || 0) + 1; e.anunciar = true; }
  e.falarEm = segundoCheio_(agora + 3000);
  e.ate = e.falarEm + ESPERA_VEZ;
  e.dono = null;
  e.n = (e.n || 0) + 1;
  return e;
}

// Começo da chamada (todos os computadores pedem no mesmo instante; o primeiro monta a fila).
function apiComecarVez(hora, inicio) {
  return comTrava_(function () {
    var e = lerVez_(hora);
    if (e && e.inicio === Number(inicio)) return e;
    var r = apiChamada(hora, inicio);
    e = { inicio: Number(inicio), conversa: temOuvinte_(), fila: r.faltam, tentativas: {}, pessoa: r.faltam[0] || null,
      falarEm: Number(inicio), ate: Number(inicio) + ESPERA_VEZ, dono: null, n: 0, volta: 0, anunciar: false };
    guardarVez_(hora, e);
    return e;
  });
}

// Como está a vez agora. Passou o tempo sem ninguém pegar (ou o computador que pegou sumiu): próxima pessoa.
function apiVez(hora) {
  var e = lerVez_(hora);
  if (!e || !e.pessoa || Date.now() <= e.ate) return e;
  return comTrava_(function () {
    var x = lerVez_(hora);
    if (x && x.pessoa && Date.now() > x.ate) { avancarVez_(hora, x, Date.now()); guardarVez_(hora, x); }
    return x;
  });
}

// O computador que ouviu a pessoa pega a vez (só um).
function apiPegarVez(hora, falarEm, aparelho) {
  return comTrava_(function () {
    var e = lerVez_(hora);
    if (!e || e.falarEm !== Number(falarEm) || (e.dono && e.dono !== aparelho)) return { ok: false, vez: e };
    e.dono = aparelho;
    e.ate = Date.now() + 120000;
    guardarVez_(hora, e);
    return { ok: true, vez: e };
  });
}

// Conversa terminada (gravou, nada, ou desistiu): todos chamam a próxima. Um botão na tela também termina a vez.
function apiFimDaVez(hora, falarEm, aparelho) {
  return comTrava_(function () {
    var e = lerVez_(hora);
    if (!e || e.falarEm !== Number(falarEm) || (e.dono && aparelho && e.dono !== aparelho)) return e;
    avancarVez_(hora, e, Date.now());
    guardarVez_(hora, e);
    return e;
  });
}


// ---------- IA da conversa (Gemini, plano gratuito do Google AI Studio) ----------
// A chave vem do segredo IA_CHAVE do GitHub (ou da propriedade IA_CHAVE do script). Sem chave, a conversa usa só as regras do app.
var IA_MODELOS = ['gemini-3.1-flash-lite-preview', 'gemini-flash-lite-latest', 'gemini-flash-latest'];
function iaChave_() {
  var segredos = (typeof SEGREDOS !== 'undefined' && SEGREDOS) ? SEGREDOS : {};
  return segredos.iaChave || PropertiesService.getScriptProperties().getProperty('IA_CHAVE') || '';
}
// Reserva: Ollama Cloud (segredo OLLAMA_API_KEY), usado quando o Gemini está ocupado ou fora do ar.
var OLLAMA_MODELO = 'gpt-oss:20b';
function ollamaChave_() {
  var segredos = (typeof SEGREDOS !== 'undefined' && SEGREDOS) ? SEGREDOS : {};
  return segredos.ollamaChave || PropertiesService.getScriptProperties().getProperty('OLLAMA_API_KEY') || '';
}

// Rode uma vez no editor do Apps Script para autorizar o acesso à internet (a IA), depois de colocar a chave.
function autorizarIA() {
  var r = apiConversar({ pessoa: 'TESTE', hora: 9, historico: [{ quem: 'pessoa', texto: 'é... hum... fiz 3 notebooks finalizados' }] });
  Logger.log(JSON.stringify(r));
  return r;
}

function promptConversa_(pessoa, faixa, categorias) {
  return [
    'Você é a voz da chamada de produção de uma equipe de assistência técnica. Está conversando por voz, num computador da sala, com '
      + pessoa + ' sobre a produção da faixa ' + faixa + '.',
    'O texto da pessoa vem de um ditado automático: pode ter erros de reconhecimento, vícios de linguagem (é, hum, tipo, né, então), pausas, '
      + 'repetições, frases cortadas e correções ("3, não, 4"). Entenda a intenção e use sempre o último valor corrigido.',
    'Objetivo: saber, para cada categoria de produto, quantos ela testou, finalizou e quantos deram defeito nesta faixa.',
    'Use só categorias desta lista, escolhendo a mais parecida com o que foi dito (plural, apelido ou erro de ditado contam); '
      + 'se não der para saber qual, pergunte: ' + categorias.join(', ') + '.',
    'Se só disser a quantidade sem dizer se foi testado ou finalizado, pergunte. Se a pessoa esquecer, ajude com perguntas curtas, uma de cada vez.',
    'Fale em português do Brasil, com frases curtas e naturais, porque a resposta é lida em voz alta: sem emojis, sem listas, sem markdown. '
      + 'Chame a pessoa pelo primeiro nome.',
    'Quando tiver a produção, repita tudo resumido e pergunte se está certo (acao "confirmar", com as linhas).',
    'Só use acao "gravar" quando a pessoa confirmar o resumo que você acabou de fazer; repita as mesmas linhas e fale algo curto como "Perfeito, enviando."',
    'Se ela disser que não fez nada nesta faixa, confirme; quando confirmar, use acao "nada". Se ela pedir para parar, use acao "desistir".',
    'Nunca invente números. Responda só o JSON pedido.'
  ].join('\n');
}

function apiConversar(pedido) {
  var chave = iaChave_(), reserva = ollamaChave_();
  if (!chave && !reserva) throw new Error('A IA da conversa não está configurada (falta o segredo IA_CHAVE ou OLLAMA_API_KEY).');
  pedido = pedido || {};
  var ss = abrir_(config_());
  var l = listas_(ss);
  var f = faixa_(pedido.hora);
  var nome = String(pedido.pessoa || '').split(/\s+/)[0];
  nome = nome.charAt(0) + nome.slice(1).toLowerCase();
  var contents = (pedido.historico || []).slice(-20).map(function (m) {
    return { role: m.quem === 'ia' ? 'model' : 'user', parts: [{ text: String(m.texto || '').slice(0, 800) }] };
  });
  // A conversa tem de começar com a pessoa.
  while (contents.length && contents[0].role === 'model') contents.shift();
  if (!contents.length) contents = [{ role: 'user', parts: [{ text: 'sou eu' }] }];
  var corpo = {
    systemInstruction: { parts: [{ text: promptConversa_(nome, f.inicio + ' às ' + f.fim, l.categorias) }] },
    contents: contents,
    generationConfig: {
      temperature: 0.3,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          fala: { type: 'STRING' },
          acao: { type: 'STRING', enum: ['perguntar', 'confirmar', 'gravar', 'nada', 'desistir'] },
          linhas: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
            categoria: { type: 'STRING' }, testados: { type: 'INTEGER' }, finalizados: { type: 'INTEGER' }, defeitos: { type: 'INTEGER' }
          }, required: ['categoria'] } }
        },
        required: ['fala', 'acao']
      }
    }
  };
  var ultimoErro = '';
  for (var i = 0; chave && i < IA_MODELOS.length; i++) {
    var resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + IA_MODELOS[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': chave }, payload: JSON.stringify(corpo)
    });
    if (resp.getResponseCode() !== 200) { ultimoErro = resp.getResponseCode() + ' ' + String(resp.getContentText()).slice(0, 200); continue; }
    try {
      var dados = JSON.parse(resp.getContentText());
      var texto = dados.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('');
      return limparRespostaIA_(JSON.parse(texto), l.categorias);
    } catch (e) { ultimoErro = 'resposta estranha: ' + e.message; }
  }
  if (reserva) {
    var mensagens = [{ role: 'system', content: corpo.systemInstruction.parts[0].text + '\nResponda só com um JSON: {"fala": texto, "acao": perguntar|confirmar|gravar|nada|desistir, "linhas": [{"categoria", "testados", "finalizados", "defeitos"}]}.' }]
      .concat(contents.map(function (c) { return { role: c.role === 'model' ? 'assistant' : 'user', content: c.parts[0].text }; }));
    var r2 = UrlFetchApp.fetch('https://ollama.com/api/chat', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + reserva },
      payload: JSON.stringify({ model: OLLAMA_MODELO, messages: mensagens, stream: false, format: 'json', options: { temperature: 0.3 } })
    });
    if (r2.getResponseCode() === 200) {
      try { return limparRespostaIA_(JSON.parse(JSON.parse(r2.getContentText()).message.content), l.categorias); }
      catch (e) { ultimoErro = 'Ollama: resposta estranha: ' + e.message; }
    } else ultimoErro = 'Ollama ' + r2.getResponseCode() + ' ' + String(r2.getContentText()).slice(0, 200);
  }
  throw new Error('A IA não respondeu (' + ultimoErro + ').');
}

// Só categorias da planilha e quantidades inteiras; categoria desconhecida vira pergunta.
function limparRespostaIA_(r, categorias) {
  var oficial = {};
  categorias.forEach(function (c) { oficial[semAcento_(c)] = c; });
  var acao = ['perguntar', 'confirmar', 'gravar', 'nada', 'desistir'].indexOf(r.acao) >= 0 ? r.acao : 'perguntar';
  var desconhecida = '';
  var linhas = (r.linhas || []).map(function (x) {
    var c = oficial[semAcento_(x.categoria)];
    if (!c) desconhecida = x.categoria;
    var n = function (v) { v = Math.round(Number(v) || 0); return v > 0 && v < 1000 ? v : 0; };
    return { categoria: c || '', testados: n(x.testados), finalizados: n(x.finalizados), defeitos: n(x.defeitos) };
  }).filter(function (x) { return x.categoria && x.testados + x.finalizados + x.defeitos > 0; });
  var fala = String(r.fala || '').replace(/[*_#`]/g, '').trim();
  if (desconhecida && (acao === 'gravar' || acao === 'confirmar')) {
    return { fala: 'Não achei a categoria ' + desconhecida + ' na planilha. Qual é o produto?', acao: 'perguntar', linhas: [] };
  }
  if ((acao === 'gravar' || acao === 'confirmar') && !linhas.length) acao = 'perguntar';
  return { fala: fala || 'Pode repetir?', acao: acao, linhas: linhas };
}
