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

var VERSAO = '1.8.5';
var LOG_ABA = 'Lançamentos por voz';
var APELIDOS_ABA = 'Apelidos';
var HORARIOS_ABA = 'Horários';
var CHAMADA_ABA = 'Chamada';
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
    supervisores: l.supervisores,
    maximos: maximos_(ss),
    habitos: h.habitos,
    ativos: h.ativos,
    hora: Number(Utilities.formatDate(new Date(), tz, 'H'))
  };
}

function faixa_(hora) {
  var h = Math.floor(Number(hora));
  if (!(h >= 0 && h <= 23)) throw new Error('Faixa de horário inválida.');
  var dois = function (n) { return ('0' + (n % 24)).slice(-2) + ':00'; };
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
      var repetidas = repetidasNaFaixa_(aba, col, ultima, tz, dataTxt, f.inicio, linhas);
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

function repetidasNaFaixa_(aba, col, ultima, tz, dataTxt, inicio, linhas) {
  if (ultima <= col.linhaCabecalho) return [];
  var de = Math.max(col.linhaCabecalho + 1, ultima - 400);
  var vals = aba.getRange(de, 1, ultima - de + 1, col.largura).getDisplayValues();
  var hoje = dataTexto_(dataTxt, tz);
  var vistas = {};
  vals.forEach(function (l) {
    if (dataTexto_(l[col.data], tz) !== hoje) return;
    var m = String(l[col.inicio]).match(/(\d{1,2}):(\d{2})/);
    if (!m || ('0' + m[1]).slice(-2) + ':' + m[2] !== inicio) return;
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
  var mapa = {}, mudou = false;
  var situacao = vals.map(function (r) {
    var nome = oficial[semAcento_(r[0])], de = horaDe(r[1]), ate = horaDe(r[2]);
    var vazio = !String(r[1]).trim() && !String(r[2]).trim();
    var txt = !String(r[0]).trim() && vazio ? '' : !nome ? 'não achei esse nome em Configurações'
      : vazio ? SEM : de === null || ate === null || ate <= de ? 'confira entrada e saída (ex.: 08:00 e 12:00)' : 'ok';
    if (txt === 'ok') {
      var horas = mapa[nome] || (mapa[nome] = []);
      for (var x = Math.floor(de); x < Math.ceil(ate); x++) if (horas.indexOf(x) < 0) horas.push(x);
    }
    if (String(r[3]) !== txt) mudou = true;
    return [txt];
  });
  if (mudou) aba.getRange(2, 4, situacao.length, 1).setValues(situacao);
  return mapa;
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
    var m = String(r[col.inicio]).match(/(\d{1,2}):/);
    if (!p || !m || d < limite) return;
    (dias[p] || (dias[p] = {}))[d] = true;
    var ph = porHora[p] || (porHora[p] = {});
    (ph[Number(m[1])] || (ph[Number(m[1])] = {}))[d] = true;
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
function apiChamada(hora) {
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
  // Só chama quem trabalha nessa hora: pela aba Horários; quem não está lá, pelo costume das últimas 2 semanas.
  var ativos = habitos_(aba, col, ultima, tz, l.pessoas).ativos.filter(function (p) {
    if (chefes.indexOf(semAcento_(p)) >= 0) return false;
    var horas = fixos[p] || aprendidos[p] || [];
    return horas.indexOf(h) >= 0;
  });
  var feitos = {};
  if (ultima > col.linhaCabecalho) {
    var de = Math.max(col.linhaCabecalho + 1, ultima - 600);
    var hoje = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    aba.getRange(de, 1, ultima - de + 1, col.largura).getDisplayValues().forEach(function (r) {
      if (dataTexto_(r[col.data], tz) !== hoje) return;
      var m = String(r[col.inicio]).match(/(\d{1,2}):(\d{2})/);
      if (m && ('0' + m[1]).slice(-2) + ':' + m[2] === f.inicio) feitos[semAcento_(r[col.pessoa])] = true;
    });
  }
  // Respostas de hoje na aba Chamada: "nada" vale para a faixa; "não veio" vale para o dia todo.
  var nada = {}, foraHoje = {}, faltou = {};
  var cham = ss.getSheetByName(CHAMADA_ABA);
  if (cham && cham.getLastRow() > 1) {
    var hojeIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var n = cham.getLastRow(), ini = Math.max(2, n - 500);
    cham.getRange(ini, 1, n - ini + 1, CHAMADA_CABECALHO.length).getDisplayValues().forEach(function (r) {
      if (dataTexto_(r[1], tz) !== hojeIso) return;
      var p = semAcento_(r[3]);
      if (r[4] === 'não veio') foraHoje[p] = true;
      if (r[4] === 'falta') { foraHoje[p] = true; faltou[p] = true; }
      if (r[4] === 'nada' && String(r[2]).indexOf(f.inicio) === 0) nada[p] = true;
    });
  }
  var lista = function (teste) { return ativos.filter(function (p) { return teste(semAcento_(p)); }).sort(); };
  return {
    faixa: f.inicio + ' – ' + f.fim,
    faltam: lista(function (p) { return !feitos[p] && !nada[p] && !foraHoje[p]; }),
    feitos: lista(function (p) { return feitos[p]; }),
    nada: lista(function (p) { return !feitos[p] && nada[p]; }),
    naoVieram: lista(function (p) { return !feitos[p] && !nada[p] && foraHoje[p] && !faltou[p]; }),
    faltas: lista(function (p) { return !feitos[p] && !nada[p] && faltou[p]; })
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
  aba.appendRow([Utilities.formatDate(agora, tz, 'HH:mm'), dia, f.inicio + ' – ' + f.fim, pessoa, pedido.resposta]);
  if (pedido.resposta !== 'não respondeu') return { ok: true };
  // Duas horas do dia sem responder: falta no dia (a chamada não chama mais essa pessoa hoje).
  var hoje = dataTexto_(dia, tz), alvo = semAcento_(pessoa), faixas = {}, jaTem = false;
  var n = aba.getLastRow(), ini = Math.max(2, n - 500);
  aba.getRange(ini, 1, n - ini + 1, CHAMADA_CABECALHO.length).getDisplayValues().forEach(function (r) {
    if (dataTexto_(r[1], tz) !== hoje || semAcento_(r[3]) !== alvo) return;
    if (r[4] === 'não respondeu') faixas[r[2]] = true;
    if (r[4] === 'falta') jaTem = true;
  });
  if (jaTem || Object.keys(faixas).length < 2) return { ok: true };
  aba.appendRow([Utilities.formatDate(agora, tz, 'HH:mm'), dia, f.inicio + ' – ' + f.fim, pessoa, 'falta']);
  return { ok: true, falta: true };
}

// Fala sincronizada: o aparelho que comanda a chamada publica o texto com a hora certa de falar (relógio do servidor);
// os outros aparelhos leem e falam no mesmo instante.
function apiAnunciar(aparelho, texto) {
  var agora = Date.now();
  var anuncio = { id: aparelho + '-' + agora, de: String(aparelho), texto: String(texto).slice(0, 300), falarEm: agora + 3000 };
  CacheService.getScriptCache().put('anuncio', JSON.stringify(anuncio), 600);
  return { agora: Date.now(), anuncio: anuncio };
}

function apiAnuncio() {
  var bruto = CacheService.getScriptCache().get('anuncio');
  return { agora: Date.now(), anuncio: bruto ? JSON.parse(bruto) : null };
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
