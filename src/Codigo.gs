/*
 * Produção por voz — app da Web que grava a produção dos colaboradores na aba "Registros"
 * da planilha "Controle de Produção". Não altera colunas nem abas existentes: só acrescenta
 * linhas em Registros e guarda quem lançou na aba LOG_ABA (criada sozinha).
 *
 * Configuração (fica fora do repositório):
 *   PLANILHA_ID   — ID da planilha de produção
 *   CODIGO_EQUIPE — código que a equipe digita uma vez no celular
 * Vêm do arquivo 00_Segredos.gs, gerado pelo GitHub Actions a partir dos segredos do repositório,
 * ou das Propriedades do script.
 */

var VERSAO = '1.1.0';
var LOG_ABA = 'Lançamentos por voz';
var LOG_CABECALHO = ['ID', 'QUANDO', 'QUEM LANÇOU', 'O QUE FOI FALADO', 'LINHAS', 'DESFEITO EM'];
var DESFAZER_MINUTOS = 30;
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
  var props = PropertiesService.getScriptProperties();
  return {
    planilhaId: segredos.planilhaId || props.getProperty('PLANILHA_ID') || '',
    codigo: String(segredos.codigoEquipe || props.getProperty('CODIGO_EQUIPE') || '')
  };
}

// Bloqueia tentativas seguidas de código errado (10 erros em 10 minutos travam por 10 minutos).
function conferirCodigo_(codigo) {
  var cfg = config_();
  if (!cfg.planilhaId || !cfg.codigo) throw new Error('O app ainda não foi configurado (planilha ou código da equipe).');
  var cache = CacheService.getScriptCache();
  var erros = Number(cache.get('erros') || 0);
  if (erros >= 10) throw new Error('Muitas tentativas. Espere 10 minutos.');
  if (normalizarCodigo_(codigo) !== normalizarCodigo_(cfg.codigo)) {
    cache.put('erros', String(erros + 1), 600);
    throw new Error('CODIGO_ERRADO');
  }
  return cfg;
}

function normalizarCodigo_(c) {
  return String(c || '').trim().toLowerCase();
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
  return { pessoas: listaDaColuna_(cfgAba, colPessoas), categorias: listaDaColuna_(cfgAba, colCategorias) };
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

function apiInicio(codigo) {
  var cfg = conferirCodigo_(codigo);
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

function apiGravar(codigo, pedido) {
  var cfg = conferirCodigo_(codigo);
  pedido = pedido || {};
  var quem = String(pedido.quem || '').trim();
  if (!quem) throw new Error('Escolha quem está lançando.');
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

    var primeira = ultima + 1;
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

function apiDesfazer(codigo, id) {
  var cfg = conferirCodigo_(codigo);
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
    if (Date.now() - quando.getTime() > DESFAZER_MINUTOS * 60000) throw new Error('Só dá para desfazer até ' + DESFAZER_MINUTOS + ' minutos depois.');

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
