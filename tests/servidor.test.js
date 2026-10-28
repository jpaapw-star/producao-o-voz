const test = require('node:test');
const assert = require('node:assert');
const { criarPlanilha, carregarServidor } = require('./gas');

const CAB = ['', 'DATA', 'HORA DE REGISTRO', 'PESSOA', 'CATEGORIA', 'QTD. TESTADOS', 'QTD. FINALIZADOS', 'QTD. DEFEITOS', 'INÍCIO', 'FIM'];

function planilha() {
  const hoje = new Date();
  return criarPlanilha({
    Registros: [
      CAB,
      ['Outubro/2026', '06/10/2026', '08:35', 'ANA LEAL', 'NOTEBOOK', 0, 1, 0, '08:30', '09:30'],
      ['Outubro/2026', hoje, '08:36', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 3, 0, '08:30', '09:30'],
      ['Outubro/2026', hoje, '09:36', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 4, 0, '09:30', '10:30']
    ],
    'Configurações': [
      ['PESSOAS', 'CARGO', '', 'CATEGORIAS'],
      ['ANA LEAL', 'FUNCIONÁRIO', '', 'NOTEBOOK'],
      ['DANIEL VARGAS', 'ESTAGIÁRIO', '', 'CAIXA DE SOM'],
      ['JULIA GOMES', 'ESTAGIÁRIO', '', 'ROTEADOR'],
      ['', '', '', 'FONE']
    ],
    Dashboard: [['x']]
  });
}

test('entra sem código; sem planilha configurada avisa', () => {
  const s = carregarServidor(planilha());
  assert.ok(s.apiInicio().pessoas.length);
  s.SEGREDOS.planilhaId = '';
  assert.throws(() => s.apiInicio(), /não foi configurado/);
});

test('início traz listas e hábitos', () => {
  const s = carregarServidor(planilha());
  const d = JSON.parse(JSON.stringify(s.apiInicio()));
  assert.deepStrictEqual(d.pessoas, ['ANA LEAL', 'DANIEL VARGAS', 'JULIA GOMES']);
  assert.deepStrictEqual(d.categorias, ['NOTEBOOK', 'CAIXA DE SOM', 'ROTEADOR', 'FONE']);
  assert.deepStrictEqual(d.habitos['DANIEL VARGAS'], { categorias: ['CAIXA DE SOM'], tipo: 'finalizados' });
  assert.deepStrictEqual(d.ativos, ['DANIEL VARGAS']);
});

test('grava no fim de Registros sem tocar no resto, e o Desfazer apaga só o que gravou', () => {
  const p = planilha();
  const s = carregarServidor(p);
  const r = s.apiGravar({
    quem: 'ANA LEAL', hora: 10, texto: 'Julia roteador 4 testados; Ana notebook 2 finalizados',
    linhas: [
      { pessoa: 'julia gomes', categoria: 'roteador', testados: 4, finalizados: 0, defeitos: 0 },
      { pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 0, finalizados: 2, defeitos: 0 }
    ]
  });
  assert.strictEqual(r.n, 2);
  assert.strictEqual(r.faixa, '10:30 – 11:30');
  const reg = p.getSheetByName('Registros').dados;
  assert.strictEqual(reg.length, 6);
  const nova = reg[4];
  assert.match(nova[0], /^[A-ZÇ][a-zç]+\/\d{4}$/);
  assert.match(nova[1], /^\d{2}\/\d{2}\/\d{4}$/);
  assert.match(nova[2], /^\d{2}:\d{2}$/);
  assert.deepStrictEqual(nova.slice(3), ['JULIA GOMES', 'ROTEADOR', 4, 0, 0, '10:30', '11:30']);

  const log = p.getSheetByName('Lançamentos por voz').dados;
  assert.deepStrictEqual(JSON.parse(JSON.stringify(log[0])), ['ID', 'QUANDO', 'QUEM LANÇOU', 'O QUE FOI FALADO', 'LINHAS', 'DESFEITO EM']);
  assert.strictEqual(log[1][2], 'ANA LEAL');

  // Outra pessoa lança depois; o Desfazer não pode apagar a linha dela.
  s.apiGravar({ quem: 'DANIEL VARGAS', hora: 10, linhas: [{ pessoa: 'DANIEL VARGAS', categoria: 'FONE', finalizados: 1 }] });
  const d = s.apiDesfazer(r.id);
  assert.strictEqual(d.apagadas, 2);
  assert.deepStrictEqual(reg.slice(3).map((l) => l[3] + ' ' + l[4]), ['DANIEL VARGAS CAIXA DE SOM', 'DANIEL VARGAS FONE']);
  assert.throws(() => s.apiDesfazer(r.id), /já foi desfeito/);
});

test('recusa pessoa ou categoria fora da lista e linha vazia', () => {
  const s = carregarServidor(planilha());
  assert.throws(() => s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'FULANO', categoria: 'NOTEBOOK', testados: 1 }] }), /pessoa fora da lista/);
  assert.throws(() => s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'ANA LEAL', categoria: 'CARRO', testados: 1 }] }), /categoria fora da lista/);
  assert.throws(() => s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'ANA LEAL', categoria: 'NOTEBOOK' }] }), /sem quantidade/);
});

test('coluna com ARRAYFORMULA fica vazia e fórmula da linha de cima é repetida', () => {
  const p = planilha();
  const reg = p.getSheetByName('Registros');
  reg._formulas['1,1'] = '=ARRAYFORMULA(TEXT(B2:B;"mmmm/yyyy"))';
  const s = carregarServidor(p);
  s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 1 }] });
  assert.strictEqual(reg.dados[4][0], '');
});

test('mesma pessoa e categoria na mesma faixa pede confirmação antes de gravar de novo', () => {
  const p = planilha();
  const s = carregarServidor(p);
  const pedido = { quem: 'DANIEL VARGAS', hora: 9, linhas: [{ pessoa: 'DANIEL VARGAS', categoria: 'CAIXA DE SOM', finalizados: 2 }] };
  const r = s.apiGravar(pedido);
  assert.strictEqual(r.ok, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r.repetidas)), ['DANIEL VARGAS · CAIXA DE SOM']);
  assert.strictEqual(p.getSheetByName('Registros').dados.length, 4);
  assert.strictEqual(s.apiGravar(Object.assign({ confirmarRepetidos: true }, pedido)).n, 1);
  assert.strictEqual(s.apiGravar(Object.assign({}, pedido, { hora: 11 })).n, 1);
});

test('coluna de quantidade faltando no cabeçalho avisa em vez de gravar errado', () => {
  const p = planilha();
  p.getSheetByName('Registros').dados[0][7] = 'OUTRA COISA';
  const s = carregarServidor(p);
  assert.throws(() => s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 1 }] }), /coluna de: defeitos/);
});

test('aba Apelidos é criada vazia e depois lida, só com nomes da lista', () => {
  const p = planilha();
  const s = carregarServidor(p);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiInicio().apelidos)), []);
  const aba = p.getSheetByName('Apelidos');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(aba.dados[0])), ['APELIDO', 'NOME OU CATEGORIA']);
  aba.dados.push(['Dani', 'daniel vargas'], ['som', 'CAIXA DE SOM'], ['Zé', 'FULANO'], ['', 'FONE']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiInicio().apelidos)), [['Dani', 'DANIEL VARGAS'], ['som', 'CAIXA DE SOM']]);
  assert.deepStrictEqual(aba.dados.map((l) => l[2]), ['SITUAÇÃO', 'ok', 'ok', 'não achei esse nome em Configurações', 'falta o apelido']);
});

test('início traz supervisores e máximos da Referência Metas', () => {
  const p = planilha();
  p.getSheetByName('Configurações').dados[2][1] = 'SUPERVISOR';
  p.insertSheet('Referência Metas').dados.push(['PESSOA', 'CATEGORIA', 'META SUGERIDA', 'MÉDIA HIST.', 'MIN HIST.', 'MAX HIST.'], ['ANA LEAL', 'NOTEBOOK', 24, 22.5, 20, 25]);
  const d = JSON.parse(JSON.stringify(carregarServidor(p).apiInicio()));
  assert.deepStrictEqual(d.supervisores, ['DANIEL VARGAS']);
  assert.deepStrictEqual(d.maximos, { 'ANA LEAL|NOTEBOOK': 25 });
});

test('grava abaixo de linha que outro app escreveu, e lista e desfaz lançamentos de hoje', () => {
  const p = planilha();
  const s = carregarServidor(p);
  const reg = p.getSheetByName('Registros');
  // Simula o outro app escrevendo na linha 5 sem a nossa leitura ver (pula a busca da última linha).
  const original = reg.getLastRow;
  let chamadas = 0;
  reg.getLastRow = () => (chamadas++ === 0 ? 4 : original());
  reg.dados.push(['Outubro/2026', 'x', '10:00', 'BRUNO', 'FONE', 1, 0, 0, '10:30', '11:30']);
  const r = s.apiGravar({ quem: 'ANA LEAL', hora: 10, linhas: [{ pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 1 }] });
  reg.getLastRow = original;
  assert.strictEqual(reg.dados[4][3], 'BRUNO');
  assert.strictEqual(reg.dados[5][3], 'ANA LEAL');
  const hoje = JSON.parse(JSON.stringify(s.apiHoje('ana leal')));
  assert.strictEqual(hoje.length, 1);
  assert.strictEqual(hoje[0].id, r.id);
  assert.deepStrictEqual(hoje[0].linhas[0], { pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 1, finalizados: 0, defeitos: 0, faixa: '10:30 – 11:30' });
  s.apiDesfazer(r.id);
  assert.strictEqual(JSON.parse(JSON.stringify(s.apiHoje('ANA LEAL')))[0].desfeito, true);
});

test('mês que aparece como número (46296) volta a ser "Outubro/2026"', () => {
  const p = planilha();
  const reg = p.getSheetByName('Registros');
  reg.dados[1][0] = 'Outubro/2026';
  reg.dados.push([46296, '27/10/2026', '11:32', 'ANA LEAL', 'NOTEBOOK', 1, 0, 0, '11:30', '12:30']);
  const s = carregarServidor(p);
  s.apiGravar({ quem: 'ANA LEAL', hora: 9, linhas: [{ pessoa: 'ANA LEAL', categoria: 'FONE', testados: 1 }] });
  assert.strictEqual(reg.dados[4][0], 'Outubro/2026');
  assert.match(String(reg.dados[5][0]), /^[A-ZÇ][a-zç]+\/\d{4}$/);
});

test('chamada chama só quem trabalha naquela hora (costume ou aba Horários), sem supervisores', () => {
  const p = planilha();
  p.getSheetByName('Configurações').dados[1][1] = 'SUPERVISORA';
  const reg = p.getSheetByName('Registros').dados;
  const dia = (n) => new Date(Date.now() - n * 864e5);
  [1, 2, 3].forEach((n) => {
    reg.push(['Outubro/2026', dia(n), '10:10', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 2, 0, '10:30', '11:30']);
    reg.push(['Outubro/2026', dia(n), '14:10', 'JULIA GOMES', 'ROTEADOR', 2, 0, 0, '14:30', '15:30']);
    reg.push(['Outubro/2026', dia(n), '10:10', 'ANA LEAL', 'NOTEBOOK', 2, 0, 0, '10:30', '11:30']);
  });
  const s = carregarServidor(p);
  let c = JSON.parse(JSON.stringify(s.apiChamada(10)));
  assert.strictEqual(c.faixa, '10:30 – 11:30');
  assert.deepStrictEqual(c.faltam, ['DANIEL VARGAS']);
  assert.ok(p.getSheetByName('Horários'), 'cria a aba Horários');

  reg.push(['Outubro/2026', new Date(), '10:20', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 2, 0, '10:30', '11:30']);
  c = JSON.parse(JSON.stringify(s.apiChamada(10)));
  assert.deepStrictEqual(c.faltam, []);
  assert.deepStrictEqual(c.feitos, ['DANIEL VARGAS']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(14))).faltam, ['JULIA GOMES']);
  // Outro app grava hora cheia (14:00 – 15:00): conta para 13:30 – 14:30 e 14:30 – 15:30.
  reg.push(['Outubro/2026', new Date(), '14:20', 'JULIA GOMES', 'ROTEADOR', 1, 0, 0, '14:00', '15:00']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(14))).faltam, []);
  reg.pop();
  // Lançado 15:40 (logo depois da chamada das 15:30) com a faixa 15:30 – 16:30: conta para 14:30 – 15:30.
  reg.push(['Outubro/2026', new Date(), '15:40', 'JULIA GOMES', 'ROTEADOR', 1, 0, 0, '15:30', '16:30']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(14))).faltam, []);
  reg.pop();

  // Lista da vez: o primeiro a pedir fixa a lista; os outros computadores recebem a mesma, mesmo com lançamento novo.
  const vez1 = JSON.parse(JSON.stringify(s.apiListaDaVez(14, 1761665405000)));
  reg.push(['Outubro/2026', new Date(), '15:31', 'JULIA GOMES', 'ROTEADOR', 1, 0, 0, '14:30', '15:30']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiListaDaVez(14, 1761665405200))).faltam, vez1.faltam, 'mesma lista no mesmo instante');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiListaDaVez(14, 1761665525000))).faltam, [], 'instante novo lê de novo');
  reg.pop();
  // Conversa pelo microfone: a lista da vez avisa até quando esperar.
  assert.strictEqual(s.apiListaDaVez(14, 1761665999000).conversaAte, 0);
  const ate = Date.now() + 15000;
  s.apiConversa(ate);
  s.apiConversa(ate - 5000);
  assert.strictEqual(s.apiListaDaVez(14, 1761665999000).conversaAte, ate, 'fica o maior');

  const hor = p.getSheetByName('Horários').dados;
  hor.push(['julia gomes', '10:00', '12:00']);
  hor.push(['FULANO', '08:00', '12:00']);
  c = JSON.parse(JSON.stringify(s.apiChamada(10)));
  assert.deepStrictEqual(c.faltam, ['JULIA GOMES']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(14))).faltam, [], 'a aba vale mais que o costume');
  const linha = (n) => hor.find((r) => r[0] === n);
  assert.strictEqual(linha('julia gomes')[3], 'ok');
  assert.match(linha('FULANO')[3], /não achei/);
  // A aba recebe sozinha as pessoas de Configurações (menos supervisores), sem horário.
  assert.deepStrictEqual(JSON.parse(JSON.stringify(linha('DANIEL VARGAS'))), ['DANIEL VARGAS', '', '', 'sem horário: chama pelo costume']);
  assert.strictEqual(linha('ANA LEAL'), undefined, 'supervisora fica fora');
  s.apiResposta({ pessoa: 'julia gomes', hora: 10, resposta: 'nada' });
  c = JSON.parse(JSON.stringify(s.apiChamada(10)));
  assert.deepStrictEqual(c.faltam, []);
  assert.deepStrictEqual(c.nada, ['JULIA GOMES']);
  const cham = p.getSheetByName('Chamada').dados;
  assert.deepStrictEqual(JSON.parse(JSON.stringify(cham[1].slice(2))), ['10:30 – 11:30', 'JULIA GOMES', 'nada']);
  hor.push(['DANIEL VARGAS', '08:00', '18:00']);
  s.apiResposta({ pessoa: 'DANIEL VARGAS', hora: 9, resposta: 'não veio' });
  c = JSON.parse(JSON.stringify(s.apiChamada(11)));
  // Julia sai 12:00: a chamada das 12:30 (faixa 11:30 – 12:30) não chama quem já foi embora.
  assert.deepStrictEqual(c.faltam, []);
  assert.deepStrictEqual(c.naoVieram, ['DANIEL VARGAS']);
  hor.find((r) => r[0] === 'julia gomes')[2] = '12:35';
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(11))).faltam, ['JULIA GOMES'], 'até 10 minutos depois da saída ainda chama');
  // 16:45 cobra 16:30 – 17:30: quem sai 16:45 é chamado; quem sai 14:45 não é chamado às 15:30.
  hor.find((r) => r[0] === 'julia gomes')[1] = '07:00';
  hor.find((r) => r[0] === 'julia gomes')[2] = '16:45';
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(16))).faltam, ['JULIA GOMES']);
  hor.find((r) => r[0] === 'julia gomes')[2] = '14:45';
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(14))).faltam, []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiChamada(13))).faltam, ['JULIA GOMES']);
  hor.find((r) => r[0] === 'julia gomes')[1] = '10:00';
  hor.find((r) => r[0] === 'julia gomes')[2] = '12:00';
  assert.throws(() => s.apiResposta({ pessoa: 'FULANO', hora: 10, resposta: 'nada' }), /Configurações/);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiResposta({ pessoa: 'JULIA GOMES', hora: 11, resposta: 'não respondeu' }))), { ok: true });
  const r2 = s.apiResposta({ pessoa: 'JULIA GOMES', hora: 12, resposta: 'não respondeu' });
  assert.strictEqual(r2.falta, undefined, 'falta nunca é marcada sozinha');
  c = JSON.parse(JSON.stringify(s.apiChamada(9)));
  assert.deepStrictEqual(c.faltam, ['JULIA GOMES'], 'continua sendo chamada');
  assert.deepStrictEqual(c.faltas, []);
  assert.ok(!p.getSheetByName('Chamada').dados.some((r) => r[4] === 'falta'));
  assert.strictEqual(s.apiChamador('Aparelho A', 10), true);
  assert.strictEqual(s.apiChamador('Aparelho B', 10), false, 'só um aparelho fala por hora');
  assert.strictEqual(s.apiChamador('Aparelho A', 10), true);
  assert.strictEqual(s.apiChamador('Aparelho B', 11), true);
  assert.throws(() => s.apiResposta({ pessoa: 'JULIA GOMES', hora: 10, resposta: 'x' }), /inválida/);
});

test('anúncio da chamada vale para todos os aparelhos', () => {
  const s = carregarServidor(planilha());
  assert.strictEqual(s.apiAnuncio().anuncio, null);
  const r = s.apiAnunciar('Aparelho A', 'Julia Gomes, falta a sua produção das 9 horas.');
  const lido = s.apiAnuncio();
  assert.strictEqual(lido.anuncio.texto, 'Julia Gomes, falta a sua produção das 9 horas.');
  assert.strictEqual(lido.anuncio.id, r.anuncio.id);
  assert.ok(lido.anuncio.falarEm > r.agora);
});

test('agenda da chamada: salva na aba, em ordem, e volta no início', () => {
  const p = planilha();
  const s = carregarServidor(p);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiInicio().agenda)), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiSalvarAgenda(['16:50', '9:50', '10h05', '09:50']))), ['09:50', '10:05', '16:50']);
  const ag = p.getSheetByName('Agenda da chamada').dados;
  assert.deepStrictEqual(JSON.parse(JSON.stringify(ag.slice(1))), [['09:50', '09:30 – 10:30'], ['10:05', '09:30 – 10:30'], ['16:50', '16:30 – 17:30']]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiInicio().agenda)), ['09:50', '10:05', '16:50']);
  s.apiSalvarAgenda(['11:50']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(s.apiInicio().agenda)), ['11:50'], 'tirar um horário apaga a linha');
  assert.throws(() => s.apiSalvarAgenda(['25:00']), /inválido/);
});

test('faixas de meia em meia hora: 07:30 – 08:30; lançamento antigo em hora cheia conta na faixa da meia hora anterior', () => {
  const s = carregarServidor(planilha());
  assert.strictEqual(s.faixaDe_('07:30'), 7);
  assert.strictEqual(s.faixaDe_('08:00'), 7);
  assert.strictEqual(s.faixaDe_('08:29'), 7);
  assert.strictEqual(s.faixaDe_('00:10'), 23);
  assert.strictEqual(s.horaDaAgenda_('07:30'), 6);
  assert.strictEqual(s.horaDaAgenda_('08:30'), 7);
  assert.strictEqual(s.horaDaAgenda_('16:30'), 15);
  assert.strictEqual(s.horaDaAgenda_('16:45'), 16);
  assert.strictEqual(s.faixaDaAgenda_('08:30'), '07:30 – 08:30');
  const r = s.apiGravar({ quem: 'ANA LEAL', hora: 7, linhas: [{ pessoa: 'ANA LEAL', categoria: 'NOTEBOOK', testados: 1 }] });
  assert.strictEqual(r.faixa, '07:30 – 08:30');
});

test('vários aparelhos: "não respondeu" da mesma pessoa e faixa entra uma vez só', () => {
  const p = planilha();
  const s = carregarServidor(p);
  s.apiResposta({ pessoa: 'JULIA GOMES', hora: 9, resposta: 'não respondeu' });
  s.apiResposta({ pessoa: 'JULIA GOMES', hora: 9, resposta: 'não respondeu' });
  const linhas = p.getSheetByName('Chamada').dados.slice(1).filter((r) => r[4] === 'não respondeu');
  assert.strictEqual(linhas.length, 1);
});

test('chamada com conversa: um de cada vez, só um computador pega a vez, quem não responde vai para o fim', () => {
  const s = carregarServidor(planilha());
  let faltam = ['ANA LEAL', 'DANIEL VARGAS', 'JULIA GOMES'];
  s.apiChamada = () => ({ faixa: 'x', faltam: faltam.slice(), feitos: [], nada: [], naoVieram: [], faltas: [] });
  // Sem computador ouvindo: chamada de todos os nomes, sem conversa.
  assert.strictEqual(s.apiComecarVez(14, 1000).conversa, false);
  s.apiAnuncio('PC1', true);
  let e = s.apiComecarVez(14, 5000);
  assert.strictEqual(e.conversa, true);
  assert.strictEqual(e.pessoa, 'ANA LEAL');
  assert.strictEqual(s.apiComecarVez(14, 5000).pessoa, 'ANA LEAL', 'o mesmo começo para todos');
  assert.strictEqual(s.apiPegarVez(14, 5000, 'PC1').ok, true);
  assert.strictEqual(s.apiPegarVez(14, 5000, 'PC2').ok, false, 'só um computador conversa');
  assert.strictEqual(s.apiFimDaVez(14, 5000, 'PC2').pessoa, 'ANA LEAL', 'outro computador não encerra a vez');
  faltam = ['DANIEL VARGAS', 'JULIA GOMES'];
  e = s.apiFimDaVez(14, 5000, 'PC1');
  assert.strictEqual(e.pessoa, 'DANIEL VARGAS');
  assert.strictEqual(e.anunciar, false);
  // Daniel não responde: vai para o fim; Julia é a próxima; depois volta ao começo e anuncia a lista.
  e = s.apiFimDaVez(14, e.falarEm, '');
  assert.strictEqual(e.pessoa, 'JULIA GOMES');
  e = s.apiFimDaVez(14, e.falarEm, '');
  assert.strictEqual(e.pessoa, 'DANIEL VARGAS');
  assert.strictEqual(e.anunciar, true, 'voltou ao começo da lista');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(e.fila)), ['DANIEL VARGAS', 'JULIA GOMES']);
  faltam = [];
  e = s.apiFimDaVez(14, e.falarEm, '');
  assert.strictEqual(e.pessoa, null, 'todos responderam');
});

test('conversa com a IA: manda o histórico, aceita só categorias da planilha', () => {
  const s = carregarServidor(planilha());
  assert.strictEqual(s.apiInicio().ia, false);
  assert.throws(() => s.apiConversar({ pessoa: 'ANA LEAL', hora: 9, historico: [] }), /IA_CHAVE/);
  s.SEGREDOS.iaChave = 'chave-teste';
  assert.strictEqual(s.apiInicio().ia, true);
  let enviado = null, resposta = { fala: 'Ana, 3 notebooks finalizados. Está certo?', acao: 'confirmar', linhas: [{ categoria: 'notebook', finalizados: 3 }] };
  const modelos = [];
  s.UrlFetchApp = { fetch: (url, o) => {
    modelos.push(url.match(/models\/([^:]+)/)[1]);
    if (modelos.length === 1) return { getResponseCode: () => 503, getContentText: () => 'ocupado' };
    enviado = JSON.parse(o.payload);
    assert.strictEqual(o.headers['x-goog-api-key'], 'chave-teste');
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: JSON.stringify(resposta) }] } }] }) };
  } };
  let r = JSON.parse(JSON.stringify(s.apiConversar({ pessoa: 'ANA LEAL', hora: 9, historico: [
    { quem: 'ia', texto: 'Ana Leal, venha lançar' }, { quem: 'pessoa', texto: 'é... hum... 3 notebook finalizado' }] })));
  assert.strictEqual(modelos.length, 2, 'modelo ocupado: tenta o próximo');
  assert.deepStrictEqual(r, { fala: 'Ana, 3 notebooks finalizados. Está certo?', acao: 'confirmar', linhas: [{ categoria: 'NOTEBOOK', testados: 0, finalizados: 3, defeitos: 0 }] });
  assert.strictEqual(enviado.contents[0].role, 'user', 'começa pela pessoa');
  assert.match(enviado.systemInstruction.parts[0].text, /CAIXA DE SOM/);
  assert.match(enviado.systemInstruction.parts[0].text, /Ana/);
  resposta = { fala: 'Perfeito, enviando.', acao: 'gravar', linhas: [{ categoria: 'GELADEIRA', finalizados: 3 }] };
  r = s.apiConversar({ pessoa: 'ANA LEAL', hora: 9, historico: [{ quem: 'pessoa', texto: 'sim' }] });
  assert.strictEqual(r.acao, 'perguntar', 'categoria fora da planilha não grava');
  assert.match(r.fala, /GELADEIRA/);
});

test('conversa com a IA: Ollama Cloud é a reserva quando o Gemini falha', () => {
  const s = carregarServidor(planilha());
  s.SEGREDOS.ollamaChave = 'ollama-teste';
  assert.strictEqual(s.apiInicio().ia, true, 'só com a chave do Ollama também conversa');
  s.SEGREDOS.iaChave = 'gemini-teste';
  const urls = [];
  s.UrlFetchApp = { fetch: (url, o) => {
    urls.push(url);
    if (/googleapis/.test(url)) return { getResponseCode: () => 429, getContentText: () => 'cota' };
    assert.strictEqual(o.headers.Authorization, 'Bearer ollama-teste');
    const p = JSON.parse(o.payload);
    assert.strictEqual(p.messages[0].role, 'system');
    assert.strictEqual(p.messages[p.messages.length - 1].content, '2 caixas de som testadas');
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ message: { content: JSON.stringify({ fala: 'Duas caixas de som testadas, certo?', acao: 'confirmar', linhas: [{ categoria: 'caixa de som', testados: 2 }] }) } }) };
  } };
  const r = JSON.parse(JSON.stringify(s.apiConversar({ pessoa: 'DANIEL VARGAS', hora: 9, historico: [{ quem: 'pessoa', texto: '2 caixas de som testadas' }] })));
  assert.strictEqual(urls.length, 4, '3 modelos do Gemini e depois o Ollama');
  assert.deepStrictEqual(r.linhas, [{ categoria: 'CAIXA DE SOM', testados: 2, finalizados: 0, defeitos: 0 }]);
});

test('site próprio: doPost chama só funções api e devolve JSON', () => {
  const s = carregarServidor(planilha());
  const post = (corpo) => JSON.parse(s.doPost({ postData: { contents: JSON.stringify(corpo) } }).texto);
  assert.ok(post({ fn: 'apiInicio', args: [] }).r.pessoas.length);
  assert.match(post({ fn: 'config_', args: [] }).erro, /desconhecida/);
  assert.match(post({ fn: 'autorizarIA' }).erro, /desconhecida/);
  s.SEGREDOS.planilhaId = '';
  assert.match(post({ fn: 'apiInicio' }).erro, /configurado/);
});
