const test = require('node:test');
const assert = require('node:assert');
const { criarPlanilha, carregarServidor } = require('./gas');

const CAB = ['', 'DATA', 'HORA DE REGISTRO', 'PESSOA', 'CATEGORIA', 'QTD. TESTADOS', 'QTD. FINALIZADOS', 'QTD. DEFEITOS', 'INÍCIO', 'FIM'];

function planilha() {
  const hoje = new Date();
  return criarPlanilha({
    Registros: [
      CAB,
      ['Outubro/2026', '06/10/2026', '08:35', 'ANA LEAL', 'NOTEBOOK', 0, 1, 0, '08:00', '09:00'],
      ['Outubro/2026', hoje, '08:36', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 3, 0, '08:00', '09:00'],
      ['Outubro/2026', hoje, '09:36', 'DANIEL VARGAS', 'CAIXA DE SOM', 0, 4, 0, '09:00', '10:00']
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
  assert.strictEqual(r.faixa, '10:00 – 11:00');
  const reg = p.getSheetByName('Registros').dados;
  assert.strictEqual(reg.length, 6);
  const nova = reg[4];
  assert.match(nova[0], /^[A-ZÇ][a-zç]+\/\d{4}$/);
  assert.match(nova[1], /^\d{2}\/\d{2}\/\d{4}$/);
  assert.match(nova[2], /^\d{2}:\d{2}$/);
  assert.deepStrictEqual(nova.slice(3), ['JULIA GOMES', 'ROTEADOR', 4, 0, 0, '10:00', '11:00']);

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
});
