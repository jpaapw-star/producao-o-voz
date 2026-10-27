const test = require('node:test');
const assert = require('node:assert');
const { carregarLeitor } = require('./carregar');

const L = carregarLeitor();
const pessoas = ['ANA LEAL', 'JULIA GOMES', 'JOÃO PEDRO', 'ANA CLARA', 'CAUÃ ENRICO', 'DANIEL VARGAS', 'JOÃO SUETO',
  'SAMUEL BRITO', 'ARTHUR SILVÉRIO', 'ARTUR ANTUNES', 'JANINE RODRIGUES', 'GABRIEL VIEIRA', 'GABRIEL MARCUCCI',
  'MARIANA DE PAULA'];
const categorias = ['NOTEBOOK', 'ROTEADOR', 'CAIXA DE SOM', 'CAIXA DE FERRAMENTAS', 'FONE', 'VOLANTE', 'MONITOR',
  'FECHADURA DIGITAL', 'PERIFERICOS', 'VIDEO GAMES', 'CABO USB', 'CABO HDMI'];

function leitor(extra) {
  return L.criarLeitor(Object.assign({ pessoas, categorias, habitos: {}, ativos: [] }, extra || {}));
}
const resumo = (r) => r.linhas.map((l) => [l.pessoa, l.categoria, l.testados, l.finalizados, l.defeitos]);

test('várias pessoas numa frase só', () => {
  const r = leitor().ler('Ana Leal notebook 3 testados 2 finalizados. Julia roteador 4 testados. Cauã fechadura 5 finalizados 1 defeito.', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(r), [
    ['ANA LEAL', 'NOTEBOOK', 3, 2, 0],
    ['JULIA GOMES', 'ROTEADOR', 4, 0, 0],
    ['CAUÃ ENRICO', 'FECHADURA DIGITAL', 0, 5, 1]
  ]);
  assert.ok(r.linhas.every((l) => l.problemas.length === 0));
});

test('números por extenso, plural e "caixa de som"', () => {
  const r = leitor().ler('Daniel caixas de som três finalizados e um com defeito', '');
  assert.deepStrictEqual(resumo(r), [['DANIEL VARGAS', 'CAIXA DE SOM', 0, 3, 1]]);
  const r2 = leitor().ler('vinte e cinco notebooks finalizados', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(r2), [['JOÃO PEDRO', 'NOTEBOOK', 0, 25, 0]]);
});

test('sem nome vai para quem está lançando', () => {
  const r = leitor().ler('testados 3 finalizados 2 roteadores', 'JULIA GOMES');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 3, 2, 0]]);
  assert.strictEqual(r.linhas[0].supostos.pessoa, true);
});

test('"eu" é quem está lançando', () => {
  const r = leitor().ler('Julia 2 roteadores testados e eu 4 notebooks finalizados', 'ANA CLARA');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 2, 0, 0], ['ANA CLARA', 'NOTEBOOK', 0, 4, 0]]);
});

test('mesma pessoa com duas categorias', () => {
  const r = leitor().ler('Cauã fechadura digital 5 finalizados roteador 2 finalizados', '');
  assert.deepStrictEqual(resumo(r), [['CAUÃ ENRICO', 'FECHADURA DIGITAL', 0, 5, 0], ['CAUÃ ENRICO', 'ROTEADOR', 0, 2, 0]]);
});

test('Arthur e Artur: pergunta, a não ser que o histórico resolva', () => {
  const r = leitor().ler('Arthur monitores 2 finalizados', '');
  assert.strictEqual(r.linhas[0].pessoa, null);
  assert.deepStrictEqual(r.linhas[0].pessoaOpcoes, ['ARTHUR SILVÉRIO', 'ARTUR ANTUNES']);
  assert.deepStrictEqual(r.linhas[0].problemas, ['Qual arthur?']);

  const r2 = leitor({ habitos: { 'ARTHUR SILVÉRIO': { categorias: ['MONITOR'] } } }).ler('Arthur monitores 2 finalizados', '');
  assert.strictEqual(r2.linhas[0].pessoa, 'ARTHUR SILVÉRIO');
  assert.strictEqual(r2.linhas[0].supostos.pessoa, true);

  const r3 = leitor().ler('Artur Antunes notebook 1 finalizado', '');
  assert.strictEqual(r3.linhas[0].pessoa, 'ARTUR ANTUNES');
});

test('Gabriel: quem trabalhou nos últimos dias decide', () => {
  const r = leitor({ ativos: ['GABRIEL VIEIRA'] }).ler('Gabriel notebook 2 finalizados', '');
  assert.strictEqual(r.linhas[0].pessoa, 'GABRIEL VIEIRA');
});

test('sem categoria nem tipo: usa o costume da pessoa', () => {
  const l = leitor({ habitos: { 'DANIEL VARGAS': { categorias: ['CAIXA DE SOM'], tipo: 'finalizados' } } });
  const r = l.ler('Daniel 4', '');
  assert.deepStrictEqual(resumo(r), [['DANIEL VARGAS', 'CAIXA DE SOM', 0, 4, 0]]);
  assert.deepStrictEqual(r.linhas[0].supostos, { categoria: true, tipo: 'finalizados' });
});

test('número sem tipo e sem costume vira pergunta', () => {
  const r = leitor().ler('Julia roteador 4', '');
  assert.deepStrictEqual(r.linhas[0].semTipo, [4]);
  assert.deepStrictEqual(r.linhas[0].problemas, ['Testados, finalizados ou defeitos?']);
});

test('"caixa" sozinha: o histórico escolhe entre as caixas', () => {
  const r = leitor().ler('Daniel caixa 3 finalizadas', '');
  assert.deepStrictEqual(r.linhas[0].categoriaOpcoes, ['CAIXA DE SOM', 'CAIXA DE FERRAMENTAS']);
  const r2 = leitor({ habitos: { 'DANIEL VARGAS': { categorias: ['VOLANTE', 'CAIXA DE SOM'] } } }).ler('Daniel caixa 3 finalizadas', '');
  assert.strictEqual(r2.linhas[0].categoria, 'CAIXA DE SOM');
});

test('erros de ditado comuns', () => {
  const r = leitor().ler('Janine Rodriguez notbook 6 finalizados; Samuel volantes 1 testado', '');
  assert.deepStrictEqual(resumo(r), [['JANINE RODRIGUES', 'NOTEBOOK', 0, 6, 0], ['SAMUEL BRITO', 'VOLANTE', 1, 0, 0]]);
});

test('palavras desconhecidas aparecem para conferir', () => {
  const r = leitor().ler('Fulano notebook 2 finalizados', 'ANA LEAL');
  assert.deepStrictEqual(r.naoEntendi, ['fulano']);
});

test('frase repetida soma na mesma linha', () => {
  const r = leitor().ler('Julia roteador 2 testados. Julia roteador 3 testados', '');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 5, 0, 0]]);
});

test('Mariana de Paula (nome com "de")', () => {
  const r = leitor().ler('Mariana de Paula periféricos 10 testados', '');
  assert.deepStrictEqual(resumo(r), [['MARIANA DE PAULA', 'PERIFERICOS', 10, 0, 0]]);
});

test('frase do print: produto fora da lista, "testados e finalizados", "de feito" e "validados"', () => {
  const r = leitor({ categorias: categorias.concat(['TECLADO']) }).ler(
    'Artur Antunes 3 webcams finalizadas 4 notebooks testados e finalizados 5 fones de feito e 3 teclados validados', '');
  assert.deepStrictEqual(resumo(r), [
    ['ARTUR ANTUNES', null, 0, 3, 0],
    ['ARTUR ANTUNES', 'NOTEBOOK', 4, 4, 0],
    ['ARTUR ANTUNES', 'FONE', 0, 0, 5],
    ['ARTUR ANTUNES', 'TECLADO', 3, 0, 0]
  ]);
  assert.match(r.linhas[0].problemas[0], /webcams/);
  assert.deepStrictEqual(r.naoEntendi, []);
});

test('"testados e finalizados 4" (tipo antes do número) vale para os dois', () => {
  const r = leitor().ler('Julia roteador testados e finalizados 4', '');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 4, 4, 0]]);
});

test('conversa paralela no meio é ignorada', () => {
  const r = leitor().ler('Ana Leal notebook 3 testados. Ô Julia, você viu onde ficou a chave do armário? ' +
    'Pega lá pra mim por favor. Daniel caixa de som 2 finalizados', '');
  assert.deepStrictEqual(resumo(r), [['ANA LEAL', 'NOTEBOOK', 3, 0, 0], ['DANIEL VARGAS', 'CAIXA DE SOM', 0, 2, 0]]);
  assert.strictEqual(r.ignorados.length, 2);
});

test('nome dito na conversa não rouba os números de quem estava sendo lançado', () => {
  const r = leitor().ler('Daniel caixa de som 2 finalizados. Julia espera um pouco. 3 testados', '');
  assert.deepStrictEqual(resumo(r), [['DANIEL VARGAS', 'CAIXA DE SOM', 3, 2, 0]]);
});

test('"esquece" descarta a frase', () => {
  const r = leitor().ler('Julia roteador 4 testados. Não, 5 testados, esquece. Ana Leal notebook 1 finalizado', '');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 4, 0, 0], ['ANA LEAL', 'NOTEBOOK', 0, 1, 0]]);
});

test('conversa com número e muitas palavras estranhas é ignorada', () => {
  const r = leitor().ler('Julia roteador 4 testados. Amanhã chego umas 8 horas porque tenho médico', '');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 4, 0, 0]]);
});
