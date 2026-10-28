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

test('palavra estranha no lugar do nome pergunta de quem é, em vez de lançar para quem está no aparelho', () => {
  const r = leitor().ler('Fulano notebook 2 finalizados', 'ANA LEAL');
  assert.strictEqual(r.linhas[0].pessoa, null);
  assert.deepStrictEqual(r.linhas[0].problemas, ['“fulano” não está nas pessoas. Toque para escolher.']);
  assert.strictEqual(leitor().ler('notebook 2 finalizados galera', 'ANA LEAL').linhas[0].pessoa, 'ANA LEAL');
});

test('frase repetida soma na mesma linha', () => {
  const r = leitor().ler('Julia roteador 2 testados. Julia roteador 3 testados', '');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'ROTEADOR', 5, 0, 0]]);
});

test('Mariana de Paula (nome com "de")', () => {
  const r = leitor().ler('Mariana de Paula periféricos 10 testados', '');
  assert.deepStrictEqual(resumo(r), [['MARIANA DE PAULA', 'PERIFERICOS', 10, 0, 0]]);
});

test('frase do print: produto fora da lista, "testados e finalizados" e "de feito"', () => {
  const r = leitor({ categorias: categorias.concat(['TECLADO']) }).ler(
    'Artur Antunes 3 webcams finalizadas 4 notebooks testados e finalizados 5 fones de feito e 3 teclados testados', '');
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

test('validação é categoria: "notebooks validados" e "validação notebook" com testados ou defeitos', () => {
  const l = leitor({ categorias: categorias.concat(['VALIDAÇÃO NOTEBOOK']) });
  assert.deepStrictEqual(resumo(l.ler('Gabriel Vieira 6 notebooks validados 1 defeito', '')),
    [['GABRIEL VIEIRA', 'VALIDAÇÃO NOTEBOOK', 6, 0, 1]]);
  assert.deepStrictEqual(resumo(l.ler('Cauã validação notebook 4 testados 2 defeitos', '')),
    [['CAUÃ ENRICO', 'VALIDAÇÃO NOTEBOOK', 4, 0, 2]]);
  assert.deepStrictEqual(resumo(l.ler('Cauã validação de notebook 5', '')),
    [['CAUÃ ENRICO', 'VALIDAÇÃO NOTEBOOK', 5, 0, 0]]);
  // Notebook comum continua notebook.
  assert.deepStrictEqual(resumo(l.ler('Ana Leal 2 notebooks finalizados', '')), [['ANA LEAL', 'NOTEBOOK', 0, 2, 0]]);
});

test('"videogame" junto casa com "VÍDEO GAME" e validações de várias categorias', () => {
  const cats = ['NOTEBOOK', 'FONE', 'VÍDEO GAME', 'VALIDAÇÃO NOTEBOOK', 'VALIDAÇÃO ROTEADOR', 'VALIDAÇÃO VÍDEO GAME', 'ROTEADOR'];
  const l = leitor({ categorias: cats });
  assert.deepStrictEqual(resumo(l.ler('Gabriel Vieira validação notebooks 2 testados um defeito Validação videogame Um testado 3 defeitos Notebooks 3 testados 2 finalizados', '')), [
    ['GABRIEL VIEIRA', 'VALIDAÇÃO NOTEBOOK', 2, 0, 1],
    ['GABRIEL VIEIRA', 'VALIDAÇÃO VÍDEO GAME', 1, 0, 3],
    ['GABRIEL VIEIRA', 'NOTEBOOK', 3, 2, 0]
  ]);
  const r = l.ler('2 notebooks testados um notebook defeito 3 validações videogames testadas uma validação videogame defeito 3 fones finalizados', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(r), [
    ['JOÃO PEDRO', 'NOTEBOOK', 2, 0, 1],
    ['JOÃO PEDRO', 'VALIDAÇÃO VÍDEO GAME', 3, 0, 1],
    ['JOÃO PEDRO', 'FONE', 0, 3, 0]
  ]);
  assert.deepStrictEqual(resumo(l.ler('Julia 4 video games testados', '')), [['JULIA GOMES', 'VÍDEO GAME', 4, 0, 0]]);
  assert.deepStrictEqual(resumo(l.ler('Julia 2 videogames validados', '')), [['JULIA GOMES', 'VALIDAÇÃO VÍDEO GAME', 2, 0, 0]]);
});

test('categoria com número no nome (PlayStation 5)', () => {
  const r = L.criarLeitor({ pessoas, categorias: categorias.concat(['PLAYSTATION 5', 'PLAYSTATION 4']), habitos: {}, ativos: [] })
    .ler('Julia playstation 5 três testados', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(r), [['JULIA GOMES', 'PLAYSTATION 5', 3, 0, 0]]);
});

test('validação com finalizados é problema e quantidade alta vira aviso', () => {
  assert.deepStrictEqual(L.problemas({ pessoa: 'ANA LEAL', categoria: 'VALIDAÇÃO NOTEBOOK', testados: 0, finalizados: 2, defeitos: 0 }),
    ['Validação é só testados ou defeitos']);
  assert.deepStrictEqual(L.avisos({ testados: 150, finalizados: 3, defeitos: 0 }), ['Confira: 150 testados é muito']);
  assert.deepStrictEqual(L.avisos({ testados: 100, finalizados: 3, defeitos: 0 }), []);
});

test('apelidos da aba Apelidos valem para pessoas e categorias', () => {
  const l = L.criarLeitor({ pessoas, categorias: categorias.concat(['PLAYSTATION 5']), habitos: {}, ativos: [],
    apelidos: [['Gabi', 'GABRIEL VIEIRA'], ['PS5', 'PLAYSTATION 5'], ['cam', 'NOTEBOOK'], ['Zé', 'FULANO FORA DA LISTA']] });
  assert.deepStrictEqual(resumo(l.ler('Gabi 3 ps5 testados', 'JOÃO PEDRO')), [['GABRIEL VIEIRA', 'PLAYSTATION 5', 3, 0, 0]]);
  assert.deepStrictEqual(resumo(l.ler('Gabi 3 PS 5 testados', 'JOÃO PEDRO')), [['GABRIEL VIEIRA', 'PLAYSTATION 5', 3, 0, 0]]);
});

test('"mouses" é MOUSE e "web cans" separado é WEBCAM (frase do João)', () => {
  const l = L.criarLeitor({ pessoas, categorias: ['NOTEBOOK', 'MOUSE', 'WEBCAM', 'VALIDAÇÃO NOTEBOOK', 'VALIDAÇÃO VIDEO GAME', 'VIDEO GAMES'], habitos: {}, ativos: [] });
  const r = l.ler('Júlia Gomes 4 notebooks testados 5 notebooks defeitos 3 mouses finalizados 4 mouses testados e finalizados 8 webcans testadas 13 web cans defeitos Validação notebook 5 testadas validação videogame 3 defeitos e 5 testadas Aquilo ali faz o seguinte', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(r), [
    ['JULIA GOMES', 'NOTEBOOK', 4, 0, 5],
    ['JULIA GOMES', 'MOUSE', 4, 7, 0],
    ['JULIA GOMES', 'WEBCAM', 8, 0, 13],
    ['JULIA GOMES', 'VALIDAÇÃO NOTEBOOK', 5, 0, 0],
    ['JULIA GOMES', 'VALIDAÇÃO VIDEO GAME', 5, 0, 3]
  ]);
  assert.ok(r.linhas.every((x) => x.problemas.length === 0));
  // "meses" continua sendo "mes", e não quebra palavras que já casavam.
  assert.deepStrictEqual(resumo(leitor().ler('Ana Leal 2 roteadores testados', '')), [['ANA LEAL', 'ROTEADOR', 2, 0, 0]]);
});

test('segunda rodada: correção falada, dúzia, horário, frase repetida e "ontem"', () => {
  const l = L.criarLeitor({ pessoas, categorias: ['NOTEBOOK', 'MOUSE', 'FONE', 'ROTEADOR'], habitos: {}, ativos: [] });
  const r = (t) => resumo(l.ler(t, 'JOÃO PEDRO'));
  assert.deepStrictEqual(r('Julia 3 notebooks finalizados não 4 notebooks finalizados'), [['JULIA GOMES', 'NOTEBOOK', 0, 4, 0]]);
  assert.deepStrictEqual(r('Julia 3 notebooks, quer dizer, 4 finalizados'), [['JULIA GOMES', 'NOTEBOOK', 0, 4, 0]]);
  assert.deepStrictEqual(r('Julia notebook não tem defeito, 3 finalizados'), [['JULIA GOMES', 'NOTEBOOK', 0, 3, 0]]);
  assert.deepStrictEqual(r('Julia 3 notebooks finalizados, não teve nenhum defeito'), [['JULIA GOMES', 'NOTEBOOK', 0, 3, 0]]);
  assert.deepStrictEqual(r('meia dúzia de fones testados'), [['JOÃO PEDRO', 'FONE', 6, 0, 0]]);
  assert.deepStrictEqual(r('Julia duas dúzias de mouses finalizados'), [['JULIA GOMES', 'MOUSE', 0, 24, 0]]);
  assert.deepStrictEqual(r('Julia três notebooks finalizados na hora das nove'), [['JULIA GOMES', 'NOTEBOOK', 0, 3, 0]]);
  assert.deepStrictEqual(r('das 9 às 10 Julia 5 fones testados'), [['JULIA GOMES', 'FONE', 5, 0, 0]]);
  assert.deepStrictEqual(r('Julia 2 fones testados às 9 horas'), [['JULIA GOMES', 'FONE', 2, 0, 0]]);
  const rep = l.ler('Julia 3 notebooks finalizados. Julia 3 notebooks finalizados', 'JOÃO PEDRO');
  assert.deepStrictEqual(resumo(rep), [['JULIA GOMES', 'NOTEBOOK', 0, 6, 0]]);
  assert.match(rep.avisos[0], /apareceu duas vezes/);
  assert.match(l.ler('Julia 2 fones testados ontem', 'JOÃO PEDRO').avisos[0], /“ontem”/);
});

test('supervisor sem dizer nome: pergunta de quem é; colaborador sem nome: vai para ele', () => {
  const l = L.criarLeitor({ pessoas, categorias, habitos: {}, ativos: [], supervisores: ['JOÃO PEDRO'] });
  const sup = l.ler('3 notebooks finalizados', 'JOÃO PEDRO');
  assert.strictEqual(sup.linhas[0].pessoa, null);
  assert.deepStrictEqual(sup.linhas[0].problemas, ['De quem? Toque para escolher.']);
  assert.strictEqual(l.ler('eu 3 notebooks finalizados', 'JOÃO PEDRO').linhas[0].pessoa, 'JOÃO PEDRO');
  assert.strictEqual(l.ler('3 notebooks finalizados', 'ANA LEAL').linhas[0].pessoa, 'ANA LEAL');
});

test('avisos: máximo da pessoa, tipo fora do costume e coisas deduzidas', () => {
  const dados = { maximos: { 'JULIA GOMES|ROTEADOR': 30 }, habitos: { 'JULIA GOMES': { categorias: ['ROTEADOR'], tipo: 'finalizados' } } };
  assert.deepStrictEqual(L.avisos({ pessoa: 'JULIA GOMES', categoria: 'ROTEADOR', testados: 40, finalizados: 0, defeitos: 0, supostos: {} }, dados), [
    'Confira: 40 testados passa do máximo dessa pessoa nessa categoria (30)',
    'Confira o tipo: Julia costuma lançar finalizados'
  ]);
  assert.deepStrictEqual(L.avisos({ pessoa: 'JULIA GOMES', categoria: 'ROTEADOR', testados: 0, finalizados: 20, defeitos: 1, supostos: {} }, dados), []);
  assert.deepStrictEqual(L.avisos({ pessoa: 'ANA LEAL', categoria: 'FONE', testados: 2, finalizados: 0, defeitos: 0, supostos: { pessoa: true, categoria: true } }, dados),
    ['Nome deduzido: confira', 'Categoria deduzida: confira']);
});

test('erros de escrita do ditado: "perfeito", "estados", "jogos", números escritos errado', () => {
  const cats = ['FONE DE OUVIDO', 'TECLADO', 'VIDEO GAMES', 'NOTEBOOK', 'MOUSE', 'FONE'];
  const l = L.criarLeitor({ pessoas: pessoas.concat(['MARIA GABRIELA']), categorias: cats, habitos: {}, ativos: [] });
  const r = (t) => resumo(l.ler(t, 'JOÃO PEDRO'));
  // Frase do print do João.
  assert.deepStrictEqual(r('Maria Gabriela 2 testados fone de ouvido Um defeito 4 finalizados 2 teclados testados um teclado perfeito 3 teclados testados e finalizados Jogos um defeito e 2 estados'), [
    ['MARIA GABRIELA', 'FONE DE OUVIDO', 2, 4, 1],
    ['MARIA GABRIELA', 'TECLADO', 5, 3, 1],
    ['MARIA GABRIELA', 'VIDEO GAMES', 2, 0, 1]
  ]);
  assert.deepStrictEqual(r('Julia 3 teclados finalisados 2 testadas 1 com efeito'), [['JULIA GOMES', 'TECLADO', 2, 3, 1]]);
  assert.deepStrictEqual(r('Julia dose fones testados e treis defeituosos'), [['JULIA GOMES', 'FONE', 12, 0, 3]]);
  assert.deepStrictEqual(r('Julia quarto notebooks testados'), [['JULIA GOMES', 'NOTEBOOK', 4, 0, 0]]);
  // "quarto" sem categoria depois não é número; "estado" solto não copia quantidade.
  assert.deepStrictEqual(r('Julia no quarto 3 notebooks testados'), [['JULIA GOMES', 'NOTEBOOK', 3, 0, 0]]);
  assert.deepStrictEqual(r('Julia 3 mouses testados, estado de São Paulo'), [['JULIA GOMES', 'MOUSE', 3, 0, 0]]);
  assert.deepStrictEqual(l.ler('tudo perfeito por aqui', 'JOÃO PEDRO').linhas, []);
});

test('bateria de erros de escrita com nomes e produtos da planilha', () => {
  const ps = ['ÁLVARO COELHO', 'ANA CLARA', 'ANA LEAL', 'ARTHUR SILVÉRIO', 'ARTUR ANTUNES', 'CAUÃ ENRICO', 'EDUARDA CRISTINA',
    'GABRIEL MARCUCCI', 'GABRIEL VIEIRA', 'HEVILLYN EMANUELE', 'IGOR GOMES', 'JANINE RODRIGUES', 'JOÃO PEDRO', 'JOÃO SUETO',
    'JULIA GOMES', 'LAVINEA LOPES', 'MARIA GABRIELA', 'SAMUEL BRITO', 'ANNA LUIZA', 'ARTHUR PRATES', 'DANIEL VARGAS'];
  const cs = ['NOTEBOOK', 'AIR COOLER', 'APPLE TV', 'CABO HDMI', 'CABO USB', 'CONFERÊNCIA DE KIT', 'NINTENDO SWITCH OLED', 'TRACKPAD',
    'HEADSET', 'PLAYSTATION 5', 'MOUSE', 'TECLADO', 'FONE', 'ANTENA', 'CAIXA DE SOM', 'CAIXA DE FERRAMENTAS', 'MONITOR', 'VOLANTE'];
  const l = L.criarLeitor({ pessoas: ps, categorias: cs, habitos: {}, ativos: [] });
  const um = (t) => { const r = l.ler(t, 'JOÃO PEDRO').linhas; assert.strictEqual(r.length, 1, t); return [r[0].pessoa, r[0].categoria, r[0].testados, r[0].finalizados, r[0].defeitos]; };
  const casos = [
    ['Kauã 3 notebooks testados', ['CAUÃ ENRICO', 'NOTEBOOK', 3, 0, 0]],
    ['Cauan 3 notebooks testados', ['CAUÃ ENRICO', 'NOTEBOOK', 3, 0, 0]],
    ['Evelyn 2 mouses finalizados', ['HEVILLYN EMANUELE', 'MOUSE', 0, 2, 0]],
    ['Hevelin 2 mauses finalizados', ['HEVILLYN EMANUELE', 'MOUSE', 0, 2, 0]],
    ['Lavínia 4 teclados testados', ['LAVINEA LOPES', 'TECLADO', 4, 0, 0]],
    ['Marcuci 2 fones testados', ['GABRIEL MARCUCCI', 'FONE', 2, 0, 0]],
    ['Alvaro 3 antenas testadas', ['ÁLVARO COELHO', 'ANTENA', 3, 0, 0]],
    ['Ana Luisa 2 monitores testados', ['ANNA LUIZA', 'MONITOR', 2, 0, 0]],
    ['Suéto 2 fones testados', ['JOÃO SUETO', 'FONE', 2, 0, 0]],
    ['Igor 2 cabos u esse bê testados', ['IGOR GOMES', 'CABO USB', 2, 0, 0]],
    ['Igor 2 cabos agá dê eme i finalizados', ['IGOR GOMES', 'CABO HDMI', 0, 2, 0]],
    ['Daniel 1 Nintendo Suite finalizado', ['DANIEL VARGAS', 'NINTENDO SWITCH OLED', 0, 1, 0]],
    ['Daniel 2 notebuques testados', ['DANIEL VARGAS', 'NOTEBOOK', 2, 0, 0]],
    ['Daniel 2 rédisete testados', ['DANIEL VARGAS', 'HEADSET', 2, 0, 0]],
    ['Daniel 1 pleisteichon 5 testado', ['DANIEL VARGAS', 'PLAYSTATION 5', 1, 0, 0]],
    ['Daniel 3 épou tivi testados', ['DANIEL VARGAS', 'APPLE TV', 3, 0, 0]],
    ['Daniel 2 trequipad finalizados', ['DANIEL VARGAS', 'TRACKPAD', 0, 2, 0]],
    ['Daniel 2 er cooler testados', ['DANIEL VARGAS', 'AIR COOLER', 2, 0, 0]],
    ['Daniel 4 conferência de quite testados', ['DANIEL VARGAS', 'CONFERÊNCIA DE KIT', 4, 0, 0]],
    ['Julia. 3 notebooks. Testados.', ['JULIA GOMES', 'NOTEBOOK', 3, 0, 0]],
    ['Pronto, Julia 3 teclados testados', ['JULIA GOMES', 'TECLADO', 3, 0, 0]],
    ['Julia 3 teclados testados, pronto.', ['JULIA GOMES', 'TECLADO', 3, 0, 0]],
    ['No final do dia Julia 3 teclados testados', ['JULIA GOMES', 'TECLADO', 3, 0, 0]],
    ['Samuel cincu fones testados', ['SAMUEL BRITO', 'FONE', 5, 0, 0]],
    ['Samuel seti fones finalizados', ['SAMUEL BRITO', 'FONE', 0, 7, 0]],
    ['Hum Samuel 2 fones testados', ['SAMUEL BRITO', 'FONE', 2, 0, 0]],
    ['Samuel 2 caixas de ferramenta testadas', ['SAMUEL BRITO', 'CAIXA DE FERRAMENTAS', 2, 0, 0]]
  ];
  casos.forEach(([frase, esperado]) => assert.deepStrictEqual(um(frase), esperado, frase));
});

test('lista grande continua rápida', () => {
  const cs = []; for (let i = 0; i < 140; i++) cs.push('PRODUTO EXEMPLO ' + String.fromCharCode(65 + (i % 26)) + String.fromCharCode(65 + Math.floor(i / 26)) + ' LED USB');
  const l = L.criarLeitor({ pessoas, categorias: cs.concat(['NOTEBOOK', 'TECLADO']), habitos: {}, ativos: [] });
  const t = Date.now();
  l.ler('Julia 3 notebooks testados 2 teclados finalizados um teclado perfeito '.repeat(4), 'JOÃO PEDRO');
  assert.ok(Date.now() - t < 500, 'demorou ' + (Date.now() - t) + ' ms');
});

test('frase real: "controle nintendo", item sem número e "atualizada" no lugar de finalizada', () => {
  const cats = ['CONTROLE', 'CONTROLE REMOTO', 'NINTENDO SWITCH OLED', 'FONE', 'FONE DE OUVIDO', 'CARREGADOR PORTÁTIL', 'ANTENA',
    'CAIXA DE SOM', 'CABO USB', 'CABO MICRO USB'];
  const l = L.criarLeitor({ pessoas: ['ANA CLARA'], categorias: cats, habitos: {}, ativos: [] });
  const r = l.ler('Ana Clara fone com fone de ouvido finalizado Uma carregador portatiu finalizada Uma antena finalizada uma caixa de som Atualizada Um cabo micro USB finalizado Um controle nintendo finalizado', '');
  assert.deepStrictEqual(r.linhas.map((x) => [x.categoria, x.testados, x.finalizados, x.defeitos]), [
    ['FONE DE OUVIDO', 0, 1, 0], ['CARREGADOR PORTÁTIL', 0, 1, 0], ['ANTENA', 0, 1, 0], ['CAIXA DE SOM', 0, 1, 0],
    ['CABO MICRO USB', 0, 1, 0], ['CONTROLE', 0, 1, 0]]);
  assert.strictEqual(l.ler('Ana Clara 2 nintendo switch finalizados', '').linhas[0].categoria, 'NINTENDO SWITCH OLED');
  assert.strictEqual(l.ler('Ana Clara 3 controles remotos testados', '').linhas[0].categoria, 'CONTROLE REMOTO');
});

test('duas categorias coladas viram a categoria que junta as duas, se existir', () => {
  const l = L.criarLeitor({ pessoas: ['ANA CLARA'], categorias: ['CONTROLE', 'NINTENDO SWITCH OLED', 'CONTROLE NINTENDO SWITCH'], habitos: {}, ativos: [] });
  const r = l.ler('Ana Clara Um controle de nitendo finalizado', '').linhas;
  assert.deepStrictEqual(r.map((x) => [x.categoria, x.finalizados]), [['CONTROLE NINTENDO SWITCH', 1]]);
});
