// Gera capturas da tela (claro e escuro) com dados de exemplo: node tests/ui/previa.js <pasta>
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const raiz = path.join(__dirname, '..', '..', 'src');
const leitor = fs.readFileSync(path.join(raiz, 'Leitor.html'), 'utf8');
const dados = {
  versao: 'prévia', hora: 9, ativos: ['GABRIEL VIEIRA'],
  pessoas: ['ANA LEAL', 'JULIA GOMES', 'JOÃO PEDRO', 'ANA CLARA', 'CAUÃ ENRICO', 'DANIEL VARGAS', 'ARTHUR SILVÉRIO', 'ARTUR ANTUNES'],
  categorias: ['NOTEBOOK', 'ROTEADOR', 'CAIXA DE SOM', 'FONE', 'MONITOR', 'FECHADURA DIGITAL', 'TECLADO'],
  habitos: { 'DANIEL VARGAS': { categorias: ['CAIXA DE SOM'], tipo: 'finalizados' } }
};
const falso = `<script>
  localStorage.setItem('producao.quem','JOÃO PEDRO');
  window.google = { script: { run: (function mk(ok){ return new Proxy({}, { get: (_, k) =>
    k === 'withSuccessHandler' ? (f) => mk(f) : k === 'withFailureHandler' ? () => mk(ok) :
    () => setTimeout(() => ok(${JSON.stringify(dados)}), 10) }); })() } };
</script>`;
const html = fs.readFileSync(path.join(raiz, 'Index.html'), 'utf8')
  .replace("<?!= incluir_('Leitor') ?>", falso + leitor);

(async () => {
  const saida = process.argv[2] || '.';
  const nav = await chromium.launch();
  for (const tema of ['light', 'dark']) {
    const pg = await nav.newPage({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 2, colorScheme: tema });
    await pg.route('http://previa.local/', (r) => r.fulfill({ contentType: 'text/html', body: html }));
    await pg.goto('http://previa.local/');
    await pg.waitForSelector('#tLancar.ativa');
    await pg.fill('#fala', 'Artur Antunes 3 webcams finalizadas 4 notebooks testados e finalizados 5 fones de feito e 3 teclados validados. Ô Julia, você viu a chave do armário? Daniel 4');
    await pg.dispatchEvent('#fala', 'input');
    await pg.waitForTimeout(400);
    await pg.screenshot({ path: path.join(saida, 'tela-' + tema + '.png'), fullPage: true });
  }
  await nav.close();
})();
