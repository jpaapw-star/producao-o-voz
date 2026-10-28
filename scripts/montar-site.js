// Monta o site próprio (GitHub Pages) a partir da mesma tela do Apps Script: node scripts/montar-site.js <pasta>
// No site o microfone funciona (a página do Apps Script fica dentro de um quadro do Google que bloqueia o microfone).
// A página conversa com o mesmo app da Web por POST (doPost em src/Codigo.gs).
const fs = require('fs');
const path = require('path');

const API = 'https://script.google.com/macros/s/AKfycbyZ1ke6fEYJ7iWRkWS-Ly7SSzr3cqe0ogvOwlxqqRP3JM_BntXhHoVeTSQ0jLo9au2lsg/exec';
const raiz = path.join(__dirname, '..', 'src');
const saida = process.argv[2] || 'site';

// Imita google.script.run: cada chamada vira um fetch para o app da Web.
const ponte = `<script>
window.SITE_API = ${JSON.stringify(API)};
window.google = { script: { run: (function feito(ok, falha) {
  return new Proxy({}, { get: function (_, nome) {
    if (nome === 'withSuccessHandler') return function (f) { return feito(f, falha); };
    if (nome === 'withFailureHandler') return function (f) { return feito(ok, f); };
    return function () {
      var args = Array.prototype.slice.call(arguments);
      fetch(window.SITE_API, { method: 'POST', redirect: 'follow', cache: 'no-store',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn: nome, args: args }) })
        .then(function (r) { if (!r.ok) throw new Error('Sem conexão (http ' + r.status + ')'); return r.json(); })
        .then(function (d) { if (d && d.erro) throw new Error(d.erro); if (ok) ok(d ? d.r : null); })
        .catch(function (e) { if (falha) falha(e); });
    };
  } });
})() } };
</script>`;

const leitor = fs.readFileSync(path.join(raiz, 'Leitor.html'), 'utf8');
let html = fs.readFileSync(path.join(raiz, 'Index.html'), 'utf8');
if (!html.includes("<?!= incluir_('Leitor') ?>")) throw new Error('Index.html mudou: não achei o incluir_ do Leitor.');
html = html
  .replace("<?!= incluir_('Leitor') ?>", () => ponte + leitor)
  .replace('<meta charset="utf-8">', '<meta charset="utf-8">\n<title>Produção</title>\n' +
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">');
if (/<\?/.test(html)) throw new Error('Sobrou marcação do Apps Script no site.');
fs.mkdirSync(saida, { recursive: true });
fs.writeFileSync(path.join(saida, 'index.html'), html);
fs.writeFileSync(path.join(saida, '.nojekyll'), '');
console.log('Site montado em ' + path.join(saida, 'index.html'));
