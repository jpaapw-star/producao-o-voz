// Carrega o script de src/Leitor.html em Node (mesmo código que roda no navegador).
const fs = require('fs');
const path = require('path');

function carregarLeitor() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'src', 'Leitor.html'), 'utf8');
  const codigo = html.replace(/^\s*<script>/, '').replace(/<\/script>\s*$/, '');
  const caixa = {};
  new Function('module', codigo).call(caixa, undefined);
  return caixa.LeitorProducao;
}

module.exports = { carregarLeitor };
