const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

test('site próprio: monta a página sem marcação do Apps Script e com a ponte para o servidor', () => {
  const pasta = fs.mkdtempSync(path.join(os.tmpdir(), 'site-'));
  execFileSync(process.execPath, [path.join(__dirname, '..', 'scripts', 'montar-site.js'), pasta]);
  const html = fs.readFileSync(path.join(pasta, 'index.html'), 'utf8');
  assert.ok(!/<\?/.test(html));
  assert.match(html, /window\.SITE_API = "https:\/\/script\.google\.com\/macros\/s\/AKfyc/);
  assert.match(html, /LeitorProducao/);
  assert.match(html, /<title>Produção<\/title>/);
});
