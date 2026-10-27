# Produção por voz

App da Web (Google Apps Script) para lançar a produção dos colaboradores falando, em vez de digitar.

Você toca no microfone do teclado do celular e fala, por exemplo:

> Ana Leal notebook 3 testados 2 finalizados. Julia roteador 4 testados. Cauã fechadura 5 finalizados 1 defeito.

O app mostra as linhas para conferir e grava na aba **Registros** da planilha de produção, com data, hora e faixa de horário.

- Nomes e categorias vêm da aba **Configurações**. O app entende apelidos ("Julia" = JULIA GOMES), plurais, números por extenso e pequenos erros de ditado.
- Sem nome, a produção vai para quem está lançando ("eu" também vale).
- Na dúvida (dois Arthur, número sem "testados/finalizados"), o app usa o histórico recente. Se ainda assim não der para saber, ele pergunta só aquilo.
- **Desfazer** qualquer lançamento do dia em "Lançamentos de hoje neste aparelho". Quem lançou fica na aba **Lançamentos por voz**.
- Não usa IA paga. O ditado é o do próprio iPhone ou Android.

Instalação: [docs/INSTALACAO.md](docs/INSTALACAO.md). Testes: `npm test`.

| Arquivo | Para que serve |
| --- | --- |
| `src/Leitor.html` | Entende a frase (roda no celular e nos testes) |
| `src/Codigo.gs` | Lê as listas, grava em Registros e desfaz |
| `src/Index.html` | Tela do app |
