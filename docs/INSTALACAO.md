# Instalação (uma vez só)

Depois disso, todo código novo que entrar na branch `main` é enviado e publicado sozinho.

## 1. Dois segredos no GitHub

No repositório: **Settings › Secrets and variables › Actions › New repository secret**.

| Nome | O que colar |
| --- | --- |
| `CLASPRC_JSON` | O mesmo conteúdo usado no Preço Médio e na Memória Global: o arquivo `.clasprc.json` da sua pasta de usuário no Windows (`C:\Users\<você>\.clasprc.json`), aberto no Bloco de Notas. |
| `PLANILHA_ID` | O link da planilha "Controle de Produção Informatica" (pode colar o link inteiro). |

Nunca cole esses valores no chat.

## 2. Primeira publicação

Em **Actions › Apps Script › Run workflow**. Na primeira vez, o projeto do Apps Script é criado na sua conta Google e o link do app aparece no resumo da execução.

## 3. Autorizar uma vez

Abra o link do app no navegador **logado na sua conta Google** e aceite o pedido de acesso à planilha. Se aparecer "O Google não verificou este app", toque em **Avançado › Acessar**. Isso é normal em app próprio.

Se a equipe receber "Você precisa de permissão", abra o projeto em script.google.com › Implantar › Gerenciar implantações › editar e escolha **Quem pode acessar: Qualquer pessoa**.

## 4. Passar para a equipe

Mande o link. Não tem senha: quem tem o link pode lançar. No iPhone: abra no Safari › Compartilhar › **Adicionar à Tela de Início**. No Android: Chrome › ⋮ › **Adicionar à tela inicial**.

## Bom saber

- O app grava como você, então a equipe não precisa ter acesso à planilha. Você precisa poder editar a planilha.
- Ele só **acrescenta linhas** na aba Registros, no mesmo formato da linha de cima. Ele não mexe nas outras abas nem no app que já existe.
- Quem lançou cada coisa fica na aba **Lançamentos por voz** (criada sozinha).
- Apelidos e siglas ficam na aba **Apelidos** (criada sozinha): coluna A o apelido ("Gabi", "PS5"), coluna B o nome ou a categoria igual à aba Configurações. Vale na próxima vez que o app abrir (ou em até 10 minutos). A coluna C mostra "ok" ou o que está errado na linha.
- **Chamada**: nos horários da chamada (07:30 … 12:30, 14:30 … 16:30 e 16:45, ou os da aba **Agenda da chamada**), todos os computadores com o app aberto falam juntos, pelo relógio do Google. O horário de cada um vem da aba **Horários** (PESSOA, ENTRADA, SAÍDA; quem já saiu não é chamado). Quem não está na aba é chamado pelo costume das últimas 2 semanas. Supervisores não são chamados. Falta nunca é marcada sozinha.
- **Conversa na chamada**: marque "Ouvir as respostas pelo microfone" nos computadores com microfone (Edge ou Chrome). Aí a chamada chama uma pessoa por vez; ela vai a qualquer computador, diz "sou eu" e fala o que fez; o computador repete, pergunta se está certo e grava. Sem nenhum computador ouvindo, a chamada fala a lista de todos que faltam e repete a cada 2 minutos.

## IA da conversa (opcional, grátis)

Com a IA do Google, a conversa entende erros de fala, "é… hum…", correções e ajuda quem esqueceu.

1. Entre em https://aistudio.google.com/apikey com a sua conta Google e clique em **Create API key** (plano gratuito; não precisa de cartão).
2. No GitHub do app: **Settings › Secrets and variables › Actions › New repository secret**, nome `IA_CHAVE`, valor = a chave. Nunca mande a chave por chat.
   - Reserva opcional: chave do Ollama Cloud (ollama.com › Settings › Keys) no segredo `OLLAMA_API_KEY`; é usada quando o Gemini está ocupado.
3. Depois da publicação: abra o projeto no Apps Script (script.google.com, "Produção por voz"), escolha a função **autorizarIA** no topo e clique em **Executar**; aceite a permissão nova ("Conectar a um serviço externo"). Sem isso o app fica pedindo autorização.

- O segredo antigo `CODIGO_EQUIPE` não é mais usado e pode ser apagado.
