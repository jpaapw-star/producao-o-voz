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
- **Chamada**: o botão Chamada fala, um de cada vez, o nome de quem está no horário e ainda não lançou a hora. Ligue "Chamar sozinho" num aparelho que fica aberto (PC da sala). O horário de cada um vem da aba **Horários** (criada sozinha: PESSOA, ENTRADA, SAÍDA; pode ter duas linhas para manhã e tarde; a coluna D mostra "ok" ou o erro). Quem não está na aba é chamado pelo costume das últimas 2 semanas. Supervisores não são chamados. Os horários da chamada ficam na aba **Agenda da chamada**, escolhidos na tela Chamada (sem nenhum, chama aos 50 minutos de cada hora).
- O segredo antigo `CODIGO_EQUIPE` não é mais usado e pode ser apagado.

## Voz natural do Google (opcional)

Sem isso, cada aparelho usa a voz que já tem. Com isso, a chamada usa as vozes Chirp 3 HD do Google. É grátis até 1 milhão de letras por mês, e a chamada gasta bem menos que isso.

1. Em [console.cloud.google.com](https://console.cloud.google.com), escolha (ou crie) um projeto e ative o **Faturamento** com um cartão.
2. Em **APIs e serviços › Biblioteca**, procure **Cloud Text-to-Speech API** e clique em **Ativar**.
3. Em **APIs e serviços › Credenciais › Criar credenciais › Chave de API**. Copie a chave. Depois, em **Editar chave › Restrições de API**, deixe só a Cloud Text-to-Speech API.
4. No GitHub, em **Settings › Secrets and variables › Actions › New repository secret**, crie o segredo `TTS_CHAVE` e cole a chave lá.
5. Rode **Actions › Apps Script › Run workflow** e abra o link do app uma vez logado na sua conta Google. Aceite a nova permissão, "conectar a serviço externo", porque é por ela que o app chama a voz do Google.

**Sem cartão:** em vez dos passos 1 a 3, crie uma chave em [aistudio.google.com](https://aistudio.google.com) › Get API key › **Create API key in new project** (projeto só para a voz, para não dividir o limite com outro app). O app percebe sozinho e usa a voz do Gemini, que tem as mesmas vozes. Ela é grátis, mas tem um limite de pedidos por dia; frases repetidas ficam guardadas 6 horas e não contam de novo. Quando o limite acaba, os aparelhos falam com a voz deles.

Na tela Chamada, em **Voz**, aparecem as vozes "natural do Google". Se a internet ou a voz falhar, o aparelho fala com a voz dele.
