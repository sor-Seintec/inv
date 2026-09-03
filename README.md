# Auditoria Escolar — Vercel

Versão preparada para publicação no Vercel.

## Como funciona

Toda a leitura das planilhas Excel é feita diretamente no navegador do usuário.
A matriz e os inventários não são enviados para o servidor nem para uma API.
Isso evita o limite de upload das Vercel Functions e mantém os arquivos no computador local.

## Publicar no Vercel pelo GitHub

1. Crie um repositório no GitHub.
2. Envie todos os arquivos desta pasta para a raiz do repositório.
3. No Vercel, escolha **Add New > Project**.
4. Importe o repositório.
5. O Vercel deve identificar o projeto como **Vite**.
6. Build Command: `npm run build`.
7. Output Directory: `dist`.
8. Clique em **Deploy**.

Depois do deploy, abra a URL `https://SEU-PROJETO.vercel.app`.

## Uso

1. Clique em **Baixar base de inventário** e baixe a pasta **Sorocaba** no SharePoint.
2. Selecione a planilha matriz.
3. Selecione a pasta Sorocaba.
4. Clique em **Processar auditoria**.
5. Clique em qualquer escola para abrir o detalhamento e a comparação com a matriz.

## Desenvolvimento local

```bash
npm install
npm run dev
```

## Observação

Como esta versão não usa banco de dados, atualizar ou fechar a página limpa o resultado da auditoria. Isso foi intencional para manter o sistema simples, portátil e sem armazenamento dos inventários na nuvem.
