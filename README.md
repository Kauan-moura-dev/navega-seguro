# Navega Seguro

Portal web educativo e acessível de segurança digital para pessoas idosas.

Projeto de Extensão em Desenvolvimento de Software – Curso de Ciência da Computação – Universidade Nove de Julho (UNINOVE), 2026.

- **Autor:** Kauan Castro Moura – RA 2223202665
- **Orientadora:** Prof.ª Debora Virgilia Canne

## Funcionalidades

- **Dicas de segurança** organizadas por categoria (senhas, golpes por mensagem, telefone e banco, compras e links), com filtro.
- **Verificador de força de senha** (fraca, média ou forte), com sugestões. Nada é enviado pela internet.
- **Desafio "É golpe?"**: 8 situações do dia a dia, com explicação após cada resposta e placar final.
- **Acessibilidade:** texto ampliado (20 px), botões A+ / A−, modo de alto contraste, navegação por teclado com foco visível e layout responsivo.

## Tecnologias

HTML5, CSS3 e JavaScript (ES6), sem frameworks e sem servidor. Hospedagem: GitHub Pages.

## Como executar

Abra o arquivo `index.html` em qualquer navegador moderno ou acesse a versão publicada no GitHub Pages.

## Estrutura

```
index.html              página principal
css/style.css           estilos (inclui alto contraste)
js/dados.js             categorias, dicas e perguntas do quiz
js/acessibilidade.js    tamanho de fonte e contraste
js/dicas.js             listagem e filtro de dicas
js/senha.js             verificador de força de senha
js/quiz.js              desafio "É golpe?"
docs/                   documentação ABNT (.docx), planilha de tarefas, diagramas e telas
```

## Referências de conteúdo

As orientações seguem a Cartilha de Segurança para Internet do CERT.br (fascículos *Autenticação* e *Golpes: não se deixe enganar*) e as Diretrizes de Acessibilidade para Conteúdo Web (WCAG) 2.2 do W3C.
