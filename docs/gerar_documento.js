// Gera o documento ABNT do Projeto de Extensão "Navega Seguro".
// Uso: node gerar_documento.js <saida.docx> [paginas.json]
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Header, AlignmentType, HeadingLevel,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageNumber, TabStopType,
  LevelFormat, PageBreak, VerticalAlign
} = require("docx");

const OUT = process.argv[2] || "documento.docx";
const PAG = process.argv[3] && fs.existsSync(process.argv[3]) ? JSON.parse(fs.readFileSync(process.argv[3])) : {};
const IMG = path.join(__dirname, "img");

// ---------- constantes de layout (ABNT / manual UNINOVE) ----------
const CM = 567;                                  // 1 cm em twips
const MARGENS = { top: 3 * CM, left: 3 * CM, bottom: 2 * CM, right: 2 * CM, header: 2 * CM };
const LARGURA_UTIL = 11906 - 5 * CM;             // A4 (11906) - 3 cm - 2 cm
const L15 = 360;                                 // entrelinhas 1,5
const L10 = 240;                                 // entrelinhas simples
const FONTE = "Arial";

// ---------- dados do projeto (fornecidos pelo aluno) ----------
const AUTOR = "KAUAN CASTRO MOURA";
const RA = "2223202665";
const TITULO = "NAVEGA SEGURO:";
const SUBTITULO = "portal de segurança digital para pessoas com pouca familiaridade com a internet";
const ORIENTADORA = "Prof.ª Debora Virgilia Canne";

// ---------- helpers ----------
const t = (text, o = {}) => new TextRun({ text, font: FONTE, size: 24, ...o });
const pend = (text) => new TextRun({ text, font: FONTE, size: 24, highlight: "yellow", bold: true });

function p(children, o = {}) {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({
    children,
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { line: o.line || L15, before: o.before || 0, after: o.after || 0 },
    indent: o.noIndent ? undefined : (o.indent || { firstLine: Math.round(1.25 * CM) }),
    keepNext: o.keepNext,
    pageBreakBefore: o.pageBreakBefore,
  });
}
const centro = (children, o = {}) => p(children, { align: AlignmentType.CENTER, noIndent: true, ...o });
const vazio = (n = 1, line = L15) => Array.from({ length: n }, () => new Paragraph({ children: [t("")], spacing: { line } }));

function h1(text, o = {}) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, pageBreakBefore: o.pageBreak !== false,
    alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    spacing: { line: L15, before: 0, after: L15 }, keepNext: true,
    children: [t(text, { bold: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, alignment: AlignmentType.LEFT, keepNext: true,
    spacing: { line: L15, before: L15, after: L15 }, children: [t(text, { bold: true })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3, alignment: AlignmentType.LEFT, keepNext: true,
    spacing: { line: L15, before: L15, after: L15 }, children: [t(text, { bold: true })],
  });
}
function item(children, ref = "marcadores") {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({ children, numbering: { reference: ref, level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L15 } });
}
// Legenda de ilustração: identificação na parte superior; fonte na parte inferior (tamanho menor)
const legenda = (txt) => new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { line: L10, before: 120, after: 60 }, children: [t(txt, { size: 20 })] });
const fonteIlustr = (txt = "Fonte: Elaborado pelo autor (2026).") => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: L10, before: 60, after: 240 }, children: [t(txt, { size: 20 })] });

function figura(rotulo, arquivo, larguraCm) {
  const buf = fs.readFileSync(path.join(IMG, arquivo));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20); // dimensões do PNG
  const largPx = Math.round(larguraCm * 37.8);
  const altPx = Math.round(largPx * h / w);
  return [
    legenda(rotulo),
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { line: L10 },
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: largPx, height: altPx },
        altText: { title: rotulo, description: rotulo, name: arquivo } })] }),
    fonteIlustr(),
  ];
}

const borda = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const bordas = { top: borda, bottom: borda, left: borda, right: borda };
function quadro(colunas, linhas, larguras) {
  const total = larguras.reduce((a, b) => a + b, 0);
  const cel = (txt, cab, w) => new TableCell({
    borders: bordas, width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: cab ? { fill: "E7E6E6", type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ spacing: { line: L10 }, alignment: cab ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: Array.isArray(txt) ? txt : [t(txt, { size: 20, bold: cab })] })],
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: larguras, alignment: AlignmentType.CENTER,
    rows: [new TableRow({ tableHeader: true, children: colunas.map((c, i) => cel(c, true, larguras[i])) }),
      ...linhas.map(l => new TableRow({ cantSplit: true, children: l.map((c, i) => cel(c, false, larguras[i])) }))],
  });
}

// Linha do sumário com pontilhado e número da página à direita
function linhaSumario(texto, chave, nivel) {
  const pag = PAG[chave] !== undefined ? String(PAG[chave]) : "?";
  return new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: LARGURA_UTIL, leader: "dot" }],
    spacing: { line: L15 },
    indent: nivel === 2 ? { left: 0 } : undefined,
    children: [t(texto, { bold: nivel === 1 }), t("\t" + pag, { bold: nivel === 1 })],
  });
}

// =====================================================================
// ELEMENTOS PRÉ-TEXTUAIS
// =====================================================================
const capa = [
  centro([t("UNIVERSIDADE NOVE DE JULHO – UNINOVE", { bold: true })]),
  centro([t("DIRETORIA DE INFORMÁTICA", { bold: true })]),
  centro([t("CURSO CIÊNCIA DA COMPUTAÇÃO", { bold: true })]),
  ...vazio(3),
  centro([t(`${AUTOR} – RA: ${RA}`)]),
  ...vazio(7),
  centro([t(TITULO, { bold: true })]),
  centro([t(SUBTITULO)]),
  ...vazio(15),
  centro([t("SÃO PAULO")]),
  centro([t("2026")]),
];

const natureza = (txt) => new Paragraph({ indent: { left: 8 * CM }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L10, after: 240 }, children: [t(txt)] });

const folhaRosto = [
  centro([t(`${AUTOR} – RA: ${RA}`)], { pageBreakBefore: true }),
  ...vazio(9),
  centro([t(TITULO, { bold: true })]),
  centro([t(SUBTITULO)]),
  ...vazio(2),
  natureza("Projeto de Extensão apresentado ao Curso de Ciência da Computação da Universidade Nove de Julho (UNINOVE), como requisito parcial para aprovação na disciplina de Projeto de Extensão em Desenvolvimento de Software."),
  natureza(`Professora Orientadora: ${ORIENTADORA}.`),
  ...vazio(10),
  centro([t("SÃO PAULO")]),
  centro([t("2026")]),
];

const assinatura = (nome, inst) => [
  ...vazio(2),
  centro([t("_______________________________________________")], { line: L10 }),
  centro([t(nome)], { line: L10 }),
  centro([t(inst)], { line: L10 }),
];

const folhaAprovacao = [
  centro([t(`${AUTOR} – RA: ${RA}`)], { pageBreakBefore: true }),
  ...vazio(3),
  centro([t(TITULO, { bold: true })]),
  centro([t(SUBTITULO)]),
  ...vazio(2),
  natureza("Projeto de Extensão apresentado ao Curso de Ciência da Computação da Universidade Nove de Julho (UNINOVE), como requisito parcial para aprovação na disciplina de Projeto de Extensão em Desenvolvimento de Software."),
  ...vazio(1),
  p([t("Aprovado em: _____ / _____ / _________.")], { noIndent: true }),
  ...vazio(1),
  centro([t("BANCA EXAMINADORA", { bold: true })]),
  ...assinatura(`${ORIENTADORA} – Orientadora`, "Universidade Nove de Julho – UNINOVE"),
  ...assinatura("Prof.(ª) _____________________________", "Instituição: __________________________"),
  ...assinatura("Prof.(ª) _____________________________", "Instituição: __________________________"),
];

const textoResumo = "O presente trabalho apresenta o desenvolvimento do Projeto de Extensão intitulado Navega Seguro, voltado a adultos e idosos com pouca familiaridade com a internet e, por isso, mais expostos a golpes digitais. O objetivo principal consiste em criar uma solução de software web, educativa e acessível, capaz de auxiliar esse público a reconhecer tentativas de fraude, criar senhas mais seguras e adotar cuidados básicos de navegação. A pesquisa parte da constatação de que golpistas exploram a engenharia social para enganar suas vítimas e de que parte da população, especialmente as pessoas com 60 anos ou mais, ainda tem acesso limitado à internet. A metodologia adotada apoia-se no método Kanban para organização das tarefas, no levantamento de requisitos funcionais e não funcionais, na modelagem por meio de diagramas de casos de uso, de classes e de entidade-relacionamento e na implementação com HTML5, CSS3 e JavaScript, sem servidor próprio. O sistema oferece dicas de segurança organizadas por categoria, um verificador de força de senha executado apenas no navegador, um questionário interativo do tipo “É golpe?” com explicações e opção de sinalizar erros, uma seção de notícias recentes sobre golpes e recursos de acessibilidade, como ampliação do texto e modo de alto contraste. Os testes funcionais automatizados confirmaram o funcionamento de todos os requisitos verificados. A validação foi realizada com duas familiares do autor, de 45 e 40 anos, que obtiveram oito e seis acertos em oito situações do questionário, avaliaram o site como relevante e informativo e sugeriram uma opção para sinalizar respostas incorretas, implementada na versão final. Conclui-se que soluções simples, construídas com tecnologias web abertas, podem contribuir para o uso mais seguro da internet, aproximando a formação universitária das necessidades da comunidade.";

const resumo = [
  h1("RESUMO", { center: true }),
  p(textoResumo, { noIndent: true }),
  ...vazio(1),
  p([t("Palavras-chave: ", { bold: true }), t("Desenvolvimento de Software. Projeto de Extensão. Segurança Digital. Inclusão Digital. Acessibilidade.")], { noIndent: true, align: AlignmentType.LEFT }),
];

const textoAbstract = "This work presents the development of the Extension Project entitled Navega Seguro, aimed at adults and older adults with little familiarity with the internet and, therefore, more exposed to digital scams. The main objective is to create an educational and accessible web software solution capable of helping this audience recognize fraud attempts, create safer passwords and adopt basic browsing precautions. The research starts from the observation that scammers exploit social engineering to deceive their victims and that part of the population, especially people aged 60 or over, still has limited access to the internet. The methodology is based on the Kanban method for task organization, on the elicitation of functional and non-functional requirements, on modeling through use case, class and entity-relationship diagrams and on the implementation with HTML5, CSS3 and JavaScript, without its own server. The system offers security tips organized by category, a password strength checker executed only in the browser, an interactive “Is it a scam?” quiz with explanations and an option to report errors, a section with recent news about scams and accessibility features, such as text enlargement and a high contrast mode. Automated functional tests confirmed the operation of all verified requirements. The validation was carried out with two of the author’s relatives, aged 45 and 40, who scored eight and six correct answers out of eight quiz situations, rated the website as relevant and informative and suggested an option to report incorrect answers, which was implemented in the final version. It is concluded that simple solutions, built with open web technologies, can contribute to a safer use of the internet, bringing university education closer to the needs of the community.";

const abstract = [
  h1("ABSTRACT", { center: true }),
  p(textoAbstract, { noIndent: true }),
  ...vazio(1),
  p([t("Keywords: ", { bold: true }), t("Software Development. Extension Project. Digital Security. Digital Inclusion. Accessibility.")], { noIndent: true, align: AlignmentType.LEFT }),
];

const SUMARIO = [
  ["1 INTRODUÇÃO", "s1", 1], ["1.1 CONTEXTUALIZAÇÃO E JUSTIFICATIVA DE EXTENSÃO", "s11", 2], ["1.2 PROBLEMATIZAÇÃO", "s12", 2], ["1.3 OBJETIVOS", "s13", 2],
  ["2 REVISÃO BIBLIOGRÁFICA E FUNDAMENTAÇÃO TEÓRICA", "s2", 1], ["2.1 EXTENSÃO UNIVERSITÁRIA", "s21", 2], ["2.2 INCLUSÃO DIGITAL", "s22", 2],
  ["2.3 SEGURANÇA DIGITAL DO USUÁRIO FINAL", "s23", 2], ["2.4 ENGENHARIA DE SOFTWARE E REQUISITOS", "s24", 2], ["2.5 TECNOLOGIAS WEB", "s25", 2], ["2.6 ACESSIBILIDADE E USABILIDADE", "s26", 2],
  ["3 COMUNIDADE ATENDIDA E IMPACTO SOCIAL", "s3", 1],
  ["4 ESPECIFICAÇÃO DO SISTEMA E REQUISITOS", "s4", 1], ["4.1 REQUISITOS FUNCIONAIS (RF)", "s41", 2], ["4.2 REQUISITOS NÃO FUNCIONAIS (RNF)", "s42", 2],
  ["4.3 ARQUITETURA DO SISTEMA E TECNOLOGIAS UTILIZADAS", "s43", 2], ["4.4 MODELAGEM E DIAGRAMAS DE SISTEMA", "s44", 2],
  ["5 METODOLOGIA DE DESENVOLVIMENTO E CRONOGRAMA", "s5", 1],
  ["6 RESULTADOS OBTIDOS E TESTES DE SOFTWARE", "s6", 1], ["6.1 TESTES FUNCIONAIS", "s61", 2], ["6.2 VALIDAÇÃO COM O PÚBLICO-ALVO", "s62", 2],
  ["7 CONSIDERAÇÕES FINAIS", "s7", 1],
  ["REFERÊNCIAS", "ref", 1],
];
const sumario = [h1("SUMÁRIO", { center: true }), ...SUMARIO.map(([a, b, c]) => linhaSumario(a, b, c))];

// =====================================================================
// ELEMENTOS TEXTUAIS
// =====================================================================
const cap1 = [
  h1("1 INTRODUÇÃO", { pageBreak: false }),
  p("A ampliação do acesso à internet trouxe benefícios importantes para a vida cotidiana, como a comunicação com familiares, o acesso a serviços bancários e a realização de compras sem sair de casa. Esse mesmo ambiente, porém, tornou-se espaço para golpes que exploram a confiança e a pouca familiaridade de parte dos usuários com as tecnologias digitais, situação que afeta tanto pessoas idosas quanto adultos que começaram a usar a internet recentemente ou que a utilizam apenas para tarefas básicas."),
  p("Diante desse cenário, este Projeto de Extensão propõe o Navega Seguro, um portal web educativo e acessível destinado a pessoas com pouca familiaridade com a internet. A aplicação reúne dicas de segurança digital, um verificador de força de senha, um questionário interativo que apresenta situações do dia a dia para que o usuário identifique se se trata de golpe ou de uma prática segura e uma seção com notícias recentes sobre golpes. O escopo do software é deliberadamente simples: funciona em qualquer navegador moderno, não exige cadastro, não envia os dados digitados pelo usuário e pode ser publicado gratuitamente."),

  h2("1.1 CONTEXTUALIZAÇÃO E JUSTIFICATIVA DE EXTENSÃO"),
  p("A extensão na educação superior é compreendida como atividade que articula ensino e pesquisa e promove a interação transformadora entre as instituições de ensino superior e os demais setores da sociedade (BRASIL, 2018). Nesse sentido, o projeto aproxima o conhecimento adquirido no curso de Ciência da Computação da Universidade Nove de Julho (UNINOVE) de uma necessidade concreta da comunidade: a proteção de pessoas com pouca experiência digital contra fraudes na internet."),
  p("Os indicadores nacionais reforçam a relevância da proposta. Segundo a pesquisa TIC Domicílios 2024, 89% da população brasileira com 10 anos ou mais era usuária de internet, enquanto na faixa de 60 anos ou mais esse percentual era de 63%, o menor entre todas as faixas etárias investigadas (NIC.BR, 2024). Esses dados mostram que uma parcela relevante da população ainda está chegando ao ambiente digital. A menor experiência com a internet, seja por idade, seja por uso recente ou pouco frequente, torna esses usuários alvos frequentes de golpistas. Ao oferecer orientação clara e prática, a tecnologia desenvolvida contribui para a autonomia e a segurança desse público."),

  h2("1.2 PROBLEMATIZAÇÃO"),
  p("Os golpes digitais costumam empregar técnicas de engenharia social, isto é, exploram sentimentos como medo, confiança, vontade de ajudar e senso de urgência para levar a vítima a agir sem refletir (CERT.BR, 2026). Mensagens de falsos parentes que pedem transferências via Pix, ligações de supostas centrais bancárias que solicitam senhas e pedidos para repassar códigos de verificação recebidos por SMS são exemplos recorrentes dessas práticas."),
  p("Somam-se a isso hábitos inseguros na criação de senhas, como o uso de sequências numéricas, datas de nascimento ou a mesma senha em vários serviços, comportamentos desaconselhados pelas orientações de segurança (CERT.BR, 2022). Muitas pessoas com pouca familiaridade digital não dispõem de um material acessível, com linguagem simples, letras grandes e exemplos práticos, que as ajude a reconhecer essas ameaças. O problema que orienta este projeto é, portanto: como apoiar pessoas com pouca familiaridade com a internet, por meio de uma solução de software simples e acessível, a identificar golpes digitais e a adotar práticas seguras?"),

  h2("1.3 OBJETIVOS"),
  h3("1.3.1 Objetivo Geral"),
  p("Desenvolver uma aplicação web funcional e acessível que auxilie adultos e idosos com pouca familiaridade com a internet a reconhecer golpes digitais, avaliar a força de senhas e adotar práticas seguras de navegação."),
  h3("1.3.2 Objetivos Específicos"),
  item("Mapear os requisitos funcionais e não funcionais da aplicação a partir das necessidades do público-alvo."),
  item("Prototipar e implementar a interface priorizando usabilidade e acessibilidade."),
  item("Implementar o catálogo de dicas, o verificador de força de senha, o questionário interativo e a seção de notícias utilizando HTML5, CSS3 e JavaScript."),
  item("Testar o funcionamento do sistema por meio de testes funcionais."),
  item("Validar a aplicação com familiares do autor pertencentes ao público-alvo, registrar suas percepções e incorporar melhorias."),
];

const cap2 = [
  h1("2 REVISÃO BIBLIOGRÁFICA E FUNDAMENTAÇÃO TEÓRICA"),
  p("Este capítulo apresenta os conceitos que embasam o desenvolvimento do projeto, abrangendo extensão universitária, inclusão digital, segurança da informação voltada ao usuário final, engenharia de software, tecnologias web e acessibilidade."),
  h2("2.1 EXTENSÃO UNIVERSITÁRIA"),
  p("A Resolução CNE/CES nº 7, de 18 de dezembro de 2018, estabelece as diretrizes para a extensão na educação superior brasileira e a caracteriza como um processo interdisciplinar, educativo, cultural, científico e político, que deve integrar-se à matriz curricular e gerar impacto na formação do estudante e na sociedade (BRASIL, 2018). O presente trabalho insere-se nessa perspectiva ao aplicar conhecimentos de desenvolvimento de software a uma demanda social identificada."),
  h2("2.2 INCLUSÃO DIGITAL"),
  p("A inclusão digital envolve não apenas o acesso à internet, mas também a capacidade de utilizá-la com autonomia e segurança. Os dados da pesquisa TIC Domicílios 2024 indicam diferenças importantes entre os grupos da população: enquanto praticamente todas as pessoas de 16 a 44 anos eram usuárias da rede, 37% das pessoas com 60 anos ou mais e 60% das pessoas analfabetas ou com educação infantil ainda não utilizavam a internet (NIC.BR, 2024). O Estatuto da Pessoa Idosa considera pessoa idosa aquela com idade igual ou superior a 60 anos (BRASIL, 2003). Esse cenário evidencia a importância de iniciativas que promovam o uso seguro e consciente das tecnologias digitais por quem ainda tem pouca experiência com elas, independentemente da idade."),
  h2("2.3 SEGURANÇA DIGITAL DO USUÁRIO FINAL"),
  p("De acordo com a Cartilha de Segurança para Internet, a engenharia social consiste em enganar uma pessoa explorando sentimentos como medo, obediência, confiança, carência, ganância, curiosidade e vontade de ajudar. Entre os sinais de alerta estão a ameaça, a pressa, o apelo afetivo, o pedido de segredo e ofertas boas demais para ser verdade; a recomendação central é confirmar a história com a pessoa ou a organização por meio de um contato oficial ou já conhecido (CERT.BR, 2026)."),
  p("No que se refere à autenticação, recomenda-se criar senhas longas, como a combinação de palavras aleatórias, misturar diferentes tipos de caracteres, evitar sequências de teclado e informações pessoais, não reutilizar a mesma senha em diferentes serviços, ativar a verificação em duas etapas e nunca compartilhar senhas ou códigos de verificação por mensagens, e-mails ou ligações (CERT.BR, 2022). Essas orientações fundamentaram o conteúdo das dicas, as regras do verificador de senha e as situações apresentadas no questionário."),
  h2("2.4 ENGENHARIA DE SOFTWARE E REQUISITOS"),
  p("A engenharia de software fornece processos, métodos e ferramentas para a construção de software de qualidade (PRESSMAN; MAXIM, 2016). Sommerville (2018) distingue os requisitos funcionais, que descrevem os serviços que o sistema deve oferecer e como deve reagir a determinadas entradas, dos requisitos não funcionais, que representam restrições e qualidades do sistema, como desempenho, usabilidade e segurança. Essa distinção foi adotada na especificação apresentada no capítulo 4."),
  p("Para a modelagem, utilizou-se a Linguagem de Modelagem Unificada (UML), por meio dos diagramas de casos de uso e de classes, além do Diagrama de Entidade-Relacionamento (DER) para representar a estrutura dos dados. A construção seguiu uma abordagem incremental, na qual o sistema é desenvolvido em partes que são testadas e incorporadas progressivamente (SOMMERVILLE, 2018)."),
  h2("2.5 TECNOLOGIAS WEB"),
  p("A aplicação foi construída com as tecnologias padrão da web: a Linguagem de Marcação de Hipertexto (HTML5), responsável pela estrutura e pela semântica do conteúdo; as Folhas de Estilo em Cascata (CSS3), responsáveis pela apresentação visual e pela adaptação a diferentes tamanhos de tela; e a linguagem JavaScript, responsável pela interatividade. Por serem executadas diretamente pelo navegador, essas tecnologias dispensam instalação de programas pelo usuário e permitem a hospedagem gratuita em serviços de páginas estáticas, como o GitHub Pages. A seção de notícias utiliza o formato RSS (Really Simple Syndication), padrão de distribuição de conteúdo usado por portais de notícias para disponibilizar suas publicações mais recentes."),
  h2("2.6 ACESSIBILIDADE E USABILIDADE"),
  p("As Diretrizes de Acessibilidade para Conteúdo Web (WCAG) 2.2, publicadas pelo World Wide Web Consortium (W3C), organizam recomendações para tornar o conteúdo web mais acessível a pessoas com deficiência, incluindo pessoas idosas com alterações de visão, audição, mobilidade ou cognição decorrentes do envelhecimento (W3C, 2025). As diretrizes estão estruturadas em quatro princípios: o conteúdo deve ser perceptível, operável, compreensível e robusto. No projeto, esses princípios orientaram decisões como o uso de texto ampliado, contraste adequado, foco visível na navegação por teclado, áreas de toque amplas e linguagem simples."),
];

const cap3 = [
  h1("3 COMUNIDADE ATENDIDA E IMPACTO SOCIAL"),
  p("O público-alvo do projeto é formado por adultos e idosos com pouca familiaridade com a internet, que utilizam celular ou computador para se comunicar, acessar serviços bancários ou realizar compras, mas conhecem pouco as práticas de segurança digital. Inicialmente, o projeto foi concebido apenas para pessoas idosas; como as participantes disponíveis para a validação tinham entre 40 e 45 anos, o público-alvo foi ampliado para manter a coerência entre a proposta e a comunidade efetivamente atendida."),
  p("Por se tratar de um projeto desenvolvido individualmente, a validação foi realizada com um grupo pequeno, composto por duas familiares do autor, o que viabilizou o contato direto com as usuárias e a coleta de percepções em curto prazo. O Quadro 1 apresenta o perfil das participantes."),
  legenda("Quadro 1 – Perfil das participantes da validação"),
  quadro(["Participante", "Parentesco", "Idade", "Familiaridade com tecnologia"], [
    ["Cintia Moura de Almeida Castro", "Mãe", "45 anos", "Iniciante"],
    ["Priscila de Almeida Rocha", "Tia", "40 anos", "Iniciante"],
  ], [3600, 1500, 1300, LARGURA_UTIL - 3600 - 1500 - 1300]),
  fonteIlustr("Fonte: Elaborado pelo autor (2026)."),
  p("Considerando o perfil do público, a interface foi projetada com letras grandes, possibilidade de ampliação do texto, modo de alto contraste, botões amplos e linguagem direta. Espera-se que a solução contribua para que os usuários reconheçam os sinais mais comuns de golpe, criem senhas mais fortes e adotem a prática de confirmar pedidos suspeitos por um canal oficial, ganhando mais autonomia e segurança no uso da internet. O material também pode ser compartilhado pelas próprias participantes com amigos e outros familiares, inclusive pessoas idosas, ampliando o alcance da ação."),
];

const RF = [
  ["RF-001", "O sistema deve exibir dicas de segurança digital organizadas por categoria."],
  ["RF-002", "O sistema deve permitir filtrar as dicas por categoria (Senhas, Golpes por mensagem, Telefone e banco, Compras e links)."],
  ["RF-003", "O sistema deve avaliar a força de uma senha digitada, em tempo real, classificando-a como fraca, média ou forte."],
  ["RF-004", "O sistema deve apresentar sugestões para tornar a senha mais forte."],
  ["RF-005", "O sistema deve permitir mostrar e ocultar a senha digitada."],
  ["RF-006", "O sistema deve apresentar um questionário com situações do cotidiano para que o usuário indique se é golpe ou se é seguro."],
  ["RF-007", "O sistema deve exibir, após cada resposta, se o usuário acertou e uma explicação sobre a situação."],
  ["RF-008", "O sistema deve exibir o total de acertos ao final do questionário e permitir refazê-lo."],
  ["RF-009", "O sistema deve permitir aumentar e diminuir o tamanho do texto."],
  ["RF-010", "O sistema deve permitir ativar e desativar o modo de alto contraste."],
  ["RF-011", "O sistema deve exibir as notícias mais recentes sobre golpes digitais, com título, veículo, data e link para a matéria original, e permitir atualizá-las."],
  ["RF-012", "O sistema deve exibir links para materiais oficiais do CERT.br quando as notícias não puderem ser carregadas."],
  ["RF-013", "O sistema deve permitir que o usuário sinalize um erro ou comente cada pergunta do questionário."],
];
const RNF = [
  ["RNF-001", "A interface deve ser responsiva, adaptando-se a telas de celulares sem rolagem horizontal."],
  ["RNF-002", "Os dados digitados pelo usuário não devem ser enviados pela rede; a análise de senhas deve ocorrer apenas no navegador."],
  ["RNF-003", "O texto base deve ter, no mínimo, 20 pixels, e os botões devem ter área de toque mínima de 48 × 48 pixels."],
  ["RNF-004", "Todas as funções devem ser acessíveis por teclado, com indicação visível do elemento em foco."],
  ["RNF-005", "A aplicação deve funcionar em navegadores modernos sem necessidade de instalação ou cadastro."],
  ["RNF-006", "A aplicação deve ser hospedada em serviço gratuito de páginas estáticas (GitHub Pages)."],
  ["RNF-007", "A interface deve utilizar linguagem simples e direta, adequada ao público-alvo."],
  ["RNF-008", "A falha no carregamento das notícias não deve impedir o funcionamento das demais funcionalidades."],
];

const cap4 = [
  h1("4 ESPECIFICAÇÃO DO SISTEMA E REQUISITOS"),
  p("Este capítulo descreve os requisitos, a arquitetura e a modelagem do sistema Navega Seguro."),
  h2("4.1 REQUISITOS FUNCIONAIS (RF)"),
  p("O Quadro 2 apresenta os requisitos funcionais do sistema. Os requisitos RF-011 a RF-013 foram incluídos na segunda versão do sistema, a partir da evolução do projeto e do feedback recebido na validação."),
  legenda("Quadro 2 – Requisitos funcionais"),
  quadro(["Código", "Descrição"], RF, [1500, LARGURA_UTIL - 1500]),
  fonteIlustr(),
  h2("4.2 REQUISITOS NÃO FUNCIONAIS (RNF)"),
  p("O Quadro 3 apresenta os requisitos não funcionais do sistema."),
  legenda("Quadro 3 – Requisitos não funcionais"),
  quadro(["Código", "Descrição"], RNF, [1500, LARGURA_UTIL - 1500]),
  fonteIlustr(),
  h2("4.3 ARQUITETURA DO SISTEMA E TECNOLOGIAS UTILIZADAS"),
  p("O Navega Seguro adota uma arquitetura cliente, composta por arquivos estáticos executados no navegador do usuário, sem servidor de aplicação próprio nem banco de dados remoto. As categorias, as dicas e as perguntas do questionário são mantidas em uma estrutura de dados JavaScript, organizada conforme o modelo apresentado no DER. A única informação armazenada é a preferência de tamanho de texto e de contraste, guardada localmente no próprio navegador (localStorage)."),
  p("A seção de notícias consulta o feed RSS do Google Notícias com uma busca por termos relacionados a golpes digitais. Como um site estático não pode ler diretamente um feed RSS de outro domínio, utiliza-se o serviço gratuito rss2json, que converte o feed para o formato JSON. Essa requisição não contém nenhum dado digitado pelo usuário, o que preserva o requisito RNF-002. Caso o serviço esteja indisponível, o sistema exibe links fixos para materiais oficiais do CERT.br (RF-012 e RNF-008). Já a função de sinalização abre, em uma nova aba, um formulário de registro de problemas (issue) no repositório do projeto no GitHub, previamente preenchido com a pergunta sinalizada. O Quadro 4 resume a pilha tecnológica."),
  legenda("Quadro 4 – Tecnologias utilizadas"),
  quadro(["Camada", "Tecnologia"], [
    ["Front-end (estrutura)", "HTML5 semântico com atributos ARIA"],
    ["Front-end (apresentação)", "CSS3 com variáveis, Grid e Flexbox (layout responsivo)"],
    ["Front-end (comportamento)", "JavaScript (ES6), organizado em classes"],
    ["Dados", "Estrutura de dados em JavaScript (arquivo dados.js) e localStorage para preferências"],
    ["Notícias", "Feed RSS do Google Notícias convertido pelo serviço rss2json"],
    ["Sinalização de erros", "GitHub Issues"],
    ["Hospedagem", "GitHub Pages (gratuito)"],
    ["Versionamento", "Git e GitHub"],
    ["Testes", "Playwright (testes funcionais automatizados no navegador Chromium)"],
  ], [3200, LARGURA_UTIL - 3200]),
  fonteIlustr(),
  p("Os arquivos do projeto estão organizados da seguinte forma: index.html (página principal), css/style.css (estilos), js/dados.js (dados), js/acessibilidade.js, js/dicas.js, js/senha.js, js/quiz.js e js/noticias.js (lógica de cada funcionalidade)."),
  p("O código-fonte está disponível no repositório https://github.com/Kauan-moura-dev/navega-seguro, e a aplicação publicada pode ser acessada em https://kauan-moura-dev.github.io/navega-seguro/."),
  h2("4.4 MODELAGEM E DIAGRAMAS DE SISTEMA"),
  p("A Figura 1 apresenta o diagrama de casos de uso. O ator Usuário interage com as funcionalidades de consulta de dicas, verificação de senha, questionário, notícias e acessibilidade."),
  ...figura("Figura 1 – Diagrama de casos de uso", "diagrama-casos-de-uso.png", 15),
  p("A Figura 2 apresenta o diagrama de classes. As classes Acessibilidade, CatalogoDicas, VerificadorSenha, Quiz e FeedNoticias concentram a lógica de cada funcionalidade, enquanto Categoria, Dica, Pergunta e Noticia representam os dados manipulados."),
  ...figura("Figura 2 – Diagrama de classes", "diagrama-classes.png", 16),
  p("A Figura 3 apresenta o Diagrama de Entidade-Relacionamento (DER). Uma categoria possui várias dicas e classifica várias perguntas, caracterizando dois relacionamentos do tipo um-para-muitos. As notícias não fazem parte do modelo de dados, pois são obtidas em tempo real de uma fonte externa e não são armazenadas. Embora os dados sejam mantidos em JavaScript, o modelo permite a migração futura para um banco de dados relacional sem alteração da estrutura."),
  ...figura("Figura 3 – Diagrama de Entidade-Relacionamento", "diagrama-der.png", 13),
  p("As Figuras 4 a 8 apresentam as telas do sistema, que também cumprem a função de protótipo navegável da interface."),
  ...figura("Figura 4 – Tela inicial com menu de navegação", "tela-inicio.png", 15),
  ...figura("Figura 5 – Verificador de força de senha", "tela-senha.png", 14),
  ...figura("Figura 6 – Questionário “É golpe?” com explicação e opção de sinalizar", "tela-quiz-explicacao.png", 14),
  ...figura("Figura 7 – Seção de notícias e alertas recentes", "tela-noticias.png", 15),
  ...figura("Figura 8 – Modo de alto contraste com texto ampliado", "tela-contraste.png", 15),
];

const cronograma = [
  ["Definição do tema, levantamento de requisitos e modelagem", "01/11/2026", "Concluída"],
  ["Desenvolvimento da primeira versão do site", "01/11/2026", "Concluída"],
  ["Testes funcionais automatizados e correções", "01/11/2026", "Concluída"],
  ["Elaboração da documentação (normas ABNT)", "01/11/2026", "Concluída"],
  ["Publicação no GitHub e no GitHub Pages", "01/11/2026", "Concluída"],
  ["Testes com familiares e coleta de feedback", "01/11/2026", "Concluída"],
  ["Segunda versão: novo design, notícias e sinalização", "01/11/2026", "Concluída"],
  ["Gravação do vídeo de apresentação", "Previsto: até 12/11/2026", "Pendente"],
  ["Revisão final e entrega no Google Sala de Aula", "Previsto: até 15/11/2026", "Pendente"],
];

const cap5 = [
  h1("5 METODOLOGIA DE DESENVOLVIMENTO E CRONOGRAMA"),
  p("Para a organização do trabalho foi adotado o método Kanban, que se baseia na visualização do fluxo de trabalho e na limitação do trabalho em andamento, favorecendo entregas contínuas e a identificação de gargalos (ANDERSON, 2010). As tarefas foram distribuídas nas etapas A fazer, Em desenvolvimento e Concluído, e o acompanhamento foi registrado na planilha de distribuição de tarefas solicitada pela disciplina."),
  p("O desenvolvimento seguiu uma abordagem incremental (SOMMERVILLE, 2018). Na primeira versão foram implementados o catálogo de dicas, o verificador de senha, o questionário e os recursos de acessibilidade. Após a validação com as participantes, foi produzida uma segunda versão, com novo design visual, seção de notícias e opção de sinalizar erros no questionário. Cada funcionalidade foi implementada e testada separadamente antes de ser integrada à página principal."),
  p("Por se tratar de um projeto individual, todas as funções foram exercidas pelo autor, Kauan Castro Moura: levantamento de requisitos, modelagem, desenvolvimento front-end, testes, documentação, publicação no GitHub, contato com as participantes e gravação do vídeo. O Quadro 5 apresenta o cronograma do projeto, de acordo com o prazo de entrega estabelecido pela disciplina."),
  legenda("Quadro 5 – Cronograma do projeto"),
  quadro(["Etapa", "Período", "Situação"], cronograma, [LARGURA_UTIL - 2900 - 1500, 2900, 1500]),
  fonteIlustr(),
];

const testes = [
  ["T01", "Exibir todas as dicas", "10 dicas exibidas", "Aprovado"],
  ["T02", "Filtrar dicas pela categoria Senhas", "3 dicas exibidas", "Aprovado"],
  ["T03", "Avaliar a senha “123456”", "Classificada como fraca (senha comum)", "Aprovado"],
  ["T04", "Avaliar a senha “abc”", "Classificada como fraca", "Aprovado"],
  ["T05", "Avaliar a senha “Casa2024”", "Classificada como média, com sugestão de incluir símbolos", "Aprovado"],
  ["T06", "Avaliar a senha “Cafe com bolo na varanda!”", "Classificada como forte", "Aprovado"],
  ["T07", "Mostrar e ocultar a senha", "Texto exibido e ocultado corretamente", "Aprovado"],
  ["T08", "Responder corretamente às 8 situações do questionário", "Explicação exibida a cada resposta; resultado “8 de 8”", "Aprovado"],
  ["T09", "Gerar o link “Sinalizar” de uma pergunta", "Link para nova issue no GitHub preenchido com a pergunta", "Aprovado"],
  ["T10", "Carregar as notícias", "6 notícias exibidas com veículo e data", "Aprovado"],
  ["T11", "Simular falha no serviço de notícias", "Aviso exibido e 3 links oficiais do CERT.br apresentados", "Aprovado"],
  ["T12", "Aumentar o texto duas vezes", "Fonte base alterada de 20 px para 24 px", "Aprovado"],
  ["T13", "Ativar o alto contraste", "Tema escuro com cores de alto contraste aplicado", "Aprovado"],
  ["T14", "Abrir em tela de celular (390 px de largura)", "Sem rolagem horizontal", "Aprovado"],
  ["T15", "Verificar erros de execução no console", "Nenhum erro registrado", "Aprovado"],
];

const cap6 = [
  h1("6 RESULTADOS OBTIDOS E TESTES DE SOFTWARE"),
  p("Como resultado, foi entregue o portal Navega Seguro, com todas as funcionalidades especificadas no capítulo 4. As telas do sistema finalizado estão apresentadas nas Figuras 4 a 8, e a Figura 9 mostra a adaptação da interface à tela de um celular."),
  ...figura("Figura 9 – Visualização em tela de celular", "tela-celular.png", 6),
  new Paragraph({ heading: HeadingLevel.HEADING_2, pageBreakBefore: true, keepNext: true, spacing: { line: L15, before: 0, after: L15 }, children: [t("6.1 TESTES FUNCIONAIS", { bold: true })] }),
  p("Os testes funcionais foram automatizados com a ferramenta Playwright, que executa a aplicação em um navegador real (Chromium) e simula as ações do usuário. O Quadro 6 apresenta os casos de teste executados na versão final e seus resultados."),
  legenda("Quadro 6 – Casos de teste funcionais"),
  quadro(["ID", "Caso de teste", "Resultado obtido", "Situação"], testes, [700, 3500, LARGURA_UTIL - 700 - 3500 - 1300, 1300]),
  fonteIlustr(),
  p("Durante os testes foram identificados dois problemas de visibilidade no modo de alto contraste: na primeira versão, o botão de ativação desse modo ficava com a mesma cor do cabeçalho; na segunda versão, a barra do menu assumia a cor amarela dos próprios links, tornando-os ilegíveis. Ambos foram corrigidos com estilos específicos, e os testes foram executados novamente com sucesso."),
  h2("6.2 VALIDAÇÃO COM O PÚBLICO-ALVO"),
  p("A validação foi realizada em 1º de novembro de 2026 com as duas participantes apresentadas no Quadro 1. Cada participante utilizou o portal, leu as dicas, respondeu às oito situações do questionário “É golpe?” e, em seguida, enviou sua opinião sobre o site por mensagem de áudio. Os áudios foram transcritos pelo autor. O Quadro 7 apresenta o desempenho no questionário."),
  legenda("Quadro 7 – Desempenho das participantes no questionário"),
  quadro(["Participante", "Acertos", "Aproveitamento"], [
    ["Cintia Moura de Almeida Castro (mãe, 45 anos)", "8 de 8", "100%"],
    ["Priscila de Almeida Rocha (tia, 40 anos)", "6 de 8", "75%"],
  ], [LARGURA_UTIL - 1800 - 2200, 1800, 2200]),
  fonteIlustr(),
  p("Ao avaliar o site, a participante Cintia Moura de Almeida Castro destacou a clareza das explicações e a importância de confirmar informações com o banco:"),
  new Paragraph({ indent: { left: 4 * CM }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L10, before: 120, after: 240 }, children: [t("Achei um site muito seguro e relevante, principalmente nas abordagens e explicações entre golpe e fraude. Primeiro consultar o banco antes de enviar qualquer tipo de documento ou site perigoso.", { size: 20 })] }),
  p("A participante Priscila de Almeida Rocha considerou o conteúdo útil e apresentou uma sugestão de melhoria:"),
  new Paragraph({ indent: { left: 4 * CM }, alignment: AlignmentType.JUSTIFIED, spacing: { line: L10, before: 120, after: 240 }, children: [
    t("Achei o site Navega Seguro muito relevante e informativo. As informações disponibilizadas foram de grande ajuda, contribuindo para ampliar os meus conhecimentos e esclarecer dúvidas. Com certeza recomendaria o site para outras pessoas, pois acredito que o conteúdo pode ser útil para todos. Como sugestão de melhoria, seria interessante implementar uma opção que permitisse ao usuário ", { size: 20 }),
    new TextRun({ text: "[CONFIRMAR: comentar]", font: FONTE, size: 20, highlight: "yellow", bold: true }),
    t(" ou sinalizar respostas incorretas, possibilitando ajustes nas informações e tornando as ferramentas ainda mais confiáveis e eficientes.", { size: 20 })] }),
  p("As duas avaliações foram positivas quanto à relevância e à clareza do conteúdo. Em resposta à sugestão recebida, foi implementado na segunda versão o recurso “Sinalizar erro ou comentar esta pergunta” (RF-013), disponível após cada resposta do questionário, além de um botão “Enviar sugestão” na seção Sobre do site. Na mesma versão, o autor também reformulou o design da página e incluiu a seção de notícias recentes sobre golpes (RF-011 e RF-012), com o objetivo de manter o conteúdo atualizado."),
];

const cap7 = [
  h1("7 CONSIDERAÇÕES FINAIS"),
  p("O objetivo geral deste projeto, desenvolver uma aplicação web funcional e acessível que auxilie adultos e idosos com pouca familiaridade com a internet a reconhecer golpes digitais e a adotar práticas seguras, foi atingido. Os requisitos foram mapeados, a interface foi implementada com recursos de acessibilidade, todos os casos de teste funcionais foram aprovados e a aplicação foi publicada gratuitamente na internet."),
  p("A validação com as participantes indicou que o conteúdo foi considerado relevante, informativo e útil, e o desempenho no questionário, com oito e seis acertos em oito situações, mostrou que as explicações ajudam a diferenciar situações de golpe de práticas seguras. A sugestão recebida foi incorporada ao sistema, evidenciando o caráter dialógico da extensão universitária. Como limitação, destaca-se que a validação envolveu apenas duas pessoas, de 40 e 45 anos, sem a participação de pessoas idosas; por isso, os resultados não podem ser generalizados para todo o público-alvo."),
  p("Do ponto de vista acadêmico, o projeto permitiu aplicar na prática conteúdos do curso de Ciência da Computação, como levantamento de requisitos, modelagem UML, desenvolvimento web, consumo de dados externos, testes de software e versionamento com Git. Do ponto de vista extensionista, evidenciou que a tecnologia precisa ser pensada a partir das necessidades reais do usuário: linguagem simples, letras grandes e exemplos do cotidiano mostraram-se tão importantes quanto o código."),
  p("Como trabalhos futuros para uma versão 3.0, sugerem-se: a validação com um grupo maior, incluindo pessoas com 60 anos ou mais; a inclusão de leitura em voz alta do conteúdo; um canal de sinalização que não exija conta no GitHub, como um formulário simples; a ampliação do banco de situações do questionário; a migração dos dados para um banco de dados relacional, seguindo o DER apresentado; e a realização de oficinas presenciais com grupos da comunidade."),
  ...vazio(1),
  p([t("LINK DO VÍDEO NO YOUTUBE: ", { bold: true }), pend("[PENDENTE – inserir o link do vídeo explicando o funcionamento do projeto]")], { noIndent: true }),
];

// =====================================================================
// REFERÊNCIAS (NBR 6023:2018) – ordem alfabética, alinhadas à esquerda,
// espaço simples e separadas por uma linha em branco; título em negrito.
// =====================================================================
const ref = (partes) => new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: L10, after: 240 },
  children: partes.map(([txt, b]) => t(txt, { bold: !!b })) });

const referencias = [
  h1("REFERÊNCIAS", { center: true }),
  ref([["ANDERSON, David J. "], ["Kanban", true], [": successful evolutionary change for your technology business. Sequim: Blue Hole Press, 2010."]]),
  ref([["ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. "], ["NBR 6023", true], [": informação e documentação: referências: elaboração. Rio de Janeiro: ABNT, 2018."]]),
  ref([["ASSOCIAÇÃO BRASILEIRA DE NORMAS TÉCNICAS. "], ["NBR 14724", true], [": informação e documentação: trabalhos acadêmicos: apresentação. Rio de Janeiro: ABNT, 2011."]]),
  ref([["BRASIL. "], ["Lei nº 10.741, de 1º de outubro de 2003", true], [". Dispõe sobre o Estatuto da Pessoa Idosa e dá outras providências. Brasília, DF: Presidência da República, 2003. Disponível em: https://www.planalto.gov.br/ccivil_03/leis/2003/l10.741.htm. Acesso em: 1 nov. 2026."]]),
  ref([["BRASIL. Ministério da Educação. Conselho Nacional de Educação. Câmara de Educação Superior. "], ["Resolução nº 7, de 18 de dezembro de 2018", true], [". Estabelece as Diretrizes para a Extensão na Educação Superior Brasileira [...]. Diário Oficial da União: seção 1, Brasília, DF, n. 243, p. 49-50, 19 dez. 2018."]]),
  ref([["CERT.BR. "], ["Autenticação", true], [". São Paulo: NIC.br, 2022. (Cartilha de Segurança para Internet: fascículo). Disponível em: https://cartilha.cert.br/fasciculos/autenticacao/fasciculo-autenticacao.pdf. Acesso em: 1 nov. 2026."]]),
  ref([["CERT.BR. "], ["Golpes", true], [": não se deixe enganar. São Paulo: NIC.br, 2026. (Cartilha de Segurança para Internet: fascículo). Disponível em: https://cartilha.cert.br/fasciculos/golpes/fasciculo-golpes.pdf. Acesso em: 1 nov. 2026."]]),
  ref([["NIC.BR. "], ["Pesquisa sobre o uso das tecnologias de informação e comunicação nos domicílios brasileiros", true], [": TIC Domicílios 2024: tabela C2A: usuários de internet: indicador ampliado. São Paulo: Cetic.br, 2024. Disponível em: https://www.cetic.br/pt/tics/domicilios/2024/individuos/C2A/. Acesso em: 1 nov. 2026."]]),
  ref([["PRESSMAN, Roger S.; MAXIM, Bruce R. "], ["Engenharia de software", true], [": uma abordagem profissional. 8. ed. Porto Alegre: AMGH, 2016."]]),
  ref([["SOMMERVILLE, Ian. "], ["Engenharia de software", true], [". 10. ed. São Paulo: Pearson Education do Brasil, 2018."]]),
  ref([["W3C. "], ["Diretrizes de Acessibilidade para Conteúdo Web (WCAG) 2.2", true], [". Tradução autorizada para o português do Brasil. [S. l.]: W3C, 2025. Disponível em: https://www.w3.org/Translations/WCAG22-pt-BR/. Acesso em: 1 nov. 2026."]]),
];

// =====================================================================
// MONTAGEM
// =====================================================================
const cabecalhoNumerado = new Header({
  children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONTE, size: 20 })] })],
});
const vazioHeader = new Header({ children: [new Paragraph({ children: [] })] });
const propsPagina = { page: { size: { width: 11906, height: 16838 }, margin: MARGENS } };

const doc = new Document({
  creator: "Kauan Castro Moura",
  title: "Navega Seguro: portal de segurança digital para pessoas com pouca familiaridade com a internet",
  styles: {
    default: { document: { run: { font: FONTE, size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONTE, size: 24, bold: true, color: "000000" }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONTE, size: 24, bold: true, color: "000000" }, paragraph: { outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONTE, size: 24, bold: true, color: "000000" }, paragraph: { outlineLevel: 2 } },
    ],
  },
  numbering: { config: [{ reference: "marcadores", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] }] },
  sections: [
    { properties: propsPagina, headers: { default: vazioHeader }, children: [...capa, ...folhaRosto, ...folhaAprovacao, ...resumo, ...abstract, ...sumario] },
    { properties: { ...propsPagina, page: { ...propsPagina.page, pageNumbers: { start: PAG.inicioNumeracao || 1 } } },
      headers: { default: cabecalhoNumerado },
      children: [...cap1, ...cap2, ...cap3, ...cap4, ...cap5, ...cap6, ...cap7, ...referencias] },
  ],
});

Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log("gerado:", OUT); });
