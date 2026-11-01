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
const SUBTITULO = "portal de segurança digital para pessoas idosas";
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

const textoResumo = "O presente trabalho apresenta o desenvolvimento do Projeto de Extensão intitulado Navega Seguro, voltado ao atendimento de pessoas idosas que utilizam a internet e estão expostas a golpes digitais. O objetivo principal consiste em criar uma solução de software web, educativa e acessível, capaz de auxiliar esse público a reconhecer tentativas de fraude, criar senhas mais seguras e adotar cuidados básicos de navegação. A pesquisa parte da constatação de que a população com 60 anos ou mais apresenta a menor proporção de usuários de internet entre as faixas etárias brasileiras e de que golpistas exploram a engenharia social para enganar suas vítimas. A metodologia adotada apoia-se no método Kanban para organização das tarefas, no levantamento de requisitos funcionais e não funcionais, na modelagem por meio de diagramas de casos de uso, de classes e de entidade-relacionamento e na implementação com HTML5, CSS3 e JavaScript, sem dependência de servidor. O sistema oferece dicas de segurança organizadas por categoria, um verificador de força de senha executado apenas no navegador, um questionário interativo do tipo “É golpe?” com explicações após cada resposta e recursos de acessibilidade, como ampliação do texto e modo de alto contraste. Os testes funcionais automatizados confirmaram o funcionamento de todos os requisitos verificados, e a validação com o público-alvo prevê a participação de dois a três familiares do autor. Conclui-se que soluções simples, construídas com tecnologias web abertas, podem contribuir para a inclusão digital segura de pessoas idosas, aproximando a formação universitária das necessidades da comunidade.";

const resumo = [
  h1("RESUMO", { center: true }),
  p(textoResumo, { noIndent: true }),
  ...vazio(1),
  p([t("Palavras-chave: ", { bold: true }), t("Desenvolvimento de Software. Projeto de Extensão. Segurança Digital. Pessoa Idosa. Acessibilidade.")], { noIndent: true, align: AlignmentType.LEFT }),
];

const textoAbstract = "This work presents the development of the Extension Project entitled Navega Seguro, aimed at serving older adults who use the internet and are exposed to digital scams. The main objective is to create an educational and accessible web software solution capable of helping this audience recognize fraud attempts, create safer passwords and adopt basic browsing precautions. The research starts from the observation that the population aged 60 or over has the lowest proportion of internet users among Brazilian age groups and that scammers exploit social engineering to deceive their victims. The methodology is based on the Kanban method for task organization, on the elicitation of functional and non-functional requirements, on modeling through use case, class and entity-relationship diagrams and on the implementation with HTML5, CSS3 and JavaScript, without dependence on a server. The system offers security tips organized by category, a password strength checker executed only in the browser, an interactive “Is it a scam?” quiz with explanations after each answer and accessibility features, such as text enlargement and a high contrast mode. Automated functional tests confirmed the operation of all verified requirements, and the validation with the target audience includes the participation of two to three of the author’s relatives. It is concluded that simple solutions, built with open web technologies, can contribute to the safe digital inclusion of older adults, bringing university education closer to the needs of the community.";

const abstract = [
  h1("ABSTRACT", { center: true }),
  p(textoAbstract, { noIndent: true }),
  ...vazio(1),
  p([t("Keywords: ", { bold: true }), t("Software Development. Extension Project. Digital Security. Older Adults. Accessibility.")], { noIndent: true, align: AlignmentType.LEFT }),
];

const SUMARIO = [
  ["1 INTRODUÇÃO", "s1", 1], ["1.1 CONTEXTUALIZAÇÃO E JUSTIFICATIVA DE EXTENSÃO", "s11", 2], ["1.2 PROBLEMATIZAÇÃO", "s12", 2], ["1.3 OBJETIVOS", "s13", 2],
  ["2 REVISÃO BIBLIOGRÁFICA E FUNDAMENTAÇÃO TEÓRICA", "s2", 1], ["2.1 EXTENSÃO UNIVERSITÁRIA", "s21", 2], ["2.2 INCLUSÃO DIGITAL E A PESSOA IDOSA", "s22", 2],
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
  p("A ampliação do acesso à internet trouxe benefícios importantes para a vida cotidiana, como a comunicação com familiares, o acesso a serviços bancários e a realização de compras sem sair de casa. Esse mesmo ambiente, porém, tornou-se espaço para golpes que exploram a confiança e a pouca familiaridade de parte dos usuários com as tecnologias digitais, situação que afeta de forma particular as pessoas idosas."),
  p("Diante desse cenário, este Projeto de Extensão propõe o Navega Seguro, um portal web educativo e acessível destinado a pessoas idosas. A aplicação reúne dicas de segurança digital, um verificador de força de senha e um questionário interativo que apresenta situações do dia a dia para que o usuário identifique se se trata de golpe ou de uma prática segura. O escopo do software é deliberadamente simples: funciona em qualquer navegador moderno, não exige cadastro e não envia dados para servidores, o que facilita o uso pelo público-alvo e sua publicação gratuita."),

  h2("1.1 CONTEXTUALIZAÇÃO E JUSTIFICATIVA DE EXTENSÃO"),
  p("A extensão na educação superior é compreendida como atividade que articula ensino e pesquisa e promove a interação transformadora entre as instituições de ensino superior e os demais setores da sociedade (BRASIL, 2018). Nesse sentido, o projeto aproxima o conhecimento adquirido no curso de Ciência da Computação da Universidade Nove de Julho (UNINOVE) de uma necessidade concreta da comunidade: a proteção de pessoas idosas contra fraudes digitais."),
  p("A relevância social da proposta é reforçada pelos indicadores nacionais. Segundo a pesquisa TIC Domicílios 2024, 89% da população brasileira com 10 anos ou mais era usuária de internet, enquanto na faixa de 60 anos ou mais esse percentual era de 63%, o menor entre todas as faixas etárias investigadas (NIC.BR, 2024). A menor familiaridade com o ambiente digital, somada ao uso crescente de serviços online, torna esse público alvo frequente de golpistas. Ao oferecer orientação clara e prática, a tecnologia desenvolvida contribui para a autonomia e a segurança dos usuários idosos."),

  h2("1.2 PROBLEMATIZAÇÃO"),
  p("Os golpes digitais costumam empregar técnicas de engenharia social, isto é, exploram sentimentos como medo, confiança, vontade de ajudar e senso de urgência para levar a vítima a agir sem refletir (CERT.BR, 2026). Mensagens de falsos parentes que pedem transferências via Pix, ligações de supostas centrais bancárias que solicitam senhas e pedidos para repassar códigos de verificação recebidos por SMS são exemplos recorrentes dessas práticas."),
  p("Somam-se a isso hábitos inseguros na criação de senhas, como o uso de sequências numéricas, datas de nascimento ou a mesma senha em vários serviços, comportamentos desaconselhados pelas orientações de segurança (CERT.BR, 2022). Muitas pessoas idosas não dispõem de um material acessível, com linguagem simples, letras grandes e exemplos práticos, que as ajude a reconhecer essas ameaças. O problema que orienta este projeto é, portanto: como apoiar pessoas idosas, por meio de uma solução de software simples e acessível, a identificar golpes digitais e a adotar práticas seguras na internet?"),

  h2("1.3 OBJETIVOS"),
  h3("1.3.1 Objetivo Geral"),
  p("Desenvolver uma aplicação web funcional e acessível que auxilie pessoas idosas a reconhecer golpes digitais, avaliar a força de senhas e adotar práticas seguras de navegação na internet."),
  h3("1.3.2 Objetivos Específicos"),
  item("Mapear os requisitos funcionais e não funcionais da aplicação a partir das necessidades do público-alvo."),
  item("Prototipar e implementar a interface priorizando usabilidade e acessibilidade."),
  item("Implementar o verificador de força de senha, o catálogo de dicas e o questionário interativo utilizando HTML5, CSS3 e JavaScript."),
  item("Testar o funcionamento do sistema por meio de testes funcionais."),
  item("Validar a aplicação com dois a três familiares do autor pertencentes ao público-alvo e registrar suas percepções."),
];

const cap2 = [
  h1("2 REVISÃO BIBLIOGRÁFICA E FUNDAMENTAÇÃO TEÓRICA"),
  p("Este capítulo apresenta os conceitos que embasam o desenvolvimento do projeto, abrangendo extensão universitária, inclusão digital da pessoa idosa, segurança da informação voltada ao usuário final, engenharia de software, tecnologias web e acessibilidade."),
  h2("2.1 EXTENSÃO UNIVERSITÁRIA"),
  p("A Resolução CNE/CES nº 7, de 18 de dezembro de 2018, estabelece as diretrizes para a extensão na educação superior brasileira e a caracteriza como um processo interdisciplinar, educativo, cultural, científico e político, que deve integrar-se à matriz curricular e gerar impacto na formação do estudante e na sociedade (BRASIL, 2018). O presente trabalho insere-se nessa perspectiva ao aplicar conhecimentos de desenvolvimento de software a uma demanda social identificada."),
  h2("2.2 INCLUSÃO DIGITAL E A PESSOA IDOSA"),
  p("O Estatuto da Pessoa Idosa considera pessoa idosa aquela com idade igual ou superior a 60 anos (BRASIL, 2003). Os dados da pesquisa TIC Domicílios 2024 indicam que, embora a maior parte desse grupo já utilize a internet, 37% das pessoas com 60 anos ou mais ainda não eram usuárias da rede, proporção muito superior à observada nas demais faixas etárias (NIC.BR, 2024). Esse cenário evidencia a importância de iniciativas que promovam não apenas o acesso, mas o uso seguro e consciente das tecnologias digitais por esse público."),
  h2("2.3 SEGURANÇA DIGITAL DO USUÁRIO FINAL"),
  p("De acordo com a Cartilha de Segurança para Internet, a engenharia social consiste em enganar uma pessoa explorando sentimentos como medo, obediência, confiança, carência, ganância, curiosidade e vontade de ajudar. Entre os sinais de alerta estão a ameaça, a pressa, o apelo afetivo, o pedido de segredo e ofertas boas demais para ser verdade; a recomendação central é confirmar a história com a pessoa ou a organização por meio de um contato oficial ou já conhecido (CERT.BR, 2026)."),
  p("No que se refere à autenticação, recomenda-se criar senhas longas, como a combinação de palavras aleatórias, misturar diferentes tipos de caracteres, evitar sequências de teclado e informações pessoais, não reutilizar a mesma senha em diferentes serviços, ativar a verificação em duas etapas e nunca compartilhar senhas ou códigos de verificação por mensagens, e-mails ou ligações (CERT.BR, 2022). Essas orientações fundamentaram o conteúdo das dicas, as regras do verificador de senha e as situações apresentadas no questionário."),
  h2("2.4 ENGENHARIA DE SOFTWARE E REQUISITOS"),
  p("A engenharia de software fornece processos, métodos e ferramentas para a construção de software de qualidade (PRESSMAN; MAXIM, 2016). Sommerville (2018) distingue os requisitos funcionais, que descrevem os serviços que o sistema deve oferecer e como deve reagir a determinadas entradas, dos requisitos não funcionais, que representam restrições e qualidades do sistema, como desempenho, usabilidade e segurança. Essa distinção foi adotada na especificação apresentada no capítulo 4."),
  p("Para a modelagem, utilizou-se a Linguagem de Modelagem Unificada (UML), por meio dos diagramas de casos de uso e de classes, além do Diagrama de Entidade-Relacionamento (DER) para representar a estrutura dos dados. A construção seguiu uma abordagem incremental, na qual o sistema é desenvolvido em partes que são testadas e incorporadas progressivamente (SOMMERVILLE, 2018)."),
  h2("2.5 TECNOLOGIAS WEB"),
  p("A aplicação foi construída com as tecnologias padrão da web: a Linguagem de Marcação de Hipertexto (HTML5), responsável pela estrutura e pela semântica do conteúdo; as Folhas de Estilo em Cascata (CSS3), responsáveis pela apresentação visual e pela adaptação a diferentes tamanhos de tela; e a linguagem JavaScript, responsável pela interatividade. Por serem executadas diretamente pelo navegador, essas tecnologias dispensam instalação de programas pelo usuário e permitem a hospedagem gratuita em serviços de páginas estáticas, como o GitHub Pages."),
  h2("2.6 ACESSIBILIDADE E USABILIDADE"),
  p("As Diretrizes de Acessibilidade para Conteúdo Web (WCAG) 2.2, publicadas pelo World Wide Web Consortium (W3C), organizam recomendações para tornar o conteúdo web mais acessível a pessoas com deficiência, incluindo pessoas idosas com alterações de visão, audição, mobilidade ou cognição decorrentes do envelhecimento (W3C, 2025). As diretrizes estão estruturadas em quatro princípios: o conteúdo deve ser perceptível, operável, compreensível e robusto. No projeto, esses princípios orientaram decisões como o uso de texto ampliado, contraste adequado, foco visível na navegação por teclado, áreas de toque amplas e linguagem simples."),
];

const cap3 = [
  h1("3 COMUNIDADE ATENDIDA E IMPACTO SOCIAL"),
  p("O público-alvo do projeto é formado por pessoas idosas, isto é, com 60 anos ou mais (BRASIL, 2003), que utilizam celular ou computador para se comunicar, acessar serviços bancários ou realizar compras e que possuem pouca familiaridade com práticas de segurança digital."),
  p("Para a validação da solução, o projeto atende um público menor, composto por dois a três familiares do autor. Por se tratar de um projeto desenvolvido individualmente, essa escolha viabiliza o contato direto com os usuários, a observação do uso da aplicação e a coleta de percepções em curto prazo."),
  p([t("Perfil dos participantes: "), pend("[PENDENTE – informar, para cada participante, o grau de parentesco, a faixa etária, o dispositivo utilizado (celular ou computador), a familiaridade com a tecnologia e eventuais necessidades de acessibilidade, como uso de óculos ou dificuldade de leitura.]")]),
  p("Considerando o perfil do público, a interface foi projetada com letras grandes, possibilidade de ampliação do texto, modo de alto contraste, botões amplos e linguagem direta. Espera-se que a solução contribua para que os participantes reconheçam os sinais mais comuns de golpe, criem senhas mais fortes e adotem a prática de confirmar pedidos suspeitos por um canal oficial, ganhando mais autonomia e segurança no uso da internet. O material também pode ser compartilhado pelos próprios participantes com amigos e outros familiares, ampliando o alcance da ação."),
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
];
const RNF = [
  ["RNF-001", "A interface deve ser responsiva, adaptando-se a telas de celulares sem rolagem horizontal."],
  ["RNF-002", "Os dados digitados pelo usuário não devem ser enviados pela rede; todo o processamento deve ocorrer no navegador."],
  ["RNF-003", "O texto base deve ter, no mínimo, 20 pixels, e os botões devem ter área de toque mínima de 48 × 48 pixels."],
  ["RNF-004", "Todas as funções devem ser acessíveis por teclado, com indicação visível do elemento em foco."],
  ["RNF-005", "A aplicação deve funcionar em navegadores modernos sem necessidade de instalação ou cadastro."],
  ["RNF-006", "A aplicação deve ser hospedada em serviço gratuito de páginas estáticas (GitHub Pages)."],
  ["RNF-007", "A interface deve utilizar linguagem simples e direta, adequada ao público idoso."],
];

const cap4 = [
  h1("4 ESPECIFICAÇÃO DO SISTEMA E REQUISITOS"),
  p("Este capítulo descreve os requisitos, a arquitetura e a modelagem do sistema Navega Seguro."),
  h2("4.1 REQUISITOS FUNCIONAIS (RF)"),
  p("O Quadro 1 apresenta os requisitos funcionais do sistema."),
  legenda("Quadro 1 – Requisitos funcionais"),
  quadro(["Código", "Descrição"], RF, [1500, LARGURA_UTIL - 1500]),
  fonteIlustr(),
  h2("4.2 REQUISITOS NÃO FUNCIONAIS (RNF)"),
  p("O Quadro 2 apresenta os requisitos não funcionais do sistema."),
  legenda("Quadro 2 – Requisitos não funcionais"),
  quadro(["Código", "Descrição"], RNF, [1500, LARGURA_UTIL - 1500]),
  fonteIlustr(),
  h2("4.3 ARQUITETURA DO SISTEMA E TECNOLOGIAS UTILIZADAS"),
  p("O Navega Seguro adota uma arquitetura cliente, composta exclusivamente por arquivos estáticos executados no navegador do usuário. Não há servidor de aplicação nem banco de dados remoto: as categorias, as dicas e as perguntas do questionário são mantidas em uma estrutura de dados JavaScript, organizada conforme o modelo apresentado no DER. Essa decisão atende ao requisito de privacidade (RNF-002), reduz custos e simplifica a manutenção. A única informação armazenada é a preferência de tamanho de texto e de contraste, guardada localmente no próprio navegador (localStorage). O Quadro 3 resume a pilha tecnológica."),
  legenda("Quadro 3 – Tecnologias utilizadas"),
  quadro(["Camada", "Tecnologia"], [
    ["Front-end (estrutura)", "HTML5 semântico com atributos ARIA"],
    ["Front-end (apresentação)", "CSS3 com variáveis, Grid e Flexbox (layout responsivo)"],
    ["Front-end (comportamento)", "JavaScript (ES6), organizado em classes"],
    ["Dados", "Estrutura de dados em JavaScript (arquivo dados.js) e localStorage para preferências"],
    ["Hospedagem", "GitHub Pages (gratuito)"],
    ["Versionamento", "Git e GitHub"],
    ["Testes", "Playwright (testes funcionais automatizados no navegador Chromium)"],
  ], [3200, LARGURA_UTIL - 3200]),
  fonteIlustr(),
  p("Os arquivos do projeto estão organizados da seguinte forma: index.html (página principal), css/style.css (estilos), js/dados.js (dados), js/acessibilidade.js, js/dicas.js, js/senha.js e js/quiz.js (lógica de cada funcionalidade)."),
  p([t("Endereço do repositório: "), pend("[PENDENTE – link do repositório no GitHub]"), t(". Endereço da aplicação publicada: "), pend("[PENDENTE – link do GitHub Pages]"), t(".")]),
  h2("4.4 MODELAGEM E DIAGRAMAS DE SISTEMA"),
  p("A Figura 1 apresenta o diagrama de casos de uso. O ator Usuário, que representa a pessoa idosa, interage com as funcionalidades de consulta de dicas, verificação de senha, questionário e acessibilidade."),
  ...figura("Figura 1 – Diagrama de casos de uso", "diagrama-casos-de-uso.png", 15),
  p("A Figura 2 apresenta o diagrama de classes. As classes Acessibilidade, CatalogoDicas, VerificadorSenha e Quiz concentram a lógica de cada funcionalidade, enquanto Categoria, Dica e Pergunta representam os dados manipulados."),
  ...figura("Figura 2 – Diagrama de classes", "diagrama-classes.png", 16),
  p("A Figura 3 apresenta o Diagrama de Entidade-Relacionamento (DER). Uma categoria possui várias dicas e classifica várias perguntas, caracterizando dois relacionamentos do tipo um-para-muitos. Embora os dados sejam mantidos em JavaScript, o modelo permite a migração futura para um banco de dados relacional sem alteração da estrutura."),
  ...figura("Figura 3 – Diagrama de Entidade-Relacionamento", "diagrama-der.png", 13),
  p("As Figuras 4 a 7 apresentam as telas do sistema, que também cumprem a função de protótipo navegável da interface."),
  ...figura("Figura 4 – Tela inicial e menu de navegação", "tela-inicio.png", 14),
  ...figura("Figura 5 – Verificador de força de senha", "tela-senha.png", 14),
  ...figura("Figura 6 – Questionário “É golpe?” com explicação após a resposta", "tela-quiz-explicacao.png", 14),
  ...figura("Figura 7 – Modo de alto contraste com texto ampliado", "tela-contraste.png", 14),
];

const cronograma = [
  ["Definição do tema, levantamento de requisitos e modelagem", "01/11/2026", "Concluída"],
  ["Desenvolvimento do site (HTML, CSS e JavaScript)", "01/11/2026", "Concluída"],
  ["Testes funcionais automatizados e correções", "01/11/2026", "Concluída"],
  ["Elaboração da documentação (normas ABNT)", "01/11/2026", "Concluída"],
  ["Publicação no GitHub e no GitHub Pages", "Previsto: até 08/11/2026", "Pendente"],
  ["Testes com familiares e coleta de feedback", "Previsto: até 10/11/2026", "Pendente"],
  ["Gravação do vídeo de apresentação", "Previsto: até 12/11/2026", "Pendente"],
  ["Revisão final e entrega no Google Sala de Aula", "Previsto: até 15/11/2026", "Pendente"],
];

const cap5 = [
  h1("5 METODOLOGIA DE DESENVOLVIMENTO E CRONOGRAMA"),
  p("Para a organização do trabalho foi adotado o método Kanban, que se baseia na visualização do fluxo de trabalho e na limitação do trabalho em andamento, favorecendo entregas contínuas e a identificação de gargalos (ANDERSON, 2010). As tarefas foram distribuídas nas etapas A fazer, Em desenvolvimento e Concluído, e o acompanhamento foi registrado na planilha de distribuição de tarefas solicitada pela disciplina."),
  p("O desenvolvimento seguiu uma abordagem incremental (SOMMERVILLE, 2018): inicialmente foram definidos os requisitos e a modelagem; em seguida, cada funcionalidade (dicas, verificador de senha, questionário e acessibilidade) foi implementada e testada separadamente antes de ser integrada à página principal."),
  p("Por se tratar de um projeto individual, todas as funções foram exercidas pelo autor, Kauan Castro Moura: levantamento de requisitos, modelagem, desenvolvimento front-end, testes, documentação, publicação no GitHub, contato com os participantes e gravação do vídeo. O Quadro 4 apresenta o cronograma do projeto, de acordo com o prazo de entrega estabelecido pela disciplina."),
  legenda("Quadro 4 – Cronograma do projeto"),
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
  ["T09", "Aumentar o texto duas vezes", "Fonte base alterada de 20 px para 24 px", "Aprovado"],
  ["T10", "Ativar o alto contraste", "Tema escuro com cores de alto contraste aplicado", "Aprovado"],
  ["T11", "Abrir em tela de celular (390 px de largura)", "Sem rolagem horizontal", "Aprovado"],
  ["T12", "Verificar erros de execução no console", "Nenhum erro registrado", "Aprovado"],
];

const cap6 = [
  h1("6 RESULTADOS OBTIDOS E TESTES DE SOFTWARE"),
  p("Como resultado, foi entregue o portal Navega Seguro, com todas as funcionalidades especificadas no capítulo 4. As telas do sistema finalizado estão apresentadas nas Figuras 4 a 7, e a Figura 8 mostra a adaptação da interface à tela de um celular."),
  ...figura("Figura 8 – Visualização em tela de celular", "tela-celular.png", 6),
  h2("6.1 TESTES FUNCIONAIS"),
  p("Os testes funcionais foram automatizados com a ferramenta Playwright, que executa a aplicação em um navegador real (Chromium) e simula as ações do usuário. O Quadro 5 apresenta os casos de teste executados e seus resultados."),
  legenda("Quadro 5 – Casos de teste funcionais"),
  quadro(["ID", "Caso de teste", "Resultado obtido", "Situação"], testes, [700, 3500, LARGURA_UTIL - 700 - 3500 - 1300, 1300]),
  fonteIlustr(),
  p("Durante os testes identificou-se que, no modo de alto contraste, o botão de ativação desse modo ficava com a mesma cor do cabeçalho, tornando-se pouco visível. O problema foi corrigido com um estilo específico para o botão ativo, e os testes foram executados novamente com sucesso."),
  h2("6.2 VALIDAÇÃO COM O PÚBLICO-ALVO"),
  p("A validação com o público-alvo foi planejada com dois a três familiares do autor. Cada participante utiliza o portal em seu próprio dispositivo, lê as dicas, testa uma senha de exemplo e responde ao questionário, sendo depois convidado a relatar suas percepções sobre a facilidade de uso, a clareza dos textos e a utilidade das informações."),
  p([pend("[PENDENTE – relatar como foram os testes com cada participante: data, dispositivo usado, número de acertos no questionário, dificuldades observadas e se precisaram de ajuda.]")]),
  p([pend("[PENDENTE – inserir os depoimentos ou feedbacks reais dos participantes, identificando-os apenas pelo grau de parentesco e pela idade, por exemplo: “Participante 1 (avó, 72 anos): ...”.]")]),
  p([pend("[PENDENTE – descrever eventuais ajustes feitos no sistema a partir do feedback recebido.]")]),
];

const cap7 = [
  h1("7 CONSIDERAÇÕES FINAIS"),
  p("O objetivo geral deste projeto, desenvolver uma aplicação web funcional e acessível que auxilie pessoas idosas a reconhecer golpes digitais e a adotar práticas seguras na internet, foi atingido no que se refere ao desenvolvimento do software: os requisitos foram mapeados, a interface foi implementada com recursos de acessibilidade e todos os casos de teste funcionais foram aprovados."),
  p([t("Quanto à validação com o público-alvo, "), pend("[PENDENTE – concluir, com base no feedback real dos familiares, se a aplicação foi considerada fácil de usar e útil, e se o objetivo específico de validação foi atingido.]")]),
  p("Do ponto de vista acadêmico, o projeto permitiu aplicar na prática conteúdos do curso de Ciência da Computação, como levantamento de requisitos, modelagem UML, desenvolvimento web, testes de software e versionamento com Git. Do ponto de vista extensionista, evidenciou que a tecnologia precisa ser pensada a partir das necessidades reais do usuário: linguagem simples, letras grandes e exemplos do cotidiano mostraram-se tão importantes quanto o código."),
  p("Como trabalhos futuros para uma versão 2.0, sugerem-se: a inclusão de leitura em voz alta do conteúdo; a ampliação do banco de situações do questionário, com atualização periódica a partir de novos golpes divulgados por fontes oficiais; a migração dos dados para um banco de dados relacional, seguindo o DER apresentado; a criação de um modo para impressão das dicas; e a realização de oficinas presenciais com grupos maiores de pessoas idosas."),
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
  title: "Navega Seguro: portal de segurança digital para pessoas idosas",
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
