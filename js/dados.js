/* Navega Seguro – dados do sistema (categorias, dicas e perguntas do quiz).
   Estrutura equivalente ao modelo de dados (DER) descrito na documentação. */

const CATEGORIAS = [
  { id: 1, nome: "Senhas" },
  { id: 2, nome: "Golpes por mensagem" },
  { id: 3, nome: "Telefone e banco" },
  { id: 4, nome: "Compras e links" }
];

const DICAS = [
  { id: 1, categoriaId: 1, titulo: "Use senhas longas", texto: "Prefira frases com várias palavras, por exemplo: \"Cafe com bolo na varanda!\". Quanto maior a senha, mais difícil de descobrir." },
  { id: 2, categoriaId: 1, titulo: "Uma senha para cada serviço", texto: "Não use a mesma senha no banco, no e-mail e nas redes sociais. Se uma vazar, as outras continuam protegidas." },
  { id: 3, categoriaId: 1, titulo: "Evite dados pessoais", texto: "Não use data de nascimento, nome de netos ou números de telefone. Essas informações são fáceis de descobrir." },
  { id: 4, categoriaId: 2, titulo: "Desconfie de pedidos de dinheiro", texto: "Se um parente pedir dinheiro por mensagem dizendo que trocou de número, ligue para o número antigo dele antes de qualquer coisa." },
  { id: 5, categoriaId: 2, titulo: "Nunca passe códigos recebidos por SMS", texto: "Códigos de verificação são pessoais. Quem pede esse código pode estar tentando roubar a sua conta do WhatsApp." },
  { id: 6, categoriaId: 2, titulo: "Ative a verificação em duas etapas", texto: "No WhatsApp, entre em Configurações > Conta > Confirmação em duas etapas e crie um PIN. Peça ajuda a alguém de confiança se precisar." },
  { id: 7, categoriaId: 3, titulo: "O banco não pede senha por telefone", texto: "Bancos não ligam pedindo senha, código ou para você fazer transferência de \"teste\". Desligue e ligue para o número do cartão." },
  { id: 8, categoriaId: 3, titulo: "Não instale aplicativos a pedido de estranhos", texto: "Golpistas pedem para instalar programas que controlam o celular. Só instale aplicativos que você mesmo procurou na loja oficial." },
  { id: 9, categoriaId: 4, titulo: "Preço bom demais é sinal de alerta", texto: "Promoções muito baratas em links recebidos por mensagem costumam ser sites falsos. Pesquise o nome da loja antes de comprar." },
  { id: 10, categoriaId: 4, titulo: "Confira o endereço do site", texto: "Veja se o endereço está escrito corretamente e se aparece o cadeado. Mesmo assim, o cadeado sozinho não garante que o site seja confiável." }
];

/* resposta: true = é golpe; false = é seguro */
const PERGUNTAS = [
  { id: 1, categoriaId: 2, situacao: "Você recebe uma mensagem: \"Oi mãe, troquei de número. Pode me fazer um Pix urgente? Depois te explico.\"", resposta: true, explicacao: "É o golpe do falso parente. Ligue para o número antigo da pessoa antes de enviar qualquer valor." },
  { id: 2, categoriaId: 3, situacao: "Uma pessoa liga dizendo ser do seu banco e pede que você digite sua senha no telefone para \"cancelar uma compra\".", resposta: true, explicacao: "O banco nunca pede senha por telefone. Desligue e ligue para o número que está no verso do seu cartão." },
  { id: 3, categoriaId: 2, situacao: "Chega um SMS com um código de 6 números e, logo depois, alguém no WhatsApp pede que você envie esse código.", resposta: true, explicacao: "Esse código serve para ativar o seu WhatsApp em outro aparelho. Nunca o compartilhe com ninguém." },
  { id: 4, categoriaId: 4, situacao: "Você mesmo abriu o aplicativo oficial do banco, que já estava instalado, para consultar o saldo.", resposta: false, explicacao: "Usar o aplicativo oficial, instalado por você, é uma forma segura de acessar o banco." },
  { id: 5, categoriaId: 4, situacao: "Um link recebido por mensagem oferece uma TV nova pela metade do preço, só por hoje.", resposta: true, explicacao: "Ofertas exageradas com pressa para decidir são sinais típicos de site falso." },
  { id: 6, categoriaId: 1, situacao: "Sua neta ajudou você a criar uma senha longa, diferente para cada aplicativo, e anotou em um caderno guardado em casa.", resposta: false, explicacao: "Senhas longas e diferentes são recomendadas. Guardar o caderno em lugar seguro, longe do celular, é uma prática aceitável." },
  { id: 7, categoriaId: 3, situacao: "Alguém liga pedindo que você instale um aplicativo para \"proteger sua conta\" e informe o código que aparecer na tela.", resposta: true, explicacao: "Esse tipo de aplicativo pode dar controle do seu celular ao golpista. Desligue imediatamente." },
  { id: 8, categoriaId: 3, situacao: "Você recebe uma notificação de compra que não fez e decide ligar para o telefone impresso no seu cartão.", resposta: false, explicacao: "Ligar você mesmo para o número oficial é a atitude correta para confirmar a situação." }
];
