/* Navega Seguro – notícias recentes sobre golpes digitais.
   Fonte: feed RSS do Google Notícias (busca por termos de golpe), convertido para JSON
   pelo serviço rss2json, pois sites estáticos não podem ler RSS de outro domínio diretamente.
   Nenhum dado do usuário é enviado nessa requisição. Se o serviço falhar, são exibidos
   links fixos para materiais oficiais do CERT.br. */
class FeedNoticias {
  constructor(elLista, elStatus, elBotao) {
    this.elLista = elLista;
    this.elStatus = elStatus;
    this.elBotao = elBotao;
    this.limite = 6;
    this.busca = "https://news.google.com/rss/search?q=" +
      encodeURIComponent('golpe ("whatsapp" OR "pix" OR "falsa central" OR "falso parente" OR "golpe digital")') +
      "&hl=pt-BR&gl=BR&ceid=BR:pt-419";
    this.endpoint = "https://api.rss2json.com/v1/api.json?rss_url=" + encodeURIComponent(this.busca);
    this.reserva = [
      { titulo: "Golpes: não se deixe enganar", fonte: "CERT.br – Cartilha de Segurança para Internet", link: "https://cartilha.cert.br/fasciculos/golpes/fasciculo-golpes.pdf" },
      { titulo: "Autenticação: proteja suas contas", fonte: "CERT.br – Cartilha de Segurança para Internet", link: "https://cartilha.cert.br/fasciculos/autenticacao/fasciculo-autenticacao.pdf" },
      { titulo: "Todos os fascículos da Cartilha", fonte: "CERT.br", link: "https://cartilha.cert.br/fasciculos/" },
    ];
  }

  async carregar() {
    this.elStatus.textContent = "Carregando notícias...";
    this.elBotao.disabled = true;
    try {
      const resp = await fetch(this.endpoint, { cache: "no-store" });
      const dados = await resp.json();
      if (dados.status !== "ok" || !dados.items || !dados.items.length) throw new Error("feed vazio");
      const itens = dados.items.slice(0, this.limite).map(i => this.converter(i));
      this.renderizar(itens);
      this.elStatus.textContent = `Atualizado em ${new Date().toLocaleString("pt-BR")}. As notícias são de responsabilidade dos veículos que as publicaram.`;
    } catch (e) {
      this.renderizar(this.reserva);
      this.elStatus.textContent = "Não foi possível carregar as notícias agora. Veja abaixo materiais oficiais sobre golpes.";
    } finally {
      this.elBotao.disabled = false;
    }
  }

  /* O título do Google Notícias vem como "Título - Veículo"; separamos as duas partes */
  converter(item) {
    const partes = item.title.split(" - ");
    const fonte = partes.length > 1 ? partes.pop() : (item.author || "");
    const data = item.pubDate ? new Date(item.pubDate.replace(" ", "T") + "Z") : null;
    return {
      titulo: partes.join(" - "),
      fonte,
      data: data && !isNaN(data) ? data.toLocaleDateString("pt-BR") : "",
      link: item.link,
    };
  }

  renderizar(itens) {
    this.elLista.innerHTML = "";
    itens.forEach(n => {
      const art = document.createElement("article");
      art.className = "noticia";
      const a = document.createElement("a");
      a.href = n.link; a.target = "_blank"; a.rel = "noopener noreferrer";
      a.textContent = n.titulo;
      const meta = document.createElement("span");
      meta.className = "meta";
      meta.textContent = [n.fonte, n.data].filter(Boolean).join(" • ");
      art.append(a, meta);
      this.elLista.appendChild(art);
    });
  }
}

const feed = new FeedNoticias(
  document.getElementById("lista-noticias"),
  document.getElementById("status-noticias"),
  document.getElementById("btn-atualizar-noticias"));
document.getElementById("btn-atualizar-noticias").addEventListener("click", () => feed.carregar());
feed.carregar();
