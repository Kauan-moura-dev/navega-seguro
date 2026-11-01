/* Navega Seguro – exibição e filtro das dicas por categoria. */
class CatalogoDicas {
  constructor(categorias, dicas, elFiltros, elLista) {
    this.categorias = categorias;
    this.dicas = dicas;
    this.elFiltros = elFiltros;
    this.elLista = elLista;
    this.categoriaAtual = 0; // 0 = todas
  }

  iniciar() {
    const opcoes = [{ id: 0, nome: "Todas" }, ...this.categorias];
    opcoes.forEach(cat => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = cat.nome;
      b.setAttribute("aria-pressed", String(cat.id === this.categoriaAtual));
      b.addEventListener("click", () => this.filtrar(cat.id));
      this.elFiltros.appendChild(b);
    });
    this.renderizar();
  }

  filtrar(categoriaId) {
    this.categoriaAtual = categoriaId;
    [...this.elFiltros.children].forEach((b, i) => {
      const id = i === 0 ? 0 : this.categorias[i - 1].id;
      b.setAttribute("aria-pressed", String(id === categoriaId));
    });
    this.renderizar();
  }

  nomeCategoria(id) {
    const c = this.categorias.find(c => c.id === id);
    return c ? c.nome : "";
  }

  renderizar() {
    const lista = this.categoriaAtual === 0
      ? this.dicas
      : this.dicas.filter(d => d.categoriaId === this.categoriaAtual);
    this.elLista.innerHTML = "";
    lista.forEach(d => {
      const art = document.createElement("article");
      art.className = "cartao";
      art.innerHTML = `<span class="categoria"></span><h3></h3><p></p>`;
      art.querySelector(".categoria").textContent = this.nomeCategoria(d.categoriaId);
      art.querySelector("h3").textContent = d.titulo;
      art.querySelector("p").textContent = d.texto;
      this.elLista.appendChild(art);
    });
  }
}

new CatalogoDicas(CATEGORIAS, DICAS,
  document.getElementById("filtro-categorias"),
  document.getElementById("lista-dicas")).iniciar();
