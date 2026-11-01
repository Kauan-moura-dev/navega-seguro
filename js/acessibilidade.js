/* Navega Seguro – controle de acessibilidade (tamanho da fonte e alto contraste). */
class Acessibilidade {
  constructor() {
    this.tamanho = 20;
    this.minimo = 16;
    this.maximo = 32;
    this.contraste = false;
    this.carregarPreferencias();
    this.aplicar();
  }

  aumentarFonte() { this.tamanho = Math.min(this.maximo, this.tamanho + 2); this.salvarEAplicar(); }
  diminuirFonte() { this.tamanho = Math.max(this.minimo, this.tamanho - 2); this.salvarEAplicar(); }
  alternarContraste() { this.contraste = !this.contraste; this.salvarEAplicar(); }

  aplicar() {
    document.documentElement.style.setProperty("--tam-base", this.tamanho + "px");
    document.body.classList.toggle("contraste", this.contraste);
    const btn = document.getElementById("btn-contraste");
    if (btn) btn.setAttribute("aria-pressed", String(this.contraste));
  }

  salvarEAplicar() {
    try { localStorage.setItem("ns-acess", JSON.stringify({ t: this.tamanho, c: this.contraste })); } catch (e) { /* armazenamento indisponível */ }
    this.aplicar();
  }

  carregarPreferencias() {
    try {
      const p = JSON.parse(localStorage.getItem("ns-acess"));
      if (p) { this.tamanho = p.t; this.contraste = p.c; }
    } catch (e) { /* usa valores padrão */ }
  }
}

const acessibilidade = new Acessibilidade();
document.getElementById("btn-aumentar").addEventListener("click", () => acessibilidade.aumentarFonte());
document.getElementById("btn-diminuir").addEventListener("click", () => acessibilidade.diminuirFonte());
document.getElementById("btn-contraste").addEventListener("click", () => acessibilidade.alternarContraste());
