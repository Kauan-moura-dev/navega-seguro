/* Navega Seguro – quiz "É golpe?" com opção de sinalizar erro na pergunta. */
const URL_ISSUES = "https://github.com/Kauan-moura-dev/navega-seguro/issues/new";

class Quiz {
  constructor(perguntas, elArea) {
    this.perguntas = perguntas;
    this.el = elArea;
    this.indice = 0;
    this.acertos = 0;
  }

  iniciar(focar = false) {
    this.indice = 0;
    this.acertos = 0;
    this.mostrarPergunta(focar);
  }

  /* Monta o link que abre uma "issue" no GitHub já preenchida com a pergunta sinalizada */
  linkSinalizar(p) {
    const titulo = `Sinalização na pergunta ${p.id} do quiz`;
    const corpo = `Situação: ${p.situacao}\n\nResposta do site: ${p.resposta ? "É golpe" : "É seguro"}\n\nO que está errado ou o que você sugere?\n`;
    return `${URL_ISSUES}?title=${encodeURIComponent(titulo)}&body=${encodeURIComponent(corpo)}&labels=sinalizacao`;
  }

  mostrarPergunta(focar = true) {
    const p = this.perguntas[this.indice];
    const pct = Math.round((this.indice / this.perguntas.length) * 100);
    this.el.innerHTML = `
      <p class="progresso">Situação ${this.indice + 1} de ${this.perguntas.length}</p>
      <div class="trilha" aria-hidden="true"><div style="width:${pct}%"></div></div>
      <h3 id="situacao"></h3>
      <div class="botoes">
        <button type="button" data-r="true">É golpe</button>
        <button type="button" data-r="false">É seguro</button>
      </div>`;
    this.el.querySelector("#situacao").textContent = p.situacao;
    this.el.querySelectorAll("[data-r]").forEach(b =>
      b.addEventListener("click", () => this.responder(b.dataset.r === "true")));
    if (focar) this.el.querySelector("[data-r]").focus();
  }

  responder(escolha) {
    const p = this.perguntas[this.indice];
    const certo = escolha === p.resposta;
    if (certo) this.acertos++;
    this.el.querySelectorAll("[data-r]").forEach(b => (b.disabled = true));

    const div = document.createElement("div");
    div.className = "explicacao " + (certo ? "certo" : "errado");
    div.innerHTML = `<p><strong></strong></p><p></p>`;
    div.querySelector("strong").textContent = certo ? "Você acertou!" : "Resposta incorreta.";
    div.querySelector("p:last-child").textContent = p.explicacao;

    const acoes = document.createElement("div");
    acoes.className = "acoes";
    const prox = document.createElement("button");
    prox.type = "button";
    const ultima = this.indice === this.perguntas.length - 1;
    prox.textContent = ultima ? "Ver resultado" : "Próxima situação";
    prox.addEventListener("click", () => {
      this.indice++;
      ultima ? this.mostrarResultado() : this.mostrarPergunta();
    });
    const sinal = document.createElement("a");
    sinal.className = "sinalizar";
    sinal.href = this.linkSinalizar(p);
    sinal.target = "_blank";
    sinal.rel = "noopener";
    sinal.textContent = "Sinalizar erro ou comentar esta pergunta";
    acoes.append(prox, sinal);

    this.el.append(div, acoes);
    prox.focus();
  }

  mostrarResultado() {
    const total = this.perguntas.length;
    const msg = this.acertos === total ? "Excelente! Você reconheceu todas as situações."
      : this.acertos >= total / 2 ? "Muito bem! Revise as dicas para ficar ainda mais protegido."
      : "Vale a pena ler as dicas com calma e tentar de novo.";
    this.el.innerHTML = `<p class="progresso">Resultado</p>
      <div class="trilha" aria-hidden="true"><div style="width:100%"></div></div>
      <h3>Você acertou ${this.acertos} de ${total}.</h3><p></p>
      <button type="button" id="btn-reiniciar">Fazer de novo</button>`;
    this.el.querySelector("p:not(.progresso)").textContent = msg;
    const b = this.el.querySelector("#btn-reiniciar");
    b.addEventListener("click", () => this.iniciar(true));
    b.focus();
  }
}

new Quiz(PERGUNTAS, document.getElementById("area-quiz")).iniciar(false);
