/* Navega Seguro – quiz "É golpe?". */
class Quiz {
  constructor(perguntas, elArea) {
    this.perguntas = perguntas;
    this.el = elArea;
    this.indice = 0;
    this.acertos = 0;
  }

  iniciar() {
    this.indice = 0;
    this.acertos = 0;
    this.mostrarPergunta();
  }

  mostrarPergunta() {
    const p = this.perguntas[this.indice];
    this.el.innerHTML = `
      <p class="progresso">Situação ${this.indice + 1} de ${this.perguntas.length}</p>
      <h3 id="situacao"></h3>
      <div class="botoes">
        <button type="button" data-r="true">É golpe</button>
        <button type="button" data-r="false">É seguro</button>
      </div>`;
    this.el.querySelector("#situacao").textContent = p.situacao;
    this.el.querySelectorAll("[data-r]").forEach(b =>
      b.addEventListener("click", () => this.responder(b.dataset.r === "true")));
    this.el.querySelector("[data-r]").focus();
  }

  responder(escolha) {
    const p = this.perguntas[this.indice];
    const certo = escolha === p.resposta;
    if (certo) this.acertos++;
    this.el.querySelectorAll("[data-r]").forEach(b => (b.disabled = true));
    const div = document.createElement("div");
    div.className = "explicacao";
    div.innerHTML = `<p><strong></strong></p><p></p>`;
    div.querySelector("strong").textContent = certo ? "Você acertou!" : "Resposta incorreta.";
    div.querySelector("p:last-child").textContent = p.explicacao;
    const prox = document.createElement("button");
    prox.type = "button";
    const ultima = this.indice === this.perguntas.length - 1;
    prox.textContent = ultima ? "Ver resultado" : "Próxima situação";
    prox.addEventListener("click", () => {
      this.indice++;
      ultima ? this.mostrarResultado() : this.mostrarPergunta();
    });
    this.el.appendChild(div);
    this.el.appendChild(prox);
    prox.focus();
  }

  mostrarResultado() {
    const total = this.perguntas.length;
    const msg = this.acertos === total ? "Excelente! Você reconheceu todas as situações."
      : this.acertos >= total / 2 ? "Muito bem! Revise as dicas para ficar ainda mais protegido."
      : "Vale a pena ler as dicas com calma e tentar de novo.";
    this.el.innerHTML = `<h3>Você acertou ${this.acertos} de ${total}.</h3><p></p>
      <button type="button" id="btn-reiniciar">Fazer de novo</button>`;
    this.el.querySelector("p").textContent = msg;
    const b = this.el.querySelector("#btn-reiniciar");
    b.addEventListener("click", () => this.iniciar());
    b.focus();
  }
}

new Quiz(PERGUNTAS, document.getElementById("area-quiz")).iniciar();
