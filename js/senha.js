/* Navega Seguro – verificador de força de senha.
   A análise é feita apenas no navegador; nada é enviado ou armazenado. */
class VerificadorSenha {
  constructor() {
    this.comuns = ["123456", "12345678", "123456789", "senha", "senha123", "qwerty",
      "abc123", "111111", "000000", "password", "admin", "brasil", "102030", "iloveyou"];
  }

  /* Retorna { pontos (0-6), nivel, sugestoes[] } */
  avaliar(senha) {
    const sugestoes = [];
    if (!senha) return { pontos: 0, nivel: "", sugestoes };

    if (this.comuns.includes(senha.toLowerCase())) {
      return { pontos: 0, nivel: "fraca", sugestoes: ["Esta senha está entre as mais usadas no mundo. Escolha outra."] };
    }

    let pontos = 0;
    if (senha.length >= 8) pontos++; else sugestoes.push("Use pelo menos 8 caracteres (o ideal é 12 ou mais).");
    if (senha.length >= 12) pontos++;
    if (/[a-z]/.test(senha)) pontos++; else sugestoes.push("Inclua letras minúsculas.");
    if (/[A-Z]/.test(senha)) pontos++; else sugestoes.push("Inclua letras maiúsculas.");
    if (/[0-9]/.test(senha)) pontos++; else sugestoes.push("Inclua números.");
    if (/[^A-Za-z0-9]/.test(senha)) pontos++; else sugestoes.push("Inclua símbolos, como ! ? @ # ou espaço.");

    if (/(.)\1\1/.test(senha) || /(012|123|234|345|456|567|678|789|abc|bcd|cde)/i.test(senha)) {
      pontos = Math.max(0, pontos - 1);
      sugestoes.push("Evite sequências ou repetições, como 123, abc ou aaa.");
    }

    let nivel = "fraca";
    if (pontos >= 5) nivel = "forte";
    else if (pontos >= 3) nivel = "media";
    return { pontos, nivel, sugestoes };
  }
}

(function () {
  const verificador = new VerificadorSenha();
  const campo = document.getElementById("campo-senha");
  const barra = document.getElementById("barra-forca");
  const saida = document.getElementById("resultado-senha");
  const lista = document.getElementById("sugestoes-senha");
  const btnMostrar = document.getElementById("btn-mostrar");
  const textos = { fraca: "Senha FRACA", media: "Senha MÉDIA", forte: "Senha FORTE" };
  const cores = { fraca: "var(--perigo)", media: "var(--alerta)", forte: "var(--ok)" };

  campo.addEventListener("input", () => {
    const r = verificador.avaliar(campo.value);
    lista.innerHTML = "";
    if (!r.nivel) {
      saida.textContent = "Digite algo para começar.";
      saida.className = "resultado";
      barra.style.width = "0";
      return;
    }
    saida.textContent = textos[r.nivel];
    saida.className = "resultado " + r.nivel;
    barra.style.width = Math.max(10, (r.pontos / 6) * 100) + "%";
    barra.style.background = cores[r.nivel];
    r.sugestoes.forEach(s => {
      const li = document.createElement("li");
      li.textContent = s;
      lista.appendChild(li);
    });
  });

  btnMostrar.addEventListener("click", () => {
    const mostrar = campo.type === "password";
    campo.type = mostrar ? "text" : "password";
    btnMostrar.textContent = mostrar ? "Ocultar" : "Mostrar";
    btnMostrar.setAttribute("aria-pressed", String(mostrar));
  });
})();
