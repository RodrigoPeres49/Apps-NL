
// ADICIONANDO O ITEM

async function adicionar(event) {
    event.preventDefault();

    const data = document.getElementById("data").value;
    const [ano, mes, dia] = data.split("-");
    const dataFormatada = `${dia}/${mes}/${ano}`
    const produto = document.getElementById("produto").value;
    const quantidade = document.getElementById("quantidade").value.replace(",", ".");
    const preco = document.getElementById("preco").value.replace(",", ".");
    let valor = quantidade * preco;
    valor = valor.toFixed(2);

    const resposta = await window.pywebview.api.adicionar_item(dataFormatada, produto, quantidade, preco, valor);
    preencherTabela(resposta);

    // RESETANDO VALORES EXCETO A DATA

    document.getElementById("produto").value = "";
    document.getElementById("quantidade").value = "";
    document.getElementById("preco").value = "";
}

// APAGANDO ITEM

async function apagarItem(id) {
    if (!confirm("Deseja apagar este item?")) return;
    await window.pywebview.api.apagar_item(id);
    carregarTabela();
}

// EDITANDO ITEM

async function editarItem(id) {
    const data = prompt("Digite a nova Data (ex: 24/07/2025)");
    if (!data) return;

    const produto = prompt("Digite o novo Produto");
    if (!produto) return;

    const quantidade = prompt("Digite a nova Quantidade (ex: 10 ou 10,5)").replace(",",".");
    if (!quantidade) return;

    const preco = prompt("Digite o novo Preço (ex: 5 ou 5,25)").replace(",",".");
    if (!preco) return;

    let valor = parseFloat(quantidade.replace(",", ".")) * parseFloat(preco.replace(",", "."));
    valor = valor.toFixed(2)

    await window.pywebview.api.editar_item(id, data, produto, quantidade, preco, valor);
    carregarTabela();
}

// CARREGANDO DADOS NO EXCEL

async function carregarTabela() {
    const dados = await window.pywebview.api.carregar_dados();
    preencherTabela(dados);
}

// ARQUIVANDO DADOS EM UM ARQUIVO

async function armazenar_dados() {
    if (!confirm("Deseja gerar um arquivo com os dados abaixo?")) return;
    const caminho = await window.pywebview.api.armazenar_planilha();
    alert("Arquivo armazenado com sucesso em:\n" + caminho);
    carregarTabela()
}

// LISTANDO ITENS NO APLICATIVO

function preencherTabela(dados) {
    const corpo = document.getElementById("tabela-corpo");
    corpo.innerHTML = "";
    dados.forEach(linha => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${linha["Data"]}</td>
            <td>${linha["Produto"]}</td>
            <td>${linha["Quantidade"]}</td>
            <td>${linha["Preço"]}</td>
            <td>${linha["Valor Total"]}</td>
            <td>
                <button onclick="editarItem(${linha["idRegistro"]})">Editar</button>
                <button onclick="apagarItem(${linha["idRegistro"]})">Apagar</button>
            </td>
        `;
        corpo.appendChild(tr);
    });
}


// CARREGANDO LOJAS NO SELECT A PARTIR DO JSON

async function carregarSelectLojas() {
  try {
    const lojas = await window.pywebview.api.carregar_lojas();
    const select = document.getElementById("loja");

    select.innerHTML = '<option value="">Selecione a loja</option>';

    lojas.forEach(loja => {
      const option = document.createElement("option");
      option.value = loja.Loja;
      option.textContent = loja.Loja;
      select.appendChild(option);
    });
  } catch (error) {
    console.error("Erro ao carregar lojas:", error);
  }
}


// CARREGANDO NOME DA LOJA

async function carregar_json() {
    const loja = await window.pywebview.api.carregar_nome_loja();
    exibir_nome(loja); // não aparece
}

// CARREGANDO NOME DA LOJA NO LAYOUT

async function exibir_nome(loja) {
    local = document.getElementById("loja_definida");
    local.innerHTML = `Loja: ${loja}` || "Não definida";
}

// DEFININDO NOME DA LOJA

async function nome_loja() {
    const nome = document.getElementById("loja").value
    console.log(nome)
    if (!confirm("Deseja atualizar ou inserir o nome da loja?")) return;
    await window.pywebview.api.definir_nome_loja(nome);
    alert("Nome da loja atualizado!");
    carregar_json();
}

window.addEventListener('pywebviewready', async () => {
    await carregarSelectLojas();
    await carregarTabela();
    await carregar_json();
});