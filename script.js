let banco = {
  vendas: [],
  estoqueVargem: [],
  campoGrupo: null,
  colGrupoVargem: null
};

// ================= VENDAS =================
function carregarVendas(event){
  const file = event.target.files[0];
  if(!file) return;

  const reader = new FileReader();

  reader.onload = function(e){
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    banco.vendas = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    preencherFiltroVendas();
    renderTabela();
  };

  reader.readAsArrayBuffer(file);
}

// ================= ESTOQUE VARGEM =================
function carregarEstoqueVargem(event){
  const file = event.target.files[0];
  if(!file) return;

  const reader = new FileReader();

  reader.onload = function(e){
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    banco.estoqueVargem = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    preencherFiltroVargem();
    renderTabelaVargem();

    setTimeout(() => {
      document.getElementById("tabelaVargem").scrollIntoView({
        behavior: "smooth"
      });
    }, 200);
  };

  reader.readAsArrayBuffer(file);
}

// ================= FILTRO VENDAS =================
function preencherFiltroVendas(){
  if(!banco.vendas.length) return;

  const campo = Object.keys(banco.vendas[0]).find(c =>
    c.toLowerCase().includes("grupo")
  );

  banco.campoGrupo = campo;

  const valores = [...new Set(banco.vendas.map(d => d[campo]))];

  const select = document.getElementById("filtroGrupoVendas");

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}

function filtrarVendas(){
  renderTabela();
}

// ================= FILTRO VARGEM =================
function preencherFiltroVargem(){
  if(!banco.estoqueVargem.length) return;

  const colunas = Object.keys(banco.estoqueVargem[0]);

  banco.colGrupoVargem = colunas.find(c =>
    c.toLowerCase().includes("grupo")
  );

  const valores = [...new Set(
    banco.estoqueVargem.map(d => d[banco.colGrupoVargem])
  )];

  const select = document.getElementById("filtroGrupoVargem");

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}

function filtrarVargem(){
  renderTabelaVargem();
}

// ================= TABELA ABC =================
function renderTabela(){

  let dados = [...banco.vendas];

  if(!dados.length){
    document.getElementById("tabela").innerHTML = "";
    return;
  }

  const filtro = document.getElementById("filtroGrupoVendas").value;

  if(filtro){
    dados = dados.filter(item =>
      String(item[banco.campoGrupo]).trim() === filtro
    );
  }

  const mapa = {
    grupo: Object.keys(dados[0]).find(c => c.toLowerCase().includes("grupo")),
    descricao: Object.keys(dados[0]).find(c => c.toLowerCase().includes("descr")),
    estoque: Object.keys(dados[0]).find(c => c.toLowerCase().includes("estoque")),
    qtd: Object.keys(dados[0]).find(c => c.toLowerCase().includes("qtd"))
  };

  let html = `
    <table>
      <tr>
        <th>Grupo</th>
        <th>Produto</th>
        <th>Estoque</th>
        <th>Vendidos</th>
        <th>Média</th>
        <th>Pedido</th>
      </tr>
  `;

  let soma = 0;

  dados.forEach(item => {

    let estoque = Number(item[mapa.estoque]) || 0;
    let vendido = Number(item[mapa.qtd]) || 0;
    let media = vendido / 30;
    let pedido = Math.max((media * 30) - estoque, 0);

    soma += vendido;

    html += `
      <tr>
        <td>${item[mapa.grupo] || ""}</td>
        <td>${corrigirTexto(item[mapa.descricao] || "")}</td>
        <td ${estoque <= 3 ? "style='background:#ff3b3b;color:#fff;font-weight:bold;'" : ""}>${estoque}</td>
        <td>${vendido}</td>
        <td>${media.toFixed(2)}</td>
        <td ${pedido > 0 ? "style='background:#00c853;color:#fff;font-weight:bold;'" : ""}>${Math.ceil(pedido)}</td>
      </tr>
    `;
  });

  html += "</table>";

  document.getElementById("tabela").innerHTML = html;

  document.getElementById("soma_loja1").innerText = soma.toFixed(0);
  document.getElementById("media_loja1").innerText = (soma/dados.length).toFixed(2);
}

// ================= TABELA VARGEM =================
function renderTabelaVargem(){

  let dados = [...banco.estoqueVargem];

  if(!dados.length){
    document.getElementById("tabelaVargem").innerHTML = "";
    return;
  }

  const filtro = document.getElementById("filtroGrupoVargem").value;

  if(filtro){
    dados = dados.filter(item =>
      String(item[banco.colGrupoVargem]).trim() === filtro
    );
  }

  const colunas = Object.keys(dados[0]);

  const colGrupo = colunas.find(c => c.toLowerCase().includes("grupo"));
  const colNome = colunas.find(c => c.toLowerCase().includes("nome"));
  const colEstoque = colunas.find(c => c.toLowerCase().includes("vargem grande"));

  let html = `
    <table>
      <tr>
        <th>Grupo Linha</th>
        <th>Nome</th>
        <th>Estoque Vargem</th>
      </tr>
  `;

  dados.forEach(item => {

    let estoque = Number(item[colEstoque]) || 0;

    html += `
      <tr>
        <td>${corrigirTexto(item[colGrupo] || "")}</td>
        <td>${corrigirTexto(item[colNome] || "")}</td>
        <td ${estoque <= 2 ? "style='background:#ff3b3b;color:#fff;font-weight:bold;'" : ""}>${estoque}</td>
      </tr>
    `;
  });

  html += "</table>";

  document.getElementById("tabelaVargem").innerHTML = html;
}

// ================= TEXTO =================
function corrigirTexto(texto){
  if(typeof texto !== "string") return texto;

  return texto
    .replace(/Ã§/g, "ç")
    .replace(/Ã£/g, "ã")
    .replace(/Ã¡/g, "á")
    .replace(/Ã©/g, "é")
    .replace(/Ã­/g, "í")
    .replace(/Ã³/g, "ó")
    .replace(/Ãº/g, "ú");
}