let banco = {
  vendas: [],
  estoqueVargem: []
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

    renderTabela();
  };

  reader.readAsArrayBuffer(file);
}

// ================= TABELA =================
function renderTabela(){

  let dados = [...banco.vendas];

  if(dados.length === 0){
    document.getElementById("tabela").innerHTML = "Sem vendas";
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

  // coluna vargem automática
  const colunaVargem = banco.estoqueVargem.length > 0
    ? Object.keys(banco.estoqueVargem[0]).find(c =>
        c.toUpperCase().includes("VARGEM")
      )
    : null;

  let html = "<table><tr>";
  html += "<th>Grupo</th>";
  html += "<th>Produto</th>";
  html += "<th>Estoque</th>";
  html += "<th>Vendidos</th>";
  html += "<th>Média</th>";
  html += "<th>Pedido</th>";
  html += "<th>Vargem</th>";
  html += "</tr>";

  let soma = 0;

  dados.forEach((item, index) => {

    let estoque = Number(item[mapa.estoque]) || 0;
    let vendido = Number(item[mapa.qtd]) || 0;

    let media = vendido / 30;
    let pedido = (media * 30) - estoque;
    if(pedido < 0) pedido = 0;

    soma += vendido;

    // 🔥 SEM MATCH (MESMA POSIÇÃO)
    let vargem = 0;

    if(banco.estoqueVargem[index] && colunaVargem){
      vargem = Number(banco.estoqueVargem[index][colunaVargem]) || 0;
    }

    html += "<tr>";

    html += `<td>${item[mapa.grupo] || ""}</td>`;
    html += `<td>${corrigirTexto(item[mapa.descricao] || "")}</td>`;

    let estiloEstoque = estoque <= 3
      ? "style='background:#ff3b3b;color:#fff;font-weight:bold;'"
      : "";

    html += `<td ${estiloEstoque}>${estoque}</td>`;
    html += `<td>${vendido}</td>`;
    html += `<td>${media.toFixed(2)}</td>`;

    let estiloPedido = pedido > 0
      ? "style='background:#00c853;color:#fff;font-weight:bold;'"
      : "";

    html += `<td ${estiloPedido}>${Math.ceil(pedido)}</td>`;
    html += `<td>${vargem}</td>`;

    html += "</tr>";
  });

  html += "</table>";

  document.getElementById("tabela").innerHTML = html;

  let mediaTotal = soma / dados.length;

  document.getElementById("soma_loja1").innerText = soma.toFixed(0);
  document.getElementById("media_loja1").innerText = mediaTotal.toFixed(2);
}

// ================= FILTRO =================
function preencherFiltroVendas(){
  const dados = banco.vendas;
  if(!dados.length) return;

  const campo = Object.keys(dados[0]).find(c =>
    c.toLowerCase().includes("grupo")
  );

  banco.campoGrupo = campo;

  const valores = [...new Set(dados.map(d => d[campo]))];

  const select = document.getElementById("filtroGrupoVendas");

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}

function filtrarVendas(){
  renderTabela();
}

function filtrarLojas(valor){
  document.getElementById("loja1").style.display =
    (valor === "todas" || valor === "loja1") ? "block" : "none";
}

// ================= TEXTO =================
function corrigirTexto(texto){
  if(typeof texto !== "string") return texto;

  return texto
    .replace(/DescriÃ§Ã£o/gi, "Descrição")
    .replace(/Ã§/g, "ç")
    .replace(/Ã£/g, "ã")
    .replace(/Ã¡/g, "á")
    .replace(/Ã©/g, "é")
    .replace(/Ã­/g, "í")
    .replace(/Ã³/g, "ó")
    .replace(/Ãº/g, "ú");
}