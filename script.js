let banco = {
  vendas: [],
  estoque: []
};

function tentarRenderizar(){
  if(banco.vendas.length > 0 && banco.estoque.length > 0){
    renderTabela();
  }
}

function carregarVendas(event){
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e){
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    banco.vendas = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    preencherFiltroVendas();
    renderTabela(); // já renderiza mesmo sem estoque
  };

  reader.readAsArrayBuffer(file);
}

function carregarEstoque(event){
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e){
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    banco.estoque = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    preencherFiltroEstoque();
    renderTabela();
  };

  reader.readAsArrayBuffer(file);
}

function tentarRenderizar(){
  if(banco.vendas.length > 0 && banco.estoque.length > 0){
    renderTabela();
  }
}

function carregarArquivo(event){
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e){
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    // ABA 1 = vendas
    banco.vendas = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[0]]
    );

    // ABA 2 = estoque filiais
    banco.estoque = XLSX.utils.sheet_to_json(
      workbook.Sheets[workbook.SheetNames[1]]
    );

    renderTabela();
  };

  reader.readAsArrayBuffer(file);
}

function renderTabela(){

  let dadosVendas = [...(banco.vendas || [])];
  let dadosEstoque = banco.estoqueFiltrado || banco.estoque || [];

  if(dadosVendas.length === 0){
    document.getElementById("tabela").innerHTML = "Sem vendas";
    return;
  }

  // 🔹 FILTRO VENDAS (SÓ VENDAS)
  const filtroVendas = document.getElementById("filtroGrupoVendas")?.value;
  if(filtroVendas){
    dadosVendas = dadosVendas.filter(item =>
      String(item[banco.campoGrupoVendas]).trim() === filtroVendas
    );
  }

  // 🔹 MAPEAMENTO
  const mapa = {
    grupo: Object.keys(dadosVendas[0]).find(c => c.toLowerCase().includes("grupo")),
    descricao: Object.keys(dadosVendas[0]).find(c => c.toLowerCase().includes("descr")),
    estoque: Object.keys(dadosVendas[0]).find(c => c.toLowerCase().includes("estoque")),
    qtd: Object.keys(dadosVendas[0]).find(c => c.toLowerCase().includes("qtd"))
  };

  const chaveNomeEstoque = dadosEstoque.length > 0
    ? Object.keys(dadosEstoque[0]).find(c => c.toLowerCase().includes("nome"))
    : null;

  let html = "<table border='1'><tr>";

  html += "<th>Grupo Linha</th>";
  html += "<th>Produto</th>";
  html += "<th>Estoque</th>";
  html += "<th>Vendidos</th>";
  html += "<th>Média</th>";
  html += "<th>Pedido</th>";
  html += "<th>Nome Estoque</th>";
  html += "<th>Vargem</th>";
  html += "<th>Agulhas</th>";
  html += "<th>Pedra</th>";

  html += "</tr>";

  dadosVendas.forEach(item => {

    let nome = item[mapa.descricao] || "";

    // 🔹 MATCH SOMENTE PELO NOME (SEM INTERFERÊNCIA)
    let estoqueFilial = {};

    if(chaveNomeEstoque){
      estoqueFilial = dadosEstoque.find(p =>
        String(p[chaveNomeEstoque]).trim().toLowerCase() ===
        String(nome).trim().toLowerCase()
      ) || {};
    }

    let vargem = Number(estoqueFilial["Estoque ATUAL FILIAL: OBOM PET - VARGEM GRANDE"]) || 0;
    let agulhas = Number(estoqueFilial["Estoque ATUAL FILIAL: OBOM PET - AGULHAS NEGRAS"]) || 0;
    let pedra = Number(estoqueFilial["Estoque ATUAL FILIAL: OBOM PET - PEDRA DE GUARATIBA"]) || 0;

    let estoque = Number(item[mapa.estoque]) || 0;
    let vendido = Number(item[mapa.qtd]) || 0;

    let media = vendido / 30;
    let pedido = (media * 30) - estoque;
    if(pedido < 0) pedido = 0;

    html += "<tr>";

    html += `<td>${item[mapa.grupo] || ""}</td>`;
    html += `<td>${corrigirTexto(nome)}</td>`;

    // 🔴 estoque baixo
    let estiloEstoque = estoque <= 3
      ? "style='background:red;color:white;font-weight:bold;'"
      : "";

    html += `<td ${estiloEstoque}>${estoque}</td>`;
    html += `<td>${vendido}</td>`;
    html += `<td>${media.toFixed(2)}</td>`;

    // 🟢 sugestão pedido
    let estiloPedido = pedido > 0
      ? "style='background:green;color:white;font-weight:bold;'"
      : "";

    html += `<td ${estiloPedido}>${Math.ceil(pedido)}</td>`;

    html += `<td>${estoqueFilial[chaveNomeEstoque] || ""}</td>`;
    html += `<td>${vargem}</td>`;
    html += `<td>${agulhas}</td>`;
    html += `<td>${pedra}</td>`;

    html += "</tr>";
  });

  html += "</table>";

  document.getElementById("tabela").innerHTML = html;
}

function calcular(loja){
  const dados = banco[loja];

  let soma = 0;
  let count = 0;

  dados.forEach(linha=>{
    Object.values(linha).forEach(v=>{
      if(!isNaN(v)){
        soma += Number(v);
        count++;
      }
    });
  });

  let media = count ? soma / count : 0;

  document.getElementById("soma_"+loja).innerText = soma.toFixed(2);
  document.getElementById("media_"+loja).innerText = media.toFixed(2);
}

function filtrarLojas(valor){
  ["loja1","loja2","loja3"].forEach(loja=>{
    document.getElementById(loja).style.display =
      (valor === "todas" || valor === loja) ? "block" : "none";
  });
}

function pegarCampo(obj, nome){
  return Object.keys(obj).find(k => k.toLowerCase().includes(nome));
}

function preencherFiltros(loja){
  const dados = banco[loja];
  if(dados.length === 0) return;

  const campoGrupoLinha = Object.keys(dados[0]).find(c =>
    c.toLowerCase().includes("grupo linha")
  );

  let valores = new Set();

  dados.forEach(item=>{
    if(item[campoGrupoLinha]) valores.add(item[campoGrupoLinha]);
  });

  preencherSelect("grupo_"+loja, valores);

  // salvar nome real da coluna
  banco[loja+"_grupo"] = campoGrupoLinha;
}

function preencherSelect(id, valores){
  const select = document.getElementById(id);
  if(!select) return;

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}

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

function filtrarDados(){
  const valor = document.getElementById("grupo_filtro").value;

  let dados = banco.vendas;

  if(valor){
    const campoGrupo = banco.campoGrupo;

    dados = banco.vendas.filter(item =>
      String(item[campoGrupo]).trim() === valor
    );
  }

  renderTabela(dados);
}

function buscarEstoqueProduto(loja, nomeProduto){
  const estoque = banco[loja+"_estoque"];
  if(!estoque || estoque.length === 0) return {};

  const chaveNome = Object.keys(estoque[0]).find(c =>
    c.toLowerCase().includes("nome")
  );

  if(!chaveNome) return {};

  const item = estoque.find(p =>
    String(p[chaveNome]).trim().toLowerCase() ===
    String(nomeProduto).trim().toLowerCase()
  );

  return item || {};
}

function filtrarGrupo(){
  const valor = document.getElementById("filtroGrupo").value;

  let dados = banco.vendas;

  if(valor){
    dados = dados.filter(item =>
      String(item[banco.campoGrupo]).trim() === valor
    );
  }

  renderTabela(dados);
}

function preencherFiltroVendas(){
  const dados = banco.vendas;
  if(!dados.length) return;

  const campo = Object.keys(dados[0]).find(c =>
    c.toLowerCase().includes("grupo linha")
  );

  banco.campoGrupoVendas = campo;

  const valores = [...new Set(dados.map(d => d[campo]))];

  const select = document.getElementById("filtroGrupoVendas");

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}
function filtrarVendas(){
  const valor = document.getElementById("filtroGrupoVendas").value;

  let dados = banco.vendas;

  if(valor){
    dados = dados.filter(item =>
      String(item[banco.campoGrupoVendas]).trim() === valor
    );
  }

  renderTabela(dados);
}

function filtrarEstoque(){
  const valor = document.getElementById("filtroGrupoEstoque").value;

  banco.estoqueFiltrado = valor
    ? banco.estoque.filter(item =>
        String(item[banco.campoGrupoEstoque]).trim() === valor
      )
    : banco.estoque;

  renderTabela(); // NÃO mexe em vendas
}

function preencherFiltroEstoque(){
  const dados = banco.estoque;
  if(!dados.length) return;

  const campo = Object.keys(dados[0]).find(c =>
    c.toLowerCase().includes("grupo linha")
  );

  banco.campoGrupoEstoque = campo;

  const valores = [...new Set(dados.map(d => d[campo]))];

  const select = document.getElementById("filtroGrupoEstoque");

  select.innerHTML = `<option value="">Todos</option>`;

  valores.forEach(v=>{
    select.innerHTML += `<option value="${v}">${v}</option>`;
  });
}