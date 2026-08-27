function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Estoque TechCorp')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// 1. CARREGAMENTO DAS LISTAS
// ==========================================
function obterListasParaFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1.1 Mapear o estoque atual da aba Resumo
  const abaResumo = ss.getSheetByName("Resumo");
  const dadosResumo = abaResumo.getDataRange().getValues();
  const estoquePorId = {};
  for (let i = 1; i < dadosResumo.length; i++) {
    if (dadosResumo[i][0]) {
      estoquePorId[dadosResumo[i][0]] = Number(dadosResumo[i][2]);
    }
  }

  // 1.2 Mapear Itens Ativos
  const abaCadastro = ss.getSheetByName("Cadastro");
  const dadosCadastro = abaCadastro.getDataRange().getValues();
  const indexStatus = dadosCadastro[0].indexOf("Status"); 
  
  const itensAtivos = [];
  for (let i = 1; i < dadosCadastro.length; i++) {
    let linha = dadosCadastro[i];
    let id = linha[0];
    let nome = linha[1];
    let statusItem = (indexStatus !== -1 && linha[indexStatus]) ? linha[indexStatus].toString().trim() : "Ativo";
    if (id && nome && statusItem === "Ativo") {
      let qtdAtual = estoquePorId[id] || 0;
      itensAtivos.push({ id: id, nome: nome, estoque: qtdAtual });
    }
  }
  
  // 1.3 Mapear RH (Técnicos e Operadores)
  const abaRH = ss.getSheetByName("RH");
  const dadosRH = abaRH.getDataRange().getValues();
  const tecnicos = [];
  const operadores = [];
  for (let i = 1; i < dadosRH.length; i++) {
    let nome = dadosRH[i][0];
    let funcao = dadosRH[i][1];
    let statusRH = dadosRH[i][3]; // Coluna D
    if (nome && statusRH === "Ativo") {
      if (funcao === "Técnico" || funcao === "Operador") tecnicos.push(nome);
      if (funcao === "Operador") operadores.push(nome);
    }
  }

  // 1.4 Mapear Configurações (Motivos e Fornecedores)
  const abaConfig = ss.getSheetByName("Configurações");
  const motivosSaida = [];
  const motivosCanc = [];
  const fornecedores = [];
  if (abaConfig) {
    const dadosConfig = abaConfig.getDataRange().getValues();
    for (let i = 1; i < dadosConfig.length; i++) {
      // Motivos Saída (A-B)
      if (dadosConfig[i][0] && dadosConfig[i][1] === "Ativo") motivosSaida.push(dadosConfig[i][0]);
      // Motivos Cancelamento (C-D)
      if (dadosConfig[i][2] && dadosConfig[i][3] === "Ativo") motivosCanc.push(dadosConfig[i][2]);
      // Fornecedores (E-F)
      if (dadosConfig[i][4] && dadosConfig[i][5] === "Ativo") fornecedores.push(dadosConfig[i][4]);
    }
  }
  
  itensAtivos.sort((a, b) => a.nome.localeCompare(b.nome));
  tecnicos.sort();
  operadores.sort();
  motivosSaida.sort();
  motivosCanc.sort();
  fornecedores.sort();
  return { 
    itens: itensAtivos, 
    tecnicos: tecnicos, 
    operadores: operadores, 
    motivos: motivosSaida, 
    motivosCanc: motivosCanc, 
    fornecedores: fornecedores 
  };
}

// ==========================================
// 2. BUSCAR COMPRAS PENDENTES (CORREÇÃO DATA)
// ==========================================
function obterComprasPendentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCompras = ss.getSheetByName("Histórico de Compras");
  if (!abaCompras) return [];
  
  const dados = abaCompras.getDataRange().getValues();
  const pendentes = [];
  for (let i = 1; i < dados.length; i++) {
    let status = dados[i][12];
    // Coluna M
    if (status === "Pendente") {
      
      // Converte objeto Date em String formatada para não quebrar o envio ao App
      let dataPrevisao = dados[i][10];
      // Coluna K
      if (dataPrevisao instanceof Date) {
        dataPrevisao = Utilities.formatDate(dataPrevisao, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      pendentes.push({
        idCompra: dados[i][0],      // A
        idItem: dados[i][1],        // B
        nomeItem: dados[i][5],      // F
        qtdComprada: Number(dados[i][6]),   // G
        fornecedor: dados[i][9],    // J
        previsao: dataPrevisao,     // K (String)
 
        qtdRecebida: Number(dados[i][11]),  // L
        operador: dados[i][15]      // P
      });
    }
  }
  return pendentes;
}

// ==========================================
// 3. CANCELAMENTO DE COMPRA
// ==========================================
function cancelarCompraPendente(dadosCancelamento) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaCompras = ss.getSheetByName("Histórico de Compras");
    const dados = abaCompras.getDataRange().getValues();
    const dataHoje = new Date();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === dadosCancelamento.idCompra) {
        let linha = i + 1;
        abaCompras.getRange(linha, 13).setValue("Cancelada"); // M
        abaCompras.getRange(linha, 14).setValue(dataHoje);
        // N
        abaCompras.getRange(linha, 15).setValue(`Op: ${dadosCancelamento.operador} | ${dadosCancelamento.motivo}`);
        // O
        return { sucesso: true, mensagem: "Pedido cancelado!" };
      }
    }
    throw new Error("Compra não encontrada.");
  } catch(e) {
    return { sucesso: false, mensagem: e.toString() };
  }
}

// ==========================================
// 4. GRAVAÇÃO DE DADOS (DIVERSAS ROTAS)
// ==========================================
function registrarMovimentacao(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestampAgora = new Date();
    const dataStr = Utilities.formatDate(timestampAgora, Session.getScriptTimeZone(), "yyyyMMdd");
    // Busca nome do item para histórico
    let nomeItem = "Desconhecido";
    const dadosCadastro = ss.getSheetByName("Cadastro").getDataRange().getValues();
    for (let i = 1; i < dadosCadastro.length; i++) {
      if (dadosCadastro[i][0] === dados.idItem) {
        nomeItem = dadosCadastro[i][1];
        break;
      }
    }

    // ROTA A: NOVA COMPRA
    if (dados.tipo === "Compra") {
      const abaCompras = ss.getSheetByName("Histórico de Compras");
      const abaConfig = ss.getSheetByName("Configurações");
      
      // Cadastro automático de novo fornecedor
      if (abaConfig && dados.fornecedor) {
        const dadosE = abaConfig.getRange("E:E").getValues();
        let ultimaLinhaE = 1;
        let existe = false;
        for(let k = 1; k < dadosE.length; k++){
          if(dadosE[k][0] !== "") {
            ultimaLinhaE = k + 1;
            if(dadosE[k][0].toString().toLowerCase().trim() === dados.fornecedor.toLowerCase().trim()) existe = true;
          }
        }
        if(!existe){
          abaConfig.getRange(ultimaLinhaE + 1, 5).setValue(dados.fornecedor.trim());
          abaConfig.getRange(ultimaLinhaE + 1, 6).setValue("Ativo");
        }
      }

      const idComp = "CP-" + dataStr + "-" + dados.idItem + "-" + timestampAgora.getMilliseconds();
      
      // TRATAMENTO NUMÉRICO PARA EVITAR TEXTO COM PONTO NA PLANILHA
      const qtdCompra = Number(dados.quantidade) || 0;
      const valorTotalCompra = Number(dados.valorTotal) || 0;
      const valorUnit = valorTotalCompra / qtdCompra;
      
      // 18 Colunas (A até R)
      const novaLinhaCompra = [
        idComp, dados.idItem, dados.dataOcorrencia, dados.horaOcorrencia, timestampAgora,
        nomeItem, qtdCompra, valorTotalCompra, valorUnit, dados.fornecedor.trim(), 
        dados.dataPrevista, 0, "Pendente", "", "", dados.operador, "", dados.observacao || ""
      ];
      abaCompras.appendRow(novaLinhaCompra);
      
      enviarEmailCompra(dados, nomeItem, valorUnit);
      return { sucesso: true, mensagem: "Compra registrada!" };
    } 
    
    // ROTA B: RECEBIMENTO DE COMPRA (ENTRADA FÍSICA)
    else if (dados.tipo === "RecebimentoCompra") {
      const abaCompras = ss.getSheetByName("Histórico de Compras");
      const abaHistorico = ss.getSheetByName("Histórico de Operações");
      const dadosCompras = abaCompras.getDataRange().getValues();
      
      let linhaCompra = -1; let recAnterior = 0;
      let totalComp = 0; let idsAnteriores = "";
      for(let i = 1; i < dadosCompras.length; i++) {
        if(dadosCompras[i][0] === dados.idCompra) {
          linhaCompra = i + 1;
          totalComp = Number(dadosCompras[i][6]);
          recAnterior = Number(dadosCompras[i][11]) || 0;
          idsAnteriores = dadosCompras[i][16] || "";
          break;
        }
      }

      const qtdRecebidaAgora = Number(dados.quantidade) || 0;
      const novoRec = recAnterior + qtdRecebidaAgora;
      let statusFim = novoRec >= totalComp ? "Concluída" : dados.forcarEncerramento ? "Encerrada Parcial" : "Pendente";
      const idOP = "OP-" + dataStr + "-" + dados.idItem + "-" + timestampAgora.getMilliseconds();
      const listaIds = idsAnteriores === "" ? idOP : idsAnteriores + ", " + idOP;
      
      // Atualiza aba Compras
      abaCompras.getRange(linhaCompra, 12).setValue(novoRec);    // L
      abaCompras.getRange(linhaCompra, 13).setValue(statusFim);
      // M
      if (statusFim !== "Pendente") abaCompras.getRange(linhaCompra, 14).setValue(new Date());
      // N
      if (dados.motivoEncerramento) abaCompras.getRange(linhaCompra, 15).setValue(dados.motivoEncerramento); // O
      abaCompras.getRange(linhaCompra, 17).setValue(listaIds);
      // Q

      // Grava Entrada no Histórico de Operações (Tratando quantidade como número)
      const obsFinal = `Pedido: ${dados.idCompra}` + (dados.observacao ? ` | ${dados.observacao}` : "");
      abaHistorico.appendRow([idOP, dados.idItem, dados.dataOcorrencia, dados.horaOcorrencia, timestampAgora, "Entrada", dados.tecnico, nomeItem, qtdRecebidaAgora, "", "Chegada de Compra", obsFinal, dados.operador]);
      
      SpreadsheetApp.flush();
      return { sucesso: true, mensagem: "Entrada realizada!" };
    }
    
    // ROTA C: OPERAÇÃO NORMAL (BANCADA)
    else {
      const abaHistorico = ss.getSheetByName("Histórico de Operações");
      const qtdOperacao = Number(dados.quantidade) || 0;

      if (dados.tipo === "Saída") {
        const est = obterListasParaFormulario().itens.find(i => i.id === dados.idItem)?.estoque || 0;
        if (qtdOperacao > est) return { sucesso: false, mensagem: "Estoque insuficiente!" };
      }
      
      const idOP = "OP-" + dataStr + "-" + dados.idItem + "-" + timestampAgora.getMilliseconds();
      abaHistorico.appendRow([idOP, dados.idItem, dados.dataOcorrencia, dados.horaOcorrencia, timestampAgora, dados.tipo, dados.tecnico, nomeItem, qtdOperacao, "", dados.motivo || "N/A", dados.observacao, dados.operador]);
      
      SpreadsheetApp.flush();
      verificarEstoqueBaixo(dados.idItem, nomeItem);
      return { sucesso: true, mensagem: "Salvo!" };
    }
  } catch (e) { return { sucesso: false, mensagem: e.toString() };
  }
}

// ==========================================
// 5. EMAILS E ALERTAS
// ==========================================
function enviarEmailCompra(dados, nomeItem, valorUnit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emails = ss.getSheetByName("RH").getDataRange().getValues()
    .filter(r => r[1] === "Operador" && r[2] && r[3] === "Ativo" && r[0] !== dados.operador).map(r => r[2]);
  if (emails.length > 0) {
    const assunto = `[TechCorp] Novo Pedido: ${nomeItem}`;
    const corpo = `Prezados,\n\nUm novo pedido de compra foi registrado:\n\nItem: ${nomeItem}\nFornecedor: ${dados.fornecedor}\nQtd: ${dados.quantidade}\nTotal: R$ ${parseFloat(dados.valorTotal).toFixed(2)}\nPrevisão: ${dados.dataPrevista}\n\nOperador: ${dados.operador}\nControle TechCorp`;
    MailApp.sendEmail({ to: emails.join(","), subject: assunto, body: corpo });
  }
}

function verificarEstoqueBaixo(id, nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = ss.getSheetByName("Resumo").getDataRange().getValues();
  for (let i = 1; i < res.length; i++) {
    if (res[i][0] === id) {
      let q = Number(res[i][2]);
      let b = Number(res[i][3]); let c = Number(res[i][4]);
      if (q <= c && c !== -1) dispararAlerta(nome, q, c, "🔴 CRÍTICO");
      else if (q <= b && b !== -1 && b !== c) dispararAlerta(nome, q, b, "🟡 BAIXO");
      break;
    }
  }
}

function dispararAlerta(nome, q, lim, tipo) {
  const emails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RH").getDataRange().getValues()
    .filter(r => r[1] === "Operador" && r[2] && r[3] === "Ativo").map(r => r[2]);
  if (emails.length > 0) {
    MailApp.sendEmail({ 
      to: emails.join(","), 
      subject: `${tipo}: Estoque - ${nome}`, 
      body: `Atenção!\n\nItem: ${nome}\nQtd Atual: ${q}\nLimite: ${lim}\n\nTechCorp` 
    });
  }
}

// ==========================================
// 6. DADOS PARA A TELA DE RESUMO
// ==========================================
function obterDadosResumo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = ss.getSheetByName("Resumo").getDataRange().getValues();
  const cad = ss.getSheetByName("Cadastro").getDataRange().getValues();
  const stById = {};
  const idxSt = cad[0].indexOf("Status");
  for (let i = 1; i < cad.length; i++) stById[cad[i][0]] = cad[i][idxSt] || "Ativo";
  
  const final = [];
  for (let i = 1; i < res.length; i++) {
    if (res[i][0] && stById[res[i][0]] === "Ativo") {
      final.push({ nome: res[i][1], qtd: Number(res[i][2]), baixo: Number(res[i][3]), critico: Number(res[i][4]) });
    }
  }
  return final;
}

function notificarTodosCriticos(op) {
  const lista = obterDadosResumo();
  let urg = [];
  let ate = [];
  lista.forEach(i => {
    if (i.qtd <= i.critico && i.critico !== -1) urg.push(`- ${i.nome}: ${i.qtd} un (Mín: ${i.critico})`);
    else if (i.qtd <= i.baixo && i.baixo !== -1 && i.baixo !== i.critico) ate.push(`- ${i.nome}: ${i.qtd} un (Mín: ${i.baixo})`);
  });
  if (urg.length === 0 && ate.length === 0) return { sucesso: false, mensagem: "Nada crítico." };
  const emails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RH").getDataRange().getValues()
    .filter(r => r[1] === "Operador" && r[2] && r[3] === "Ativo").map(r => r[2]);
  let msg = `Prezados,\n\nRelatório de estoque bancada:\n\n`;
  if (urg.length > 0) msg += `🔴 URGENTE:\n${urg.join("\n")}\n\n`;
  if (ate.length > 0) msg += `🟡 ATENÇÃO:\n${ate.join("\n")}\n\n`;
  msg += `Disparado por: ${op}\nTechCorp`;
  MailApp.sendEmail({ to: emails.join(","), subject: `📊 Alerta de Estoque - Por ${op}`, body: msg });
  return { sucesso: true, mensagem: "E-mail enviado!" };
}

// ==========================================
// 7. SISTEMA DE AUTENTICAÇÃO (NOVO)
// ==========================================

/**
 * Autentica o usuário comparando login e hash da senha na aba RH (Col E e F).
 * Agora também verifica nível ADM na Coluna G.
 */
function autenticarUsuario(login, senha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaRH = ss.getSheetByName("RH");
    const dados = abaRH.getDataRange().getValues();
    const hashSenha = gerarHash(senha);
    
    for (let i = 1; i < dados.length; i++) {
      let nome = dados[i][0];
      let funcao = dados[i][1];
      let status = dados[i][3]; // Coluna D
      let userLogin = dados[i][4]; // Coluna E
      let userPassHash = dados[i][5]; // Coluna F
      let adm = dados[i][6]; // Coluna G
      
      if (status === "Ativo" && funcao === "Operador" && userLogin === login && userPassHash === hashSenha) {
        let isAdm = (adm && adm.toString().trim().toUpperCase() === "SIM");
        return { sucesso: true, nome: nome, adm: isAdm };
      }
    }
    return { sucesso: false, mensagem: "Usuário ou senha incorretos." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao autenticar: " + e.toString() };
  }
}

/**
 * Gera hash SHA-256 para armazenamento seguro de senhas.
 */
function gerarHash(texto) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

// ==========================================
// 8. GESTÃO DE RH
// ==========================================

function obterListaRH() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRH = ss.getSheetByName("RH");
  const dados = abaRH.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][3] === "Ativo" && dados[i][0]) {
      lista.push({
        nome: dados[i][0],
        funcao: dados[i][1]
      });
    }
  }
  return lista;
}

function adicionarMembroRH(dadosRH) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaRH = ss.getSheetByName("RH");
    const dadosExistentes = abaRH.getDataRange().getValues();
    
    // Verifica se já existe um membro ativo com o mesmo nome
    for(let i = 1; i < dadosExistentes.length; i++) {
      if(dadosExistentes[i][0].toString().toLowerCase().trim() === dadosRH.nome.toLowerCase().trim() && dadosExistentes[i][3] === "Ativo") {
        return { sucesso: false, mensagem: "Já existe um membro ativo com esse nome!" };
      }
    }

    let novaLinha = [];
    if (dadosRH.tipo === "Técnico") {
      // Nome(A), Função(B), Email(C), Status(D), Login(E), Senha(F), ADM(G)
      novaLinha = [dadosRH.nome.trim(), "Técnico", "", "Ativo", "", "", ""];
    } else if (dadosRH.tipo === "Operador") {
      const hash = gerarHash(dadosRH.senha);
      // Preenche coluna G (ADM) vazia, exigindo elevação manual pela planilha
      novaLinha = [dadosRH.nome.trim(), "Operador", dadosRH.email.trim(), "Ativo", dadosRH.login.trim(), hash, ""];
    }
    
    abaRH.appendRow(novaLinha);
    return { sucesso: true, mensagem: dadosRH.tipo + " cadastrado com sucesso!" };
  } catch(e) {
    return { sucesso: false, mensagem: e.toString() };
  }
}

function desativarMembroRH(nome) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaRH = ss.getSheetByName("RH");
    const dados = abaRH.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === nome && dados[i][3] === "Ativo") {
        abaRH.getRange(i + 1, 4).setValue("Inativo"); // Coluna D
        return { sucesso: true, mensagem: nome + " foi desativado!" };
      }
    }
    return { sucesso: false, mensagem: "Membro ativo não encontrado." };
  } catch(e) {
    return { sucesso: false, mensagem: e.toString() };
  }
}

// ==========================================
// 9. DASHBOARD DE CONSUMO
// ==========================================

function obterDadosDashboard(dataInicio, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Converter strings recebidas do Front-end para objetos Date
  const dtInicio = new Date(dataInicio + "T00:00:00");
  const dtFim = new Date(dataFim + "T23:59:59");
  
  let totalDias = Math.ceil((dtFim - dtInicio) / (1000 * 60 * 60 * 24));
  if (totalDias <= 0) totalDias = 1; // Prevenção para divisão por zero

  // 1. Mapear Itens Ativos da aba Cadastro
  const abaCadastro = ss.getSheetByName("Cadastro");
  const dadosCadastro = abaCadastro.getDataRange().getValues();
  const itensMap = {}; 
  const itensList = [];

  for (let i = 1; i < dadosCadastro.length; i++) {
    let id = dadosCadastro[i][0];
    let nome = dadosCadastro[i][1];
    let status = dadosCadastro[i][dadosCadastro[0].indexOf("Status")] || "Ativo";
    
    if (id && status === "Ativo") {
      itensMap[id] = { 
        id: id, 
        nome: nome, 
        somaValorGeral: 0, 
        qtdGeral: 0, 
        custoMedioUnitario: 0, 
        gastoTotal: 0, 
        usoPeriodo: 0, 
        frequenciaSaida: 0 
      };
      itensList.push(itensMap[id]);
    }
  }

  // 2. Histórico de Compras (Apenas status 'Concluída')
  const abaCompras = ss.getSheetByName("Histórico de Compras");
  if (abaCompras) {
    const dadosCompras = abaCompras.getDataRange().getValues();
    for (let i = 1; i < dadosCompras.length; i++) {
      let statusCompra = dadosCompras[i][12]; // Coluna M
      if (statusCompra !== "Concluída") continue; // TRAVA: Ignora se for Pendente ou Cancelada

      let idItem = dadosCompras[i][1]; // Coluna B
      if (!itensMap[idItem]) continue;
      
      let dataStr = dadosCompras[i][2]; // Coluna C (Data Ocorrência)
      let dataObj = (dataStr instanceof Date) ? dataStr : new Date(dataStr + "T12:00:00");
      
      let qtd = Number(dadosCompras[i][6]) || 0; // Coluna G
      let valor = Number(dadosCompras[i][7]) || 0; // Coluna H
      
      // Acumular para calcular custo médio geral do item
      itensMap[idItem].somaValorGeral += valor;
      itensMap[idItem].qtdGeral += qtd;
      
      // Somar gasto apenas se estiver dentro do período filtrado
      if (dataObj >= dtInicio && dataObj <= dtFim) {
        itensMap[idItem].gastoTotal += valor;
      }
    }
  }

  // Finalizar cálculo do Custo Médio Unitário (Baseado em toda a história para precisão)
  for (let id in itensMap) {
    if (itensMap[id].qtdGeral > 0) {
      itensMap[id].custoMedioUnitario = itensMap[id].somaValorGeral / itensMap[id].qtdGeral;
    }
  }

  // 3. Histórico de Operações (Para Uso, Frequência e Timeline de Barras)
  const abaHistorico = ss.getSheetByName("Histórico de Operações");
  const operacoesPeriodo = [];
  
  if (abaHistorico) {
    const dadosHist = abaHistorico.getDataRange().getValues();
    for (let i = 1; i < dadosHist.length; i++) {
      let idItem = dadosHist[i][1]; // Coluna B
      if (!itensMap[idItem]) continue;
      
      let dataStr = dadosHist[i][2]; // Coluna C (Data)
      let dataObj = (dataStr instanceof Date) ? dataStr : new Date(dataStr + "T12:00:00");
      
      // Processar apenas movimentos dentro do período escolhido
      if (dataObj >= dtInicio && dataObj <= dtFim) {
        let tipo = dadosHist[i][5]; // Coluna F (Entrada ou Saída)
        let qtd = Number(dadosHist[i][8]) || 0; // Coluna I (Quantidade)
        
        if (tipo === "Saída") {
          itensMap[idItem].usoPeriodo += qtd; // Soma quantidade total consumida
          itensMap[idItem].frequenciaSaida += 1; // Conta +1 evento/ida ao estoque
        }
        
        operacoesPeriodo.push({
          data: dataObj,
          idItem: idItem,
          tipo: tipo,
          qtd: qtd
        });
      }
    }
  }

  // 4. Montar a tabela final com os cálculos diários
  const itensFinal = [];
  for (let i = 0; i < itensList.length; i++) {
    let item = itensList[i];
    item.usoMedio = item.usoPeriodo / totalDias; // Quantidade média / dia
    item.custoMedioDia = item.usoMedio * item.custoMedioUnitario; // R$ médio / dia
    itensFinal.push(item);
  }

  // 5. Agrupamento Dinâmico para o Gráfico de Barras (Timeline)
  let tipoAgrupamento = "dia";
  if (totalDias > 90) tipoAgrupamento = "mes";
  else if (totalDias > 31) tipoAgrupamento = "semana";

  const labels = [];
  const dadosPorItem = {}; // Estrutura: idItem -> { entradas: [], saidas: [] }
  const buckets = [];
  let dataAtual = new Date(dtInicio);
  
  // Criar os intervalos de tempo (Buckets) com base na inteligência definida
  if (tipoAgrupamento === "dia") {
    while (dataAtual <= dtFim) {
      let l = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM");
      labels.push(l);
      buckets.push({ 
        inicio: new Date(dataAtual.getFullYear(), dataAtual.getMonth(), dataAtual.getDate(), 0, 0, 0), 
        fim: new Date(dataAtual.getFullYear(), dataAtual.getMonth(), dataAtual.getDate(), 23, 59, 59) 
      });
      dataAtual.setDate(dataAtual.getDate() + 1);
    }
  } else if (tipoAgrupamento === "semana") {
    let numSemana = 1;
    while (dataAtual <= dtFim) {
      let inicioSemana = new Date(dataAtual.getFullYear(), dataAtual.getMonth(), dataAtual.getDate(), 0, 0, 0);
      let fimSemana = new Date(dataAtual.getFullYear(), dataAtual.getMonth(), dataAtual.getDate(), 23, 59, 59);
      fimSemana.setDate(fimSemana.getDate() + 6);
      if (fimSemana > dtFim) fimSemana = new Date(dtFim);
      
      labels.push("Sem " + numSemana);
      buckets.push({ inicio: inicioSemana, fim: fimSemana });
      
      dataAtual.setDate(dataAtual.getDate() + 7);
      numSemana++;
    }
  } else {
    while (dataAtual <= dtFim) {
      let m = dataAtual.getMonth();
      let y = dataAtual.getFullYear();
      let nomeMes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"][m];
      labels.push(nomeMes + "/" + y.toString().substring(2));
      
      let inicioMes = new Date(y, m, 1, 0, 0, 0);
      let fimMes = new Date(y, m + 1, 0, 23, 59, 59);
      if (fimMes > dtFim) fimMes = new Date(dtFim);
      
      buckets.push({ inicio: inicioMes, fim: fimMes });
      dataAtual.setMonth(dataAtual.getMonth() + 1);
    }
  }

  // Preencher matrizes de dados do gráfico com Zeros (0)
  for (let i = 0; i < itensFinal.length; i++) {
    dadosPorItem[itensFinal[i].id] = {
      entradas: new Array(labels.length).fill(0),
      saidas: new Array(labels.length).fill(0)
    };
  }

  // Distribuir as operações dentro dos intervalos (Buckets) gerados
  for (let i = 0; i < operacoesPeriodo.length; i++) {
    let op = operacoesPeriodo[i];
    let d = op.data;
    
    for (let b = 0; b < buckets.length; b++) {
      if (d >= buckets[b].inicio && d <= buckets[b].fim) {
        if (op.tipo === "Entrada") {
          dadosPorItem[op.idItem].entradas[b] += op.qtd;
        } else if (op.tipo === "Saída") {
          dadosPorItem[op.idItem].saidas[b] += op.qtd;
        }
        break; // Achou o bucket, vai para a próxima operação
      }
    }
  }

  // Devolver pacote empacotado para o HTML
  return {
    itens: itensFinal,
    timeline: {
      labels: labels,
      dadosPorItem: dadosPorItem
    }
  };
}