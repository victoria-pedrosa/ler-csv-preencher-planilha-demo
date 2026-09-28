// --- CONFIGURAÇÕES ---
const CONFIG = {
  ROOT_FOLDER_ID: 'ID_EXEMPLO', 
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/ID_EXEMPLO/edit?gid=0#gid=0',
  SHEET_NAME: 'Página1', 
  
  COL_CODIGO: 1,       // A
  COL_NOME_EMPRESA: 2, // B
  COL_NOME_LOJA: 3,    // C
  COL_CNPJ_ID: 4,      // D
  COL_SAIDA: 10,       // J
  COL_ENTRADA: 11,     // K
  COL_DAE: 12          // L
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🟢 MENU FISCAL')
    .addItem('▶️ Rodar Atualização', 'iniciarProcessamento')
    .addToUi();
}

function iniciarProcessamento() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  
  const range = sheet.getActiveRange();
  let startRow = range.getRow();
  let numRows = range.getNumRows();
  let endRow = startRow + numRows - 1;

  if (startRow < 2) {
    if (endRow < 2) {
      ui.alert("⚠️ Por favor, selecione uma linha abaixo do cabeçalho.");
      return;
    }
    startRow = 2; 
    numRows = endRow - startRow + 1;
  }

  const resposta = ui.alert(
    'Modo de Execução',
    `Calcular APENAS as linhas selecionadas (${startRow} a ${endRow}) na aba "${sheet.getName()}"?\n\n(Clique "Sim" para focar, ou "Não" para a planilha toda)`,
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (resposta == ui.Button.CANCEL) return;
  const apenasSelecao = (resposta == ui.Button.YES);
  
  processarArquivosFiscais(sheet, apenasSelecao, startRow, numRows, ui);
}

function processarArquivosFiscais(sheet, apenasSelecao, startRow, numRows, ui) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let mapaEmpresas = mapearPlanilha(sheet, apenasSelecao ? startRow : 2, apenasSelecao ? numRows : sheet.getLastRow() - 1);
  if (mapaEmpresas.length === 0) return ss.toast("Nenhuma empresa encontrada.", "Aviso");

  ss.toast(`Acessando a pasta do Drive...`, "Aguarde");
  
  let pastaAlvo;
  try {
    pastaAlvo = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  } catch (e) {
    return ui.alert("Erro ao acessar a pasta raiz. Verifique o ID.");
  }

  const arquivosIterator = pastaAlvo.getFiles();
  const listaArquivos = [];
  while (arquivosIterator.hasNext()) {
    listaArquivos.push(arquivosIterator.next());
  }

  mapaEmpresas.forEach(emp => {
    atualizarStatus(sheet, emp.row, "🔍 Localizando CSV...");

    let arquivosDaEmpresa = listaArquivos.filter(file => {
      const nomeArquivo = file.getName().toUpperCase();
      
      // NOVA LÓGICA DE BUSCA: 
      // Verifica se o código existe no nome e se não é parte de outro número maior
      // Ex: Se o código é 1107, aceita "rpt_1107_matriz", mas recusa "21107"
      const regexCodigo = new RegExp(`(?:\\D|^)${emp.codigo}(?:\\D|$)`);
      if (!regexCodigo.test(nomeArquivo)) return false;

      const lojaPlanilha = emp.lojaNome.toUpperCase();
      
      // Tratamento de Filiais Específicas (2, 3...)
      if (lojaPlanilha.includes('FILIAL 2') || lojaPlanilha.includes('FILIAL 02')) {
        return nomeArquivo.includes('FILIAL 2') || nomeArquivo.includes('FILIAL 02') || nomeArquivo.includes('FILIAL2');
      }
      if (lojaPlanilha.includes('FILIAL 3') || lojaPlanilha.includes('FILIAL 03')) {
        return nomeArquivo.includes('FILIAL 3') || nomeArquivo.includes('FILIAL 03') || nomeArquivo.includes('FILIAL3');
      }

      // Caso Geral (Matriz ou Empresa Única)
      if (lojaPlanilha.includes('FILIAL')) {
        // Se a planilha diz "FILIAL", pega arquivos que tenham "FILIAL" mas ignora as específicas tratadas acima
        return nomeArquivo.includes('FILIAL') && !nomeArquivo.includes('FILIAL 2') && !nomeArquivo.includes('FILIAL 3');
      } else {
        // Se a planilha não menciona filial, aceita se tiver "MATRIZ" ou se NÃO tiver "FILIAL" no nome
        return nomeArquivo.includes('MATRIZ') || !nomeArquivo.includes('FILIAL');
      }
    });

    if (arquivosDaEmpresa.length > 0) {
      let leuComSucesso = false;
      emp.totalEntrada = 0; emp.totalSaida = 0;

      for (let file of arquivosDaEmpresa) {
        try {
          const texto = file.getBlob().getDataAsString('ISO-8859-1');
          if (!texto || texto.trim() === '') continue;
          
          const primeiraLinha = texto.split(/\r\n|\n|\r/)[0].toUpperCase();
          
          // Se o arquivo é de EMISSÃO da empresa (Saída/Faturamento)
          // No CSV Domínio, se a 1ª linha tem "DESTINATARIO", a empresa é a emitente.
          if (primeiraLinha.includes('DESTINATARIO') || primeiraLinha.includes('DESTINATÁRIO') || file.getName().toUpperCase().includes('EMIT')) {
              lerCsvEmitente(texto, emp);
          } else {
              // Se não, é arquivo de compras (Entrada)
              lerCsvDestinatario(texto, emp);
          }
          leuComSucesso = true;
        } catch (err) {
          emp.erroLeitura = "⚠️ Erro leitura";
        }
      }

      if (leuComSucesso) {
        escreverNaPlanilha(sheet, emp);
      } else {
        finalizarComErro(sheet, emp, "⚠️ Falha no CSV");
      }

    } else {
      finalizarComErro(sheet, emp, "⚠️ Arq. não encontrado");
    }
  });
  
  ss.toast('Processamento Concluído!', '✅ Fim', 5);
}

// --- CÁLCULO E ANULAÇÃO ---

function lerCsvEmitente(texto, empresa) {
  const csvData = parseCSV(texto);
  csvData.forEach(row => {
    const valor = lerValor(row);
    if (valor === 0) return;
    if (verificarCancelada(row)) return;

    // Se o tipo for Entrada (ex: devolução de venda), soma na entrada. Senão, saída.
    if (verificarOperacaoEntrada(row)) empresa.totalEntrada += valor;
    else empresa.totalSaida += valor;
  });
}

function lerCsvDestinatario(texto, empresa) {
  const csvData = parseCSV(texto);
  let notasValidas = [];

  csvData.forEach(row => {
    const valor = lerValor(row);
    if (valor === 0) return;
    if (verificarCancelada(row)) return;

    notasValidas.push({
      valor: valor,
      cnpj: limparNumeros(row['CNPJ EMITENTE'] || row['CNPJ/CPF EMITENTE'] || row['CNPJ'] || ''),
      tipo: verificarOperacaoEntrada(row) ? 'ENTRADA' : 'SAIDA',
      uf: (row['UF EMIT.'] || row['UF'] || '').toUpperCase(),
      anulada: false
    });
  });

  // Anulação de Devoluções/Estornos (Valores iguais, CNPJ igual, tipos opostos)
  for (let i = 0; i < notasValidas.length; i++) {
    let notaA = notasValidas[i];
    if (notaA.anulada) continue;
    for (let j = i + 1; j < notasValidas.length; j++) {
      let notaB = notasValidas[j];
      if (notaB.anulada) continue;
      if (notaA.valor === notaB.valor && notaA.cnpj === notaB.cnpj && notaA.tipo !== notaB.tipo) {
        notaA.anulada = true; notaB.anulada = true;
        break; 
      }
    }
  }

  notasValidas.forEach(nota => {
    if (!nota.anulada) {
      empresa.totalEntrada += nota.valor;
      if (nota.tipo === 'SAIDA' && nota.uf !== 'BA' && nota.uf !== '') {
        empresa.gerarDAE = true;
      }
    }
  });
}

// --- STATUS E GRAVAÇÃO ---

function atualizarStatus(sheet, linha, mensagem) {
  sheet.getRange(linha, CONFIG.COL_SAIDA).setValue(mensagem).setFontColor("#0000FF");
  SpreadsheetApp.flush(); 
}

function finalizarComErro(sheet, emp, mensagemErro) {
  sheet.getRange(emp.row, CONFIG.COL_SAIDA).setValue(0);
  sheet.getRange(emp.row, CONFIG.COL_ENTRADA).setValue(0);
  sheet.getRange(emp.row, CONFIG.COL_DAE).setValue(mensagemErro).setFontColor("red");
}

function escreverNaPlanilha(sheet, dados) {
  sheet.getRange(dados.row, CONFIG.COL_SAIDA).setValue(dados.totalSaida).setFontColor("black").setNumberFormat("#,##0.00");
  sheet.getRange(dados.row, CONFIG.COL_ENTRADA).setValue(dados.totalEntrada).setFontColor("black").setNumberFormat("#,##0.00");
  
  const textoDAE = dados.gerarDAE ? "GERAR DAE" : " - ";
  const corDAE = dados.gerarDAE ? "blue" : "black";
  sheet.getRange(dados.row, CONFIG.COL_DAE).setValue(textoDAE).setFontColor(corDAE);
}

function mapearPlanilha(sheet, startRow, numRows) {
  const dados = [];
  const rangeDados = sheet.getRange(startRow, 1, numRows, 4).getValues();

  for (let i = 0; i < numRows; i++) {
    const codigoRaw = String(rangeDados[i][CONFIG.COL_CODIGO - 1]).trim();
    const nomeRaw = String(rangeDados[i][CONFIG.COL_NOME_EMPRESA - 1]).toUpperCase().trim();
    const lojaRaw = String(rangeDados[i][CONFIG.COL_NOME_LOJA - 1]).toUpperCase().trim();
    
    if (!codigoRaw || codigoRaw === "" || codigoRaw === "undefined") continue;

    dados.push({
      row: startRow + i, 
      codigo: codigoRaw,
      nome: nomeRaw,
      lojaNome: lojaRaw, 
      totalSaida: 0,
      totalEntrada: 0,
      gerarDAE: false,
      erroLeitura: null
    });
  }
  return dados;
}

// --- AUXILIARES ---

function parseCSV(text) {
  const lines = text.split(/\r\n|\n|\r/);
  if (lines.length < 2) return [];
  const separador = lines[0].includes(';') ? ';' : ',';
  const headers = lines[0].split(separador).map(h => h.trim().toUpperCase().replace(/["']/g, ''));
  const res = [];
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const cells = line.split(separador);
    const obj = {};
    headers.forEach((h, idx) => obj[h] = cells[idx] ? cells[idx].replace(/["']/g, '').trim() : '');
    res.push(obj);
  }
  return res;
}

function lerValor(row) {
  // Procura por colunas que contenham "VALOR" no nome (ex: "Valor", "Valor (R$)", "Valor Total")
  for (const k in row) {
    if (k.includes('VALOR')) return parseValor(row[k]);
  }
  return 0;
}

function parseValor(str) {
  if (!str) return 0;
  // Converte formato brasileiro (1.200,50) para decimal (1200.50)
  return parseFloat(String(str).replace(/\./g, '').replace(',', '.')) || 0;
}

function limparNumeros(str) { return str ? str.replace(/\D/g, '') : ''; }

function verificarCancelada(row) { 
  const sit = (row['SITUACAO'] || row['SITUAÇÃO'] || row['SITUACAO ATUAL'] || '').toUpperCase();
  return sit.includes('CANCELADA'); 
}

function verificarOperacaoEntrada(row) { 
  const t = (row['TIPO OPERACAO'] || row['TIPO OPERAÇÃO'] || '').toUpperCase();
  // No CSV Domínio, '0' costuma ser entrada e '1' saída, ou vem o texto por extenso
  return t.includes('ENTRADA') || t === '0'; 
}