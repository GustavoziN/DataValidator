let dadosConvertidos = null;

function ConverterParaJson(arquivo) {
  if (!arquivo) return;

  const leitor = new FileReader();
  leitor.onload = function (evento) {
    const dados = new Uint8Array(evento.target.result);
    const workbook = XLSX.read(dados, { type: 'array' });
    const nomeAba = workbook.SheetNames[0];
    const planilha = workbook.Sheets[nomeAba];
    dadosConvertidos = XLSX.utils.sheet_to_json(planilha);
    console.log('Arquivo convertido, clique no botão para ver.');
  };
  leitor.readAsArrayBuffer(arquivo);
}

function MostrarNoConsole() {
  if (!dadosConvertidos) {
    console.log('Nenhum arquivo foi convertido ainda.');
    return;
  }

  // remove a linha fantasma (o cabeçalho que "vazou" para dentro dos dados)
  const dados = dadosConvertidos.filter(function (linha) {
    return linha['__EMPTY'] !== 'TP Movimento';
  });

  // agrupa por produto, somando quantidade produzida (010) e movimentada (999)
  const porProduto = {};

  dados.forEach(function (linha) {
    const produto = linha['__EMPTY_1'];
    const tipo = linha['__EMPTY'];
    const quantidade = Number(linha['__EMPTY_3']) || 0;

    if (!porProduto[produto]) {
      porProduto[produto] = { produzido: 0, movimentado: 0 };
    }

    if (tipo === '010') {
      porProduto[produto].produzido += quantidade;
    } else if (tipo === '999') {
      porProduto[produto].movimentado += quantidade;
    }
  });

  // monta o relatório final
  const relatorio = Object.keys(porProduto).map(function (produto) {
    const item = porProduto[produto];
    const bate = item.produzido === item.movimentado;
    return {
      Produto: produto,
      Produzido: item.produzido,
      Movimentado: item.movimentado,
      Diferenca: item.produzido - item.movimentado,
      Bate: bate ? 'SIM' : 'NAO'
    };
  });

  console.log(JSON.stringify(relatorio, null, 2));
}

