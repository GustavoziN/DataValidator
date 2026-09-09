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
  console.log(JSON.stringify(dadosConvertidos, null, 2));
}


