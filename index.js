
let dadosConvertidos = null;
let usandoExemplo = true;

const inputArquivo = document.getElementById('inputArquivo');
const btnMostrar = document.getElementById('btnMostrar');
const status = document.getElementById('status');
const corpoTabela = document.getElementById('corpoTabela');
const resumoTotal = document.getElementById('resumoTotal');
const resumoBatem = document.getElementById('resumoBatem');
const resumoNaoBatem = document.getElementById('resumoNaoBatem');

inputArquivo.addEventListener('change', function (evento) {
  const arquivo = evento.target.files[0];
  ConverterParaJson(arquivo);
});

btnMostrar.addEventListener('click', MostrarNaTela);

function ConverterParaJson(arquivo) {
  if (!arquivo) return;

  status.textContent = 'Lendo arquivo...';
  btnMostrar.disabled = true;

  const leitor = new FileReader();
  leitor.onload = function (evento) {
    try {
      const dados = new Uint8Array(evento.target.result);
      const workbook = XLSX.read(dados, { type: 'array' });
      const nomeAba = workbook.SheetNames[0];
      const planilha = workbook.Sheets[nomeAba];
      dadosConvertidos = XLSX.utils.sheet_to_json(planilha);
      usandoExemplo = false;

      status.textContent = 'Arquivo "' + arquivo.name + '" convertido. Clique em "Mostrar relatório".';
      btnMostrar.disabled = false;
    } catch (erro) {
      status.textContent = 'Erro ao ler o arquivo: ' + erro.message;
    }
  };
  leitor.onerror = function () {
    status.textContent = 'Não foi possível ler esse arquivo.';
  };
  leitor.readAsArrayBuffer(arquivo);
}

function MostrarNaTela() {
  if (!dadosConvertidos) {
    status.textContent = 'Nenhum arquivo foi convertido ainda.';
    return;
  }

  // remove a linha fantasma (o cabeçalho que "vazou" para dentro dos dados)
  const dados = dadosConvertidos.filter(function (linha) {
    return linha['__EMPTY'] !== 'TP Movimento';
  });

  // agrupa por produto + período, somando quantidade produzida (010) e movimentada (999)
  const porProdutoPeriodo = {};

  dados.forEach(function (linha) {
    const produto = linha['__EMPTY_1'];
    const descricao = linha['__EMPTY_2'];
    const tipo = linha['__EMPTY'];
    const quantidade = Number(linha['__EMPTY_3']) || 0;
    const periodo = linha['__EMPTY_8'];

    const chave = produto + '|' + periodo;

    if (!porProdutoPeriodo[chave]) {
      porProdutoPeriodo[chave] = {
        produto: produto,
        descricao: descricao,
        periodo: periodo,
        produzido: 0,
        movimentado: 0
      };
    }

    if (tipo === '010') {
      porProdutoPeriodo[chave].produzido += quantidade;
    } else if (tipo === '999') {
      porProdutoPeriodo[chave].movimentado += quantidade;
    }
  });

  const relatorio = Object.keys(porProdutoPeriodo).map(function (chave) {
    const item = porProdutoPeriodo[chave];
    return {
      Produto: item.produto,
      Descricao: item.descricao,
      Periodo: item.periodo,
      Produzido: item.produzido,
      Movimentado: item.movimentado,
      Diferenca: item.produzido - item.movimentado,
      Bate: item.produzido === item.movimentado
    };
  });

  renderizarTabela(relatorio);

  if (!usandoExemplo) {
    status.textContent = relatorio.length + ' grupo(s) encontrados no arquivo enviado.';
  }
}

function renderizarTabela(relatorio) {
  corpoTabela.innerHTML = '';

  let batem = 0;
  relatorio.forEach(function (linha) {
    if (linha.Bate) batem++;

    const tr = document.createElement('tr');
    const corDot = linha.Bate ? 'var(--ok-dot)' : 'var(--bad-dot)';
    const corBg = linha.Bate ? 'var(--ok-bg)' : 'var(--bad-bg)';
    const corTexto = linha.Bate ? 'var(--ok-text)' : 'var(--bad-text)';

    tr.innerHTML =
      '<td class="px-3 py-2.5 font-mono-num" style="color:var(--text); border-bottom:1px solid var(--border);">' + escapeHtml(linha.Produto) + '</td>' +
      '<td class="px-3 py-2.5" style="color:var(--text); border-bottom:1px solid var(--border);">' + escapeHtml(linha.Descricao) + '</td>' +
      '<td class="px-3 py-2.5" style="color:var(--text-muted); border-bottom:1px solid var(--border);">' + escapeHtml(linha.Periodo) + '</td>' +
      '<td class="px-3 py-2.5 text-right font-mono-num" style="color:var(--text); border-bottom:1px solid var(--border);">' + linha.Produzido + '</td>' +
      '<td class="px-3 py-2.5 text-right font-mono-num" style="color:var(--text); border-bottom:1px solid var(--border);">' + linha.Movimentado + '</td>' +
      '<td class="px-3 py-2.5 text-right font-mono-num" style="color:var(--text); border-bottom:1px solid var(--border);">' + (linha.Diferenca > 0 ? '+' : '') + linha.Diferenca + '</td>' +
      '<td class="px-3 py-2.5 text-center" style="border-bottom:1px solid var(--border);">' +
        '<span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold" style="background:' + corBg + '; color:' + corTexto + ';">' +
          '<span class="status-dot" style="background:' + corDot + ';"></span>' + (linha.Bate ? 'SIM' : 'NÃO') +
        '</span>' +
      '</td>';

    corpoTabela.appendChild(tr);
  });

  resumoTotal.textContent = String(relatorio.length);
  resumoBatem.textContent = String(batem);
  resumoNaoBatem.textContent = String(relatorio.length - batem);
}

function escapeHtml(valor) {
  if (valor === undefined || valor === null) return '';
  return String(valor)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

MostrarNaTela();