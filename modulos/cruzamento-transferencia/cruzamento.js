(function () {
  var cruzamentoState = {
    status: 'inicial',
    modulo: 'cruzamento-transferencia',
    currentUser: null,
    canAccess: false,
    resultadoPronto: false,
    excesso: {
      fileName: '',
      filial: '',
      headers: [],
      rows: [],
      items: [],
      valid: false,
      message: 'Aguardando arquivo de excesso.'
    },
    ruptura: {
      fileName: '',
      filial: '',
      headers: [],
      rows: [],
      items: [],
      valid: false,
      message: 'Aguardando arquivo de ruptura.'
    },
    filtros: {
      classeGeral: '',
      classeExcesso: '',
      classeRuptura: '',
      valorRupturaMin: 0,
      qtdExcessoMin: 0,
      coberturaMin: '',
      coberturaMax: '',
      busca: '',
      somenteSugestao: false
    },
    resultados: [],
    resultadosFiltrados: [],
    indices: {
      excesso: {},
      ruptura: {}
    },
    rupturaSemOrigem: 0,
    sort: {
      key: 'prioridade',
      dir: 'asc'
    }
  };

  var cruzamentoHeaderAliasGroups = {
    codigoDescricao: ['CODIGO DESCRICAO', 'CODIGO PRODUTO', 'CODIGO_PRODUTO', 'CODIGO'],
    codigoProduto: ['CODIGO_PRODUTO', 'CODIGO PRODUTO', 'CODIGO'],
    descricao: ['DESCRICAO'],
    estoque: ['ESTOQUE'],
    mediaDia: ['MEDIA DIA', 'MEDIA_DIA'],
    cobertura: ['COBERTURA'],
    excessoQtd: ['EXCESSO'],
    percExcesso: ['% EXCESSO', 'PERCENTUAL EXCESSO', 'PERC EXCESSO'],
    valorExcesso: ['R$ EXCESSO', 'RS EXCESSO', 'VALOR EXCESSO'],
    valorEstoque: ['R$ ESTOQUE', 'RS ESTOQUE', 'VALOR ESTOQUE'],
    valorRuptura: ['R$ RUPTURA', 'RS RUPTURA', 'VALOR RUPTURA'],
    dez: ['DEZ'],
    custo: ['CUSTO'],
    classeGeral: ['CLASSE GERAL'],
    classeFilial: ['CLASSE FILIAL']
  };

  var cruzamentoExcessoRequiredHeaders = [
    { label: 'CODIGO DESCRICAO', key: 'codigoDescricao' },
    { label: 'ESTOQUE', key: 'estoque' },
    { label: 'MEDIA DIA', key: 'mediaDia' },
    { label: 'COBERTURA', key: 'cobertura' },
    { label: 'EXCESSO', key: 'excessoQtd' },
    { label: 'R$ EXCESSO', key: 'valorExcesso' },
    { label: 'R$ ESTOQUE', key: 'valorEstoque' },
    { label: 'CLASSE GERAL', key: 'classeGeral' },
    { label: 'CLASSE FILIAL', key: 'classeFilial' }
  ];

  var cruzamentoRupturaRequiredHeaders = [
    { label: 'CODIGO_PRODUTO', key: 'codigoProduto' },
    { label: 'DESCRICAO', key: 'descricao' },
    { label: 'R$ RUPTURA', key: 'valorRuptura' },
    { label: 'DEZ', key: 'dez' },
    { label: 'MEDIA DIA', key: 'mediaDia' },
    { label: 'CUSTO', key: 'custo' },
    { label: 'CLASSE GERAL', key: 'classeGeral' },
    { label: 'CLASSE FILIAL', key: 'classeFilial' }
  ];

  function cruzamentoGetCurrentUser() {
    var operatorName;

    try {
      if (window.parent && window.parent !== window && window.parent.CURRENT_USER) {
        return window.parent.CURRENT_USER;
      }
    } catch (err) {
      return null;
    }

    try {
      if (window.parent && window.parent !== window && window.parent.document) {
        operatorName = cruzamentoGetParentOperatorName(window.parent.document);
        if (operatorName) {
          return {
            name: operatorName,
            role: operatorName === 'GERENTE' ? 'viewer' : 'admin'
          };
        }
      }
    } catch (err2) {
      return null;
    }

    return null;
  }

  function cruzamentoGetParentOperatorName(parentDocument) {
    var userEl = parentDocument.getElementById('seg-user');
    var topEl = parentDocument.getElementById('tb-op');
    var segOp = parentDocument.getElementById('seg-op');
    var value = '';

    if (userEl) value = userEl.textContent || '';
    if (!value && topEl) value = topEl.textContent || '';
    if (!value && segOp) value = (segOp.textContent || '').replace(/^Operador\s*:\s*/i, '');

    value = String(value || '').trim().toUpperCase();
    if (!value || value === 'JHONNY' || value === '-') return '';
    return value;
  }

  function cruzamentoCanAccess(user) {
    if (!user) return false;
    return user.role === 'admin' || user.name === 'TRANSFERENCIA';
  }

  function cruzamentoSetVisible(el, visible) {
    if (!el) return;
    el.classList.toggle('is-hidden', !visible);
  }

  function cruzamentoGetXLSX() {
    try {
      if (window.parent && window.parent.XLSX) return window.parent.XLSX;
    } catch (err) {
      return null;
    }
    return window.XLSX || null;
  }

  function cruzamentoPickSheetName(workbook) {
    var sheets = workbook && workbook.SheetNames ? workbook.SheetNames : [];
    var exportSheet = sheets.find(function (name) {
      return cruzamentoNormalizeHeader(name) === 'EXPORT';
    });

    return exportSheet || sheets[0] || '';
  }

  function cruzamentoWorkbookRows(XLSX, workbook) {
    var sheetName = cruzamentoPickSheetName(workbook);
    var rows;

    if (!sheetName) return null;

    rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    return rows.filter(function (row) {
      return row.some(function (cell) { return String(cell == null ? '' : cell).trim() !== ''; });
    });
  }

  function cruzamentoNormalizeHeader(value) {
    return String(value || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/[_\r\n\t]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  function cruzamentoAliases(keyOrName) {
    return cruzamentoHeaderAliasGroups[keyOrName] || [keyOrName];
  }

  function cruzamentoHeaderIndex(map, keyOrName) {
    var aliases = cruzamentoAliases(keyOrName);
    var idx;

    for (var i = 0; i < aliases.length; i += 1) {
      idx = map[cruzamentoNormalizeHeader(aliases[i])];
      if (idx !== undefined) return idx;
    }

    return undefined;
  }

  function cruzamentoFindHeaderRow(rows) {
    for (var i = 0; i < rows.length; i += 1) {
      var map = cruzamentoHeaderMap(rows[i]);
      if (
        cruzamentoHeaderIndex(map, 'codigoDescricao') !== undefined ||
        cruzamentoHeaderIndex(map, 'codigoProduto') !== undefined ||
        cruzamentoHeaderIndex(map, 'valorRuptura') !== undefined ||
        cruzamentoHeaderIndex(map, 'excessoQtd') !== undefined
      ) {
        return i;
      }
    }
    return 0;
  }

  function cruzamentoValidateHeaders(headers, required) {
    var map = cruzamentoHeaderMap(headers);
    var missing = required.filter(function (item) {
      return cruzamentoHeaderIndex(map, item.key) === undefined;
    }).map(function (item) {
      return item.label;
    });
    return {
      valid: missing.length === 0,
      missing: missing
    };
  }

  function cruzamentoHeaderMap(headers) {
    var map = {};
    headers.forEach(function (name, idx) {
      map[cruzamentoNormalizeHeader(name)] = idx;
    });
    return map;
  }

  function cruzamentoCell(row, map, name) {
    var idx = cruzamentoHeaderIndex(map, name);
    return idx === undefined ? '' : row[idx];
  }

  function cruzamentoCode(value) {
    return String(value == null ? '' : value)
      .replace(/[\u200B-\u200D\uFEFF]/g, '')
      .replace(/[\s\r\n\t]+/g, '')
      .replace(/\D/g, '');
  }

  function cruzamentoHasNumber(value) {
    return value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
  }

  function cruzamentoRoundQty(value) {
    if (!cruzamentoHasNumber(value)) return null;
    return Math.round(Number(value) * 100) / 100;
  }

  function cruzamentoClasseRank(value) {
    var cls = String(value || '').trim().toUpperCase();
    if (cls === 'A') return 1;
    if (cls === 'B') return 2;
    if (cls === 'C') return 3;
    return 9;
  }

  function cruzamentoIndexByCodigo(items) {
    var index = {};
    (items || []).forEach(function (item) {
      var codigo = cruzamentoCode(item && item.codigo);
      if (!codigo) return;
      if (!index[codigo]) index[codigo] = item;
    });
    return index;
  }

  function cruzamentoSplitCodigoDescricao(value) {
    var text = String(value || '').trim();
    var match = text.match(/^(\d+)\s*[-\u2013\u2014]?\s*(.*)$/);
    return {
      codigo: match ? match[1] : cruzamentoCode(text),
      descricao: match ? match[2].trim() : text
    };
  }

  function cruzamentoNumber(value, opts) {
    var options = opts || {};
    var text = String(value == null ? '' : value).trim();
    var negative = /^\(.*\)$/.test(text) || /^-/.test(text);
    var isPercent = text.indexOf('%') >= 0 || options.percent;
    var normalized;

    if (typeof value === 'number') {
      return options.percent && value > 1 ? value / 100 : value;
    }
    if (!text) return null;

    normalized = text.replace(/[R$\s%]/g, '').replace(/[()]/g, '').replace(/^[+-]/, '');
    if (normalized.indexOf(',') >= 0) {
      normalized = normalized.replace(/\./g, '').replace(',', '.');
    }

    normalized = Number(normalized);
    if (!Number.isFinite(normalized)) return null;
    if (isPercent) normalized = normalized / 100;
    return negative ? -normalized : normalized;
  }

  function cruzamentoFmtNumber(value) {
    if (value === null || value === undefined || value === '') return '';
    return (Number(value) || 0).toLocaleString('pt-BR', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    });
  }

  function cruzamentoFmtMoney(value) {
    if (value === null || value === undefined || value === '') return '';
    return (Number(value) || 0).toLocaleString('pt-BR', {
      style: 'currency',
      currency: 'BRL'
    });
  }

  function cruzamentoEscape(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, function (ch) {
      return {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      }[ch];
    });
  }

  function cruzamentoSetStatusMessage(message, state) {
    var status = document.getElementById('cruzamento-status');
    if (status) {
      status.textContent = message;
      status.setAttribute('data-state', state || 'info');
    }
  }

  function cruzamentoNow() {
    return new Date();
  }

  function cruzamentoFormatDateTime(date) {
    return date.toLocaleString('pt-BR');
  }

  function cruzamentoFileDate(date) {
    function pad(value) {
      return String(value).padStart(2, '0');
    }

    return String(date.getFullYear())
      + pad(date.getMonth() + 1)
      + pad(date.getDate())
      + '_'
      + pad(date.getHours())
      + pad(date.getMinutes());
  }

  function cruzamentoSafeFilePart(value, fallback) {
    var text = String(value || fallback || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-zA-Z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .toUpperCase();

    return text || fallback || 'NAO_INFORMADO';
  }

  function cruzamentoFiltroResumo() {
    var filtros = cruzamentoState.filtros;
    return [
      'Classe Geral: ' + (filtros.classeGeral || 'Todos'),
      'Classe Filial Excesso: ' + (filtros.classeExcesso || 'Todos'),
      'Classe Filial Ruptura: ' + (filtros.classeRuptura || 'Todos'),
      'R$ Ruptura minimo: ' + cruzamentoFmtMoney(filtros.valorRupturaMin || 0),
      'Excesso minimo: ' + cruzamentoFmtNumber(filtros.qtdExcessoMin || 0),
      'Cobertura minima: ' + (filtros.coberturaMin === '' ? 'Sem limite' : cruzamentoFmtNumber(filtros.coberturaMin)),
      'Cobertura maxima: ' + (filtros.coberturaMax === '' ? 'Sem limite' : cruzamentoFmtNumber(filtros.coberturaMax)),
      'Busca: ' + (filtros.busca || 'Sem busca'),
      'Somente sugestao maior que zero: ' + (filtros.somenteSugestao ? 'Sim' : 'Nao')
    ];
  }

  function cruzamentoReportColumns() {
    return [
      { label: 'Codigo', key: 'codigo', type: 'text' },
      { label: 'Descricao', key: 'descricao', type: 'text' },
      { label: 'Filial Origem', key: 'filialOrigem', type: 'text' },
      { label: 'Filial Destino', key: 'filialDestino', type: 'text' },
      { label: 'Classe Geral', key: 'classeGeral', type: 'text' },
      { label: 'Classe Filial Excesso', key: 'classeExcesso', type: 'text' },
      { label: 'Classe Filial Ruptura', key: 'classeRuptura', type: 'text' },
      { label: 'Estoque Origem', key: 'estoqueOrigem', type: 'number' },
      { label: 'Excesso Qtd', key: 'excessoQtd', type: 'number' },
      { label: 'Cobertura', key: 'cobertura', type: 'number' },
      { label: 'Media Dia Excesso', key: 'mediaDiaExcesso', type: 'number' },
      { label: 'R$ Excesso', key: 'excessoValor', type: 'money' },
      { label: 'R$ Estoque', key: 'estoqueValor', type: 'money' },
      { label: 'R$ Ruptura', key: 'rupturaValor', type: 'money' },
      { label: 'DEZ', key: 'dez', type: 'number' },
      { label: 'Media Dia Ruptura', key: 'mediaDiaRuptura', type: 'number' },
      { label: 'Custo', key: 'custo', type: 'money' },
      { label: 'Qtd Necessaria Estimada', key: 'qtdNecessaria', type: 'number' },
      { label: 'Qtd Sugerida Transferir', key: 'qtdSugerida', type: 'number' },
      { label: 'Valor Estimado Transferencia', key: 'valorTransferencia', type: 'money' },
      { label: 'Observacao', key: 'observacao', type: 'text' }
    ];
  }

  function cruzamentoReportValue(item, col, formatted) {
    var value = item[col.key];
    if (value === null || value === undefined || value === '') return '';
    if (!formatted) return value;
    if (col.type === 'money') return cruzamentoFmtMoney(value);
    if (col.type === 'number') return cruzamentoFmtNumber(value);
    return value;
  }

  function cruzamentoReportRows(formatted) {
    var columns = cruzamentoReportColumns();
    return cruzamentoState.resultadosFiltrados.map(function (item) {
      return columns.map(function (col) {
        return cruzamentoReportValue(item, col, formatted);
      });
    });
  }

  function cruzamentoReportMeta(now) {
    return {
      titulo: 'Cruzamento de Excesso x Ruptura',
      dataHora: cruzamentoFormatDateTime(now),
      filialOrigem: cruzamentoState.excesso.filial || '-',
      filialDestino: cruzamentoState.ruptura.filial || '-',
      arquivoExcesso: cruzamentoState.excesso.fileName || '-',
      arquivoRuptura: cruzamentoState.ruptura.fileName || '-',
      filtros: cruzamentoFiltroResumo()
    };
  }

  function cruzamentoRenderExcesso() {
    var fileEl = document.getElementById('cruzamento-excesso-file');
    var linesEl = document.getElementById('cruzamento-excesso-lines');
    var statusEl = document.getElementById('cruzamento-excesso-status');
    var summaryEl = document.getElementById('cruzamento-excesso-summary');
    var kpiEl = document.getElementById('cruzamento-kpi-excesso');
    var excesso = cruzamentoState.excesso;

    if (fileEl) fileEl.textContent = excesso.fileName || 'Nenhum arquivo selecionado';
    if (linesEl) linesEl.textContent = 'Linhas lidas: ' + excesso.rows.length;
    if (statusEl) statusEl.textContent = excesso.message;
    if (kpiEl) kpiEl.textContent = String(excesso.rows.length);

    if (summaryEl) {
      summaryEl.classList.toggle('import-empty', !excesso.fileName);
      summaryEl.classList.toggle('import-ok', excesso.valid);
      summaryEl.classList.toggle('import-error', !!excesso.fileName && !excesso.valid);
    }

    cruzamentoRenderKpis();
  }

  function cruzamentoSetExcessoError(fileName, message) {
    cruzamentoState.excesso.fileName = fileName || '';
    cruzamentoState.excesso.headers = [];
    cruzamentoState.excesso.rows = [];
    cruzamentoState.excesso.items = [];
    cruzamentoState.excesso.valid = false;
    cruzamentoState.excesso.message = message;
    cruzamentoState.status = 'excesso-invalido';
    cruzamentoClearResultados();
    cruzamentoRenderExcesso();
    cruzamentoRenderMainStatus();
    cruzamentoRenderAcoes();
  }

  function cruzamentoRenderRuptura() {
    var fileEl = document.getElementById('cruzamento-ruptura-file');
    var linesEl = document.getElementById('cruzamento-ruptura-lines');
    var statusEl = document.getElementById('cruzamento-ruptura-status');
    var summaryEl = document.getElementById('cruzamento-ruptura-summary');
    var kpiEl = document.getElementById('cruzamento-kpi-ruptura');
    var ruptura = cruzamentoState.ruptura;

    if (fileEl) fileEl.textContent = ruptura.fileName || 'Nenhum arquivo selecionado';
    if (linesEl) linesEl.textContent = 'Linhas lidas: ' + ruptura.rows.length;
    if (statusEl) statusEl.textContent = ruptura.message;
    if (kpiEl) kpiEl.textContent = String(ruptura.rows.length);

    if (summaryEl) {
      summaryEl.classList.toggle('import-empty', !ruptura.fileName);
      summaryEl.classList.toggle('import-ok', ruptura.valid);
      summaryEl.classList.toggle('import-error', !!ruptura.fileName && !ruptura.valid);
    }

    cruzamentoRenderKpis();
  }

  function cruzamentoSetRupturaError(fileName, message) {
    cruzamentoState.ruptura.fileName = fileName || '';
    cruzamentoState.ruptura.headers = [];
    cruzamentoState.ruptura.rows = [];
    cruzamentoState.ruptura.items = [];
    cruzamentoState.ruptura.valid = false;
    cruzamentoState.ruptura.message = message;
    cruzamentoState.status = 'ruptura-invalido';
    cruzamentoClearResultados();
    cruzamentoRenderRuptura();
    cruzamentoRenderMainStatus();
    cruzamentoRenderAcoes();
  }

  function cruzamentoRenderMainStatus() {
    var status = document.getElementById('cruzamento-status');
    if (!status) return;
    status.setAttribute('data-state', cruzamentoState.status);
    if (cruzamentoState.status === 'excesso-lendo') {
      status.textContent = 'Lendo arquivo de excesso...';
    } else if (cruzamentoState.status === 'ruptura-lendo') {
      status.textContent = 'Lendo arquivo de ruptura...';
    } else if (cruzamentoState.status === 'analise-ok') {
      status.textContent = 'Analise concluida. Resultados filtrados: '
        + cruzamentoState.resultadosFiltrados.length
        + '. Rupturas sem origem em excesso: '
        + cruzamentoState.rupturaSemOrigem
        + '.';
    } else if (cruzamentoState.status === 'excesso-invalido') {
      status.textContent = 'Corrija o arquivo de excesso.';
    } else if (cruzamentoState.status === 'ruptura-invalido') {
      status.textContent = 'Corrija o arquivo de ruptura.';
    } else if (cruzamentoState.excesso.valid && cruzamentoState.ruptura.valid) {
      status.textContent = 'Excesso e ruptura importados. Pronto para proximas etapas.';
    } else if (cruzamentoState.excesso.valid) {
      status.textContent = 'Excesso importado. Aguardando ruptura.';
    } else if (cruzamentoState.ruptura.valid) {
      status.textContent = 'Ruptura importada. Aguardando excesso.';
    } else {
      status.textContent = 'Aguardando importacoes';
    }
  }

  function cruzamentoReadExcessoFile(file) {
    var XLSX = cruzamentoGetXLSX();
    var reader;

    if (!file) {
      cruzamentoSetExcessoError('', 'Aguardando arquivo de excesso.');
      return;
    }

    if (!/\.xlsx$/i.test(file.name)) {
      cruzamentoSetExcessoError(file.name, 'Selecione um arquivo .xlsx valido.');
      return;
    }

    if (!XLSX) {
      cruzamentoSetExcessoError(file.name, 'Biblioteca de leitura XLSX nao encontrada no sistema principal.');
      return;
    }

    cruzamentoState.excesso.fileName = file.name;
    cruzamentoState.excesso.headers = [];
    cruzamentoState.excesso.rows = [];
    cruzamentoState.excesso.items = [];
    cruzamentoState.excesso.valid = false;
    cruzamentoState.excesso.message = 'Lendo arquivo de excesso...';
    cruzamentoState.status = 'excesso-lendo';
    cruzamentoClearResultados();
    cruzamentoRenderExcesso();
    cruzamentoRenderTabela();
    cruzamentoRenderMainStatus();

    reader = new FileReader();
    reader.onload = function (evt) {
      var workbook;
      var rows;
      var headerIndex;
      var headers;
      var dataRows;
      var validation;

      try {
        workbook = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
        rows = cruzamentoWorkbookRows(XLSX, workbook);
        if (!rows) {
          cruzamentoSetExcessoError(file.name, 'A planilha nao possui abas para leitura.');
          return;
        }

        if (!rows.length) {
          cruzamentoSetExcessoError(file.name, 'A planilha de excesso esta vazia.');
          return;
        }

        headerIndex = cruzamentoFindHeaderRow(rows);
        headers = rows[headerIndex] || [];
        dataRows = rows.slice(headerIndex + 1);
        validation = cruzamentoValidateHeaders(headers, cruzamentoExcessoRequiredHeaders);

        cruzamentoState.excesso.fileName = file.name;
        cruzamentoState.excesso.headers = headers;
        cruzamentoState.excesso.rows = dataRows;
        cruzamentoState.excesso.valid = validation.valid;
        cruzamentoState.excesso.items = validation.valid ? cruzamentoParseExcessoRows(dataRows, headers) : [];
        cruzamentoState.excesso.message = validation.valid
          ? 'Arquivo de excesso validado com sucesso.'
          : 'Cabecalhos ausentes: ' + validation.missing.join(', ');
        cruzamentoState.status = validation.valid ? 'excesso-ok' : 'excesso-invalido';

        cruzamentoRenderExcesso();
        cruzamentoRenderMainStatus();
        cruzamentoRenderAcoes();
      } catch (err) {
        cruzamentoSetExcessoError(file.name, 'Falha ao ler a planilha de excesso.');
      }
    };
    reader.onerror = function () {
      cruzamentoSetExcessoError(file.name, 'Falha ao abrir o arquivo selecionado.');
    };
    reader.readAsArrayBuffer(file);
  }

  function cruzamentoReadRupturaFile(file) {
    var XLSX = cruzamentoGetXLSX();
    var reader;

    if (!file) {
      cruzamentoSetRupturaError('', 'Aguardando arquivo de ruptura.');
      return;
    }

    if (!/\.xlsx$/i.test(file.name)) {
      cruzamentoSetRupturaError(file.name, 'Selecione um arquivo .xlsx valido.');
      return;
    }

    if (!XLSX) {
      cruzamentoSetRupturaError(file.name, 'Biblioteca de leitura XLSX nao encontrada no sistema principal.');
      return;
    }

    cruzamentoState.ruptura.fileName = file.name;
    cruzamentoState.ruptura.headers = [];
    cruzamentoState.ruptura.rows = [];
    cruzamentoState.ruptura.items = [];
    cruzamentoState.ruptura.valid = false;
    cruzamentoState.ruptura.message = 'Lendo arquivo de ruptura...';
    cruzamentoState.status = 'ruptura-lendo';
    cruzamentoClearResultados();
    cruzamentoRenderRuptura();
    cruzamentoRenderTabela();
    cruzamentoRenderMainStatus();

    reader = new FileReader();
    reader.onload = function (evt) {
      var workbook;
      var rows;
      var headerIndex;
      var headers;
      var dataRows;
      var validation;

      try {
        workbook = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
        rows = cruzamentoWorkbookRows(XLSX, workbook);
        if (!rows) {
          cruzamentoSetRupturaError(file.name, 'A planilha nao possui abas para leitura.');
          return;
        }

        if (!rows.length) {
          cruzamentoSetRupturaError(file.name, 'A planilha de ruptura esta vazia.');
          return;
        }

        headerIndex = cruzamentoFindHeaderRow(rows);
        headers = rows[headerIndex] || [];
        dataRows = rows.slice(headerIndex + 1);
        validation = cruzamentoValidateHeaders(headers, cruzamentoRupturaRequiredHeaders);

        cruzamentoState.ruptura.fileName = file.name;
        cruzamentoState.ruptura.headers = headers;
        cruzamentoState.ruptura.rows = dataRows;
        cruzamentoState.ruptura.valid = validation.valid;
        cruzamentoState.ruptura.items = validation.valid ? cruzamentoParseRupturaRows(dataRows, headers) : [];
        cruzamentoState.ruptura.message = validation.valid
          ? 'Arquivo de ruptura validado com sucesso.'
          : 'Cabecalhos ausentes: ' + validation.missing.join(', ');
        cruzamentoState.status = validation.valid ? 'ruptura-ok' : 'ruptura-invalido';

        cruzamentoRenderRuptura();
        cruzamentoRenderMainStatus();
        cruzamentoRenderAcoes();
      } catch (err) {
        cruzamentoSetRupturaError(file.name, 'Falha ao ler a planilha de ruptura.');
      }
    };
    reader.onerror = function () {
      cruzamentoSetRupturaError(file.name, 'Falha ao abrir o arquivo selecionado.');
    };
    reader.readAsArrayBuffer(file);
  }

  function cruzamentoParseExcessoRows(rows, headers) {
    var sourceRows = rows || cruzamentoState.excesso.rows;
    var map = cruzamentoHeaderMap(headers || cruzamentoState.excesso.headers);
    return sourceRows.map(function (row) {
      var codDesc = cruzamentoSplitCodigoDescricao(cruzamentoCell(row, map, 'codigoDescricao'));
      var percExcesso = cruzamentoNumber(cruzamentoCell(row, map, 'percExcesso'), { percent: true });
      var valorExcesso = cruzamentoNumber(cruzamentoCell(row, map, 'valorExcesso'));
      var valorEstoque = cruzamentoNumber(cruzamentoCell(row, map, 'valorEstoque'));
      return {
        codigo: cruzamentoCode(codDesc.codigo),
        descricao: codDesc.descricao,
        estoque: cruzamentoNumber(cruzamentoCell(row, map, 'estoque')),
        mediaDia: cruzamentoNumber(cruzamentoCell(row, map, 'mediaDia')),
        cobertura: cruzamentoNumber(cruzamentoCell(row, map, 'cobertura')),
        excessoQtd: cruzamentoNumber(cruzamentoCell(row, map, 'excessoQtd')),
        percExcesso: percExcesso,
        valorExcesso: valorExcesso,
        valorEstoque: valorEstoque,
        excessoValor: valorExcesso,
        estoqueValor: valorEstoque,
        percentDisplay: percExcesso === null ? '' : cruzamentoFmtNumber(percExcesso * 100) + '%',
        display: {
          estoque: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'estoque'))),
          mediaDia: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'mediaDia'))),
          cobertura: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'cobertura'))),
          excessoQtd: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'excessoQtd'))),
          percExcesso: percExcesso === null ? '' : cruzamentoFmtNumber(percExcesso * 100) + '%',
          valorExcesso: cruzamentoFmtMoney(valorExcesso),
          valorEstoque: cruzamentoFmtMoney(valorEstoque)
        },
        classeGeral: String(cruzamentoCell(row, map, 'classeGeral') || '').trim().toUpperCase(),
        classeFilialExcesso: String(cruzamentoCell(row, map, 'classeFilial') || '').trim().toUpperCase()
      };
    }).filter(function (item) {
      return item.codigo;
    });
  }

  function cruzamentoParseRupturaRows(rows, headers) {
    var sourceRows = rows || cruzamentoState.ruptura.rows;
    var map = cruzamentoHeaderMap(headers || cruzamentoState.ruptura.headers);
    return sourceRows.map(function (row) {
      return {
        codigo: cruzamentoCode(cruzamentoCell(row, map, 'codigoProduto')),
        descricao: String(cruzamentoCell(row, map, 'descricao') || '').trim(),
        valorRuptura: cruzamentoNumber(cruzamentoCell(row, map, 'valorRuptura')),
        rupturaValor: cruzamentoNumber(cruzamentoCell(row, map, 'valorRuptura')),
        dez: cruzamentoNumber(cruzamentoCell(row, map, 'dez')),
        mediaDiaRuptura: cruzamentoNumber(cruzamentoCell(row, map, 'mediaDia')),
        custo: cruzamentoNumber(cruzamentoCell(row, map, 'custo')),
        display: {
          valorRuptura: cruzamentoFmtMoney(cruzamentoNumber(cruzamentoCell(row, map, 'valorRuptura'))),
          dez: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'dez'))),
          mediaDiaRuptura: cruzamentoFmtNumber(cruzamentoNumber(cruzamentoCell(row, map, 'mediaDia'))),
          custo: cruzamentoFmtMoney(cruzamentoNumber(cruzamentoCell(row, map, 'custo')))
        },
        classeGeral: String(cruzamentoCell(row, map, 'classeGeral') || '').trim().toUpperCase(),
        classeFilialRuptura: String(cruzamentoCell(row, map, 'classeFilial') || '').trim().toUpperCase()
      };
    }).filter(function (item) {
      return item.codigo;
    });
  }

  function cruzamentoBuildResultado(excesso, ruptura) {
    var observacoes = [];
    var necessidade = null;
    var excessoQtd = cruzamentoHasNumber(excesso.excessoQtd) ? Number(excesso.excessoQtd) : 0;
    var valorRuptura = cruzamentoHasNumber(ruptura.valorRuptura) ? Number(ruptura.valorRuptura) : null;
    var custo = cruzamentoHasNumber(ruptura.custo) ? Number(ruptura.custo) : null;
    var dez = cruzamentoHasNumber(ruptura.dez) ? Number(ruptura.dez) : null;
    var mediaDiaRuptura = cruzamentoHasNumber(ruptura.mediaDiaRuptura) ? Number(ruptura.mediaDiaRuptura) : null;
    var sugerida = 0;
    var valorTransferencia = null;

    if (valorRuptura !== null && custo !== null && custo > 0) {
      necessidade = cruzamentoRoundQty(valorRuptura / custo);
      observacoes.push('Necessidade por ruptura/custo');
    } else if (dez !== null && mediaDiaRuptura !== null) {
      necessidade = cruzamentoRoundQty(dez * mediaDiaRuptura);
      observacoes.push('Necessidade por DEZ x media dia');
    } else {
      observacoes.push('Necessidade nao calculada');
    }

    if (excessoQtd <= 0) {
      sugerida = 0;
      observacoes.push('Sem excesso disponivel');
    } else if (necessidade === null) {
      sugerida = 0;
      observacoes.push('Sugestao zerada sem necessidade calculada');
    } else {
      sugerida = cruzamentoRoundQty(Math.max(0, Math.min(excessoQtd, necessidade)));
    }

    if (custo !== null && custo > 0) {
      valorTransferencia = cruzamentoRoundQty(sugerida * custo);
    } else {
      observacoes.push('Valor sem custo valido');
    }

    if (sugerida > 0) {
      observacoes.push('Sugestao calculada');
    } else {
      observacoes.push('Sem quantidade sugerida');
    }

    return {
      codigo: excesso.codigo,
      descricao: excesso.descricao || ruptura.descricao,
      filialOrigem: cruzamentoState.excesso.filial || '-',
      filialDestino: cruzamentoState.ruptura.filial || '-',
      classeGeral: excesso.classeGeral || ruptura.classeGeral,
      classeExcesso: excesso.classeFilialExcesso,
      classeRuptura: ruptura.classeFilialRuptura,
      estoqueOrigem: excesso.estoque,
      excessoQtd: excesso.excessoQtd,
      cobertura: excesso.cobertura,
      mediaDiaExcesso: excesso.mediaDia,
      excessoValor: excesso.valorExcesso,
      estoqueValor: excesso.valorEstoque,
      rupturaValor: ruptura.valorRuptura,
      dez: ruptura.dez,
      mediaDiaRuptura: ruptura.mediaDiaRuptura,
      custo: ruptura.custo,
      qtdNecessaria: necessidade,
      qtdSugerida: sugerida,
      valorTransferencia: valorTransferencia,
      observacao: observacoes.join('; ')
    };
  }

  function cruzamentoAnalisar() {
    var excessoRows = cruzamentoState.excesso.items.length ? cruzamentoState.excesso.items : cruzamentoParseExcessoRows();
    var rupturaRows = cruzamentoState.ruptura.items.length ? cruzamentoState.ruptura.items : cruzamentoParseRupturaRows();
    var excessoPorCodigo = cruzamentoIndexByCodigo(excessoRows);
    var rupturaPorCodigo = cruzamentoIndexByCodigo(rupturaRows);
    var resultados = [];
    var semOrigem = 0;

    if (!cruzamentoState.excesso.valid || !cruzamentoState.ruptura.valid) return;

    cruzamentoState.indices.excesso = excessoPorCodigo;
    cruzamentoState.indices.ruptura = rupturaPorCodigo;

    rupturaRows.forEach(function (ruptura) {
      var codigo = cruzamentoCode(ruptura.codigo);
      var excesso = excessoPorCodigo[codigo];
      if (excesso) {
        resultados.push(cruzamentoBuildResultado(excesso, ruptura));
      } else {
        semOrigem += 1;
      }
    });

    cruzamentoState.resultados = resultados;
    cruzamentoState.rupturaSemOrigem = semOrigem;
    cruzamentoState.resultadoPronto = true;
    cruzamentoState.sort = {
      key: 'prioridade',
      dir: 'asc'
    };
    cruzamentoState.status = 'analise-ok';
    cruzamentoApplyFiltros();
    cruzamentoRenderAcoes();
    cruzamentoRenderMainStatus();
  }

  function cruzamentoPassaFiltros(item) {
    var filtros = cruzamentoState.filtros;
    var busca = String(filtros.busca || '').toUpperCase();

    if (filtros.classeGeral && item.classeGeral !== filtros.classeGeral) return false;
    if (filtros.classeExcesso && item.classeExcesso !== filtros.classeExcesso) return false;
    if (filtros.classeRuptura && item.classeRuptura !== filtros.classeRuptura) return false;
    if ((item.rupturaValor || 0) < (filtros.valorRupturaMin || 0)) return false;
    if ((item.excessoQtd || 0) < (filtros.qtdExcessoMin || 0)) return false;
    if (filtros.coberturaMin !== '' && (item.cobertura || 0) < filtros.coberturaMin) return false;
    if (filtros.coberturaMax !== '' && (item.cobertura || 0) > filtros.coberturaMax) return false;
    if (filtros.somenteSugestao && (item.qtdSugerida || 0) <= 0) return false;
    if (busca && (String(item.codigo).toUpperCase().indexOf(busca) < 0 && String(item.descricao).toUpperCase().indexOf(busca) < 0)) return false;
    return true;
  }

  function cruzamentoSortResultados(rows) {
    var sort = cruzamentoState.sort;
    var dir = sort.dir === 'desc' ? -1 : 1;
    var numeric = {
      rupturaValor: 1,
      excessoQtd: 1,
      qtdSugerida: 1,
      cobertura: 1
    };

    return rows.slice().sort(function (a, b) {
      var av = a[sort.key];
      var bv = b[sort.key];

      if (sort.key === 'prioridade') {
        var ac = cruzamentoClasseRank(a.classeGeral);
        var bc = cruzamentoClasseRank(b.classeGeral);
        if (ac !== bc) return ac - bc;
        if ((b.rupturaValor || 0) !== (a.rupturaValor || 0)) return (b.rupturaValor || 0) - (a.rupturaValor || 0);
        if ((b.qtdSugerida || 0) !== (a.qtdSugerida || 0)) return (b.qtdSugerida || 0) - (a.qtdSugerida || 0);
        return (a.cobertura || 0) - (b.cobertura || 0);
      }

      if (numeric[sort.key]) return ((Number(av) || 0) - (Number(bv) || 0)) * dir;
      av = String(av == null ? '' : av).toUpperCase();
      bv = String(bv == null ? '' : bv).toUpperCase();
      if (av < bv) return -1 * dir;
      if (av > bv) return 1 * dir;
      return 0;
    });
  }

  function cruzamentoApplyFiltros() {
    cruzamentoState.resultadosFiltrados = cruzamentoSortResultados(
      cruzamentoState.resultados.filter(cruzamentoPassaFiltros)
    );
    cruzamentoRenderKpis();
    cruzamentoRenderTabela();
    cruzamentoRenderAcoes();
  }

  function cruzamentoClearResultados() {
    cruzamentoState.resultadoPronto = false;
    cruzamentoState.resultados = [];
    cruzamentoState.resultadosFiltrados = [];
    cruzamentoState.indices = {
      excesso: {},
      ruptura: {}
    };
    cruzamentoState.rupturaSemOrigem = 0;
  }

  function cruzamentoTotals() {
    return cruzamentoState.resultadosFiltrados.reduce(function (acc, item) {
      acc.rupturaValor += item.rupturaValor || 0;
      acc.excessoValor += item.excessoValor || 0;
      acc.transferencia += item.valorTransferencia || 0;
      return acc;
    }, {
      rupturaValor: 0,
      excessoValor: 0,
      transferencia: 0
    });
  }

  function cruzamentoRenderKpis() {
    var totals = cruzamentoTotals();
    var values = {
      'cruzamento-kpi-excesso': cruzamentoState.excesso.items.length || cruzamentoState.excesso.rows.length,
      'cruzamento-kpi-ruptura': cruzamentoState.ruptura.items.length || cruzamentoState.ruptura.rows.length,
      'cruzamento-kpi-comum': cruzamentoState.resultados.length,
      'cruzamento-kpi-filtrados': cruzamentoState.resultadosFiltrados.length,
      'cruzamento-kpi-rup-valor': cruzamentoFmtMoney(totals.rupturaValor),
      'cruzamento-kpi-exc-valor': cruzamentoFmtMoney(totals.excessoValor),
      'cruzamento-kpi-transferencia': cruzamentoFmtMoney(totals.transferencia)
    };

    Object.keys(values).forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.textContent = values[id];
    });
  }

  function cruzamentoRenderTabela() {
    var body = document.getElementById('cruzamento-resultado-body');
    var rows = cruzamentoState.resultadosFiltrados;
    var html = '';

    if (!body) return;

    if (!cruzamentoState.resultadoPronto) {
      body.innerHTML = '<tr><td class="empty-cell" colspan="21">Importe as planilhas para visualizar o cruzamento.</td></tr>';
      return;
    }

    if (!rows.length) {
      body.innerHTML = '<tr><td class="empty-cell" colspan="21">Nenhum item encontrado para os filtros atuais.</td></tr>';
      return;
    }

    rows.forEach(function (item) {
      html += '<tr>';
      html += '<td>' + cruzamentoEscape(item.codigo) + '</td>';
      html += '<td>' + cruzamentoEscape(item.descricao) + '</td>';
      html += '<td>' + cruzamentoEscape(item.filialOrigem) + '</td>';
      html += '<td>' + cruzamentoEscape(item.filialDestino) + '</td>';
      html += '<td>' + cruzamentoEscape(item.classeGeral) + '</td>';
      html += '<td>' + cruzamentoEscape(item.classeExcesso) + '</td>';
      html += '<td>' + cruzamentoEscape(item.classeRuptura) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.estoqueOrigem) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.excessoQtd) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.cobertura) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.mediaDiaExcesso) + '</td>';
      html += '<td class="num">' + cruzamentoFmtMoney(item.excessoValor) + '</td>';
      html += '<td class="num">' + cruzamentoFmtMoney(item.estoqueValor) + '</td>';
      html += '<td class="num">' + cruzamentoFmtMoney(item.rupturaValor) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.dez) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.mediaDiaRuptura) + '</td>';
      html += '<td class="num">' + cruzamentoFmtMoney(item.custo) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.qtdNecessaria) + '</td>';
      html += '<td class="num">' + cruzamentoFmtNumber(item.qtdSugerida) + '</td>';
      html += '<td class="num">' + cruzamentoFmtMoney(item.valorTransferencia) + '</td>';
      html += '<td>' + cruzamentoEscape(item.observacao) + '</td>';
      html += '</tr>';
    });

    body.innerHTML = html;
  }

  function cruzamentoBindExcessoImport() {
    var input = document.getElementById('cruzamento-excesso-input');
    var filial = document.getElementById('cruzamento-excesso-filial');

    if (input) {
      input.addEventListener('change', function () {
        cruzamentoReadExcessoFile(input.files && input.files[0]);
      });
    }

    if (filial) {
      filial.addEventListener('input', function () {
        cruzamentoState.excesso.filial = filial.value.trim();
      });
    }

    cruzamentoRenderExcesso();
  }

  function cruzamentoBindRupturaImport() {
    var input = document.getElementById('cruzamento-ruptura-input');
    var filial = document.getElementById('cruzamento-ruptura-filial');

    if (input) {
      input.addEventListener('change', function () {
        cruzamentoReadRupturaFile(input.files && input.files[0]);
      });
    }

    if (filial) {
      filial.addEventListener('input', function () {
        cruzamentoState.ruptura.filial = filial.value.trim();
      });
    }

    cruzamentoRenderRuptura();
  }

  function cruzamentoNumberValue(id, emptyValue) {
    var el = document.getElementById(id);
    var value = el ? el.value : '';
    if (value === '') return emptyValue;
    return Number(value) || 0;
  }

  function cruzamentoRenderFiltros() {
    var filtros = cruzamentoState.filtros;
    var fields = {
      'cruzamento-filtro-classe-geral': filtros.classeGeral,
      'cruzamento-filtro-classe-excesso': filtros.classeExcesso,
      'cruzamento-filtro-classe-ruptura': filtros.classeRuptura,
      'cruzamento-filtro-valor-ruptura': filtros.valorRupturaMin,
      'cruzamento-filtro-qtd-excesso': filtros.qtdExcessoMin,
      'cruzamento-filtro-cobertura-min': filtros.coberturaMin,
      'cruzamento-filtro-cobertura-max': filtros.coberturaMax,
      'cruzamento-filtro-busca': filtros.busca
    };
    var check = document.getElementById('cruzamento-filtro-sugestao');

    Object.keys(fields).forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.value = fields[id];
    });

    if (check) check.checked = !!filtros.somenteSugestao;
  }

  function cruzamentoSyncFiltrosFromInputs() {
    var filtros = cruzamentoState.filtros;
    var el;

    el = document.getElementById('cruzamento-filtro-classe-geral');
    filtros.classeGeral = el ? el.value : '';
    el = document.getElementById('cruzamento-filtro-classe-excesso');
    filtros.classeExcesso = el ? el.value : '';
    el = document.getElementById('cruzamento-filtro-classe-ruptura');
    filtros.classeRuptura = el ? el.value : '';
    filtros.valorRupturaMin = cruzamentoNumberValue('cruzamento-filtro-valor-ruptura', 0);
    filtros.qtdExcessoMin = cruzamentoNumberValue('cruzamento-filtro-qtd-excesso', 0);
    filtros.coberturaMin = cruzamentoNumberValue('cruzamento-filtro-cobertura-min', '');
    filtros.coberturaMax = cruzamentoNumberValue('cruzamento-filtro-cobertura-max', '');
    el = document.getElementById('cruzamento-filtro-busca');
    filtros.busca = el ? el.value.trim() : '';
    el = document.getElementById('cruzamento-filtro-sugestao');
    filtros.somenteSugestao = !!(el && el.checked);

    if (cruzamentoState.resultadoPronto) {
      cruzamentoApplyFiltros();
      cruzamentoRenderMainStatus();
    }
  }

  function cruzamentoBindFiltros() {
    var ids = [
      'cruzamento-filtro-classe-geral',
      'cruzamento-filtro-classe-excesso',
      'cruzamento-filtro-classe-ruptura',
      'cruzamento-filtro-valor-ruptura',
      'cruzamento-filtro-qtd-excesso',
      'cruzamento-filtro-cobertura-min',
      'cruzamento-filtro-cobertura-max',
      'cruzamento-filtro-busca',
      'cruzamento-filtro-sugestao'
    ];

    ids.forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('change', cruzamentoSyncFiltrosFromInputs);
      el.addEventListener('input', cruzamentoSyncFiltrosFromInputs);
    });

    cruzamentoRenderFiltros();
    cruzamentoSyncFiltrosFromInputs();
  }

  function cruzamentoSetDisabled(id, disabled) {
    var el = document.getElementById(id);
    if (el) el.disabled = !!disabled;
  }

  function cruzamentoRenderAcoes() {
    var canAnalyze = cruzamentoState.excesso.valid && cruzamentoState.ruptura.valid;
    var hasResult = cruzamentoState.resultadoPronto;
    var hasFilteredRows = cruzamentoState.resultadosFiltrados.length > 0;

    cruzamentoSetDisabled('cruzamento-btn-analisar', !canAnalyze);
    cruzamentoSetDisabled('cruzamento-btn-exportar', !hasResult);
    cruzamentoSetDisabled('cruzamento-btn-pdf', !hasResult);
    cruzamentoSetDisabled('cruzamento-btn-copiar', !hasFilteredRows);
  }

  function cruzamentoExportarExcel() {
    var XLSX = cruzamentoGetXLSX();
    var now = cruzamentoNow();
    var meta = cruzamentoReportMeta(now);
    var totals = cruzamentoTotals();
    var columns = cruzamentoReportColumns();
    var rows = [
      [meta.titulo],
      ['Gerado em', meta.dataHora],
      ['Filial origem', meta.filialOrigem],
      ['Filial destino', meta.filialDestino],
      ['Arquivo excesso', meta.arquivoExcesso],
      ['Arquivo ruptura', meta.arquivoRuptura],
      ['Filtros aplicados', meta.filtros.join(' | ')],
      [],
      ['Resumo'],
      ['Itens em excesso', cruzamentoState.excesso.items.length || cruzamentoState.excesso.rows.length],
      ['Itens em ruptura', cruzamentoState.ruptura.items.length || cruzamentoState.ruptura.rows.length],
      ['Codigos em comum', cruzamentoState.resultados.length],
      ['Itens apos filtros', cruzamentoState.resultadosFiltrados.length],
      ['R$ Ruptura filtrada', totals.rupturaValor],
      ['R$ Excesso filtrado', totals.excessoValor],
      ['Transferencia sugerida', totals.transferencia],
      [],
      columns.map(function (col) { return col.label; })
    ];
    var fileName;
    var wb;
    var ws;

    if (!cruzamentoState.resultadoPronto) {
      cruzamentoSetStatusMessage('Gere a analise antes de exportar.', 'acao-pendente');
      return;
    }

    if (!XLSX) {
      cruzamentoSetStatusMessage('Biblioteca XLSX nao encontrada para exportar.', 'acao-pendente');
      return;
    }

    rows = rows.concat(cruzamentoReportRows(false));
    ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = columns.map(function (col) {
      return { wch: col.key === 'descricao' || col.key === 'observacao' ? 34 : 16 };
    });
    wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Cruzamento');

    fileName = 'Cruzamento_Excesso_Ruptura_'
      + cruzamentoSafeFilePart(meta.filialOrigem, 'ORIGEM')
      + '_para_'
      + cruzamentoSafeFilePart(meta.filialDestino, 'DESTINO')
      + '_'
      + cruzamentoFileDate(now)
      + '.xlsx';

    XLSX.writeFile(wb, fileName);
    cruzamentoSetStatusMessage('Excel gerado: ' + fileName, 'acao-ok');
  }

  function cruzamentoPrintTableHtml(rows) {
    var columns = cruzamentoReportColumns();
    var html = '<table><thead><tr>';

    columns.forEach(function (col) {
      html += '<th>' + cruzamentoEscape(col.label) + '</th>';
    });
    html += '</tr></thead><tbody>';

    if (!rows.length) {
      html += '<tr><td colspan="' + columns.length + '">Nenhum item encontrado para os filtros atuais.</td></tr>';
    } else {
      rows.forEach(function (row) {
        html += '<tr>';
        row.forEach(function (value) {
          html += '<td>' + cruzamentoEscape(value) + '</td>';
        });
        html += '</tr>';
      });
    }

    return html + '</tbody></table>';
  }

  function cruzamentoImprimirPdf() {
    var now = cruzamentoNow();
    var meta = cruzamentoReportMeta(now);
    var totals = cruzamentoTotals();
    var rows = cruzamentoReportRows(true);
    var win;
    var html;

    if (!cruzamentoState.resultadoPronto) {
      cruzamentoSetStatusMessage('Gere a analise antes de imprimir.', 'acao-pendente');
      return;
    }

    win = window.open('', '_blank');
    if (!win) {
      cruzamentoSetStatusMessage('O navegador bloqueou a janela de impressao.', 'acao-pendente');
      return;
    }

    html = '<!doctype html><html><head><meta charset="utf-8"><title>'
      + cruzamentoEscape(meta.titulo)
      + '</title><style>'
      + 'body{font-family:Arial,sans-serif;font-size:10px;color:#111;margin:14px;}'
      + 'h1{font-size:18px;margin:0 0 8px;} h2{font-size:13px;margin:14px 0 6px;}'
      + '.meta,.kpis{display:grid;grid-template-columns:repeat(2,1fr);gap:4px 18px;margin-bottom:10px;}'
      + '.filters{margin:8px 0 12px;} .filters span{display:inline-block;margin:0 8px 4px 0;}'
      + 'table{width:100%;border-collapse:collapse;font-size:8px;} th,td{border:1px solid #999;padding:3px;vertical-align:top;}'
      + 'th{background:#eee;} @page{size:A4 landscape;margin:8mm;} @media print{body{margin:0;} table{page-break-inside:auto;} tr{page-break-inside:avoid;page-break-after:auto;}}'
      + '</style></head><body>';
    html += '<h1>' + cruzamentoEscape(meta.titulo) + '</h1>';
    html += '<div class="meta">'
      + '<div><strong>Gerado em:</strong> ' + cruzamentoEscape(meta.dataHora) + '</div>'
      + '<div><strong>Origem:</strong> ' + cruzamentoEscape(meta.filialOrigem) + '</div>'
      + '<div><strong>Destino:</strong> ' + cruzamentoEscape(meta.filialDestino) + '</div>'
      + '<div><strong>Excesso:</strong> ' + cruzamentoEscape(meta.arquivoExcesso) + '</div>'
      + '<div><strong>Ruptura:</strong> ' + cruzamentoEscape(meta.arquivoRuptura) + '</div>'
      + '</div>';
    html += '<h2>Filtros</h2><div class="filters">';
    meta.filtros.forEach(function (filtro) {
      html += '<span>' + cruzamentoEscape(filtro) + '</span>';
    });
    html += '</div><h2>Resumo</h2><div class="kpis">'
      + '<div><strong>Itens em excesso:</strong> ' + cruzamentoEscape(cruzamentoState.excesso.items.length || cruzamentoState.excesso.rows.length) + '</div>'
      + '<div><strong>Itens em ruptura:</strong> ' + cruzamentoEscape(cruzamentoState.ruptura.items.length || cruzamentoState.ruptura.rows.length) + '</div>'
      + '<div><strong>Codigos em comum:</strong> ' + cruzamentoEscape(cruzamentoState.resultados.length) + '</div>'
      + '<div><strong>Itens apos filtros:</strong> ' + cruzamentoEscape(cruzamentoState.resultadosFiltrados.length) + '</div>'
      + '<div><strong>R$ Ruptura filtrada:</strong> ' + cruzamentoEscape(cruzamentoFmtMoney(totals.rupturaValor)) + '</div>'
      + '<div><strong>R$ Excesso filtrado:</strong> ' + cruzamentoEscape(cruzamentoFmtMoney(totals.excessoValor)) + '</div>'
      + '<div><strong>Transferencia sugerida:</strong> ' + cruzamentoEscape(cruzamentoFmtMoney(totals.transferencia)) + '</div>'
      + '</div><h2>Itens sugeridos</h2>';
    html += cruzamentoPrintTableHtml(rows);
    html += '</body></html>';

    win.document.open();
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(function () {
      win.print();
    }, 250);
    cruzamentoSetStatusMessage('Relatorio de impressao aberto.', 'acao-ok');
  }

  function cruzamentoCopyFallback(text) {
    var area = document.createElement('textarea');
    area.value = text;
    area.setAttribute('readonly', 'readonly');
    area.style.position = 'fixed';
    area.style.left = '-9999px';
    document.body.appendChild(area);
    area.select();
    try {
      document.execCommand('copy');
      return true;
    } catch (err) {
      return false;
    } finally {
      document.body.removeChild(area);
    }
  }

  function cruzamentoCopiarCodigos() {
    var codigos = cruzamentoState.resultadosFiltrados.map(function (item) {
      return item.codigo;
    }).filter(Boolean);
    var text = codigos.join('\n');

    function done(ok) {
      cruzamentoSetStatusMessage(
        ok ? codigos.length + ' codigos copiados.' : 'Nao foi possivel copiar os codigos.',
        ok ? 'acao-ok' : 'acao-pendente'
      );
    }

    if (!codigos.length) {
      cruzamentoSetStatusMessage('Nenhum codigo filtrado para copiar.', 'acao-pendente');
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () {
        done(true);
      }).catch(function () {
        done(cruzamentoCopyFallback(text));
      });
      return;
    }

    done(cruzamentoCopyFallback(text));
  }

  function cruzamentoResetImportState() {
    cruzamentoState.status = 'inicial';
    cruzamentoState.resultadoPronto = false;
    cruzamentoState.excesso = {
      fileName: '',
      filial: '',
      headers: [],
      rows: [],
      items: [],
      valid: false,
      message: 'Aguardando arquivo de excesso.'
    };
    cruzamentoState.ruptura = {
      fileName: '',
      filial: '',
      headers: [],
      rows: [],
      items: [],
      valid: false,
      message: 'Aguardando arquivo de ruptura.'
    };
    cruzamentoState.filtros = {
      classeGeral: '',
      classeExcesso: '',
      classeRuptura: '',
      valorRupturaMin: 0,
      qtdExcessoMin: 0,
      coberturaMin: '',
      coberturaMax: '',
      busca: '',
      somenteSugestao: false
    };
    cruzamentoState.resultados = [];
    cruzamentoState.resultadosFiltrados = [];
    cruzamentoState.indices = {
      excesso: {},
      ruptura: {}
    };
    cruzamentoState.rupturaSemOrigem = 0;
    cruzamentoState.sort = {
      key: 'prioridade',
      dir: 'asc'
    };
  }

  function cruzamentoClearFileInputs() {
    var ids = [
      'cruzamento-excesso-input',
      'cruzamento-ruptura-input',
      'cruzamento-excesso-filial',
      'cruzamento-ruptura-filial'
    ];

    ids.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.value = '';
    });
  }

  function cruzamentoLimparImportacoes() {
    if (cruzamentoState.resultadoPronto && !window.confirm('Limpar importacoes, filtros e resultado da analise?')) {
      return;
    }

    cruzamentoResetImportState();
    cruzamentoClearFileInputs();
    cruzamentoRenderExcesso();
    cruzamentoRenderRuptura();
    cruzamentoRenderFiltros();
    cruzamentoRenderTabela();
    cruzamentoRenderMainStatus();
    cruzamentoRenderAcoes();
  }

  function cruzamentoBindAcoes() {
    var analisar = document.getElementById('cruzamento-btn-analisar');
    var limpar = document.getElementById('cruzamento-btn-limpar');
    var exportar = document.getElementById('cruzamento-btn-exportar');
    var pdf = document.getElementById('cruzamento-btn-pdf');
    var copiar = document.getElementById('cruzamento-btn-copiar');

    if (analisar) {
      analisar.addEventListener('click', cruzamentoAnalisar);
    }
    if (limpar) limpar.addEventListener('click', cruzamentoLimparImportacoes);
    if (exportar) exportar.addEventListener('click', cruzamentoExportarExcel);
    if (pdf) pdf.addEventListener('click', cruzamentoImprimirPdf);
    if (copiar) copiar.addEventListener('click', cruzamentoCopiarCodigos);

    cruzamentoRenderAcoes();
  }

  function cruzamentoBindOrdenacaoTabela() {
    document.querySelectorAll('[data-sort]').forEach(function (th) {
      th.addEventListener('click', function () {
        var key = th.getAttribute('data-sort');
        if (!key) return;
        if (cruzamentoState.sort.key === key) {
          cruzamentoState.sort.dir = cruzamentoState.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          cruzamentoState.sort.key = key;
          cruzamentoState.sort.dir = 'asc';
        }
        if (cruzamentoState.resultadoPronto) cruzamentoApplyFiltros();
      });
    });
  }

  function cruzamentoRenderAccess() {
    var app = document.getElementById('cruzamento-app');
    var denied = document.getElementById('cruzamento-denied');

    cruzamentoState.currentUser = cruzamentoGetCurrentUser();
    cruzamentoState.canAccess = cruzamentoCanAccess(cruzamentoState.currentUser);

    cruzamentoSetVisible(app, cruzamentoState.canAccess);
    cruzamentoSetVisible(denied, !cruzamentoState.canAccess);

    cruzamentoRenderMainStatus();
  }

  function cruzamentoWatchAccess() {
    var tries = 0;
    var timer = setInterval(function () {
      tries += 1;
      cruzamentoRenderAccess();
      if (cruzamentoState.canAccess || tries >= 120) {
        clearInterval(timer);
      }
    }, 500);

    window.addEventListener('focus', cruzamentoRenderAccess);
    document.addEventListener('visibilitychange', cruzamentoRenderAccess);
  }

  cruzamentoBindExcessoImport();
  cruzamentoBindRupturaImport();
  cruzamentoBindFiltros();
  cruzamentoBindAcoes();
  cruzamentoBindOrdenacaoTabela();
  cruzamentoRenderKpis();
  cruzamentoRenderTabela();
  cruzamentoRenderAccess();
  cruzamentoWatchAccess();
})();
