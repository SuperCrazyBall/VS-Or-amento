(function () {
  var cruzamentoState = {
    status: 'inicial',
    modulo: 'cruzamento-transferencia',
    currentUser: null,
    canAccess: false,
    excesso: {
      fileName: '',
      filial: '',
      headers: [],
      rows: [],
      valid: false,
      message: 'Aguardando arquivo de excesso.'
    }
  };

  var cruzamentoExcessoRequiredHeaders = [
    'CODIGO DESCRICAO',
    'ESTOQUE',
    'MEDIA DIA',
    'COBERTURA',
    'EXCESSO'
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

  function cruzamentoNormalizeHeader(value) {
    return String(value || '').trim().toUpperCase().replace(/\s+/g, ' ');
  }

  function cruzamentoFindHeaderRow(rows) {
    for (var i = 0; i < rows.length; i += 1) {
      var normalized = rows[i].map(cruzamentoNormalizeHeader);
      if (normalized.indexOf('CODIGO DESCRICAO') >= 0 || normalized.indexOf('ESTOQUE') >= 0) {
        return i;
      }
    }
    return 0;
  }

  function cruzamentoValidateHeaders(headers, required) {
    var normalized = headers.map(cruzamentoNormalizeHeader);
    var missing = required.filter(function (name) {
      return normalized.indexOf(name) < 0;
    });
    return {
      valid: missing.length === 0,
      missing: missing
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
  }

  function cruzamentoSetExcessoError(fileName, message) {
    cruzamentoState.excesso.fileName = fileName || '';
    cruzamentoState.excesso.headers = [];
    cruzamentoState.excesso.rows = [];
    cruzamentoState.excesso.valid = false;
    cruzamentoState.excesso.message = message;
    cruzamentoState.status = 'excesso-invalido';
    cruzamentoRenderExcesso();
    cruzamentoRenderMainStatus();
  }

  function cruzamentoRenderMainStatus() {
    var status = document.getElementById('cruzamento-status');
    if (!status) return;
    status.setAttribute('data-state', cruzamentoState.status);
    if (cruzamentoState.status === 'excesso-lendo') {
      status.textContent = 'Lendo arquivo de excesso...';
    } else if (cruzamentoState.status === 'excesso-invalido') {
      status.textContent = 'Corrija o arquivo de excesso.';
    } else if (cruzamentoState.excesso.valid) {
      status.textContent = 'Excesso importado. Aguardando ruptura.';
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
    cruzamentoState.excesso.valid = false;
    cruzamentoState.excesso.message = 'Lendo arquivo de excesso...';
    cruzamentoState.status = 'excesso-lendo';
    cruzamentoRenderExcesso();
    cruzamentoRenderMainStatus();

    reader = new FileReader();
    reader.onload = function (evt) {
      var workbook;
      var sheetName;
      var rows;
      var headerIndex;
      var headers;
      var dataRows;
      var validation;

      try {
        workbook = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
        sheetName = workbook.SheetNames[0];
        if (!sheetName) {
          cruzamentoSetExcessoError(file.name, 'A planilha nao possui abas para leitura.');
          return;
        }

        rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
        rows = rows.filter(function (row) {
          return row.some(function (cell) { return String(cell || '').trim() !== ''; });
        });

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
        cruzamentoState.excesso.message = validation.valid
          ? 'Arquivo de excesso validado com sucesso.'
          : 'Cabecalhos ausentes: ' + validation.missing.join(', ');
        cruzamentoState.status = validation.valid ? 'excesso-ok' : 'excesso-invalido';

        cruzamentoRenderExcesso();
        cruzamentoRenderMainStatus();
      } catch (err) {
        cruzamentoSetExcessoError(file.name, 'Falha ao ler a planilha de excesso.');
      }
    };
    reader.onerror = function () {
      cruzamentoSetExcessoError(file.name, 'Falha ao abrir o arquivo selecionado.');
    };
    reader.readAsArrayBuffer(file);
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
  cruzamentoRenderAccess();
  cruzamentoWatchAccess();
})();
