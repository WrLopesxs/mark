(function () {
  'use strict';

  const STORAGE_KEY = 'https://script.google.com/macros/s/AKfycbycPVMOMELCcc63j0BWvbqytnwDFzjgnjpEQ1lsqaiJCwgDpnoa_5kf5rxmVDAI8KBlgQ/exec';
  const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbycPVMOMELCcc63j0BWvbqytnwDFzjgnjpEQ1lsqaiJCwgDpnoa_5kf5rxmVDAI8KBlgQ/exec';
  const STANDARD_IMAGE_WIDTH = 1200;
  const STANDARD_IMAGE_HEIGHT = 800;
  const AUTO_LOAD_DELAY_MS = 250;
  const AUTO_LOAD_MAX_ATTEMPTS = 8;

  const state = {
    drawing: false,
    imageLoaded: false,
    loadingImage: false,
    autoLoadTimer: 0,
    brushColor: '#e11d48',
    brushSize: 8,
    renderBox: {
      x: 0,
      y: 0,
      width: STANDARD_IMAGE_WIDTH,
      height: STANDARD_IMAGE_HEIGHT
    },
    query: new URLSearchParams(window.location.search)
  };

  const els = {
    statusText: document.getElementById('statusText'),
    recordId: document.getElementById('recordId'),
    partNumber: document.getElementById('partNumber'),
    defectType: document.getElementById('defectType'),
    sheetRow: document.getElementById('sheetRow'),
    userName: document.getElementById('userName'),
    apiUrl: document.getElementById('apiUrl'),
    loadImageButton: document.getElementById('loadImageButton'),
    loadManualUrlButton: document.getElementById('loadManualUrlButton'),
    saveButton: document.getElementById('saveButton'),
    clearButton: document.getElementById('clearButton'),
    reloadButton: document.getElementById('reloadButton'),
    manualImageUrl: document.getElementById('manualImageUrl'),
    brushColor: document.getElementById('brushColor'),
    brushSize: document.getElementById('brushSize'),
    brushPreview: document.getElementById('brushPreview'),
    canvasSubtitle: document.getElementById('canvasSubtitle'),
    canvasMode: document.getElementById('canvasMode'),
    canvasStage: document.getElementById('canvasStage'),
    canvasEmpty: document.getElementById('canvasEmpty'),
    baseImage: document.getElementById('baseImage'),
    annotationCanvas: document.getElementById('annotationCanvas')
  };

  const ctx = els.annotationCanvas.getContext('2d');

  function init() {
    fillFromQuery();
    hydrateSavedApiUrl();
    bindEvents();
    resetCanvasStage();
    updateBrushPreview();
    scheduleInitialImageLoad();
  }

  function bindEvents() {
    if (els.loadImageButton) {
      els.loadImageButton.addEventListener('click', loadPartImage);
    }

    if (els.loadManualUrlButton) {
      els.loadManualUrlButton.addEventListener('click', loadManualImageFromUrl);
    }

    if (els.saveButton) {
      els.saveButton.addEventListener('click', saveAnnotation);
    }

    if (els.clearButton) {
      els.clearButton.addEventListener('click', clearCanvas);
    }

    if (els.reloadButton) {
      els.reloadButton.addEventListener('click', function () {
        window.location.reload();
      });
    }

    if (els.brushSize) {
      els.brushSize.addEventListener('input', updateBrushPreview);
    }

    if (els.apiUrl) {
      els.apiUrl.addEventListener('change', persistApiUrl);
      els.apiUrl.addEventListener('blur', persistApiUrl);
    }

    els.annotationCanvas.addEventListener('pointerdown', onPointerDown);
    els.annotationCanvas.addEventListener('pointermove', onPointerMove);
    els.annotationCanvas.addEventListener('pointerleave', onPointerUp);
    els.annotationCanvas.addEventListener('pointercancel', onPointerUp);
    window.addEventListener('pointerup', onPointerUp);
  }

  function fillFromQuery() {
    els.recordId.value = readQuery('id');
    els.partNumber.value = readQuery('pn');
    els.defectType.value = readQuery('defeito') || readQuery('defect');
    els.sheetRow.value = readQuery('row');
    els.userName.value = readQuery('user');
    els.apiUrl.value = readQuery('api') || '';
  }

  function hydrateSavedApiUrl() {
    if (els.apiUrl.value) {
      persistApiUrl();
      return;
    }

    const cached = window.localStorage.getItem(STORAGE_KEY);
    els.apiUrl.value = cached || DEFAULT_API_URL;
  }

  function persistApiUrl() {
    const apiUrl = getApiUrl();

    if (apiUrl) {
      window.localStorage.setItem(STORAGE_KEY, apiUrl);
    }
  }

  function readQuery(key) {
    return (state.query.get(key) || '').trim();
  }

  function getApiUrl() {
    return (els.apiUrl.value || '').trim();
  }

  function updateBrushPreview() {
    els.brushPreview.textContent = 'Espessura atual: ' + els.brushSize.value + ' px';
  }

  function scheduleInitialImageLoad() {
    window.clearTimeout(state.autoLoadTimer);
    attemptInitialImageLoad(AUTO_LOAD_MAX_ATTEMPTS);
  }

  function attemptInitialImageLoad(remainingAttempts) {
    const apiUrl = getApiUrl();
    const pn = els.partNumber.value.trim();

    if (apiUrl && pn) {
      state.autoLoadTimer = window.setTimeout(function () {
        loadPartImage({ auto: true });
      }, AUTO_LOAD_DELAY_MS);
      return;
    }

    if (remainingAttempts <= 0) {
      if (!pn) {
        setStatus('Informe os parametros pela URL da planilha para localizar a peca.', 'warning');
      } else {
        setStatus('Preencha a URL da API antes de buscar a imagem da peca.', 'warning');
      }
      return;
    }

    state.autoLoadTimer = window.setTimeout(function () {
      attemptInitialImageLoad(remainingAttempts - 1);
    }, AUTO_LOAD_DELAY_MS);
  }

  function setStatus(message, mode) {
    els.statusText.textContent = message;
    els.canvasMode.textContent =
      mode === 'error' ? 'Erro' :
      mode === 'success' ? 'Pronto' :
      mode === 'loading' ? 'Carregando' :
      mode === 'warning' ? 'Atencao' :
      'Pronto';
  }

  function resetCanvasStage() {
    syncStageDimensions(STANDARD_IMAGE_WIDTH, STANDARD_IMAGE_HEIGHT);
    clearCanvasSurface();

    els.baseImage.removeAttribute('src');
    els.baseImage.style.display = 'none';
    els.annotationCanvas.style.display = 'none';
    els.canvasEmpty.style.display = 'grid';

    state.drawing = false;
    state.imageLoaded = false;
    state.renderBox = {
      x: 0,
      y: 0,
      width: STANDARD_IMAGE_WIDTH,
      height: STANDARD_IMAGE_HEIGHT
    };
  }

  function syncStageDimensions(width, height) {
    els.baseImage.width = width;
    els.baseImage.height = height;
    els.baseImage.style.width = width + 'px';
    els.baseImage.style.height = height + 'px';

    els.annotationCanvas.width = width;
    els.annotationCanvas.height = height;
    els.annotationCanvas.style.width = width + 'px';
    els.annotationCanvas.style.height = height + 'px';

    els.canvasStage.style.height = height + 'px';
    els.canvasStage.style.minHeight = height + 'px';

    configureCanvasContext();
  }

  function configureCanvasContext() {
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
  }

  function clearCanvasSurface() {
    ctx.clearRect(0, 0, els.annotationCanvas.width, els.annotationCanvas.height);
    ctx.beginPath();
  }

  async function loadPartImage(options) {
    const settings = options && options.auto ? options : {};
    const apiUrl = getApiUrl();
    const pn = els.partNumber.value.trim();

    if (state.loadingImage) {
      return;
    }

    if (!apiUrl) {
      setStatus('Preencha a URL da API antes de buscar a imagem da peca.', 'warning');
      els.apiUrl.focus();
      return;
    }

    if (!pn) {
      setStatus('Part Number nao foi recebido pela URL.', 'warning');
      return;
    }

    setStatus('Buscando imagem cadastrada para o PN ' + pn + '...', 'loading');
    state.loadingImage = true;

    try {
      const url = new URL(apiUrl);
      url.searchParams.set('action', 'partImage');
      url.searchParams.set('pn', pn);

      const response = await fetch(url.toString(), {
        method: 'GET',
        cache: 'no-store'
      });
      const payload = await response.json();
      const imageSource = payload.imageDataUrl || payload.imageUrl;

      if (!response.ok || !payload.ok || !imageSource) {
        throw new Error(payload.message || payload.error || 'Imagem nao encontrada para o PN informado.');
      }

      await drawBaseImage(imageSource);
      els.canvasSubtitle.textContent = 'PN ' + pn + ' carregado para marcacao.';
      setStatus(
        settings.auto
          ? 'Imagem carregada automaticamente. Agora voce ja pode marcar o defeito.'
          : 'Imagem carregada. Agora voce ja pode marcar o defeito.',
        'success'
      );
    } catch (error) {
      setStatus('Falha ao buscar a imagem: ' + error.message, 'error');
    } finally {
      state.loadingImage = false;
    }
  }

  async function loadManualImageFromUrl() {
    const manualUrl = (els.manualImageUrl.value || '').trim();

    if (!manualUrl) {
      setStatus('Preencha a URL da imagem manual antes de carregar.', 'warning');
      return;
    }

    setStatus('Carregando imagem manual pela URL...', 'loading');

    try {
      await drawBaseImage(manualUrl);
      els.canvasSubtitle.textContent = 'Imagem manual carregada por URL.';
      setStatus('Imagem manual pronta para marcacao.', 'success');
    } catch (error) {
      setStatus('Falha ao carregar a imagem manual: ' + error.message, 'error');
    }
  }

  function drawBaseImage(imageUrl) {
    return new Promise(function (resolve, reject) {
      const image = new Image();
      image.crossOrigin = 'anonymous';

      image.onload = function () {
        syncStageDimensions(STANDARD_IMAGE_WIDTH, STANDARD_IMAGE_HEIGHT);
        clearCanvasSurface();

        state.renderBox = calculateContainBox(
          image.naturalWidth,
          image.naturalHeight,
          STANDARD_IMAGE_WIDTH,
          STANDARD_IMAGE_HEIGHT
        );

        els.baseImage.src = imageUrl;
        els.baseImage.style.display = 'block';
        els.annotationCanvas.style.display = 'block';
        els.canvasEmpty.style.display = 'none';

        state.drawing = false;
        state.imageLoaded = true;

        resolve();
      };

      image.onerror = function () {
        reject(new Error('Nao foi possivel abrir a imagem.'));
      };

      image.src = imageUrl;
    });
  }

  function calculateContainBox(sourceWidth, sourceHeight, targetWidth, targetHeight) {
    if (!sourceWidth || !sourceHeight) {
      return {
        x: 0,
        y: 0,
        width: targetWidth,
        height: targetHeight
      };
    }

    const scale = Math.min(targetWidth / sourceWidth, targetHeight / sourceHeight);
    const width = Math.round(sourceWidth * scale);
    const height = Math.round(sourceHeight * scale);

    return {
      x: Math.round((targetWidth - width) / 2),
      y: Math.round((targetHeight - height) / 2),
      width: width,
      height: height
    };
  }

  function clearCanvas() {
    if (!state.imageLoaded) {
      setStatus('Carregue uma imagem antes de limpar ou desenhar.', 'warning');
      return;
    }

    clearCanvasSurface();
    setStatus('Desenho limpo. Voce pode marcar novamente.', 'success');
  }

  function onPointerDown(event) {
    if (!state.imageLoaded) {
      return;
    }

    event.preventDefault();

    state.drawing = true;
    state.brushColor = els.brushColor.value;
    state.brushSize = Number(els.brushSize.value);

    ctx.strokeStyle = state.brushColor;
    ctx.lineWidth = state.brushSize;

    const point = getCanvasPoint(event);
    ctx.beginPath();
    ctx.moveTo(point.x, point.y);

    if (typeof els.annotationCanvas.setPointerCapture === 'function' && event.pointerId != null) {
      els.annotationCanvas.setPointerCapture(event.pointerId);
    }
  }

  function onPointerMove(event) {
    if (!state.drawing || !state.imageLoaded) {
      return;
    }

    event.preventDefault();

    const point = getCanvasPoint(event);
    ctx.lineTo(point.x, point.y);
    ctx.stroke();
  }

  function onPointerUp(event) {
    if (!state.drawing) {
      return;
    }

    state.drawing = false;
    ctx.closePath();

    if (
      event &&
      event.pointerId != null &&
      typeof els.annotationCanvas.hasPointerCapture === 'function' &&
      els.annotationCanvas.hasPointerCapture(event.pointerId)
    ) {
      els.annotationCanvas.releasePointerCapture(event.pointerId);
    }
  }

  function getCanvasPoint(event) {
    const rect = els.annotationCanvas.getBoundingClientRect();
    const scaleX = els.annotationCanvas.width / rect.width;
    const scaleY = els.annotationCanvas.height / rect.height;

    return {
      x: (event.clientX - rect.left) * scaleX,
      y: (event.clientY - rect.top) * scaleY
    };
  }

  async function saveAnnotation() {
    const apiUrl = getApiUrl();
    const id = els.recordId.value.trim();
    const pn = els.partNumber.value.trim();
    const defect = els.defectType.value.trim();
    const user = els.userName.value.trim();

    if (!apiUrl) {
      setStatus('Preencha a URL da API antes de salvar.', 'warning');
      els.apiUrl.focus();
      return;
    }

    if (!id || !pn || !defect) {
      setStatus('Faltam dados do registro. Abra esta pagina pelo link gerado na planilha.', 'warning');
      return;
    }

    if (!user) {
      setStatus('Preencha o nome do usuario antes de salvar a marcacao.', 'warning');
      els.userName.focus();
      return;
    }

    if (!state.imageLoaded) {
      setStatus('Carregue a imagem da peca antes de salvar a marcacao.', 'warning');
      return;
    }

    setStatus('Gerando imagem final e enviando para a API...', 'loading');

    try {
      const payload = {
        action: 'saveAnnotation',
        id: id,
        pn: pn,
        defect: defect,
        row: els.sheetRow.value.trim(),
        user: user,
        imageDataUrl: exportMergedImage()
      };

      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify(payload)
      });
      const result = await response.json();

      if (!response.ok || !result.ok) {
        throw new Error(result.message || result.error || 'A API nao confirmou o salvamento.');
      }

      setStatus('Marcacao salva com sucesso na linha ' + result.row + '.', 'success');
    } catch (error) {
      setStatus('Erro ao salvar a marcacao: ' + error.message, 'error');
    }
  }

  function exportMergedImage() {
    const scale = 2;
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = els.annotationCanvas.width * scale;
    exportCanvas.height = els.annotationCanvas.height * scale;

    const exportCtx = exportCanvas.getContext('2d');
    exportCtx.imageSmoothingEnabled = true;
    exportCtx.imageSmoothingQuality = 'high';
    exportCtx.scale(scale, scale);
    exportCtx.fillStyle = '#ffffff';
    exportCtx.fillRect(0, 0, els.annotationCanvas.width, els.annotationCanvas.height);

    exportCtx.drawImage(
      els.baseImage,
      state.renderBox.x,
      state.renderBox.y,
      state.renderBox.width,
      state.renderBox.height
    );

    exportCtx.drawImage(
      els.annotationCanvas,
      0,
      0,
      els.annotationCanvas.width,
      els.annotationCanvas.height
    );

    return exportCanvas.toDataURL('image/png');
  }

  init();
})();
