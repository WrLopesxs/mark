(function () {
  'use strict';

  const STORAGE_KEY = 'https://script.google.com/macros/s/AKfycbycPVMOMELCcc63j0BWvbqytnwDFzjgnjpEQ1lsqaiJCwgDpnoa_5kf5rxmVDAI8KBlgQ/exec';
  const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbycPVMOMELCcc63j0BWvbqytnwDFzjgnjpEQ1lsqaiJCwgDpnoa_5kf5rxmVDAI8KBlgQ/exec';
  const MARKER_STORAGE_PREFIX = 'toyota-qc-markers';
  const STANDARD_IMAGE_WIDTH = 1200;
  const STANDARD_IMAGE_HEIGHT = 800;
  const AUTO_LOAD_DELAY_MS = 250;
  const AUTO_LOAD_MAX_ATTEMPTS = 8;
  const MIN_MARKER_WIDTH = 0.055;
  const MIN_MARKER_HEIGHT = 0.055;

  const DEFECTS = [
    { key: 'amassado', label: 'Amassado', short: 'AM', color: '#c56a1b', iconType: 'ring', defaultWidth: 0.12, defaultHeight: 0.12, description: 'Marca circular para deformacao localizada.' },
    { key: 'caroco', label: 'Caroco', short: 'CA', color: '#a16207', iconType: 'dot-ring', defaultWidth: 0.1, defaultHeight: 0.1, description: 'Ponto de destaque para saliencia pontual.' },
    { key: 'estiramento', label: 'Estiramento', short: 'ES', color: '#2563eb', iconType: 'stretch', defaultWidth: 0.18, defaultHeight: 0.09, description: 'Indicacao linear para alongamento da chapa.' },
    { key: 'limalha', label: 'Limalha', short: 'LI', color: '#92400e', iconType: 'shard', defaultWidth: 0.11, defaultHeight: 0.11, description: 'Triangulo de alerta para particulas e cavacos.' },
    { key: 'marca-ferramenta', label: 'Marca de ferramenta', short: 'MF', color: '#7c4a03', iconType: 'tool', defaultWidth: 0.14, defaultHeight: 0.12, description: 'Icone tecnico para marca de contato com ferramenta.' },
    { key: 'poro', label: 'Poro', short: 'PO', color: '#0f766e', iconType: 'pore', defaultWidth: 0.1, defaultHeight: 0.1, description: 'Anel com centro destacado para porosidade.' },
    { key: 'rebarba', label: 'Rebarba', short: 'RE', color: '#15803d', iconType: 'burr', defaultWidth: 0.14, defaultHeight: 0.1, description: 'Traco serrilhado para excesso de material em borda.' },
    { key: 'risco', label: 'Risco', short: 'RI', color: '#ca8a04', iconType: 'scratch', defaultWidth: 0.2, defaultHeight: 0.08, description: 'Linha fina para risco superficial.' },
    { key: 'ruga', label: 'Ruga', short: 'RU', color: '#7c6543', iconType: 'wrinkle', defaultWidth: 0.18, defaultHeight: 0.1, description: 'Linha ondulada para rugas e ondulacoes.' },
    { key: 'trinca', label: 'Trinca', short: 'TR', color: '#b91c1c', iconType: 'crack', defaultWidth: 0.2, defaultHeight: 0.1, description: 'Linha vermelha irregular para trinca.' },
    { key: 'vinco', label: 'Vinco', short: 'VI', color: '#c2410c', iconType: 'crease', defaultWidth: 0.2, defaultHeight: 0.09, description: 'Linha forte para vinco ou dobra marcada.' },
    { key: 'transbordo', label: 'Transbordo', short: 'TB', color: '#be123c', iconType: 'overflow', defaultWidth: 0.18, defaultHeight: 0.1, description: 'Marcador de fluxo para excesso ou transbordo.' },
    { key: 'corrosao', label: 'Corrosao', short: 'CO', color: '#9a3412', iconType: 'corrosion', defaultWidth: 0.12, defaultHeight: 0.12, description: 'Sinalizacao oxidada para pontos de corrosao.' },
    { key: 'palete-kaizen-ok', label: 'Palete para Kaizen, Pecas OK', short: 'OK', color: '#166534', iconType: 'ok', defaultWidth: 0.15, defaultHeight: 0.11, description: 'Selo verde para palete de Kaizen e peca aprovada.' },
    { key: 'marca-oleo', label: 'Marca de oleo', short: 'MO', color: '#1d4ed8', iconType: 'oil', defaultWidth: 0.11, defaultHeight: 0.11, description: 'Gota destacada para residuo ou marca de oleo.' },
    { key: 'furo-obstruido', label: 'Furo obstruido', short: 'FO', color: '#334155', iconType: 'blocked-hole', defaultWidth: 0.11, defaultHeight: 0.11, description: 'Circulo bloqueado para furo obstruido.' },
    { key: 'esfoliamento', label: 'Esfoliamento', short: 'EF', color: '#0f766e', iconType: 'peel', defaultWidth: 0.13, defaultHeight: 0.11, description: 'Sinal de camada destacando esfoliamento.' },
    { key: 'checagem-100', label: 'Checagem de qualidade 100%', short: '100', color: '#1e40af', iconType: 'inspection', defaultWidth: 0.15, defaultHeight: 0.11, description: 'Selo azul para inspecao total.' },
    { key: 'amassado-ventosa', label: 'Amassado de ventosa', short: 'AV', color: '#0369a1', iconType: 'suction', defaultWidth: 0.13, defaultHeight: 0.13, description: 'Duplo anel para marcas de ventosa.' }
  ];

  const DEFECTS_BY_KEY = DEFECTS.reduce(function (map, defect) {
    map[defect.key] = defect;
    return map;
  }, {});

  const state = {
    imageLoaded: false,
    loadingImage: false,
    autoLoadTimer: 0,
    renderBox: {
      x: 0,
      y: 0,
      width: STANDARD_IMAGE_WIDTH,
      height: STANDARD_IMAGE_HEIGHT
    },
    query: new URLSearchParams(window.location.search),
    selectedDefectKey: DEFECTS[0].key,
    selectedMarkerId: '',
    activePointer: null,
    markers: [],
    preventPlacementUntil: 0
  };

  const els = {
    statusText: document.getElementById('statusText'),
    recordId: document.getElementById('recordId'),
    partNumber: document.getElementById('partNumber'),
    defectType: document.getElementById('defectType'),
    sheetRow: document.getElementById('sheetRow'),
    userName: document.getElementById('userName'),
    apiUrl: document.getElementById('apiUrl'),
    manualImageUrl: document.getElementById('manualImageUrl'),
    loadImageButton: document.getElementById('loadImageButton'),
    saveButton: document.getElementById('saveButton'),
    clearButton: document.getElementById('clearButton'),
    removeSelectedButton: document.getElementById('removeSelectedButton'),
    reloadButton: document.getElementById('reloadButton'),
    canvasSubtitle: document.getElementById('canvasSubtitle'),
    canvasMode: document.getElementById('canvasMode'),
    canvasStage: document.getElementById('canvasStage'),
    stageSurface: document.getElementById('stageSurface'),
    canvasEmpty: document.getElementById('canvasEmpty'),
    baseImage: document.getElementById('baseImage'),
    markerLayer: document.getElementById('markerLayer'),
    defectPalette: document.getElementById('defectPalette'),
    selectedDefectCard: document.getElementById('selectedDefectCard'),
    selectedDefectIcon: document.getElementById('selectedDefectIcon'),
    selectedDefectName: document.getElementById('selectedDefectName'),
    selectedDefectHint: document.getElementById('selectedDefectHint'),
    markerCount: document.getElementById('markerCount')
  };

  function init() {
    fillFromQuery();
    hydrateSavedApiUrl();
    bindEvents();
    renderDefectPalette();
    syncSelectedDefectCard();
    resetCanvasStage();
    updateMarkerSummary();
    scheduleInitialImageLoad();
  }

  function bindEvents() {
    if (els.loadImageButton) {
      els.loadImageButton.addEventListener('click', loadPartImage);
    }

    if (els.saveButton) {
      els.saveButton.addEventListener('click', saveAnnotation);
    }

    if (els.clearButton) {
      els.clearButton.addEventListener('click', clearMarkers);
    }

    if (els.removeSelectedButton) {
      els.removeSelectedButton.addEventListener('click', removeSelectedMarker);
    }

    if (els.reloadButton) {
      els.reloadButton.addEventListener('click', function () {
        window.location.reload();
      });
    }

    if (els.apiUrl) {
      els.apiUrl.addEventListener('change', persistApiUrl);
      els.apiUrl.addEventListener('blur', persistApiUrl);
    }

    els.markerLayer.addEventListener('pointerdown', onMarkerLayerPointerDown);
    els.markerLayer.addEventListener('click', onMarkerLayerClick);
    document.addEventListener('pointermove', onGlobalPointerMove);
    document.addEventListener('pointerup', onGlobalPointerUp);
    document.addEventListener('pointercancel', onGlobalPointerUp);
    document.addEventListener('keydown', onKeyDown);
  }

  function fillFromQuery() {
    const incomingDefect = readQuery('defeito') || readQuery('defect');

    els.recordId.value = readQuery('id');
    els.partNumber.value = readQuery('pn');
    els.defectType.value = incomingDefect;
    els.sheetRow.value = readQuery('row');
    els.userName.value = readQuery('user');
    els.apiUrl.value = readQuery('api') || '';

    const matchedDefect = findDefect(incomingDefect);
    if (matchedDefect) {
      state.selectedDefectKey = matchedDefect.key;
    }
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

  function renderDefectPalette() {
    els.defectPalette.innerHTML = '';

    DEFECTS.forEach(function (defect) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'defect-option' + (defect.key === state.selectedDefectKey ? ' is-selected' : '');
      button.dataset.defectKey = defect.key;
      button.style.setProperty('--defect-color', defect.color);
      button.innerHTML =
        '<span class="defect-icon">' + buildDefectIcon(defect) + '</span>' +
        '<span>' +
          '<span class="defect-option-name">' + defect.label + '</span>' +
          '<span class="defect-option-code">' + defect.short + '</span>' +
        '</span>';

      button.addEventListener('click', function () {
        selectDefect(defect.key);
      });

      els.defectPalette.appendChild(button);
    });
  }

  function selectDefect(defectKey) {
    if (!DEFECTS_BY_KEY[defectKey]) {
      return;
    }

    state.selectedDefectKey = defectKey;
    syncSelectedDefectCard();
    syncPaletteSelection();
    setStatus('Defeito "' + DEFECTS_BY_KEY[defectKey].label + '" selecionado para marcacao.', 'success');
  }

  function syncSelectedDefectCard() {
    const defect = getSelectedDefect();

    els.selectedDefectName.textContent = defect.label;
    els.selectedDefectHint.textContent = defect.description;
    els.selectedDefectIcon.innerHTML = buildDefectIcon(defect);
    els.selectedDefectCard.style.setProperty('--defect-color', defect.color);
  }

  function syncPaletteSelection() {
    const buttons = els.defectPalette.querySelectorAll('.defect-option');

    buttons.forEach(function (button) {
      button.classList.toggle('is-selected', button.dataset.defectKey === state.selectedDefectKey);
    });
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
      mode === 'loading' ? 'Carregando' :
      mode === 'warning' ? 'Atencao' :
      'Pronto';
  }

  function resetCanvasStage() {
    state.imageLoaded = false;
    state.loadingImage = false;
    state.renderBox = {
      x: 0,
      y: 0,
      width: STANDARD_IMAGE_WIDTH,
      height: STANDARD_IMAGE_HEIGHT
    };

    els.baseImage.removeAttribute('src');
    els.baseImage.style.display = 'none';
    els.markerLayer.style.display = 'none';
    els.canvasEmpty.style.display = 'grid';

    clearMarkersState(true);
    applyRenderBoxStyles();
  }

  function applyRenderBoxStyles() {
    applyBoxStyle(els.baseImage, state.renderBox, STANDARD_IMAGE_WIDTH, STANDARD_IMAGE_HEIGHT);
    applyBoxStyle(els.markerLayer, state.renderBox, STANDARD_IMAGE_WIDTH, STANDARD_IMAGE_HEIGHT);
  }

  function applyBoxStyle(element, box, totalWidth, totalHeight) {
    element.style.left = box.x / totalWidth * 100 + '%';
    element.style.top = box.y / totalHeight * 100 + '%';
    element.style.width = box.width / totalWidth * 100 + '%';
    element.style.height = box.height / totalHeight * 100 + '%';
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

      const incomingMarkers = resolveIncomingMarkers(payload);
      restoreMarkers(incomingMarkers, incomingMarkers !== '');

      els.canvasSubtitle.textContent = 'PN ' + pn + ' carregado para marcacao padronizada.';
      setStatus(
        settings.auto
          ? 'Imagem carregada automaticamente. Selecione e aplique os marcadores.'
          : 'Imagem carregada. Selecione e aplique os marcadores.',
        'success'
      );
    } catch (error) {
      resetCanvasStage();
      setStatus('Falha ao buscar a imagem: ' + error.message, 'error');
    } finally {
      state.loadingImage = false;
    }
  }

  function drawBaseImage(imageUrl) {
    return new Promise(function (resolve, reject) {
      const image = new Image();
      image.crossOrigin = 'anonymous';

      image.onload = function () {
        state.renderBox = calculateContainBox(
          image.naturalWidth,
          image.naturalHeight,
          STANDARD_IMAGE_WIDTH,
          STANDARD_IMAGE_HEIGHT
        );

        els.baseImage.src = imageUrl;
        els.baseImage.style.display = 'block';
        els.markerLayer.style.display = 'block';
        els.canvasEmpty.style.display = 'none';

        applyRenderBoxStyles();
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
  function resolveIncomingMarkers(payload) {
    return (
      payload.markers ||
      payload.annotationMarkers ||
      (payload.annotation && payload.annotation.markers) ||
      (payload.annotation && payload.annotation.items) ||
      ''
    );
  }

  function restoreMarkers(incomingMarkers, preferRemote) {
    const remoteMarkers = normalizeMarkersInput(incomingMarkers);
    const localMarkers = loadMarkersFromLocalStorage();
    const markersToUse = preferRemote ? remoteMarkers : (remoteMarkers.length ? remoteMarkers : localMarkers);

    state.markers = markersToUse;
    state.selectedMarkerId = markersToUse.length ? markersToUse[markersToUse.length - 1].id : '';

    renderMarkers();
    updateMarkerSummary();

    if (remoteMarkers.length) {
      persistMarkersLocally();
    }
  }

  function onMarkerLayerPointerDown(event) {
    const markerElement = event.target.closest('.marker-item');
    if (!markerElement) {
      return;
    }

    if (event.target.closest('.marker-remove')) {
      return;
    }

    const markerId = markerElement.dataset.markerId;
    const marker = getMarkerById(markerId);
    if (!marker) {
      return;
    }

    const action = event.target.closest('.marker-resize') ? 'resize' : 'drag';

    event.preventDefault();
    selectMarker(markerId);
    startPointerInteraction(marker, markerElement, action, event);
  }

  function onMarkerLayerClick(event) {
    const removeButton = event.target.closest('.marker-remove');
    if (removeButton) {
      event.preventDefault();
      removeMarkerById(removeButton.dataset.markerId);
      return;
    }

    const markerElement = event.target.closest('.marker-item');
    if (markerElement) {
      selectMarker(markerElement.dataset.markerId);
      return;
    }

    if (!state.imageLoaded) {
      return;
    }

    if (Date.now() < state.preventPlacementUntil) {
      return;
    }

    placeMarkerFromEvent(event);
  }

  function startPointerInteraction(marker, element, action, event) {
    state.activePointer = {
      pointerId: event.pointerId,
      markerId: marker.id,
      action: action,
      startClientX: event.clientX,
      startClientY: event.clientY,
      originX: marker.x,
      originY: marker.y,
      originWidth: marker.width,
      originHeight: marker.height,
      layerRect: els.markerLayer.getBoundingClientRect(),
      moved: false,
      element: element
    };

    if (typeof els.markerLayer.setPointerCapture === 'function') {
      try {
        els.markerLayer.setPointerCapture(event.pointerId);
      } catch (captureError) {
        // Ignore capture failures and continue with document listeners.
      }
    }
  }

  function onGlobalPointerMove(event) {
    if (!state.activePointer || event.pointerId !== state.activePointer.pointerId) {
      return;
    }

    const active = state.activePointer;
    const marker = getMarkerById(active.markerId);
    if (!marker) {
      return;
    }

    const dx = (event.clientX - active.startClientX) / active.layerRect.width;
    const dy = (event.clientY - active.startClientY) / active.layerRect.height;

    active.moved = active.moved || Math.abs(dx) > 0.002 || Math.abs(dy) > 0.002;

    if (active.action === 'drag') {
      marker.x = clamp(active.originX + dx, 0, 1 - marker.width);
      marker.y = clamp(active.originY + dy, 0, 1 - marker.height);
    } else {
      marker.width = clamp(active.originWidth + dx, MIN_MARKER_WIDTH, 1 - marker.x);
      marker.height = clamp(active.originHeight + dy, MIN_MARKER_HEIGHT, 1 - marker.y);
    }

    applyMarkerElementStyle(active.element, marker, DEFECTS_BY_KEY[marker.defectKey]);
  }

  function onGlobalPointerUp(event) {
    if (!state.activePointer || event.pointerId !== state.activePointer.pointerId) {
      return;
    }

    if (
      typeof els.markerLayer.hasPointerCapture === 'function' &&
      els.markerLayer.hasPointerCapture(event.pointerId)
    ) {
      els.markerLayer.releasePointerCapture(event.pointerId);
    }

    if (state.activePointer.moved) {
      state.preventPlacementUntil = Date.now() + 150;
      persistMarkersLocally();
    }

    state.activePointer = null;
  }

  function onKeyDown(event) {
    if (event.key === 'Escape') {
      selectMarker('');
      return;
    }

    if (event.key !== 'Delete' && event.key !== 'Backspace') {
      return;
    }

    const targetTag = event.target && event.target.tagName ? event.target.tagName.toLowerCase() : '';
    if (targetTag === 'input' || targetTag === 'textarea') {
      return;
    }

    if (state.selectedMarkerId) {
      event.preventDefault();
      removeMarkerById(state.selectedMarkerId);
    }
  }

  function placeMarkerFromEvent(event) {
    const defect = getSelectedDefect();
    const rect = els.markerLayer.getBoundingClientRect();
    const pointX = (event.clientX - rect.left) / rect.width;
    const pointY = (event.clientY - rect.top) / rect.height;

    if (pointX < 0 || pointX > 1 || pointY < 0 || pointY > 1) {
      return;
    }

    const width = defect.defaultWidth;
    const height = defect.defaultHeight;
    const marker = {
      id: createMarkerId(),
      defectKey: defect.key,
      x: clamp(pointX - width / 2, 0, 1 - width),
      y: clamp(pointY - height / 2, 0, 1 - height),
      width: width,
      height: height
    };

    state.markers.push(marker);
    state.selectedMarkerId = marker.id;

    renderMarkers();
    updateMarkerSummary();
    persistMarkersLocally();
    flashPlacement(pointX, pointY, defect.color);
    setStatus('Marcador "' + defect.label + '" adicionado.', 'success');
  }

  function flashPlacement(x, y, color) {
    const flash = document.createElement('div');
    flash.className = 'marker-flash';
    flash.style.left = x * 100 + '%';
    flash.style.top = y * 100 + '%';
    flash.style.setProperty('--flash-color', color);

    els.markerLayer.appendChild(flash);
    window.setTimeout(function () {
      if (flash.parentNode) {
        flash.parentNode.removeChild(flash);
      }
    }, 450);
  }

  function renderMarkers() {
    els.markerLayer.innerHTML = '';

    state.markers.forEach(function (marker) {
      const defect = DEFECTS_BY_KEY[marker.defectKey];
      if (!defect) {
        return;
      }

      const element = document.createElement('div');
      element.className = 'marker-item' + (marker.id === state.selectedMarkerId ? ' is-selected' : '');
      element.dataset.markerId = marker.id;

      applyMarkerElementStyle(element, marker, defect);

      element.innerHTML =
        '<div class="marker-body" data-action="drag">' +
          '<div class="marker-symbol">' + buildDefectIcon(defect) + '</div>' +
          '<span class="marker-code">' + defect.short + '</span>' +
        '</div>' +
        '<button class="marker-remove" type="button" data-marker-id="' + marker.id + '" aria-label="Remover marcador">x</button>' +
        '<div class="marker-resize" data-action="resize" aria-hidden="true"></div>';

      els.markerLayer.appendChild(element);
    });
  }

  function applyMarkerElementStyle(element, marker, defect) {
    element.style.left = marker.x * 100 + '%';
    element.style.top = marker.y * 100 + '%';
    element.style.width = marker.width * 100 + '%';
    element.style.height = marker.height * 100 + '%';
    element.style.setProperty('--marker-color', defect.color);
    element.style.setProperty('--marker-soft', hexToRgba(defect.color, 0.18));
    element.style.setProperty('--marker-glow', hexToRgba(defect.color, 0.38));
  }

  function selectMarker(markerId) {
    state.selectedMarkerId = markerId || '';

    const elements = els.markerLayer.querySelectorAll('.marker-item');
    elements.forEach(function (element) {
      element.classList.toggle('is-selected', element.dataset.markerId === state.selectedMarkerId);
    });
  }

  function clearMarkers() {
    if (!state.imageLoaded) {
      setStatus('Carregue uma imagem antes de limpar os marcadores.', 'warning');
      return;
    }

    clearMarkersState();
    setStatus('Todos os marcadores foram removidos.', 'success');
  }

  function clearMarkersState(skipPersist) {
    state.markers = [];
    state.selectedMarkerId = '';
    renderMarkers();
    updateMarkerSummary();

    if (!skipPersist) {
      persistMarkersLocally();
    }
  }

  function removeSelectedMarker() {
    if (!state.selectedMarkerId) {
      setStatus('Selecione um marcador para remover.', 'warning');
      return;
    }

    removeMarkerById(state.selectedMarkerId);
  }

  function removeMarkerById(markerId) {
    const nextMarkers = state.markers.filter(function (marker) {
      return marker.id !== markerId;
    });

    if (nextMarkers.length === state.markers.length) {
      return;
    }

    state.markers = nextMarkers;
    state.selectedMarkerId = nextMarkers.length ? nextMarkers[nextMarkers.length - 1].id : '';

    renderMarkers();
    updateMarkerSummary();
    persistMarkersLocally();
    setStatus('Marcador removido.', 'success');
  }

  function updateMarkerSummary() {
    const count = state.markers.length;
    els.markerCount.textContent = count + (count === 1 ? ' marcador' : ' marcadores');
  }

  function normalizeMarkersInput(input) {
    if (!input) {
      return [];
    }

    let source = input;
    if (typeof source === 'string') {
      try {
        source = JSON.parse(source);
      } catch (parseError) {
        return [];
      }
    }

    if (!Array.isArray(source) && source && Array.isArray(source.markers)) {
      source = source.markers;
    }

    if (!Array.isArray(source)) {
      return [];
    }

    return source.map(function (item) {
      const defect = findDefect(item.defectKey || item.type || item.defect || item.label);
      if (!defect) {
        return null;
      }

      const width = clamp(normalizePercent(item.width != null ? item.width : item.w, defect.defaultWidth), MIN_MARKER_WIDTH, 1);
      const height = clamp(normalizePercent(item.height != null ? item.height : item.h, defect.defaultHeight), MIN_MARKER_HEIGHT, 1);

      return {
        id: String(item.id || createMarkerId()),
        defectKey: defect.key,
        x: clamp(normalizePercent(item.x != null ? item.x : item.left, 0), 0, 1 - width),
        y: clamp(normalizePercent(item.y != null ? item.y : item.top, 0), 0, 1 - height),
        width: width,
        height: height
      };
    }).filter(Boolean);
  }

  function normalizePercent(value, fallback) {
    if (value == null || value === '') {
      return fallback;
    }

    const numeric = Number(String(value).replace(',', '.'));
    if (!Number.isFinite(numeric)) {
      return fallback;
    }

    if (numeric > 1) {
      return numeric / 100;
    }

    return numeric;
  }

  function loadMarkersFromLocalStorage() {
    const key = getMarkerStorageKey();
    if (!key) {
      return [];
    }

    try {
      const stored = window.localStorage.getItem(key);
      if (!stored) {
        return [];
      }

      const payload = JSON.parse(stored);
      return normalizeMarkersInput(payload.markers || payload);
    } catch (error) {
      return [];
    }
  }

  function persistMarkersLocally() {
    const key = getMarkerStorageKey();
    if (!key) {
      return;
    }

    window.localStorage.setItem(key, JSON.stringify({
      recordId: els.recordId.value.trim(),
      partNumber: els.partNumber.value.trim(),
      markers: serializeMarkers()
    }));
  }

  function getMarkerStorageKey() {
    const id = els.recordId.value.trim();
    const pn = els.partNumber.value.trim();

    if (!id && !pn) {
      return '';
    }

    return MARKER_STORAGE_PREFIX + '::' + id + '::' + pn;
  }

  function serializeMarkers() {
    const recordId = els.recordId.value.trim();
    const partNumber = els.partNumber.value.trim();

    return state.markers.map(function (marker) {
      const defect = DEFECTS_BY_KEY[marker.defectKey];
      return {
        id: marker.id,
        defectKey: marker.defectKey,
        type: defect ? defect.label : marker.defectKey,
        x: round(marker.x),
        y: round(marker.y),
        width: round(marker.width),
        height: round(marker.height),
        recordId: recordId,
        partNumber: partNumber
      };
    });
  }
  async function saveAnnotation() {
    const apiUrl = getApiUrl();
    const id = els.recordId.value.trim();
    const pn = els.partNumber.value.trim();
    const recordDefect = els.defectType.value.trim();
    const user = els.userName.value.trim();

    if (!apiUrl) {
      setStatus('Preencha a URL da API antes de salvar.', 'warning');
      return;
    }

    if (!id || !pn) {
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

    if (!state.markers.length) {
      setStatus('Insira ao menos um marcador antes de salvar.', 'warning');
      return;
    }

    setStatus('Gerando imagem final e enviando os marcadores para a API...', 'loading');

    try {
      const markers = serializeMarkers();
      const payload = {
        action: 'saveAnnotation',
        id: id,
        pn: pn,
        defect: recordDefect,
        selectedDefect: getSelectedDefect().label,
        row: els.sheetRow.value.trim(),
        user: user,
        markerCount: markers.length,
        markerSummary: summarizeMarkers(markers),
        markers: markers,
        markersJson: JSON.stringify(markers),
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

      persistMarkersLocally();
      setStatus('Marcacao salva com sucesso na linha ' + result.row + '.', 'success');
    } catch (error) {
      setStatus('Erro ao salvar a marcacao: ' + error.message, 'error');
    }
  }

  function summarizeMarkers(markers) {
    const summary = {};

    markers.forEach(function (marker) {
      summary[marker.type] = (summary[marker.type] || 0) + 1;
    });

    return Object.keys(summary).map(function (label) {
      return label + ': ' + summary[label];
    }).join(' | ');
  }

  function exportMergedImage() {
    const scale = 2;
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = STANDARD_IMAGE_WIDTH * scale;
    exportCanvas.height = STANDARD_IMAGE_HEIGHT * scale;

    const exportCtx = exportCanvas.getContext('2d');
    exportCtx.scale(scale, scale);
    exportCtx.imageSmoothingEnabled = true;
    exportCtx.imageSmoothingQuality = 'high';
    exportCtx.fillStyle = '#ffffff';
    exportCtx.fillRect(0, 0, STANDARD_IMAGE_WIDTH, STANDARD_IMAGE_HEIGHT);

    exportCtx.drawImage(
      els.baseImage,
      state.renderBox.x,
      state.renderBox.y,
      state.renderBox.width,
      state.renderBox.height
    );

    state.markers.forEach(function (marker) {
      const defect = DEFECTS_BY_KEY[marker.defectKey];
      if (!defect) {
        return;
      }

      drawMarkerOnCanvas(exportCtx, marker, defect);
    });

    return exportCanvas.toDataURL('image/png');
  }

  function drawMarkerOnCanvas(ctx, marker, defect) {
    const x = state.renderBox.x + marker.x * state.renderBox.width;
    const y = state.renderBox.y + marker.y * state.renderBox.height;
    const width = marker.width * state.renderBox.width;
    const height = marker.height * state.renderBox.height;
    const padding = Math.min(width, height) * 0.14;

    ctx.save();
    ctx.shadowColor = 'rgba(36, 23, 15, 0.18)';
    ctx.shadowBlur = 18;
    ctx.shadowOffsetY = 8;
    ctx.fillStyle = 'rgba(255, 255, 255, 0.92)';
    ctx.strokeStyle = defect.color;
    ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.06);
    drawRoundRect(ctx, x, y, width, height, Math.min(width, height) * 0.22);
    ctx.fill();
    ctx.stroke();
    ctx.restore();

    ctx.save();
    ctx.translate(x + padding, y + padding);
    drawCanvasIcon(ctx, defect, width - padding * 2, height - padding * 2);
    ctx.restore();

    drawCanvasLabel(ctx, defect, x, y, width, height);
  }

  function drawCanvasLabel(ctx, defect, x, y, width, height) {
    const badgeWidth = Math.max(42, width * 0.34);
    const badgeHeight = Math.max(18, height * 0.2);
    const badgeX = x + 8;
    const badgeY = y + height - badgeHeight - 8;

    ctx.save();
    ctx.fillStyle = 'rgba(255, 255, 255, 0.96)';
    ctx.strokeStyle = 'rgba(36, 23, 15, 0.08)';
    ctx.lineWidth = 1.5;
    drawRoundRect(ctx, badgeX, badgeY, badgeWidth, badgeHeight, badgeHeight / 2);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = defect.color;
    ctx.font = '700 ' + Math.max(11, badgeHeight * 0.55) + 'px Barlow, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(defect.short, badgeX + badgeWidth / 2, badgeY + badgeHeight / 2 + 0.5);
    ctx.restore();
  }

  function drawCanvasIcon(ctx, defect, width, height) {
    const color = defect.color;
    const fillSoft = hexToRgba(color, 0.18);

    ctx.save();
    ctx.strokeStyle = color;
    ctx.fillStyle = fillSoft;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';

    switch (defect.iconType) {
      case 'ring':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.12);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.28, 0, Math.PI * 2);
        ctx.stroke();
        break;
      case 'dot-ring':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.1);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.25, 0, Math.PI * 2);
        ctx.stroke();
        ctx.beginPath();
        ctx.fillStyle = color;
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.09, 0, Math.PI * 2);
        ctx.fill();
        break;
      case 'stretch':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.12);
        drawLine(ctx, width * 0.18, height * 0.5, width * 0.82, height * 0.5);
        drawArrowHead(ctx, width * 0.18, height * 0.5, -1, color, Math.min(width, height) * 0.14);
        drawArrowHead(ctx, width * 0.82, height * 0.5, 1, color, Math.min(width, height) * 0.14);
        break;
      case 'shard':
        ctx.lineWidth = Math.max(2, Math.min(width, height) * 0.08);
        drawPolygon(ctx, [[width * 0.22, height * 0.76], [width * 0.48, height * 0.18], [width * 0.78, height * 0.78]], true);
        drawPolygon(ctx, [[width * 0.62, height * 0.36], [width * 0.78, height * 0.18], [width * 0.86, height * 0.42]], true);
        break;
      case 'tool':
        ctx.lineWidth = Math.max(2.5, Math.min(width, height) * 0.08);
        drawRoundRect(ctx, width * 0.18, height * 0.24, width * 0.64, height * 0.52, Math.min(width, height) * 0.1);
        ctx.stroke();
        drawLine(ctx, width * 0.26, height * 0.66, width * 0.74, height * 0.34);
        break;
      case 'pore':
        ctx.lineWidth = Math.max(2.5, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.26, 0, Math.PI * 2);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.12, 0, Math.PI * 2);
        ctx.fill();
        break;
      case 'burr':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.moveTo(width * 0.15, height * 0.68);
        ctx.lineTo(width * 0.3, height * 0.36);
        ctx.lineTo(width * 0.42, height * 0.66);
        ctx.lineTo(width * 0.56, height * 0.34);
        ctx.lineTo(width * 0.7, height * 0.64);
        ctx.lineTo(width * 0.85, height * 0.3);
        ctx.stroke();
        break;
      case 'scratch':
        ctx.lineWidth = Math.max(2.2, Math.min(width, height) * 0.07);
        drawLine(ctx, width * 0.18, height * 0.75, width * 0.82, height * 0.25);
        break;
      case 'wrinkle':
        ctx.lineWidth = Math.max(2.8, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.moveTo(width * 0.12, height * 0.62);
        ctx.bezierCurveTo(width * 0.26, height * 0.12, width * 0.42, height * 0.9, width * 0.56, height * 0.38);
        ctx.bezierCurveTo(width * 0.68, height * 0.14, width * 0.8, height * 0.76, width * 0.88, height * 0.28);
        ctx.stroke();
        break;
      case 'crack':
        ctx.lineWidth = Math.max(3.2, Math.min(width, height) * 0.09);
        ctx.beginPath();
        ctx.moveTo(width * 0.16, height * 0.2);
        ctx.lineTo(width * 0.34, height * 0.46);
        ctx.lineTo(width * 0.44, height * 0.28);
        ctx.lineTo(width * 0.58, height * 0.66);
        ctx.lineTo(width * 0.72, height * 0.42);
        ctx.lineTo(width * 0.84, height * 0.76);
        ctx.stroke();
        break;
      case 'crease':
        ctx.lineWidth = Math.max(4, Math.min(width, height) * 0.11);
        drawLine(ctx, width * 0.18, height * 0.74, width * 0.82, height * 0.26);
        break;
      case 'overflow':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.moveTo(width * 0.18, height * 0.58);
        ctx.bezierCurveTo(width * 0.36, height * 0.18, width * 0.56, height * 0.82, width * 0.82, height * 0.38);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(width * 0.22, height * 0.76);
        ctx.bezierCurveTo(width * 0.42, height * 0.36, width * 0.62, height * 0.94, width * 0.86, height * 0.56);
        ctx.stroke();
        break;
      case 'corrosion':
        ctx.lineWidth = Math.max(2.6, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.18, 0, Math.PI * 2);
        ctx.fill();
        drawBurst(ctx, width / 2, height / 2, Math.min(width, height) * 0.34, 8, color);
        break;
      case 'ok':
        ctx.lineWidth = Math.max(2.6, Math.min(width, height) * 0.08);
        drawRoundRect(ctx, width * 0.16, height * 0.18, width * 0.68, height * 0.56, Math.min(width, height) * 0.1);
        ctx.stroke();
        ctx.beginPath();
        ctx.moveTo(width * 0.28, height * 0.48);
        ctx.lineTo(width * 0.44, height * 0.64);
        ctx.lineTo(width * 0.74, height * 0.3);
        ctx.stroke();
        break;
      case 'oil':
        ctx.lineWidth = Math.max(2.6, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.moveTo(width * 0.5, height * 0.14);
        ctx.bezierCurveTo(width * 0.3, height * 0.36, width * 0.22, height * 0.52, width * 0.22, height * 0.66);
        ctx.bezierCurveTo(width * 0.22, height * 0.84, width * 0.36, height * 0.9, width * 0.5, height * 0.9);
        ctx.bezierCurveTo(width * 0.64, height * 0.9, width * 0.78, height * 0.84, width * 0.78, height * 0.66);
        ctx.bezierCurveTo(width * 0.78, height * 0.52, width * 0.7, height * 0.36, width * 0.5, height * 0.14);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        break;
      case 'blocked-hole':
        ctx.lineWidth = Math.max(3, Math.min(width, height) * 0.09);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.26, 0, Math.PI * 2);
        ctx.stroke();
        drawLine(ctx, width * 0.28, height * 0.72, width * 0.72, height * 0.28);
        break;
      case 'peel':
        ctx.lineWidth = Math.max(2.5, Math.min(width, height) * 0.08);
        drawPolygon(ctx, [[width * 0.22, height * 0.18], [width * 0.7, height * 0.18], [width * 0.82, height * 0.32], [width * 0.82, height * 0.8], [width * 0.22, height * 0.8]], true);
        drawPolygon(ctx, [[width * 0.58, height * 0.18], [width * 0.82, height * 0.18], [width * 0.82, height * 0.42]], true);
        break;
      case 'inspection':
        ctx.lineWidth = Math.max(2.5, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.arc(width / 2, height * 0.46, Math.min(width, height) * 0.28, 0, Math.PI * 2);
        ctx.stroke();
        ctx.fillStyle = color;
        ctx.font = '700 ' + Math.max(12, height * 0.26) + 'px Barlow, sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('100', width / 2, height * 0.48);
        break;
      case 'suction':
        ctx.lineWidth = Math.max(2.5, Math.min(width, height) * 0.08);
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.28, 0, Math.PI * 2);
        ctx.stroke();
        ctx.beginPath();
        ctx.arc(width / 2, height / 2, Math.min(width, height) * 0.14, 0, Math.PI * 2);
        ctx.stroke();
        break;
    }

    ctx.restore();
  }

  function drawRoundRect(ctx, x, y, width, height, radius) {
    const clampedRadius = Math.min(radius, width / 2, height / 2);
    ctx.beginPath();
    ctx.moveTo(x + clampedRadius, y);
    ctx.lineTo(x + width - clampedRadius, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + clampedRadius);
    ctx.lineTo(x + width, y + height - clampedRadius);
    ctx.quadraticCurveTo(x + width, y + height, x + width - clampedRadius, y + height);
    ctx.lineTo(x + clampedRadius, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - clampedRadius);
    ctx.lineTo(x, y + clampedRadius);
    ctx.quadraticCurveTo(x, y, x + clampedRadius, y);
    ctx.closePath();
  }

  function drawLine(ctx, x1, y1, x2, y2) {
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  }

  function drawArrowHead(ctx, x, y, direction, color, size) {
    ctx.save();
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + direction * size, y - size * 0.7);
    ctx.lineTo(x + direction * size, y + size * 0.7);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  }

  function drawPolygon(ctx, points, fill) {
    ctx.beginPath();
    ctx.moveTo(points[0][0], points[0][1]);
    for (let index = 1; index < points.length; index += 1) {
      ctx.lineTo(points[index][0], points[index][1]);
    }
    ctx.closePath();
    if (fill) {
      ctx.fill();
    }
    ctx.stroke();
  }

  function drawBurst(ctx, x, y, radius, spikes, color) {
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = Math.max(2, radius * 0.16);
    for (let index = 0; index < spikes; index += 1) {
      const angle = Math.PI * 2 * index / spikes;
      drawLine(
        ctx,
        x + Math.cos(angle) * radius * 0.58,
        y + Math.sin(angle) * radius * 0.58,
        x + Math.cos(angle) * radius,
        y + Math.sin(angle) * radius
      );
    }
    ctx.restore();
  }
  function buildDefectIcon(defect) {
    const color = defect.color;
    const fillSoft = hexToRgba(color, 0.18);

    return '<svg viewBox="0 0 100 100" aria-hidden="true">' + getDefectSvgContent(defect, color, fillSoft) + '</svg>';
  }

  function getDefectSvgContent(defect, color, fillSoft) {
    switch (defect.iconType) {
      case 'ring':
        return '<circle cx="50" cy="50" r="28" fill="none" stroke="' + color + '" stroke-width="10" />';
      case 'dot-ring':
        return '<circle cx="50" cy="50" r="26" fill="none" stroke="' + color + '" stroke-width="9" /><circle cx="50" cy="50" r="9" fill="' + color + '" />';
      case 'stretch':
        return '<line x1="18" y1="50" x2="82" y2="50" stroke="' + color + '" stroke-width="10" stroke-linecap="round" /><polygon points="18,50 34,38 34,62" fill="' + color + '" /><polygon points="82,50 66,38 66,62" fill="' + color + '" />';
      case 'shard':
        return '<polygon points="24,78 48,18 78,78" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="6" /><polygon points="62,36 78,18 86,42" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="5" />';
      case 'tool':
        return '<rect x="18" y="24" width="64" height="52" rx="12" fill="none" stroke="' + color + '" stroke-width="7" /><line x1="28" y1="66" x2="72" y2="34" stroke="' + color + '" stroke-width="7" stroke-linecap="round" />';
      case 'pore':
        return '<circle cx="50" cy="50" r="26" fill="none" stroke="' + color + '" stroke-width="8" /><circle cx="50" cy="50" r="12" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="6" />';
      case 'burr':
        return '<polyline points="16,68 30,34 42,64 56,32 70,62 84,28" fill="none" stroke="' + color + '" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" />';
      case 'scratch':
        return '<line x1="20" y1="78" x2="80" y2="22" stroke="' + color + '" stroke-width="6" stroke-linecap="round" />';
      case 'wrinkle':
        return '<path d="M12 64 C26 14, 42 90, 56 38 S80 76, 88 28" fill="none" stroke="' + color + '" stroke-width="7" stroke-linecap="round" />';
      case 'crack':
        return '<polyline points="18,18 34,44 44,28 58,66 72,42 84,76" fill="none" stroke="' + color + '" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" />';
      case 'crease':
        return '<line x1="20" y1="76" x2="80" y2="24" stroke="' + color + '" stroke-width="10" stroke-linecap="round" />';
      case 'overflow':
        return '<path d="M18 58 C36 18, 56 82, 82 38" fill="none" stroke="' + color + '" stroke-width="7" stroke-linecap="round" /><path d="M22 76 C42 36, 62 94, 86 56" fill="none" stroke="' + color + '" stroke-width="7" stroke-linecap="round" />';
      case 'corrosion':
        return '<circle cx="50" cy="50" r="16" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="6" /><path d="M50 8 L54 24 M78 22 L66 34 M92 50 L76 50 M78 78 L66 66 M50 92 L54 76 M22 78 L34 66 M8 50 L24 50 M22 22 L34 34" fill="none" stroke="' + color + '" stroke-width="6" stroke-linecap="round" />';
      case 'ok':
        return '<rect x="16" y="18" width="68" height="56" rx="12" fill="none" stroke="' + color + '" stroke-width="7" /><path d="M28 50 L44 64 L74 30" fill="none" stroke="' + color + '" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" />';
      case 'oil':
        return '<path d="M50 14 C30 36 22 52 22 66 C22 84 36 90 50 90 C64 90 78 84 78 66 C78 52 70 36 50 14 Z" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="6" />';
      case 'blocked-hole':
        return '<circle cx="50" cy="50" r="26" fill="none" stroke="' + color + '" stroke-width="8" /><line x1="28" y1="72" x2="72" y2="28" stroke="' + color + '" stroke-width="8" stroke-linecap="round" />';
      case 'peel':
        return '<path d="M22 18 H70 L82 30 V80 H22 Z" fill="' + fillSoft + '" stroke="' + color + '" stroke-width="6" /><path d="M58 18 H82 V42 Z" fill="#ffffff" stroke="' + color + '" stroke-width="6" />';
      case 'inspection':
        return '<circle cx="50" cy="46" r="28" fill="none" stroke="' + color + '" stroke-width="7" /><text x="50" y="52" text-anchor="middle" font-size="24" font-weight="700" fill="' + color + '">100</text>';
      case 'suction':
        return '<circle cx="50" cy="50" r="28" fill="none" stroke="' + color + '" stroke-width="7" /><circle cx="50" cy="50" r="14" fill="none" stroke="' + color + '" stroke-width="7" />';
      default:
        return '<circle cx="50" cy="50" r="26" fill="none" stroke="' + color + '" stroke-width="8" />';
    }
  }

  function findDefect(value) {
    if (!value) {
      return null;
    }

    const normalizedValue = normalizeText(value);

    return DEFECTS.find(function (defect) {
      return (
        normalizeText(defect.key) === normalizedValue ||
        normalizeText(defect.label) === normalizedValue ||
        normalizeText(defect.short) === normalizedValue
      );
    }) || null;
  }

  function normalizeText(value) {
    return String(value)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  function getSelectedDefect() {
    return DEFECTS_BY_KEY[state.selectedDefectKey] || DEFECTS[0];
  }

  function getMarkerById(markerId) {
    return state.markers.find(function (marker) {
      return marker.id === markerId;
    }) || null;
  }

  function createMarkerId() {
    return 'marker_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function round(value) {
    return Number(value.toFixed(4));
  }

  function hexToRgba(hex, alpha) {
    const cleaned = String(hex).replace('#', '');
    const normalized = cleaned.length === 3
      ? cleaned.split('').map(function (char) { return char + char; }).join('')
      : cleaned;

    const red = parseInt(normalized.slice(0, 2), 16);
    const green = parseInt(normalized.slice(2, 4), 16);
    const blue = parseInt(normalized.slice(4, 6), 16);

    return 'rgba(' + red + ', ' + green + ', ' + blue + ', ' + alpha + ')';
  }

  init();
})();
