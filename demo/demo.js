/**
 * ============================================================================
 * UNIVERSAL TABLE MANAGER - IN-BROWSER SIMULATION & LIVE SYNC ENGINE
 * Demostración interactiva para GitHub Pages (Office.js Mock Engine)
 * ============================================================================
 */

(function () {
  'use strict';

  // ============================================================================
  // 1. BASE DE DATOS EN MEMORIA DE LAS TABLAS DEL LIBRO DE EXCEL
  // ============================================================================

  const TABLAS_LIBRO = {
    TablaVentas: {
      nombre: "TablaVentas",
      archivo: "Plantilla_Superpoderes_Excel_TablaPersonal.xlsx",
      encabezados: [
        "ID Venta", "Fecha", "Cliente", "Producto", 
        "Cantidad", "Precio Unitario", "Subtotal", "IVA (16%)", "Total Venta", "Vendedor"
      ],
      tipos: ["text", "date", "text", "text", "number", "number", "number", "number", "number", "text"],
      esCalculada: [false, false, false, false, false, false, true, true, true, false],
      formulas: [
        null, null, null, null, null, null, 
        "=[@Cantidad]*[@[Precio Unitario]]", "=[@Subtotal]*0.16", "=[@Subtotal]+[@[IVA (16%)]]", null
      ],
      opciones: {
        3: ["Licencia Office 365 E5", "Power BI Pro", "Azure Cloud Server", "Copilot Pro Studio", "Surface Laptop 5"],
        9: ["Mateo Navarro", "Valeria Guzmán", "Sebastián Cruz", "Isabella Delgado", "Alicia Paz"]
      },
      filas: [
        ["VTA-801", "2024-05-02", "TechCorp Soluciones", "Licencia Office 365 E5", 25, 420, 10500, 1680, 12180, "Mateo Navarro"],
        ["VTA-802", "2024-05-03", "Innovaciones Alfa", "Power BI Pro", 40, 110, 4400, 704, 5104, "Valeria Guzmán"],
        ["VTA-803", "2024-05-04", "Grupo Financiero Sur", "Azure Cloud Server", 5, 2800, 14000, 2240, 16240, "Isabella Delgado"],
        ["VTA-804", "2024-05-06", "Consultores Nexus", "Copilot Pro Studio", 15, 360, 5400, 864, 6264, "Alicia Paz"],
        ["VTA-805", "2024-05-07", "Logística Express", "Surface Laptop 5", 8, 1450, 11600, 1856, 13456, "Sebastián Cruz"],
        ["VTA-806", "2024-05-08", "BioSalud Farmacias", "Licencia Office 365 E5", 50, 420, 21000, 3360, 24360, "Isabella Delgado"],
        ["VTA-807", "2024-05-09", "Distribuidora Bajío", "Power BI Pro", 20, 110, 2200, 352, 2552, "Mateo Navarro"],
        ["VTA-808", "2024-05-10", "Manufacturas del Norte", "Azure Cloud Server", 3, 2800, 8400, 1344, 9744, "Valeria Guzmán"]
      ]
    },

    TablaInventario: {
      nombre: "TablaInventario",
      archivo: "Plantilla_Superpoderes_Excel_TablaPersonal.xlsx",
      encabezados: [
        "SKU", "Descripción", "Categoría", "Stock Actual", "Stock Mínimo", "Costo Unitario", "Valor Total", "Estado"
      ],
      tipos: ["text", "text", "text", "number", "number", "number", "number", "text"],
      esCalculada: [false, false, false, false, false, false, true, false],
      formulas: [null, null, null, null, null, null, "=[@[Stock Actual]]*[@Costo Unitario]", null],
      opciones: {
        2: ["Cómputo", "Periféricos", "Redes & Servidores", "Audio & Video"],
        7: ["Óptimo", "Alerta Mínimo", "Agotado", "Sobrestock"]
      },
      filas: [
        ["SKU-101", "Monitor Dell 27 4K UltraHD", "Periféricos", 48, 15, 380, 18240, "Óptimo"],
        ["SKU-102", "Teclado Mecánico Ergonómico MX", "Periféricos", 120, 30, 95, 11400, "Óptimo"],
        ["SKU-103", "Mouse Inalámbrico Master 3S", "Periféricos", 14, 25, 78, 1092, "Alerta Mínimo"],
        ["SKU-104", "Switch Gigabit Cisco 24 Puertos", "Redes & Servidores", 8, 5, 450, 3600, "Óptimo"],
        ["SKU-105", "Servidor Rack Dell PowerEdge 1U", "Redes & Servidores", 4, 3, 2600, 10400, "Óptimo"],
        ["SKU-106", "Auriculares Cancelación Ruido ANC", "Audio & Video", 0, 20, 160, 0, "Agotado"],
        ["SKU-107", "Webcam 4K Streaming Pro", "Audio & Video", 32, 10, 140, 4480, "Óptimo"]
      ]
    },

    TablaClientes: {
      nombre: "TablaClientes",
      archivo: "Plantilla_Superpoderes_Excel_TablaPersonal.xlsx",
      encabezados: [
        "ID Cliente", "Empresa", "Contacto", "Email", "Teléfono", "Ciudad", "Segmento", "Estado"
      ],
      tipos: ["text", "text", "text", "email", "text", "text", "text", "text"],
      esCalculada: [false, false, false, false, false, false, false, false],
      formulas: [null, null, null, null, null, null, null, null],
      opciones: {
        6: ["Enterprise", "Mid-Market", "PyME", "Gobierno"],
        7: ["Activo", "Prospecto", "Inactivo", "Lead Calificado"]
      },
      filas: [
        ["CLI-001", "TechCorp Soluciones SA", "Ing. Roberto Díaz", "rdiaz@techcorp.com", "+52 55 4123 8900", "Ciudad de México", "Enterprise", "Activo"],
        ["CLI-002", "Innovaciones Alfa SC", "Lic. Mónica Silva", "msilva@alfa-innova.com", "+52 33 2890 1234", "Guadalajara", "Mid-Market", "Activo"],
        ["CLI-003", "Grupo Financiero Sur", "Mtro. Arturo Vaca", "arturo.vaca@finsur.mx", "+52 81 9012 3456", "Monterrey", "Enterprise", "Activo"],
        ["CLI-004", "Consultores Nexus Global", "Dra. Carmen Lara", "clara@nexusglobal.org", "+52 44 2345 6789", "Querétaro", "PyME", "Lead Calificado"],
        ["CLI-005", "BioSalud Farmacéutica", "QFB. Andrés Paz", "andres.paz@biosalud.com", "+52 22 8912 3450", "Puebla", "Enterprise", "Activo"]
      ]
    }
  };

  // ============================================================================
  // 2. ESTADO GENERAL DE LA APLICACIÓN
  // ============================================================================

  const Estado = {
    tablaActiva: "TablaVentas",
    encabezados: [],
    filas: [],
    tipos: [],
    esCalculada: [],
    formulas: [],
    opciones: {},
    columnasVisibles: [],
    indicesFiltrados: [],
    busqueda: "",
    colOrden: null,
    dirOrden: "asc",
    filaSeleccionada: null,
    filaEnEdicion: null,
    filaAEliminar: null
  };

  // Referencias al DOM
  let el = {};

  function initDOM() {
    el = {
      // Demo Topbar & Window
      btnOpenVideoModal: document.getElementById('btnOpenVideoModal'),
      btnBannerVideo: document.getElementById('btnBannerVideo'),
      videoModalBackdrop: document.getElementById('videoModalBackdrop'),
      btnCloseVideoModal: document.getElementById('btnCloseVideoModal'),
      videoIframe: document.getElementById('videoIframe'),
      excelWindowTitle: document.getElementById('excelWindowTitle'),
      excelGridTable: document.getElementById('excelGridTable'),
      excelSheetTabsBar: document.getElementById('excelSheetTabsBar'),
      btnAddSheetTab: document.getElementById('btnAddSheetTab'),
      cellAddressBox: document.getElementById('cellAddressBox'),
      formulaInputBox: document.getElementById('formulaInputBox'),
      syncStatusText: document.getElementById('syncStatusText'),
      gridStatsBadge: document.getElementById('gridStatsBadge'),

      // Task Pane Elements (Universal Table Manager)
      selectTables: document.getElementById('selectTables'),
      btnRefreshTables: document.getElementById('btnRefreshTables'),
      btnMockRefresh: document.getElementById('btnMockRefresh'),
      inputSearch: document.getElementById('inputSearch'),
      btnClearSearch: document.getElementById('btnClearSearch'),
      btnNewRecord: document.getElementById('btnNewRecord'),
      counterText: document.getElementById('counterText'),
      activeTableBadge: document.getElementById('activeTableBadge'),
      recordsContainer: document.getElementById('recordsContainer'),

      // Column picker
      btnToggleColumnsPicker: document.getElementById('btnToggleColumnsPicker'),
      columnsDropdown: document.getElementById('columnsDropdown'),
      columnsPickerLabel: document.getElementById('columnsPickerLabel'),
      btnToggleAllColumns: document.getElementById('btnToggleAllColumns'),
      columnsListContainer: document.getElementById('columnsListContainer'),

      // Save View
      btnOpenSaveView: document.getElementById('btnOpenSaveView'),
      modalSaveView: document.getElementById('modalSaveView'),
      inputViewName: document.getElementById('inputViewName'),
      saveViewSummaryBox: document.getElementById('saveViewSummaryBox'),
      btnCloseModalSaveView: document.getElementById('btnCloseModalSaveView'),
      btnCancelSaveView: document.getElementById('btnCancelSaveView'),
      btnConfirmSaveView: document.getElementById('btnConfirmSaveView'),

      // Form Modal
      modalForm: document.getElementById('modalForm'),
      modalFormTitle: document.getElementById('modalFormTitle'),
      modalFormIcon: document.getElementById('modalFormIcon'),
      dynamicRecordForm: document.getElementById('dynamicRecordForm'),
      dynamicFormFields: document.getElementById('dynamicFormFields'),
      btnSaveText: document.getElementById('btnSaveText'),
      btnCloseModalForm: document.getElementById('btnCloseModalForm'),
      btnCancelForm: document.getElementById('btnCancelForm'),

      // Detail Modal
      modalDetail: document.getElementById('modalDetail'),
      modalDetailTitle: document.getElementById('modalDetailTitle'),
      detailContent: document.getElementById('detailContent'),
      btnCloseModalDetail: document.getElementById('btnCloseModalDetail'),
      btnCloseDetail: document.getElementById('btnCloseDetail'),
      btnEditFromDetail: document.getElementById('btnEditFromDetail'),

      // Confirm Delete Modal
      modalConfirm: document.getElementById('modalConfirm'),
      btnCancelDelete: document.getElementById('btnCancelDelete'),
      btnConfirmDelete: document.getElementById('btnConfirmDelete'),

      // Toast
      toastContainer: document.getElementById('toastContainer')
    };
  }

  // ============================================================================
  // 3. INICIALIZACIÓN
  // ============================================================================

  function iniciarDemo() {
    initDOM();
    vincularEventos();
    cargarTabla("TablaVentas");
  }

  function vincularEventos() {
    // Selector de tabla
    if (el.selectTables) {
      el.selectTables.addEventListener('change', (e) => {
        cargarTabla(e.target.value);
      });
    }

    // Refrescar tablas
    const onRefresh = () => {
      mostrarToast("Tablas detectadas en el libro (3 tablas oficiales)", "info");
      cargarTabla(Estado.tablaActiva);
    };
    if (el.btnRefreshTables) el.btnRefreshTables.addEventListener('click', onRefresh);
    if (el.btnMockRefresh) el.btnMockRefresh.addEventListener('click', onRefresh);

    // Buscador interactivo
    if (el.inputSearch) {
      el.inputSearch.addEventListener('input', (e) => {
        Estado.busqueda = e.target.value.trim().toLowerCase();
        if (el.btnClearSearch) {
          el.btnClearSearch.classList.toggle('hidden', Estado.busqueda.length === 0);
        }
        filtrarYRenderizar();
      });
    }

    if (el.btnClearSearch) {
      el.btnClearSearch.addEventListener('click', () => {
        el.inputSearch.value = "";
        Estado.busqueda = "";
        el.btnClearSearch.classList.add('hidden');
        filtrarYRenderizar();
      });
    }

    // Botón Nuevo
    if (el.btnNewRecord) {
      el.btnNewRecord.addEventListener('click', abrirModalNuevo);
    }

    // Cerrar modales
    if (el.btnCloseModalForm) el.btnCloseModalForm.addEventListener('click', cerrarModalForm);
    if (el.btnCancelForm) el.btnCancelForm.addEventListener('click', cerrarModalForm);
    if (el.btnCloseModalDetail) el.btnCloseModalDetail.addEventListener('click', cerrarModalDetail);
    if (el.btnCloseDetail) el.btnCloseDetail.addEventListener('click', cerrarModalDetail);
    if (el.btnCancelDelete) el.btnCancelDelete.addEventListener('click', cerrarModalConfirm);
    if (el.btnConfirmDelete) el.btnConfirmDelete.addEventListener('click', ejecutarEliminar);

    // Editar desde modal de detalle
    if (el.btnEditFromDetail) {
      el.btnEditFromDetail.addEventListener('click', () => {
        const idx = Estado.filaSeleccionada;
        cerrarModalDetail();
        if (idx !== null) abrirModalEditar(idx);
      });
    }

    // Guardar formulario
    if (el.dynamicRecordForm) {
      el.dynamicRecordForm.addEventListener('submit', guardarRegistroForm);
    }

    // Popover de Columnas
    if (el.btnToggleColumnsPicker) {
      el.btnToggleColumnsPicker.addEventListener('click', (e) => {
        e.stopPropagation();
        el.columnsDropdown?.classList.toggle('hidden');
      });
    }

    if (el.btnToggleAllColumns) {
      el.btnToggleAllColumns.addEventListener('click', (e) => {
        e.stopPropagation();
        alternarTodasLasColumnas();
      });
    }

    document.addEventListener('click', (e) => {
      if (!e.target.closest('.columns-picker-wrapper')) {
        el.columnsDropdown?.classList.add('hidden');
      }
    });

    // Guardar como Vista
    if (el.btnOpenSaveView) el.btnOpenSaveView.addEventListener('click', abrirModalSaveView);
    if (el.btnCloseModalSaveView) el.btnCloseModalSaveView.addEventListener('click', cerrarModalSaveView);
    if (el.btnCancelSaveView) el.btnCancelSaveView.addEventListener('click', cerrarModalSaveView);
    if (el.btnConfirmSaveView) el.btnConfirmSaveView.addEventListener('click', ejecutarGuardarVista);

    // Modal de Video YouTube
    const openVideo = () => {
      el.videoModalBackdrop?.classList.remove('hidden');
    };
    const closeVideo = () => {
      el.videoModalBackdrop?.classList.add('hidden');
      // Pausar video reseteando src
      if (el.videoIframe) {
        const curSrc = el.videoIframe.src;
        el.videoIframe.src = curSrc;
      }
    };
    if (el.btnOpenVideoModal) el.btnOpenVideoModal.addEventListener('click', openVideo);
    if (el.btnBannerVideo) el.btnBannerVideo.addEventListener('click', openVideo);
    if (el.btnCloseVideoModal) el.btnCloseVideoModal.addEventListener('click', closeVideo);
    if (el.videoModalBackdrop) {
      el.videoModalBackdrop.addEventListener('click', (e) => {
        if (e.target === el.videoModalBackdrop) closeVideo();
      });
    }

    // Botón agregar pestaña en Excel
    if (el.btnAddSheetTab) {
      el.btnAddSheetTab.addEventListener('click', () => {
        abrirModalSaveView();
      });
    }

    // Escuchar Escape para cerrar modales
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        cerrarModalForm();
        cerrarModalDetail();
        cerrarModalConfirm();
        cerrarModalSaveView();
        closeVideo();
      }
    });
  }

  // ============================================================================
  // 4. GESTIÓN DE TABLAS Y SINCRONIZACIÓN CON EXCEL
  // ============================================================================

  function cargarTabla(nombreTabla) {
    const info = TABLAS_LIBRO[nombreTabla];
    if (!info) return;

    Estado.tablaActiva = nombreTabla;
    Estado.encabezados = [...info.encabezados];
    Estado.filas = info.filas.map(r => [...r]);
    Estado.tipos = [...info.tipos];
    Estado.esCalculada = [...info.esCalculada];
    Estado.formulas = [...info.formulas];
    Estado.opciones = info.opciones || {};
    Estado.columnasVisibles = Estado.encabezados.map((_, i) => i);
    Estado.colOrden = null;
    Estado.dirOrden = "asc";
    Estado.filaSeleccionada = null;

    if (el.selectTables) el.selectTables.value = nombreTabla;
    if (el.activeTableBadge) el.activeTableBadge.textContent = nombreTabla;
    if (el.excelWindowTitle) el.excelWindowTitle.textContent = `Plantilla_Superpoderes_Excel_${nombreTabla}.xlsx - Excel`;

    // Actualizar pestañas inferiores del mockup
    actualizarPestañasHojasExcel(nombreTabla);

    renderizarSelectorColumnas();
    filtrarYRenderizar();
    renderizarCuadriculaExcel();

    mostrarToast(`Tabla "${nombreTabla}" cargada (${Estado.filas.length} registros)`, "info");
  }

  function actualizarPestañasHojasExcel(nombreActivo) {
    if (!el.excelSheetTabsBar) return;
    const tabs = el.excelSheetTabsBar.querySelectorAll('.sheet-tab');
    tabs.forEach(tab => {
      const tName = tab.getAttribute('data-table');
      if (tName === nombreActivo) {
        tab.classList.add('active');
      } else {
        tab.classList.remove('active');
      }
      tab.onclick = () => {
        cargarTabla(tName);
      };
    });
  }

  // ============================================================================
  // 5. RENDERIZADO DE LA HOJA DE CÁLCULO DE EXCEL (PANEL IZQUIERDO)
  // ============================================================================

  function obtenerLetraColumna(index) {
    let letra = "";
    while (index >= 0) {
      letra = String.fromCharCode((index % 26) + 65) + letra;
      index = Math.floor(index / 26) - 1;
    }
    return letra;
  }

  function renderizarCuadriculaExcel(filaDestacadaIndex = null) {
    if (!el.excelGridTable) return;
    el.excelGridTable.innerHTML = "";

    const totalCols = Estado.encabezados.length;

    // 1. Fila de Encabezados de Columna de Excel (A, B, C...)
    const trLetters = document.createElement('tr');
    trLetters.className = 'col-header-row';

    const thEmptyCorner = document.createElement('th');
    thEmptyCorner.className = 'row-index-cell';
    thEmptyCorner.textContent = '';
    trLetters.appendChild(thEmptyCorner);

    for (let c = 0; c < totalCols; c++) {
      const thLetter = document.createElement('th');
      thLetter.textContent = obtenerLetraColumna(c);
      trLetters.appendChild(thLetter);
    }
    el.excelGridTable.appendChild(trLetters);

    // 2. Fila 1: Encabezados Oficiales de la Tabla (con flecha de autofiltro)
    const trTableHeader = document.createElement('tr');
    trTableHeader.className = 'table-header-row';

    const tdRowIndex1 = document.createElement('td');
    tdRowIndex1.className = 'row-index-cell';
    tdRowIndex1.textContent = '1';
    trTableHeader.appendChild(tdRowIndex1);

    Estado.encabezados.forEach((headerName, colIdx) => {
      const th = document.createElement('th');
      th.innerHTML = `
        <div class="th-content">
          <span>${headerName}</span>
          <span class="filter-arrow">▾</span>
        </div>
      `;
      th.addEventListener('click', () => {
        actualizarBarraFormulas(obtenerLetraColumna(colIdx) + '1', headerName);
      });
      trTableHeader.appendChild(th);
    });
    el.excelGridTable.appendChild(trTableHeader);

    // 3. Filas de Datos
    Estado.filas.forEach((fila, rIdx) => {
      const excelRowNumber = rIdx + 2;
      const tr = document.createElement('tr');
      tr.className = 'excel-data-row';
      tr.dataset.rowIndex = rIdx;

      if (filaDestacadaIndex === rIdx) {
        tr.classList.add('new-row-animation');
      }

      // Si hay búsqueda activa y esta fila coincide
      if (Estado.busqueda && Estado.indicesFiltrados.includes(rIdx)) {
        tr.classList.add('highlight-match');
      }

      // Número de fila en la cuadrícula
      const tdRowIndex = document.createElement('td');
      tdRowIndex.className = 'row-index-cell';
      tdRowIndex.textContent = String(excelRowNumber);
      tr.appendChild(tdRowIndex);

      fila.forEach((val, cIdx) => {
        const td = document.createElement('td');
        const esFormula = Estado.esCalculada[cIdx];
        const formulaTexto = Estado.formulas[cIdx];
        const cellCoord = obtenerLetraColumna(cIdx) + excelRowNumber;

        if (esFormula) {
          td.classList.add('formula-cell');
        }

        // Formato visual de celda
        if (Estado.tipos[cIdx] === 'number' && typeof val === 'number') {
          td.style.textAlign = 'right';
          td.textContent = val.toLocaleString('es-MX');
        } else {
          td.textContent = val !== null && val !== undefined ? String(val) : '';
        }

        // Evento al hacer clic en cualquier celda de Excel
        td.addEventListener('click', (e) => {
          e.stopPropagation();
          // Remover selección previa en Excel
          el.excelGridTable.querySelectorAll('td').forEach(c => c.style.outline = '');
          td.style.outline = '2px solid var(--excel-selected-cell)';
          td.style.outlineOffset = '-2px';

          // Actualizar barra de fórmulas de Excel
          const formulaValue = esFormula ? formulaTexto : (val !== null && val !== undefined ? String(val) : '');
          actualizarBarraFormulas(cellCoord, formulaValue);

          // También seleccionar la fila correspondiente en el Task Pane
          seleccionarFilaEnTaskPane(rIdx);
        });

        tr.appendChild(td);
      });

      el.excelGridTable.appendChild(tr);
    });

    // Actualizar badge de estado
    if (el.gridStatsBadge) {
      el.gridStatsBadge.textContent = `${Estado.filas.length} filas • Fórmulas automáticas activas`;
    }
  }

  function actualizarBarraFormulas(direccion, contenido) {
    if (el.cellAddressBox) el.cellAddressBox.textContent = direccion;
    if (el.formulaInputBox) el.formulaInputBox.value = contenido;
  }

  function seleccionarFilaEnTaskPane(rowIndex) {
    Estado.filaSeleccionada = rowIndex;
    if (!el.recordsContainer) return;
    const rows = el.recordsContainer.querySelectorAll('tbody tr');
    rows.forEach(r => {
      const dataIdx = Number(r.getAttribute('data-index'));
      if (dataIdx === rowIndex) {
        r.classList.add('selected-row');
        r.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      } else {
        r.classList.remove('selected-row');
      }
    });
  }

  // ============================================================================
  // 6. MOTOR DE BÚSQUEDA, ORDENAMIENTO Y FILTRADO (TASK PANE)
  // ============================================================================

  function filtrarYRenderizar() {
    const q = Estado.busqueda;

    // 1. Filtrar
    if (!q) {
      Estado.indicesFiltrados = Estado.filas.map((_, i) => i);
    } else {
      Estado.indicesFiltrados = [];
      Estado.filas.forEach((fila, idx) => {
        const coincide = fila.some((val) => {
          if (val === null || val === undefined) return false;
          return String(val).toLowerCase().includes(q);
        });
        if (coincide) Estado.indicesFiltrados.push(idx);
      });
    }

    // 2. Ordenar
    if (Estado.colOrden !== null && Estado.indicesFiltrados.length > 1) {
      const c = Estado.colOrden;
      const dir = Estado.dirOrden === 'asc' ? 1 : -1;
      const tipo = Estado.tipos[c];

      Estado.indicesFiltrados.sort((a, b) => {
        const valA = Estado.filas[a][c];
        const valB = Estado.filas[b][c];

        if (valA === valB) return 0;
        if (valA === null || valA === undefined || valA === '') return 1;
        if (valB === null || valB === undefined || valB === '') return -1;

        if (tipo === 'number') {
          return (Number(valA) - Number(valB)) * dir;
        }

        if (tipo === 'date') {
          return (new Date(valA).getTime() - new Date(valB).getTime()) * dir;
        }

        return String(valA).localeCompare(String(valB), 'es', { numeric: true, sensitivity: 'base' }) * dir;
      });
    }

    renderizarVistaTabularTaskPane();
    actualizarContador();
    sincronizarResaltadoEnExcel();
  }

  function alternarOrden(colIdx) {
    if (Estado.colOrden === colIdx) {
      Estado.dirOrden = Estado.dirOrden === 'asc' ? 'desc' : 'asc';
    } else {
      Estado.colOrden = colIdx;
      Estado.dirOrden = 'asc';
    }
    filtrarYRenderizar();
  }

  function actualizarContador() {
    if (!el.counterText) return;
    const total = Estado.filas.length;
    const filtrados = Estado.indicesFiltrados.length;
    if (Estado.busqueda) {
      el.counterText.textContent = `${filtrados} de ${total} registros`;
    } else {
      el.counterText.textContent = `${total} registros`;
    }
  }

  function sincronizarResaltadoEnExcel() {
    if (!el.excelGridTable) return;
    const excelRows = el.excelGridTable.querySelectorAll('.excel-data-row');
    excelRows.forEach(tr => {
      const rIdx = Number(tr.dataset.rowIndex);
      if (Estado.busqueda && Estado.indicesFiltrados.includes(rIdx)) {
        tr.classList.add('highlight-match');
      } else {
        tr.classList.remove('highlight-match');
      }
    });
  }

  function resaltarCoincidencia(texto, query) {
    if (!query || texto === null || texto === undefined || texto === '') return String(texto || '');
    const str = String(texto);
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`(${escaped})`, 'gi');
    return str.replace(regex, '<mark class="search-highlight">$1</mark>');
  }

  // ============================================================================
  // 7. RENDERIZADO DEL TASK PANE (VISTA TABULAR)
  // ============================================================================

  function renderizarVistaTabularTaskPane() {
    if (!el.recordsContainer) return;

    if (Estado.filas.length === 0) {
      el.recordsContainer.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">📭</div>
          <p class="empty-title">Sin registros</p>
          <p class="empty-desc">Esta tabla está vacía en Excel. Usa el botón "Nuevo" para agregar uno.</p>
        </div>
      `;
      return;
    }

    if (Estado.indicesFiltrados.length === 0) {
      el.recordsContainer.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">🔍</div>
          <p class="empty-title">No hay coincidencias</p>
          <p class="empty-desc">No se encontró ningún registro para "${Estado.busqueda}".</p>
        </div>
      `;
      return;
    }

    el.recordsContainer.innerHTML = "";

    const wrapper = document.createElement('div');
    wrapper.className = 'table-responsive-wrapper';

    const table = document.createElement('table');
    table.className = 'modern-data-table';

    // THEAD
    const thead = document.createElement('thead');
    const trHead = document.createElement('tr');

    Estado.encabezados.forEach((colName, cIdx) => {
      if (!Estado.columnasVisibles.includes(cIdx)) return;

      const th = document.createElement('th');
      th.className = 'th-sortable';
      if (Estado.tipos[cIdx] === 'number') th.classList.add('col-num-header');

      let sortIconHtml = '<span class="sort-idle-icon">⇅</span>';
      if (Estado.colOrden === cIdx) {
        const arrow = Estado.dirOrden === 'asc' ? '▲' : '▼';
        sortIconHtml = `<span class="sort-icon active">${arrow}</span>`;
      }

      th.innerHTML = `
        <div class="th-content">
          <span>${colName}</span>
          ${sortIconHtml}
        </div>
      `;
      th.title = `Ordenar por ${colName}`;
      th.addEventListener('click', () => alternarOrden(cIdx));
      trHead.appendChild(th);
    });

    const thActions = document.createElement('th');
    thActions.className = 'th-actions';
    thActions.textContent = 'Acciones';
    trHead.appendChild(thActions);

    thead.appendChild(trHead);
    table.appendChild(thead);

    // TBODY
    const tbody = document.createElement('tbody');

    Estado.indicesFiltrados.forEach((origIdx) => {
      const row = Estado.filas[origIdx];
      const tr = document.createElement('tr');
      tr.setAttribute('data-index', origIdx);

      if (Estado.filaSeleccionada === origIdx) {
        tr.classList.add('selected-row');
      }

      Estado.encabezados.forEach((colName, cIdx) => {
        if (!Estado.columnasVisibles.includes(cIdx)) return;

        const td = document.createElement('td');
        const val = row[cIdx];
        const tipo = Estado.tipos[cIdx];
        const esId = cIdx === 0;

        if (esId) {
          td.classList.add('col-primary');
          td.title = `${colName}: ${val} (Clic para copiar)`;
          td.style.cursor = 'pointer';
          td.addEventListener('click', (e) => {
            e.stopPropagation();
            if (navigator.clipboard) {
              navigator.clipboard.writeText(String(val)).then(() => {
                mostrarToast(`📋 Copiado: ${val}`, 'info');
              });
            }
          });
        }

        if (tipo === 'number') {
          td.classList.add('col-number');
          const formatted = typeof val === 'number' ? val.toLocaleString('es-MX') : val;
          td.innerHTML = resaltarCoincidencia(formatted, Estado.busqueda);
        } else if (tipo === 'email' && val) {
          td.innerHTML = `<a href="mailto:${val}" class="email-link" onclick="event.stopPropagation()">${resaltarCoincidencia(val, Estado.busqueda)}</a>`;
        } else {
          td.innerHTML = resaltarCoincidencia(val, Estado.busqueda);
        }

        tr.appendChild(td);
      });

      // Columna de Acciones
      const tdActions = document.createElement('td');
      tdActions.className = 'td-actions';
      tdActions.innerHTML = `
        <div class="table-action-btns">
          <button type="button" class="action-icon-btn view-btn" title="Ver ficha 360°">👁️</button>
          <button type="button" class="action-icon-btn edit-btn" title="Editar registro">✏️</button>
          <button type="button" class="action-icon-btn delete-btn" title="Eliminar registro">🗑️</button>
        </div>
      `;

      tdActions.querySelector('.view-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        abrirModalDetail(origIdx);
      });

      tdActions.querySelector('.edit-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        abrirModalEditar(origIdx);
      });

      tdActions.querySelector('.delete-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        abrirModalConfirm(origIdx);
      });

      tr.appendChild(tdActions);

      // Clic en la fila
      tr.addEventListener('click', () => {
        tbody.querySelectorAll('tr').forEach(r => r.classList.remove('selected-row'));
        tr.classList.add('selected-row');
        Estado.filaSeleccionada = origIdx;

        // Resaltar en la cuadrícula de Excel
        resaltarFilaEnExcel(origIdx);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);
    el.recordsContainer.appendChild(wrapper);
  }

  function resaltarFilaEnExcel(origIdx) {
    if (!el.excelGridTable) return;
    const excelRows = el.excelGridTable.querySelectorAll('.excel-data-row');
    excelRows.forEach(tr => {
      if (Number(tr.dataset.rowIndex) === origIdx) {
        tr.style.backgroundColor = '#EFF6FF';
        tr.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        // Posicionar barra de fórmulas
        const firstVal = Estado.filas[origIdx][0];
        actualizarBarraFormulas(`A${origIdx + 2}`, firstVal);
      } else {
        tr.style.backgroundColor = '';
      }
    });
  }

  // ============================================================================
  // 8. SELECTOR DE COLUMNAS (POPOVER)
  // ============================================================================

  function renderizarSelectorColumnas() {
    if (!el.columnsListContainer || !el.columnsPickerLabel) return;

    el.columnsPickerLabel.textContent = `Columnas (${Estado.columnasVisibles.length}/${Estado.encabezados.length})`;
    el.columnsListContainer.innerHTML = "";

    const nonIdIndices = Estado.encabezados.map((_, i) => i).filter(i => i !== 0);
    const allSelected = nonIdIndices.every(i => Estado.columnasVisibles.includes(i));
    if (el.btnToggleAllColumns) {
      el.btnToggleAllColumns.textContent = allSelected ? "Quitar todo" : "Seleccionar todo";
    }

    Estado.encabezados.forEach((colName, cIdx) => {
      const isId = cIdx === 0;
      const isVisible = Estado.columnasVisibles.includes(cIdx);

      const label = document.createElement('label');
      label.className = `column-checkbox-item ${isId ? 'disabled' : ''}`;

      const chk = document.createElement('input');
      chk.type = 'checkbox';
      chk.checked = isVisible;
      if (isId) chk.disabled = true;

      chk.addEventListener('change', () => {
        if (isId) return;
        if (chk.checked) {
          if (!Estado.columnasVisibles.includes(cIdx)) {
            Estado.columnasVisibles.push(cIdx);
            Estado.columnasVisibles.sort((a, b) => a - b);
          }
        } else {
          Estado.columnasVisibles = Estado.columnasVisibles.filter(i => i !== cIdx);
        }
        renderizarSelectorColumnas();
        renderizarVistaTabularTaskPane();
      });

      const span = document.createElement('span');
      span.className = 'col-name';
      span.textContent = colName;

      label.appendChild(chk);
      label.appendChild(span);

      if (isId) {
        const badge = document.createElement('span');
        badge.className = 'id-lock-tag';
        badge.textContent = '🔒 ID';
        label.appendChild(badge);
      }

      el.columnsListContainer.appendChild(label);
    });
  }

  function alternarTodasLasColumnas() {
    const nonIdIndices = Estado.encabezados.map((_, i) => i).filter(i => i !== 0);
    const allSelected = nonIdIndices.every(i => Estado.columnasVisibles.includes(i));

    if (allSelected) {
      Estado.columnasVisibles = [0]; // Solo dejar ID
    } else {
      Estado.columnasVisibles = Estado.encabezados.map((_, i) => i);
    }
    renderizarSelectorColumnas();
    renderizarVistaTabularTaskPane();
  }

  // ============================================================================
  // 9. FORMULARIO DINÁMICO & CÁLCULO DE FÓRMULAS VIVAS EN TIEMPO REAL
  // ============================================================================

  function generarSiguienteId(colIdx) {
    const vals = Estado.filas.map(r => String(r[colIdx] || '')).filter(Boolean);
    if (vals.length === 0) return "ID-101";

    let maxNum = -1;
    let prefix = "ID-";
    let pad = 3;

    vals.forEach(v => {
      const match = v.match(/^(.*?)(\d+)$/);
      if (match) {
        const num = parseInt(match[2], 10);
        if (num > maxNum) {
          maxNum = num;
          prefix = match[1];
          pad = match[2].length;
        }
      }
    });

    if (maxNum !== -1) {
      return `${prefix}${String(maxNum + 1).padStart(pad, '0')}`;
    }
    return `ID-${vals.length + 1}`;
  }

  function abrirModalNuevo() {
    Estado.filaEnEdicion = null;
    if (el.modalFormTitle) el.modalFormTitle.textContent = "Nuevo Registro";
    if (el.modalFormIcon) el.modalFormIcon.textContent = "➕";
    if (el.btnSaveText) el.btnSaveText.textContent = "Guardar en Excel";

    construirCamposFormulario(null);
    el.modalForm?.classList.remove('hidden');
  }

  function abrirModalEditar(rowIndex) {
    Estado.filaEnEdicion = rowIndex;
    if (el.modalFormTitle) el.modalFormTitle.textContent = `Editar Registro (Fila ${rowIndex + 1})`;
    if (el.modalFormIcon) el.modalFormIcon.textContent = "✏️";
    if (el.btnSaveText) el.btnSaveText.textContent = "Actualizar en Excel";

    const datosFila = Estado.filas[rowIndex];
    construirCamposFormulario(datosFila);
    el.modalForm?.classList.remove('hidden');
  }

  function cerrarModalForm() {
    el.modalForm?.classList.add('hidden');
    Estado.filaEnEdicion = null;
  }

  function construirCamposFormulario(datosFila) {
    if (!el.dynamicFormFields) return;
    el.dynamicFormFields.innerHTML = "";

    Estado.encabezados.forEach((colName, cIdx) => {
      const esFormula = Estado.esCalculada[cIdx];
      const tipo = Estado.tipos[cIdx];
      const esId = cIdx === 0;
      const opciones = Estado.opciones[cIdx];

      let valInicial = datosFila ? (datosFila[cIdx] !== undefined ? datosFila[cIdx] : '') : '';
      if (!datosFila && esId && !esFormula) {
        valInicial = generarSiguienteId(cIdx);
      }

      const group = document.createElement('div');
      group.className = 'input-group';

      const label = document.createElement('label');
      label.className = 'input-label';
      label.innerHTML = `
        <span>${colName}</span>
        ${esFormula ? '<span class="formula-tag">⚙️ Calculada</span>' : ''}
        ${opciones && !esFormula ? '<span class="formula-tag" style="background:#E0E7FF; color:#4338CA;">📋 Lista</span>' : ''}
        ${esId && !esFormula && !datosFila ? '<span class="formula-tag" style="background:#DCFCE7; color:#15803D;">⚡ Auto-ID</span>' : ''}
        ${esId && !esFormula && datosFila ? '<span class="formula-tag" style="background:#F1F5F9; color:#64748B;">🔒 ID Inmutable</span>' : ''}
      `;
      group.appendChild(label);

      // Si es campo de fórmula de Excel -> Inhabilitado con preview vivo
      if (esFormula) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'modal-input';
        input.name = `field_${cIdx}`;
        input.value = valInicial !== '' ? (typeof valInicial === 'number' ? valInicial.toLocaleString('es-MX') : valInicial) : '(Calculado por Excel)';
        input.disabled = true;
        input.title = `Fórmula de Excel: ${Estado.formulas[cIdx]}`;
        group.appendChild(input);
      }
      // Si es ID en edición -> Inmutable
      else if (esId && datosFila) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'modal-input';
        input.name = `field_${cIdx}`;
        input.value = valInicial;
        input.disabled = true;
        group.appendChild(input);
      }
      // Si es lista desplegable / catálogo
      else if (opciones && opciones.length > 0) {
        const select = document.createElement('select');
        select.className = 'modal-input custom-select';
        select.name = `field_${cIdx}`;

        const defOpt = document.createElement('option');
        defOpt.value = '';
        defOpt.textContent = `-- Seleccionar ${colName} --`;
        if (!valInicial) defOpt.selected = true;
        select.appendChild(defOpt);

        opciones.forEach(opt => {
          const o = document.createElement('option');
          o.value = opt;
          o.textContent = opt;
          if (String(valInicial).toLowerCase() === opt.toLowerCase()) o.selected = true;
          select.appendChild(o);
        });

        group.appendChild(select);
      }
      // Campo de texto, número, fecha, etc.
      else {
        const input = document.createElement('input');
        input.className = 'modal-input';
        input.name = `field_${cIdx}`;

        if (tipo === 'number') {
          input.type = 'number';
          input.step = 'any';
          input.placeholder = '0.00';
          input.value = valInicial;

          // Si es TablaVentas y estamos editando Cantidad o Precio, dar feedback de fórmulas en vivo
          if (Estado.tablaActiva === 'TablaVentas' && (colName === 'Cantidad' || colName === 'Precio Unitario')) {
            input.addEventListener('input', () => {
              const cantInput = el.dynamicFormFields.querySelector('[name="field_4"]');
              const precInput = el.dynamicFormFields.querySelector('[name="field_5"]');
              const cant = cantInput ? parseFloat(cantInput.value) || 0 : 0;
              const prec = precInput ? parseFloat(precInput.value) || 0 : 0;
              const subtotal = Math.round((cant * prec) * 100) / 100;
              const iva = Math.round((subtotal * 0.16) * 100) / 100;
              const total = Math.round((subtotal + iva) * 100) / 100;

              const subtotalField = el.dynamicFormFields.querySelector('[name="field_6"]');
              const ivaField = el.dynamicFormFields.querySelector('[name="field_7"]');
              const totalField = el.dynamicFormFields.querySelector('[name="field_8"]');

              if (subtotalField) subtotalField.value = subtotal.toLocaleString('es-MX');
              if (ivaField) ivaField.value = iva.toLocaleString('es-MX');
              if (totalField) totalField.value = total.toLocaleString('es-MX');
            });
          }
        } else if (tipo === 'date') {
          input.type = 'date';
          input.value = valInicial;
        } else if (tipo === 'email') {
          input.type = 'email';
          input.placeholder = 'contacto@empresa.com';
          input.value = valInicial;
        } else {
          input.type = 'text';
          input.placeholder = esId ? `Folio ID` : `Ingresar ${colName.toLowerCase()}...`;
          input.value = valInicial;
        }

        group.appendChild(input);
      }

      el.dynamicFormFields.appendChild(group);
    });
  }

  function guardarRegistroForm(e) {
    e.preventDefault();

    const nuevaFila = [];
    let cantidadVal = 0;
    let precioVal = 0;
    let stockVal = 0;
    let costoVal = 0;

    for (let c = 0; c < Estado.encabezados.length; c++) {
      const esFormula = Estado.esCalculada[c];
      const tipo = Estado.tipos[c];

      if (esFormula) {
        nuevaFila.push(0); // Se calculará abajo según la fórmula
        continue;
      }

      const input = el.dynamicFormFields.querySelector(`[name="field_${c}"]`);
      let val = input ? input.value.trim() : "";

      if (tipo === 'number') {
        val = val === "" ? 0 : parseFloat(val);
        if (Estado.encabezados[c] === 'Cantidad') cantidadVal = val;
        if (Estado.encabezados[c] === 'Precio Unitario') precioVal = val;
        if (Estado.encabezados[c] === 'Stock Actual') stockVal = val;
        if (Estado.encabezados[c] === 'Costo Unitario') costoVal = val;
      }

      nuevaFila.push(val);
    }

    // Cálculo automático de fórmulas según la tabla activa
    if (Estado.tablaActiva === 'TablaVentas') {
      const subtotal = Math.round((cantidadVal * precioVal) * 100) / 100;
      const iva = Math.round((subtotal * 0.16) * 100) / 100;
      const totalVenta = Math.round((subtotal + iva) * 100) / 100;
      nuevaFila[6] = subtotal;
      nuevaFila[7] = iva;
      nuevaFila[8] = totalVenta;
    } else if (Estado.tablaActiva === 'TablaInventario') {
      const valorTotal = Math.round((stockVal * costoVal) * 100) / 100;
      nuevaFila[6] = valorTotal;
    }

    let indiceModificado = null;

    if (Estado.filaEnEdicion === null) {
      // INSERTAR NUEVO REGISTRO
      Estado.filas.unshift(nuevaFila);
      TABLAS_LIBRO[Estado.tablaActiva].filas.unshift(nuevaFila);
      indiceModificado = 0;
      mostrarToast(`✅ Registro ${nuevaFila[0]} agregado y fórmulas calculadas`, "success");
    } else {
      // ACTUALIZAR REGISTRO EXISTENTE
      const idx = Estado.filaEnEdicion;
      Estado.filas[idx] = nuevaFila;
      TABLAS_LIBRO[Estado.tablaActiva].filas[idx] = nuevaFila;
      indiceModificado = idx;
      mostrarToast(`✅ Registro ${nuevaFila[0]} actualizado con éxito`, "success");
    }

    cerrarModalForm();
    filtrarYRenderizar();
    renderizarCuadriculaExcel(indiceModificado);

    // Actualizar texto del callout de sincronización
    if (el.syncStatusText) {
      el.syncStatusText.textContent = `Escritura exitosa en Excel: Registro ${nuevaFila[0]}`;
      setTimeout(() => {
        el.syncStatusText.textContent = "Sincronización Bidireccional Activa (Office.js Mock Engine)";
      }, 4000);
    }
  }

  // ============================================================================
  // 10. FICHA 360° Y ELIMINACIÓN
  // ============================================================================

  function abrirModalDetail(rowIndex) {
    Estado.filaSeleccionada = rowIndex;
    const fila = Estado.filas[rowIndex];
    const primerVal = fila[0] || `Fila #${rowIndex + 1}`;

    if (el.modalDetailTitle) {
      el.modalDetailTitle.textContent = `${primerVal} • Ficha Detallada`;
    }

    if (el.detailContent) {
      el.detailContent.innerHTML = "";
      Estado.encabezados.forEach((colName, cIdx) => {
        const val = fila[cIdx];
        const tipo = Estado.tipos[cIdx];
        const esFormula = Estado.esCalculada[cIdx];

        let formatted = val;
        if (val === null || val === undefined || val === '') {
          formatted = '<span style="color:#94A3B8;">-</span>';
        } else if (tipo === 'number' && typeof val === 'number') {
          formatted = val.toLocaleString('es-MX');
        }

        const item = document.createElement('div');
        item.className = 'detail-item';
        item.innerHTML = `
          <div class="detail-key">
            <span>${colName}</span>
            ${esFormula ? '<span class="formula-tag">⚙️ Calculada</span>' : ''}
          </div>
          <div class="detail-val">${formatted}</div>
        `;
        el.detailContent.appendChild(item);
      });
    }

    el.modalDetail?.classList.remove('hidden');
  }

  function cerrarModalDetail() {
    el.modalDetail?.classList.add('hidden');
  }

  function abrirModalConfirm(rowIndex) {
    Estado.filaAEliminar = rowIndex;
    el.modalConfirm?.classList.remove('hidden');
  }

  function cerrarModalConfirm() {
    el.modalConfirm?.classList.add('hidden');
    Estado.filaAEliminar = null;
  }

  function ejecutarEliminar() {
    if (Estado.filaAEliminar === null) return;
    const idx = Estado.filaAEliminar;
    const eliminado = Estado.filas.splice(idx, 1);
    TABLAS_LIBRO[Estado.tablaActiva].filas.splice(idx, 1);
    cerrarModalConfirm();

    filtrarYRenderizar();
    renderizarCuadriculaExcel();

    mostrarToast(`🗑️ Registro ${eliminado[0]?.[0] || ''} eliminado de Excel`, "danger");
  }

  // ============================================================================
  // 11. GUARDAR COMO VISTA EJECUTIVA EN EXCEL
  // ============================================================================

  function generarNombreVistaDefault() {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const dd = pad(now.getDate());
    const mm = pad(now.getMonth() + 1);
    const yy = String(now.getFullYear()).slice(-2);
    const hh = pad(now.getHours());
    const min = pad(now.getMinutes());
    return `Vista_${dd}${mm}${yy}_${hh}${min}`;
  }

  function abrirModalSaveView() {
    if (!el.inputViewName || !el.saveViewSummaryBox) return;

    el.inputViewName.value = generarNombreVistaDefault();

    const filtrados = Estado.indicesFiltrados.length;
    const colsVisibles = Estado.columnasVisibles.length;
    const filtroStr = Estado.busqueda ? `"${Estado.busqueda}"` : "Ninguno (Todos)";

    el.saveViewSummaryBox.innerHTML = `
      <div class="view-stat-row">
        <span>Tabla Origen:</span>
        <strong>${Estado.tablaActiva}</strong>
      </div>
      <div class="view-stat-row">
        <span>Filas a Exportar:</span>
        <strong>${filtrados} registros</strong>
      </div>
      <div class="view-stat-row">
        <span>Columnas Visibles:</span>
        <strong>${colsVisibles} de ${Estado.encabezados.length}</strong>
      </div>
      <div class="view-stat-row">
        <span>Filtro de Búsqueda:</span>
        <strong>${filtroStr}</strong>
      </div>
    `;

    el.modalSaveView?.classList.remove('hidden');
  }

  function cerrarModalSaveView() {
    el.modalSaveView?.classList.add('hidden');
  }

  function ejecutarGuardarVista() {
    const rawName = el.inputViewName ? el.inputViewName.value.trim() : "";
    const cleanName = rawName.replace(/[^a-zA-Z0-9_-]/g, '') || generarNombreVistaDefault();

    // Crear la nueva pestaña en el objeto de tablas
    const visibleHeaders = Estado.columnasVisibles.map(c => Estado.encabezados[c]);
    const visibleTypes = Estado.columnasVisibles.map(c => Estado.tipos[c]);
    const visibleCalculada = Estado.columnasVisibles.map(c => Estado.esCalculada[c]);
    const visibleFormulas = Estado.columnasVisibles.map(c => Estado.formulas[c]);
    const visibleRows = Estado.indicesFiltrados.map(rIdx => {
      const fullRow = Estado.filas[rIdx];
      return Estado.columnasVisibles.map(cIdx => fullRow[cIdx]);
    });

    TABLAS_LIBRO[cleanName] = {
      nombre: cleanName,
      archivo: "Plantilla_Superpoderes_Excel_TablaPersonal.xlsx",
      encabezados: visibleHeaders,
      tipos: visibleTypes,
      esCalculada: visibleCalculada,
      formulas: visibleFormulas,
      opciones: {},
      filas: visibleRows
    };

    // Agregar opción al select de tablas del Task Pane
    if (el.selectTables) {
      const opt = document.createElement('option');
      opt.value = cleanName;
      opt.textContent = `📑 ${cleanName}`;
      el.selectTables.appendChild(opt);
    }

    // Agregar pestaña al mockup de Excel
    if (el.excelSheetTabsBar && el.btnAddSheetTab) {
      const tab = document.createElement('div');
      tab.className = 'sheet-tab';
      tab.setAttribute('data-table', cleanName);
      tab.innerHTML = `<span>📑 ${cleanName}</span>`;
      el.excelSheetTabsBar.insertBefore(tab, el.btnAddSheetTab);
    }

    cerrarModalSaveView();
    mostrarToast(`✅ Hoja y tabla ejecutiva "${cleanName}" creada en Excel`, "success");

    // Cambiar automáticamente a la nueva tabla
    cargarTabla(cleanName);
  }

  // ============================================================================
  // 12. SISTEMA DE NOTIFICACIONES TOAST
  // ============================================================================

  function mostrarToast(mensaje, tipo = "info") {
    if (!el.toastContainer) return;

    const toast = document.createElement('div');
    toast.className = `toast ${tipo}`;

    let icon = "⚡";
    if (tipo === 'success') icon = "✅";
    if (tipo === 'danger') icon = "⚠️";

    toast.innerHTML = `
      <span class="toast-icon">${icon}</span>
      <span class="toast-message">${mensaje}</span>
    `;

    el.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }

  // ============================================================================
  // AUTO-INICIALIZACIÓN AL CARGAR LA PÁGINA
  // ============================================================================
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', iniciarDemo);
  } else {
    iniciarDemo();
  }

})();
