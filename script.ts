/**
 * ============================================================================
 * 📊 UNIVERSAL TABLE MANAGER - SCRIPT LAB (OFFICE.JS & EXCEL 365)
 * ============================================================================
 * Aplicación universal para interactuar, gestionar y visualizar cualquier tabla
 * de Excel con capacidades avanzadas tipo Power BI, formularios dinámicos,
 * autoincremento de folios, protección de fórmulas y exportación ejecutiva.
 * 
 * Desarrollado para Microsoft Script Lab & Office.js
 * ============================================================================
 */

// ============================================================================
// 1️⃣ TIPOS, MODELOS Y ESTADO DE LA APLICACIÓN
// ============================================================================

type TipoColumna = 'date' | 'number' | 'text' | 'email';
type DireccionOrden = 'asc' | 'desc';
type TipoNotificacion = 'info' | 'success' | 'danger';
type ValorCelda = string | number | boolean | null | undefined;
type DatosFila = ValorCelda[];

/**
 * Modelo de Estado Centralizado de la Aplicación.
 */
interface ModeloEstadoApp {
  isOfficeReady: boolean;
  tablasDisponibles: string[];
  nombreTablaActiva: string;
  encabezados: string[];
  filasDatos: DatosFila[];
  datosFormulas: string[][];
  datosFormatosNumero: string[][];
  esColumnaCalculada: boolean[];
  tiposColumna: TipoColumna[];
  opcionesColumna: (string[] | null)[];
  indicesColumnasVisibles: number[];
  indicesFiltrados: number[];
  consultaBusquedaActiva: string;
  indiceColumnaOrden: number | null;
  direccionOrden: DireccionOrden;
  indiceFilaSeleccionada: number | null;
  indiceFilaEnEdicion: number | null;
  indiceFilaAEliminar: number | null;
}

/**
 * Referencias Fuertemente Tipadas a los Elementos del DOM.
 */
interface ElementosDOM {
  // Panel Superior y Filtros
  selectTables: HTMLSelectElement;
  btnRefreshTables: HTMLButtonElement;
  inputSearch: HTMLInputElement;
  btnClearSearch: HTMLButtonElement;
  btnNewRecord: HTMLButtonElement;
  counterText: HTMLElement;
  activeTableBadge: HTMLElement;
  recordsContainer: HTMLElement;
  
  // Selector de Columnas Visibles
  btnToggleColumnsPicker: HTMLButtonElement;
  columnsDropdown: HTMLElement;
  columnsPickerLabel: HTMLElement;
  btnToggleAllColumns: HTMLButtonElement;
  columnsListContainer: HTMLElement;

  // Botón Guardar Vista
  btnOpenSaveView: HTMLButtonElement;

  // Modal 1: Formulario Dinámico (Crear / Editar)
  modalForm: HTMLElement;
  modalFormTitle: HTMLElement;
  modalFormIcon: HTMLElement;
  dynamicRecordForm: HTMLFormElement;
  dynamicFormFields: HTMLElement;
  btnSaveForm: HTMLButtonElement;
  btnSaveText: HTMLElement;
  btnCloseModalForm: HTMLButtonElement;
  btnCancelForm: HTMLButtonElement;
  
  // Modal 2: Ficha 360° (Detalle)
  modalDetail: HTMLElement;
  modalDetailTitle: HTMLElement;
  detailContent: HTMLElement;
  btnCloseModalDetail: HTMLButtonElement;
  btnCloseDetail: HTMLButtonElement;
  btnEditFromDetail: HTMLButtonElement;
  
  // Modal 3: Confirmación de Eliminación
  modalConfirm: HTMLElement;
  btnCancelDelete: HTMLButtonElement;
  btnConfirmDelete: HTMLButtonElement;
  
  // Modal 4: Guardar como Vista en Excel
  modalSaveView: HTMLElement;
  inputViewName: HTMLInputElement;
  saveViewSummaryBox: HTMLElement;
  btnCloseModalSaveView: HTMLButtonElement;
  btnCancelSaveView: HTMLButtonElement;
  btnConfirmSaveView: HTMLButtonElement;

  // Contenedor Toast
  toastContainer: HTMLElement;
}

const EstadoApp: ModeloEstadoApp = {
  isOfficeReady: false,
  tablasDisponibles: [],
  nombreTablaActiva: '',
  encabezados: [],
  filasDatos: [],
  datosFormulas: [],
  datosFormatosNumero: [],
  esColumnaCalculada: [],
  tiposColumna: [],
  opcionesColumna: [],
  indicesColumnasVisibles: [],
  indicesFiltrados: [],
  consultaBusquedaActiva: '',
  indiceColumnaOrden: null,
  direccionOrden: 'asc',
  indiceFilaSeleccionada: null,
  indiceFilaEnEdicion: null,
  indiceFilaAEliminar: null
};

let DOM: ElementosDOM;

function obtenerElementosDOM(): ElementosDOM {
  return {
    selectTables: document.getElementById('selectTables') as HTMLSelectElement,
    btnRefreshTables: document.getElementById('btnRefreshTables') as HTMLButtonElement,
    inputSearch: document.getElementById('inputSearch') as HTMLInputElement,
    btnClearSearch: document.getElementById('btnClearSearch') as HTMLButtonElement,
    btnNewRecord: document.getElementById('btnNewRecord') as HTMLButtonElement,
    counterText: document.getElementById('counterText') as HTMLElement,
    activeTableBadge: document.getElementById('activeTableBadge') as HTMLElement,
    recordsContainer: document.getElementById('recordsContainer') as HTMLElement,
    
    btnToggleColumnsPicker: document.getElementById('btnToggleColumnsPicker') as HTMLButtonElement,
    columnsDropdown: document.getElementById('columnsDropdown') as HTMLElement,
    columnsPickerLabel: document.getElementById('columnsPickerLabel') as HTMLElement,
    btnToggleAllColumns: document.getElementById('btnToggleAllColumns') as HTMLButtonElement,
    columnsListContainer: document.getElementById('columnsListContainer') as HTMLElement,

    btnOpenSaveView: document.getElementById('btnOpenSaveView') as HTMLButtonElement,

    modalForm: document.getElementById('modalForm') as HTMLElement,
    modalFormTitle: document.getElementById('modalFormTitle') as HTMLElement,
    modalFormIcon: document.getElementById('modalFormIcon') as HTMLElement,
    dynamicRecordForm: document.getElementById('dynamicRecordForm') as HTMLFormElement,
    dynamicFormFields: document.getElementById('dynamicFormFields') as HTMLElement,
    btnSaveForm: document.getElementById('btnSaveForm') as HTMLButtonElement,
    btnSaveText: document.getElementById('btnSaveText') as HTMLElement,
    btnCloseModalForm: document.getElementById('btnCloseModalForm') as HTMLButtonElement,
    btnCancelForm: document.getElementById('btnCancelForm') as HTMLButtonElement,
    
    modalDetail: document.getElementById('modalDetail') as HTMLElement,
    modalDetailTitle: document.getElementById('modalDetailTitle') as HTMLElement,
    detailContent: document.getElementById('detailContent') as HTMLElement,
    btnCloseModalDetail: document.getElementById('btnCloseModalDetail') as HTMLButtonElement,
    btnCloseDetail: document.getElementById('btnCloseDetail') as HTMLButtonElement,
    btnEditFromDetail: document.getElementById('btnEditFromDetail') as HTMLButtonElement,
    
    modalConfirm: document.getElementById('modalConfirm') as HTMLElement,
    btnCancelDelete: document.getElementById('btnCancelDelete') as HTMLButtonElement,
    btnConfirmDelete: document.getElementById('btnConfirmDelete') as HTMLButtonElement,
    
    modalSaveView: document.getElementById('modalSaveView') as HTMLElement,
    inputViewName: document.getElementById('inputViewName') as HTMLInputElement,
    saveViewSummaryBox: document.getElementById('saveViewSummaryBox') as HTMLElement,
    btnCloseModalSaveView: document.getElementById('btnCloseModalSaveView') as HTMLButtonElement,
    btnCancelSaveView: document.getElementById('btnCancelSaveView') as HTMLButtonElement,
    btnConfirmSaveView: document.getElementById('btnConfirmSaveView') as HTMLButtonElement,

    toastContainer: document.getElementById('toastContainer') as HTMLElement
  };
}

// ============================================================================
// 2️⃣ INICIALIZACIÓN Y VINCULACIÓN DE EVENTOS
// ============================================================================

function inicializarAdministradorUniversal(): void {
  DOM = obtenerElementosDOM();
  vincularEventosUI();
  cargarTablasDelLibro();
}

// Inicialización de Office.js garantizada en Script Lab
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      EstadoApp.isOfficeReady = true;
      inicializarAdministradorUniversal();
    } else {
      inicializarAdministradorUniversal();
    }
  });
} else {
  setTimeout(inicializarAdministradorUniversal, 100);
}

function vincularEventosUI(): void {
  // Cambio de tabla activa
  DOM.selectTables?.addEventListener('change', (e: Event) => {
    const target = e.target as HTMLSelectElement;
    const nombre = target.value;
    if (nombre) cambiarTablaActiva(nombre);
  });

  // Refrescar lista de tablas
  DOM.btnRefreshTables?.addEventListener('click', () => {
    mostrarNotificacionToast("Detectando tablas en el libro...", "info");
    cargarTablasDelLibro();
  });

  // Buscador universal en tiempo real
  DOM.inputSearch?.addEventListener('input', (e: Event) => {
    const target = e.target as HTMLInputElement;
    EstadoApp.consultaBusquedaActiva = target.value.trim().toLowerCase();
    DOM.btnClearSearch?.classList.toggle('hidden', EstadoApp.consultaBusquedaActiva.length === 0);
    filtraryOrdenarRegistros();
  });

  DOM.btnClearSearch?.addEventListener('click', () => {
    if (DOM.inputSearch) DOM.inputSearch.value = '';
    EstadoApp.consultaBusquedaActiva = '';
    DOM.btnClearSearch?.classList.add('hidden');
    filtraryOrdenarRegistros();
  });

  // Botón Nuevo Registro
  DOM.btnNewRecord?.addEventListener('click', abrirModalNuevoRegistro);

  // Cerrar Ventanas Modales
  DOM.btnCloseModalForm?.addEventListener('click', cerrarModalFormulario);
  DOM.btnCancelForm?.addEventListener('click', cerrarModalFormulario);
  DOM.btnCloseModalDetail?.addEventListener('click', cerrarModalFicha360);
  DOM.btnCloseDetail?.addEventListener('click', cerrarModalFicha360);
  DOM.btnCancelDelete?.addEventListener('click', cerrarModalConfirmacionEliminar);

  // Selector de Columnas Popover
  DOM.btnToggleColumnsPicker?.addEventListener('click', (e: MouseEvent) => {
    e.stopPropagation();
    DOM.columnsDropdown?.classList.toggle('hidden');
  });

  DOM.btnToggleAllColumns?.addEventListener('click', (e: MouseEvent) => {
    e.stopPropagation();
    alternarTodasLasColumnas();
  });

  // Cerrar Popover de Columnas al hacer clic fuera
  document.addEventListener('click', (e: MouseEvent) => {
    const target = e.target as HTMLElement;
    if (!target.closest('.columns-picker-wrapper')) {
      DOM.columnsDropdown?.classList.add('hidden');
    }
  });

  // Enviar Formulario
  DOM.dynamicRecordForm?.addEventListener('submit', procesarEnvioFormulario);
  DOM.btnSaveForm?.addEventListener('click', (e: MouseEvent) => {
    e.preventDefault();
    procesarEnvioFormulario();
  });

  // Editar desde Ficha de Detalle
  DOM.btnEditFromDetail?.addEventListener('click', () => {
    const idx = EstadoApp.indiceFilaSeleccionada;
    cerrarModalFicha360();
    if (idx !== null) abrirModalEditarRegistro(idx);
  });

  // Modal Guardar Vista
  DOM.btnOpenSaveView?.addEventListener('click', abrirModalGuardarVista);
  DOM.btnCloseModalSaveView?.addEventListener('click', cerrarModalGuardarVista);
  DOM.btnCancelSaveView?.addEventListener('click', cerrarModalGuardarVista);
  DOM.btnConfirmSaveView?.addEventListener('click', confirmarGuardarVista);

  // Confirmar Eliminación
  DOM.btnConfirmDelete?.addEventListener('click', confirmarEliminarRegistro);
}

// ============================================================================
// 3️⃣ SUPERPODER 1: ADAPTABILIDAD UNIVERSAL (Lectura con Office.js)
// ============================================================================

/**
 * Detecta dinámicamente todas las tablas de Excel en el libro activo.
 */
async function cargarTablasDelLibro(): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const tables = context.workbook.tables;
      tables.load(["items/name"]);
      await context.sync();

      EstadoApp.tablasDisponibles = tables.items.map(t => t.name);

      if (EstadoApp.tablasDisponibles.length === 0) {
        if (DOM.selectTables) DOM.selectTables.innerHTML = '<option value="">No hay tablas en el libro</option>';
        if (DOM.activeTableBadge) DOM.activeTableBadge.textContent = 'Sin Tablas';
        renderizarEstadoVacio("No se encontraron tablas de Excel en este libro.");
        return;
      }

      poblarSelectorDeTablas();

      const actual = EstadoApp.nombreTablaActiva;
      if (actual && EstadoApp.tablasDisponibles.includes(actual)) {
        cambiarTablaActiva(actual);
      } else {
        cambiarTablaActiva(EstadoApp.tablasDisponibles[0]);
      }
    });
  } catch (error) {
    console.error("Error al cargar tablas del libro:", error);
    mostrarNotificacionToast("Error al conectar con las tablas de Excel", "danger");
  }
}

/**
 * Lee datos, encabezados, fórmulas, formatos y reglas de validación de la tabla seleccionada.
 */
async function cargarDatosTablaActiva(nombreTabla: string): Promise<void> {
  if (!nombreTabla) return;
  mostrarEsqueletosDeCarga();

  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const table = context.workbook.tables.getItem(nombreTabla);
      const headerRange = table.getHeaderRowRange();
      const bodyRange = table.getDataBodyRange();

      headerRange.load(["values"]);
      bodyRange.load(["values", "formulas", "numberFormat"]);
      await context.sync();

      EstadoApp.nombreTablaActiva = nombreTabla;
      EstadoApp.encabezados = (headerRange.values[0] || []).map(h => String(h).trim());
      EstadoApp.filasDatos = bodyRange.values || [];
      EstadoApp.datosFormulas = bodyRange.formulas || [];
      EstadoApp.datosFormatosNumero = bodyRange.numberFormat || [];
      EstadoApp.opcionesColumna = new Array(EstadoApp.encabezados.length).fill(null);

      // Cargar validaciones de datos (Listas desplegables de Excel)
      try {
        const columns = table.columns;
        columns.load(["items/name"]);
        await context.sync();

        const validaciones: { colIdx: number; validation: Excel.DataValidation }[] = [];
        for (let c = 0; c < EstadoApp.encabezados.length; c++) {
          const colRange = table.columns.getItemAt(c).getDataBodyRange();
          colRange.dataValidation.load(["rule"]);
          validaciones.push({ colIdx: c, validation: colRange.dataValidation });
        }
        await context.sync();

        for (const item of validaciones) {
          const rule = item.validation?.rule;
          if (rule && rule.list && rule.list.source) {
            const sourceStr = String(rule.list.source).trim();

            if (sourceStr.startsWith('=')) {
              // Rango referenciado (ej: =Catalogos!$A$2:$A$10)
              try {
                const cleanRef = sourceStr.substring(1).trim();
                let referencedRange: Excel.Range | null = null;

                if (cleanRef.includes('!')) {
                  const parts = cleanRef.split('!');
                  const sheetName = parts[0].replace(/^['"]|['"]$/g, '').trim();
                  const address = parts[1].trim();
                  const targetSheet = context.workbook.worksheets.getItem(sheetName);
                  referencedRange = targetSheet.getRange(address);
                } else if (/^[A-Za-z$]+[0-9]+/.test(cleanRef)) {
                  referencedRange = table.worksheet.getRange(cleanRef);
                } else {
                  const namedItem = context.workbook.names.getItem(cleanRef);
                  referencedRange = namedItem.getRange();
                }

                if (referencedRange) {
                  referencedRange.load(["values"]);
                  await context.sync();

                  const listValues: string[] = [];
                  if (referencedRange.values) {
                    referencedRange.values.forEach((r: any[]) => {
                      r.forEach((val: any) => {
                        if (val !== null && val !== undefined && String(val).trim() !== '') {
                          listValues.push(String(val).trim());
                        }
                      });
                    });
                  }
                  if (listValues.length > 0) {
                    EstadoApp.opcionesColumna[item.colIdx] = Array.from(new Set(listValues));
                  }
                }
              } catch (refErr) {
                console.warn(`No se pudo resolver el catálogo referenciado "${sourceStr}":`, refErr);
              }
            } else {
              // Lista de valores literales separados por comas
              const cleanedSource = sourceStr.replace(/^["']|["']$/g, '');
              const listValues = cleanedSource
                .split(/[,;]/)
                .map(s => s.trim().replace(/^["']|["']$/g, ''))
                .filter(s => s.length > 0);
              if (listValues.length > 0) {
                EstadoApp.opcionesColumna[item.colIdx] = listValues;
              }
            }
          }
        }
      } catch (valErr) {
        console.warn("Validaciones de lista no disponibles:", valErr);
      }

      // Inferir tipos reales de columnas
      inferirTiposDeColumna();

      // Columnas visibles por defecto
      EstadoApp.indicesColumnasVisibles = EstadoApp.encabezados.map((_, i) => i);
      renderizarSelectorColumnas();

      filtraryOrdenarRegistros();
    });
  } catch (error) {
    console.error(`Error al cargar datos de la tabla ${nombreTabla}:`, error);
    mostrarNotificacionToast(`Error al leer tabla ${nombreTabla}`, "danger");
  }
}

function cambiarTablaActiva(nombreTabla: string): void {
  if (DOM.selectTables) DOM.selectTables.value = nombreTabla;
  if (DOM.activeTableBadge) DOM.activeTableBadge.textContent = nombreTabla;
  EstadoApp.indiceColumnaOrden = null;
  EstadoApp.direccionOrden = 'asc';
  cargarDatosTablaActiva(nombreTabla);
}

function poblarSelectorDeTablas(): void {
  if (!DOM.selectTables) return;
  DOM.selectTables.innerHTML = '';
  EstadoApp.tablasDisponibles.forEach((nombre: string) => {
    const opt = document.createElement('option');
    opt.value = nombre;
    opt.textContent = nombre;
    DOM.selectTables.appendChild(opt);
  });
}

// ============================================================================
// 4️⃣ SUPERPODER 2: BUSCADOR MULTICOLUMNA & ORDENAMIENTO POWER BI
// ============================================================================

/**
 * Filtra y ordena los registros en memoria a través de todas las columnas.
 */
function filtraryOrdenarRegistros(): void {
  const query = EstadoApp.consultaBusquedaActiva;

  // 1. Filtrar coincidencias en cualquier columna
  if (!query) {
    EstadoApp.indicesFiltrados = EstadoApp.filasDatos.map((_, i) => i);
  } else {
    EstadoApp.indicesFiltrados = [];
    EstadoApp.filasDatos.forEach((fila: DatosFila, index: number) => {
      const match = fila.some((valorCelda: ValorCelda, colIdx: number) => {
        if (valorCelda === null || valorCelda === undefined) return false;
        if (EstadoApp.tiposColumna[colIdx] === 'date') {
          const fechaFormateada = formatearFechaExcel(valorCelda);
          if (fechaFormateada.toLowerCase().includes(query)) return true;
        }
        return String(valorCelda).toLowerCase().includes(query);
      });
      if (match) EstadoApp.indicesFiltrados.push(index);
    });
  }

  // 2. Ordenar según columna seleccionada
  if (EstadoApp.indiceColumnaOrden !== null && EstadoApp.indicesFiltrados.length > 1) {
    const colIdx = EstadoApp.indiceColumnaOrden;
    const dir = EstadoApp.direccionOrden === 'asc' ? 1 : -1;
    const tipoCol = EstadoApp.tiposColumna[colIdx];

    EstadoApp.indicesFiltrados.sort((idxA: number, idxB: number) => {
      const valA = EstadoApp.filasDatos[idxA][colIdx];
      const valB = EstadoApp.filasDatos[idxB][colIdx];

      if (valA === valB) return 0;
      if (valA === null || valA === undefined || valA === '') return 1;
      if (valB === null || valB === undefined || valB === '') return -1;

      if (tipoCol === 'number') {
        const numA = Number(valA);
        const numB = Number(valB);
        if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * dir;
      }

      if (tipoCol === 'date') {
        const aTimestamp = (v: ValorCelda) => {
          if (typeof v === 'number' && v > 10000) return v;
          const dStr = formatearFechaExcel(v);
          const t = new Date(dStr).getTime();
          return isNaN(t) ? 0 : t;
        };
        return (aTimestamp(valA) - aTimestamp(valB)) * dir;
      }

      return String(valA).localeCompare(String(valB), 'es', { numeric: true, sensitivity: 'base' }) * dir;
    });
  }

  renderizarVistaTabular();
  actualizarContador();
}

function alternarOrdenColumna(colIdx: number): void {
  if (EstadoApp.indiceColumnaOrden === colIdx) {
    EstadoApp.direccionOrden = EstadoApp.direccionOrden === 'asc' ? 'desc' : 'asc';
  } else {
    EstadoApp.indiceColumnaOrden = colIdx;
    EstadoApp.direccionOrden = 'asc';
  }
  filtraryOrdenarRegistros();
}

function resaltarCoincidencia(texto: ValorCelda, query: string): string {
  if (!query || texto === null || texto === undefined || texto === '') return String(texto || '');
  const str = String(texto);
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`(${escaped})`, 'gi');
  return str.replace(regex, '<mark class="search-highlight">$1</mark>');
}

// ============================================================================
// 5️⃣ SUPERPODER 3: SELECTOR DINÁMICO DE COLUMNAS AL VUELO
// ============================================================================

function renderizarSelectorColumnas(): void {
  if (!DOM.columnsListContainer || !DOM.columnsPickerLabel) return;

  DOM.columnsPickerLabel.textContent = `Columnas (${EstadoApp.indicesColumnasVisibles.length}/${EstadoApp.encabezados.length})`;
  DOM.columnsListContainer.innerHTML = '';

  const indicesNoId = EstadoApp.encabezados
    .map((h, i) => ({ encabezado: h, idx: i }))
    .filter(item => !esColumnaIdentificador(item.encabezado) && item.idx !== 0)
    .map(item => item.idx);

  const todosSeleccionados = indicesNoId.every(idx => EstadoApp.indicesColumnasVisibles.includes(idx));
  if (DOM.btnToggleAllColumns) {
    DOM.btnToggleAllColumns.textContent = todosSeleccionados ? 'Quitar todo' : 'Seleccionar todo';
  }

  EstadoApp.encabezados.forEach((encabezado: string, colIdx: number) => {
    const esId = esColumnaIdentificador(encabezado) || colIdx === 0;
    const esVisible = EstadoApp.indicesColumnasVisibles.includes(colIdx);

    const itemLabel = document.createElement('label');
    itemLabel.className = `column-checkbox-item ${esId ? 'disabled' : ''}`;
    if (esId) itemLabel.title = 'La columna ID es obligatoria y no se puede ocultar';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = esVisible;
    if (esId) checkbox.disabled = true;

    checkbox.addEventListener('change', () => {
      if (esId) return;
      if (checkbox.checked) {
        if (!EstadoApp.indicesColumnasVisibles.includes(colIdx)) {
          EstadoApp.indicesColumnasVisibles.push(colIdx);
          EstadoApp.indicesColumnasVisibles.sort((a, b) => a - b);
        }
      } else {
        EstadoApp.indicesColumnasVisibles = EstadoApp.indicesColumnasVisibles.filter(i => i !== colIdx);
      }
      renderizarSelectorColumnas();
      renderizarVistaTabular();
      actualizarContador();
    });

    const nameSpan = document.createElement('span');
    nameSpan.className = 'col-name';
    nameSpan.textContent = encabezado;

    itemLabel.appendChild(checkbox);
    itemLabel.appendChild(nameSpan);

    if (esId) {
      const lockBadge = document.createElement('span');
      lockBadge.className = 'id-lock-tag';
      lockBadge.textContent = '🔒 ID';
      itemLabel.appendChild(lockBadge);
    }

    DOM.columnsListContainer.appendChild(itemLabel);
  });
}

function alternarTodasLasColumnas(): void {
  const idIndices = EstadoApp.encabezados
    .map((h, i) => ({ encabezado: h, idx: i }))
    .filter(item => esColumnaIdentificador(item.encabezado) || item.idx === 0)
    .map(item => item.idx);

  const nonIdIndices = EstadoApp.encabezados
    .map((h, i) => ({ encabezado: h, idx: i }))
    .filter(item => !esColumnaIdentificador(item.encabezado) && item.idx !== 0)
    .map(item => item.idx);

  const todosSeleccionados = nonIdIndices.every(idx => EstadoApp.indicesColumnasVisibles.includes(idx));

  if (todosSeleccionados) {
    EstadoApp.indicesColumnasVisibles = [...idIndices];
  } else {
    EstadoApp.indicesColumnasVisibles = EstadoApp.encabezados.map((_, i) => i);
  }

  renderizarSelectorColumnas();
  renderizarVistaTabular();
  actualizarContador();
}

// ============================================================================
// 6️⃣ SUPERPODER 4: GENERADOR DE FORMULARIOS Y PROTECCIÓN DE FÓRMULAS
// ============================================================================

function esColumnaIdentificador(encabezado: string): boolean {
  const clean = (encabezado || '').trim().toLowerCase();
  return clean === 'id' || clean === 'id_' || clean === 'código' || clean === 'codigo' || clean === 'sku' || clean === 'folio' || clean === 'clave';
}

function generarSiguienteIdAuto(colIdx: number): string {
  const valoresExistentes = EstadoApp.filasDatos
    .map(r => String(r[colIdx] || '').trim())
    .filter(Boolean);

  if (valoresExistentes.length === 0) return 'ID-101';

  let maxNum = -1;
  let samplePrefix = '';
  let padLength = 1;

  for (const val of valoresExistentes) {
    const match = val.match(/^(.*?)(\d+)$/);
    if (match) {
      const prefix = match[1];
      const numStr = match[2];
      const num = parseInt(numStr, 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
        samplePrefix = prefix;
        padLength = numStr.length;
      }
    }
  }

  if (maxNum !== -1) {
    const nextNum = maxNum + 1;
    const padded = String(nextNum).padStart(padLength, '0');
    return `${samplePrefix}${padded}`;
  }

  return `ID-${valoresExistentes.length + 1}`;
}

function abrirModalNuevoRegistro(): void {
  EstadoApp.indiceFilaEnEdicion = null;
  if (DOM.modalFormTitle) DOM.modalFormTitle.textContent = 'Nuevo Registro';
  if (DOM.modalFormIcon) DOM.modalFormIcon.textContent = '➕';
  if (DOM.btnSaveText) DOM.btnSaveText.textContent = 'Guardar en Excel';
  
  construirCamposFormularioDinamico(null);
  DOM.modalForm?.classList.remove('hidden');
}

function abrirModalEditarRegistro(rowIndex: number): void {
  EstadoApp.indiceFilaEnEdicion = rowIndex;
  if (DOM.modalFormTitle) DOM.modalFormTitle.textContent = `Editar Registro (Fila ${rowIndex + 1})`;
  if (DOM.modalFormIcon) DOM.modalFormIcon.textContent = '✏️';
  if (DOM.btnSaveText) DOM.btnSaveText.textContent = 'Actualizar en Excel';

  const datosFilaActual = EstadoApp.filasDatos[rowIndex];
  construirCamposFormularioDinamico(datosFilaActual);
  DOM.modalForm?.classList.remove('hidden');
}

function cerrarModalFormulario(): void {
  DOM.modalForm?.classList.add('hidden');
  EstadoApp.indiceFilaEnEdicion = null;
}

/**
 * Construye dinámicamente inputs tipados, listas de catálogos y bloquea fórmulas.
 */
function construirCamposFormularioDinamico(datosFila: DatosFila | null): void {
  if (!DOM.dynamicFormFields) return;
  DOM.dynamicFormFields.innerHTML = '';

  EstadoApp.encabezados.forEach((encabezado: string, colIdx: number) => {
    const esFormula = EstadoApp.esColumnaCalculada[colIdx];
    const tipoCol = EstadoApp.tiposColumna[colIdx];
    const opciones = EstadoApp.opcionesColumna[colIdx];
    const esId = esColumnaIdentificador(encabezado);

    let valorInicial = datosFila ? (datosFila[colIdx] !== undefined && datosFila[colIdx] !== null ? String(datosFila[colIdx]) : '') : '';

    if (!datosFila && esId && !esFormula) {
      valorInicial = generarSiguienteIdAuto(colIdx);
    }

    const group = document.createElement('div');
    group.className = 'input-group';

    const label = document.createElement('label');
    label.className = 'input-label';
    label.innerHTML = `
      <span>${encabezado}</span>
      ${esFormula ? '<span class="formula-tag">⚙️ Calculada</span>' : ''}
      ${opciones && !esFormula ? '<span class="formula-tag" style="background:#E0E7FF; color:#4338CA;">📋 Lista</span>' : ''}
      ${esId && !esFormula && !datosFila ? '<span class="formula-tag" style="background:#DCFCE7; color:#15803D;">⚡ Auto-ID</span>' : ''}
      ${esId && !esFormula && datosFila ? '<span class="formula-tag" style="background:#F1F5F9; color:#64748B;">🔒 ID Inmutable</span>' : ''}
    `;

    if (esFormula) {
      const input = document.createElement('input');
      input.name = `field_${colIdx}`;
      input.className = 'modal-input';
      input.dataset.colIndex = String(colIdx);
      input.type = 'text';
      input.value = valorInicial || '(Calculado automáticamente)';
      input.disabled = true;

      group.appendChild(label);
      group.appendChild(input);
    } else if (esId && datosFila) {
      const input = document.createElement('input');
      input.name = `field_${colIdx}`;
      input.className = 'modal-input';
      input.dataset.colIndex = String(colIdx);
      input.type = 'text';
      input.value = valorInicial;
      input.disabled = true;
      input.title = 'El identificador principal es inmutable';

      group.appendChild(label);
      group.appendChild(input);
    } else if (opciones && opciones.length > 0) {
      const select = document.createElement('select');
      select.name = `field_${colIdx}`;
      select.className = 'modal-input custom-select';
      select.dataset.colIndex = String(colIdx);

      const defaultOpt = document.createElement('option');
      defaultOpt.value = '';
      defaultOpt.textContent = `-- Seleccionar ${encabezado} --`;
      if (!valorInicial) defaultOpt.selected = true;
      select.appendChild(defaultOpt);

      opciones.forEach((optVal: string) => {
        const opt = document.createElement('option');
        opt.value = optVal;
        opt.textContent = optVal;
        if (valorInicial && String(valorInicial).toLowerCase() === optVal.toLowerCase()) {
          opt.selected = true;
        }
        select.appendChild(opt);
      });

      group.appendChild(label);
      group.appendChild(select);
    } else {
      const input = document.createElement('input');
      input.name = `field_${colIdx}`;
      input.className = 'modal-input';
      input.dataset.colIndex = String(colIdx);

      if (tipoCol === 'number') {
        input.type = 'number';
        input.step = 'any';
        input.value = valorInicial;
        input.placeholder = `0.00`;
      } else if (tipoCol === 'date') {
        input.type = 'date';
        input.value = formatearFechaExcel(valorInicial);
      } else if (tipoCol === 'email') {
        input.type = 'email';
        input.value = valorInicial;
        input.placeholder = `contacto@empresa.com`;
      } else {
        input.type = 'text';
        input.value = valorInicial;
        input.placeholder = esId ? `Folio ID (ej: ${generarSiguienteIdAuto(colIdx)})` : `Ingresar ${encabezado.toLowerCase()}...`;
      }

      group.appendChild(label);
      group.appendChild(input);
    }

    DOM.dynamicFormFields.appendChild(group);
  });
}

function marcarCampoInvalido(colIdx: number): void {
  const inputEl = DOM.dynamicFormFields.querySelector(`[name="field_${colIdx}"]`) as HTMLElement;
  if (inputEl) {
    inputEl.focus();
    inputEl.style.borderColor = '#EF4444';
    inputEl.style.boxShadow = '0 0 0 3px rgba(239, 68, 68, 0.2)';
  }
}

// ============================================================================
// 7️⃣ SUPERPODER 5: ESCRITURA Y CRUD BIDIRECCIONAL EN EXCEL (WRITE-BACK)
// ============================================================================

async function insertarFilaEnExcel(valoresFila: DatosFila): Promise<boolean> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const table = context.workbook.tables.getItem(EstadoApp.nombreTablaActiva);
      const rows = table.rows;
      rows.load(["count"]);
      await context.sync();

      const rowCount = rows.count;
      const sanitized = valoresFila.map(v => (v === undefined ? null : v));

      // 1. Insertar la nueva fila con los valores proporcionados
      const newRow = table.rows.add(null, [sanitized]);

      // 2. Si la tabla ya tiene filas, asegurar que las columnas calculadas hereden la fórmula de la fila anterior
      if (rowCount > 0) {
        const prevRow = table.rows.getItemAt(rowCount - 1);
        const prevRange = prevRow.getRange();
        const newRange = newRow.getRange();

        for (let colIdx = 0; colIdx < EstadoApp.encabezados.length; colIdx++) {
          if (EstadoApp.esColumnaCalculada[colIdx]) {
            const prevCell = prevRange.getCell(0, colIdx);
            const newCell = newRange.getCell(0, colIdx);
            newCell.copyFrom(prevCell, Excel.RangeCopyType.formulas);
          }
        }
      }

      await context.sync();

      mostrarNotificacionToast("Registro agregado a Excel correctamente", "success");
      await cargarDatosTablaActiva(EstadoApp.nombreTablaActiva);
    });
    return true;
  } catch (error: any) {
    console.error("Error al insertar fila en Excel:", error);
    const detail = error?.message || "Validación rechazada por Excel";
    mostrarNotificacionToast(`⚠️ Error al guardar: ${detail}`, "danger");
    return false;
  }
}

async function actualizarFilaEnExcel(rowIndex: number, valoresFila: DatosFila): Promise<boolean> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const table = context.workbook.tables.getItem(EstadoApp.nombreTablaActiva);
      const row = table.rows.getItemAt(rowIndex);
      const rowRange = row.getRange();

      // Actualizar únicamente las celdas no calculadas para no sobreescribir las fórmulas de Excel
      for (let colIdx = 0; colIdx < EstadoApp.encabezados.length; colIdx++) {
        if (!EstadoApp.esColumnaCalculada[colIdx]) {
          const val = valoresFila[colIdx];
          rowRange.getCell(0, colIdx).values = [[val === undefined ? null : val]];
        }
      }

      await context.sync();

      mostrarNotificacionToast("Registro actualizado en Excel", "success");
      await cargarDatosTablaActiva(EstadoApp.nombreTablaActiva);
    });
    return true;
  } catch (error: any) {
    console.error("Error al actualizar fila en Excel:", error);
    const detail = error?.message || "Validación rechazada por Excel";
    mostrarNotificacionToast(`⚠️ Error al actualizar: ${detail}`, "danger");
    return false;
  }
}

async function eliminarFilaEnExcel(rowIndex: number): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const table = context.workbook.tables.getItem(EstadoApp.nombreTablaActiva);
      const row = table.rows.getItemAt(rowIndex);
      row.delete();
      await context.sync();

      mostrarNotificacionToast("Registro eliminado de Excel", "success");
      await cargarDatosTablaActiva(EstadoApp.nombreTablaActiva);
    });
  } catch (error) {
    console.error("Error al eliminar fila en Excel:", error);
    mostrarNotificacionToast("Error al eliminar registro", "danger");
  }
}

async function procesarEnvioFormulario(e?: Event): Promise<void> {
  if (e) e.preventDefault();

  try {
    DOM.dynamicFormFields.querySelectorAll('.modal-input').forEach((el) => {
      const htmlEl = el as HTMLElement;
      htmlEl.style.borderColor = '';
      htmlEl.style.boxShadow = '';
    });

    const valoresFila: DatosFila = [];

    for (let colIdx = 0; colIdx < EstadoApp.encabezados.length; colIdx++) {
      const esFormula = EstadoApp.esColumnaCalculada[colIdx];
      const tipoCol = EstadoApp.tiposColumna[colIdx];
      const opciones = EstadoApp.opcionesColumna[colIdx];
      const encabezado = EstadoApp.encabezados[colIdx];
      const esId = esColumnaIdentificador(encabezado);

      if (esFormula) {
        valoresFila.push(null);
      } else if (esId && EstadoApp.indiceFilaEnEdicion !== null) {
        const origId = EstadoApp.filasDatos[EstadoApp.indiceFilaEnEdicion]?.[colIdx];
        valoresFila.push(origId !== undefined && origId !== null ? origId : '');
      } else {
        const inputEl = DOM.dynamicFormFields.querySelector(`[name="field_${colIdx}"]`) as HTMLInputElement | HTMLSelectElement | null;
        const rawVal = inputEl ? inputEl.value.trim() : '';

        // Validación 1: Lista Desplegable
        if (opciones && opciones.length > 0 && rawVal !== '') {
          const isValidOption = opciones.some(opt => opt.toLowerCase() === rawVal.toLowerCase());
          if (!isValidOption) {
            mostrarNotificacionToast(`El valor para "${encabezado}" debe pertenecer a la lista`, "danger");
            marcarCampoInvalido(colIdx);
            return;
          }
        }

        // Validación 2: Unicidad de ID
        if (esId) {
          if (rawVal === '') {
            mostrarNotificacionToast(`El campo "${encabezado}" no puede estar vacío`, "danger");
            marcarCampoInvalido(colIdx);
            return;
          }

          const isDuplicate = EstadoApp.filasDatos.some((r, idx) => {
            if (EstadoApp.indiceFilaEnEdicion !== null && idx === EstadoApp.indiceFilaEnEdicion) {
              return false;
            }
            return String(r[colIdx] || '').trim().toLowerCase() === rawVal.toLowerCase();
          });

          if (isDuplicate) {
            mostrarNotificacionToast(`El ID "${rawVal}" ya existe. Los IDs deben ser únicos.`, "danger");
            marcarCampoInvalido(colIdx);
            return;
          }
        }

        // Validación 3: Formato Correo
        if (tipoCol === 'email' && rawVal !== '') {
          const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
          if (!emailRegex.test(rawVal)) {
            mostrarNotificacionToast(`El campo "${encabezado}" debe tener formato de correo válido`, "danger");
            marcarCampoInvalido(colIdx);
            return;
          }
        }

        // Validación 4: Formato Numérico
        if (tipoCol === 'number' && rawVal !== '') {
          const num = Number(rawVal);
          if (isNaN(num)) {
            mostrarNotificacionToast(`El campo "${encabezado}" debe ser numérico`, "danger");
            marcarCampoInvalido(colIdx);
            return;
          }
          valoresFila.push(num);
        } else {
          valoresFila.push(rawVal);
        }
      }
    }

    mostrarNotificacionToast("Enviando datos a Excel...", "info");

    const btnSave = DOM.dynamicRecordForm.querySelector('button[type="submit"]') as HTMLButtonElement;
    if (btnSave) btnSave.disabled = true;

    let success = false;
    if (EstadoApp.indiceFilaEnEdicion === null) {
      success = await insertarFilaEnExcel(valoresFila);
    } else {
      success = await actualizarFilaEnExcel(EstadoApp.indiceFilaEnEdicion, valoresFila);
    }

    if (btnSave) btnSave.disabled = false;
    if (success) cerrarModalFormulario();
  } catch (err: any) {
    console.error("Error al procesar formulario:", err);
    mostrarNotificacionToast(`Error al procesar: ${err?.message || err}`, "danger");
  }
}

// ============================================================================
// 8️⃣ SUPERPODER 6: GUARDAR COMO VISTA EJECUTIVA EN EXCEL
// ============================================================================

/**
 * Crea una nueva hoja y tabla oficial en Excel con encabezado y estilo ejecutivo.
 */
async function generarVistaEjecutivaExcel(nombreVista: string): Promise<void> {
  try {
    mostrarNotificacionToast(`Creando reporte ejecutivo "${nombreVista}"...`, "info");

    await Excel.run(async (context: Excel.RequestContext) => {
      const visibleHeaders = EstadoApp.indicesColumnasVisibles.map(i => EstadoApp.encabezados[i]);

      const visibleRows = EstadoApp.indicesFiltrados.map(rowIdx => {
        const row = EstadoApp.filasDatos[rowIdx];
        return EstadoApp.indicesColumnasVisibles.map(colIdx => {
          const val = row[colIdx];
          return val === undefined ? null : val;
        });
      });

      // Crear nueva hoja con color corporativo
      const newSheet = context.workbook.worksheets.add(nombreVista);
      newSheet.tabColor = "#4F46E5";

      const nowStr = new Date().toLocaleString('es-MX', { 
        year: 'numeric', month: 'short', day: 'numeric', 
        hour: '2-digit', minute: '2-digit' 
      });
      const filterInfo = EstadoApp.consultaBusquedaActiva ? `Filtro: "${EstadoApp.consultaBusquedaActiva}"` : 'Filtro: Ninguno (Todos)';
      const sortInfo = EstadoApp.indiceColumnaOrden !== null 
        ? `Orden: ${EstadoApp.encabezados[EstadoApp.indiceColumnaOrden]} (${EstadoApp.direccionOrden.toUpperCase()})` 
        : 'Orden original';

      // Fila 2: Título Ejecutivo
      const titleCell = newSheet.getRange("B2");
      titleCell.values = [[`📊 REPORTE DE VISTA: ${nombreVista.toUpperCase()}`]];
      titleCell.format.font.size = 14;
      titleCell.format.font.bold = true;
      titleCell.format.font.color = "#0F172A";

      // Fila 3: Metadatos
      const subtitleCell = newSheet.getRange("B3");
      subtitleCell.values = [[`Generado: ${nowStr}  |  Origen: ${EstadoApp.nombreTablaActiva}  |  ${filterInfo}  |  ${sortInfo}`]];
      subtitleCell.format.font.size = 9;
      subtitleCell.format.font.color = "#64748B";

      // Fila 5: Tabla Oficial
      const startRow = 4;
      const startCol = 1;
      const numRows = visibleRows.length > 0 ? visibleRows.length + 1 : 2;
      const numCols = visibleHeaders.length;

      const targetRange = newSheet.getRangeByIndexes(startRow, startCol, numRows, numCols);
      const dataMatrix = [
        visibleHeaders,
        ...(visibleRows.length > 0 ? visibleRows : [visibleHeaders.map(() => '')])
      ];
      targetRange.values = dataMatrix;

      const newTable = newSheet.tables.add(targetRange, true);
      newTable.name = nombreVista;
      newTable.style = "TableStyleMedium9";

      newSheet.getUsedRange().format.autofitColumns();
      newSheet.activate();

      await context.sync();

      mostrarNotificacionToast(`✅ Vista "${nombreVista}" creada con éxito`, "success");

      await cargarTablasDelLibro();
      cambiarTablaActiva(nombreVista);
    });
  } catch (error: any) {
    console.error("Error al guardar vista en Excel:", error);
    const msg = error?.message || "Error al crear la hoja en Excel";
    mostrarNotificacionToast(`⚠️ Error al guardar vista: ${msg}`, "danger");
  }
}

// ============================================================================
// 9️⃣ MODALES AUXILIARES: FICHA 360°, CONFIRMACIÓN Y VISTA
// ============================================================================

function abrirModalFicha360(rowIndex: number): void {
  EstadoApp.indiceFilaSeleccionada = rowIndex;
  const row = EstadoApp.filasDatos[rowIndex];
  const primaryTitle = row[0] !== undefined && row[0] !== null ? String(row[0]) : `Registro #${rowIndex + 1}`;

  if (DOM.modalDetailTitle) DOM.modalDetailTitle.textContent = `${primaryTitle} • Detalle 360°`;
  if (DOM.detailContent) {
    DOM.detailContent.innerHTML = '';
    EstadoApp.encabezados.forEach((encabezado: string, colIdx: number) => {
      const val = row[colIdx];
      const tipoCol = EstadoApp.tiposColumna[colIdx];
      const formattedVal = formatearValorCelda(val, tipoCol);

      const item = document.createElement('div');
      item.className = 'detail-item';
      item.innerHTML = `
        <span class="detail-key">${encabezado}</span>
        <span class="detail-val">${formattedVal}</span>
      `;
      DOM.detailContent.appendChild(item);
    });
  }

  DOM.modalDetail?.classList.remove('hidden');
}

function cerrarModalFicha360(): void {
  DOM.modalDetail?.classList.add('hidden');
  EstadoApp.indiceFilaSeleccionada = null;
}

function abrirModalConfirmacionEliminar(rowIndex: number): void {
  EstadoApp.indiceFilaAEliminar = rowIndex;
  DOM.modalConfirm?.classList.remove('hidden');
}

function cerrarModalConfirmacionEliminar(): void {
  DOM.modalConfirm?.classList.add('hidden');
  EstadoApp.indiceFilaAEliminar = null;
}

function confirmarEliminarRegistro(): void {
  if (EstadoApp.indiceFilaAEliminar !== null) {
    const idx = EstadoApp.indiceFilaAEliminar;
    cerrarModalConfirmacionEliminar();
    eliminarFilaEnExcel(idx);
  }
}

function generarNombreVistaPorDefecto(): string {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  const dd = pad(now.getDate());
  const mm = pad(now.getMonth() + 1);
  const yy = String(now.getFullYear()).slice(-2);
  const hh = pad(now.getHours());
  const min = pad(now.getMinutes());
  const ss = pad(now.getSeconds());
  return `Vista_${dd}${mm}${yy}${hh}${min}${ss}`;
}

function abrirModalGuardarVista(): void {
  if (!DOM.inputViewName || !DOM.saveViewSummaryBox) return;

  const defaultName = generarNombreVistaPorDefecto();
  DOM.inputViewName.value = defaultName;
  DOM.inputViewName.style.borderColor = '';

  const totalFiltered = EstadoApp.indicesFiltrados.length;
  const visibleCols = EstadoApp.indicesColumnasVisibles.length;
  const totalCols = EstadoApp.encabezados.length;
  const searchFilter = EstadoApp.consultaBusquedaActiva ? `"${EstadoApp.consultaBusquedaActiva}"` : 'Ninguno (Todos)';

  DOM.saveViewSummaryBox.innerHTML = `
    <div class="view-stat-row">
      <span class="view-stat-label">Tabla Origen:</span>
      <span class="view-stat-val">${EstadoApp.nombreTablaActiva}</span>
    </div>
    <div class="view-stat-row">
      <span class="view-stat-label">Registros a Exportar:</span>
      <span class="view-stat-val">${totalFiltered} filas</span>
    </div>
    <div class="view-stat-row">
      <span class="view-stat-label">Columnas Incluidas:</span>
      <span class="view-stat-val">${visibleCols} de ${totalCols} cols</span>
    </div>
    <div class="view-stat-row">
      <span class="view-stat-label">Filtro de Búsqueda:</span>
      <span class="view-stat-val">${searchFilter}</span>
    </div>
  `;

  DOM.modalSaveView?.classList.remove('hidden');
  setTimeout(() => DOM.inputViewName.focus(), 100);
}

function cerrarModalGuardarVista(): void {
  DOM.modalSaveView?.classList.add('hidden');
}

async function confirmarGuardarVista(): Promise<void> {
  const rawName = (DOM.inputViewName?.value || '').trim();

  if (!rawName) {
    mostrarNotificacionToast("Debes ingresar un nombre para la vista", "danger");
    if (DOM.inputViewName) DOM.inputViewName.focus();
    return;
  }

  const invalidCharsRegex = /[\\/?*:[\]]/;
  if (invalidCharsRegex.test(rawName)) {
    mostrarNotificacionToast('El nombre no puede contener: \\ / ? * : [ ]', 'danger');
    if (DOM.inputViewName) DOM.inputViewName.focus();
    return;
  }

  if (rawName.length > 31) {
    mostrarNotificacionToast('El nombre no puede exceder 31 caracteres', 'danger');
    if (DOM.inputViewName) DOM.inputViewName.focus();
    return;
  }

  if (DOM.btnConfirmSaveView) DOM.btnConfirmSaveView.disabled = true;

  cerrarModalGuardarVista();
  await generarVistaEjecutivaExcel(rawName);

  if (DOM.btnConfirmSaveView) DOM.btnConfirmSaveView.disabled = false;
}

// ============================================================================
// 🔟 FORMATEO, RENDERIZADO TABULAR Y UTILIDADES UI
// ============================================================================

function formatearFechaExcel(val: ValorCelda): string {
  if (val === null || val === undefined || val === '') return '';
  
  if (typeof val === 'number') {
    if (val > 1000 && val < 100000) {
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      const y = date.getUTCFullYear();
      const m = String(date.getUTCMonth() + 1).padStart(2, '0');
      const d = String(date.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return String(val);
  }

  const str = String(val).trim();
  if (str.includes('T')) return str.split('T')[0];
  if (str.includes(' ')) {
    const part = str.split(' ')[0];
    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(part) || /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(part)) {
      return formatearFechaExcel(part);
    }
  }

  const numVal = Number(str);
  if (!isNaN(numVal) && numVal > 1000 && numVal < 100000 && !str.includes('-') && !str.includes('/')) {
    const date = new Date(Math.round((numVal - 25569) * 86400 * 1000));
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const d = String(date.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const matchISO = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (matchISO) return `${matchISO[1]}-${matchISO[2].padStart(2, '0')}-${matchISO[3].padStart(2, '0')}`;

  const matchLat = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (matchLat) return `${matchLat[3]}-${matchLat[2].padStart(2, '0')}-${matchLat[1].padStart(2, '0')}`;

  return str;
}

function inferirTiposDeColumna(): void {
  const colCount = EstadoApp.encabezados.length;
  EstadoApp.tiposColumna = new Array(colCount).fill('text');
  EstadoApp.esColumnaCalculada = new Array(colCount).fill(false);

  for (let c = 0; c < colCount; c++) {
    const headerName = (EstadoApp.encabezados[c] || '').toLowerCase();
    const isHeaderDateHint = headerName.includes('fecha') || headerName.includes('date') || headerName.includes('fec') || headerName.includes('emision');

    const hasFormulas = EstadoApp.datosFormulas.some(row => {
      const cellFormula = row[c];
      return typeof cellFormula === 'string' && cellFormula.startsWith('=');
    });
    EstadoApp.esColumnaCalculada[c] = hasFormulas;

    let isExcelDateFormat = false;
    if (EstadoApp.datosFormatosNumero && EstadoApp.datosFormatosNumero.length > 0) {
      const formatSamples = EstadoApp.datosFormatosNumero.map(r => String(r[c] || '')).filter(Boolean);
      isExcelDateFormat = formatSamples.some(fmt => {
        const lower = fmt.toLowerCase();
        return (lower.includes('yy') || lower.includes('mm') || lower.includes('dd') || lower.includes('aaaa')) && !lower.includes('general');
      });
    }

    const nonNullValues = EstadoApp.filasDatos
      .map(row => row[c])
      .filter(val => val !== null && val !== undefined && String(val).trim() !== '');

    if (nonNullValues.length === 0) {
      EstadoApp.tiposColumna[c] = (isExcelDateFormat || isHeaderDateHint) ? 'date' : 'text';
      continue;
    }

    let numberCount = 0;
    let dateCount = 0;
    let emailCount = 0;
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    for (const val of nonNullValues) {
      if (typeof val === 'number') {
        if ((isExcelDateFormat || isHeaderDateHint) && val > 10000 && val < 90000) {
          dateCount++;
        } else {
          numberCount++;
        }
      } else if (typeof val === 'string') {
        const trimmed = val.trim();
        if ((isExcelDateFormat || isHeaderDateHint) && !isNaN(Number(trimmed)) && Number(trimmed) > 10000 && Number(trimmed) < 90000) {
          dateCount++;
        } else if (emailRegex.test(trimmed)) {
          emailCount++;
        } else if (trimmed.includes('T') && !isNaN(Date.parse(trimmed))) {
          dateCount++;
        } else if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(trimmed) || /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(trimmed)) {
          dateCount++;
        } else if (!isNaN(Number(trimmed)) && trimmed !== '') {
          numberCount++;
        }
      }
    }

    const total = nonNullValues.length;
    if (isExcelDateFormat || isHeaderDateHint || dateCount / total >= 0.5) {
      EstadoApp.tiposColumna[c] = 'date';
    } else if (emailCount / total >= 0.5) {
      EstadoApp.tiposColumna[c] = 'email';
    } else if (numberCount / total >= 0.6) {
      EstadoApp.tiposColumna[c] = 'number';
    } else {
      EstadoApp.tiposColumna[c] = 'text';
    }
  }
}

function formatearValorCelda(val: ValorCelda, tipo: TipoColumna, query: string = ''): string {
  if (val === null || val === undefined || val === '') return '<span style="color: var(--text-muted);">-</span>';

  if (tipo === 'email' && val) {
    const emailStr = String(val);
    const highlightedVal = query ? resaltarCoincidencia(emailStr, query) : emailStr;
    return `<a href="mailto:${emailStr}" class="email-link" onclick="event.stopPropagation()">${highlightedVal}</a>`;
  }

  if (tipo === 'date') {
    const dateStr = formatearFechaExcel(val);
    return query ? resaltarCoincidencia(dateStr, query) : dateStr;
  }

  if (tipo === 'number' && typeof val === 'number') {
    const formatted = val.toLocaleString('es-MX');
    return query ? resaltarCoincidencia(formatted, query) : formatted;
  }

  const strVal = String(val);
  return query ? resaltarCoincidencia(strVal, query) : strVal;
}

function renderizarVistaTabular(): void {
  if (!DOM.recordsContainer) return;

  if (EstadoApp.filasDatos.length === 0) {
    renderizarEstadoVacio("Esta tabla no contiene registros.");
    return;
  }

  if (EstadoApp.indicesFiltrados.length === 0) {
    renderizarEstadoVacio(`No se encontraron coincidencias para "${EstadoApp.consultaBusquedaActiva}"`);
    return;
  }

  DOM.recordsContainer.innerHTML = '';

  const wrapper = document.createElement('div');
  wrapper.className = 'table-responsive-wrapper';

  const table = document.createElement('table');
  table.className = 'modern-data-table';

  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');

  EstadoApp.encabezados.forEach((encabezado: string, colIdx: number) => {
    if (!EstadoApp.indicesColumnasVisibles.includes(colIdx)) return;

    const th = document.createElement('th');
    th.className = 'th-sortable';
    if (EstadoApp.tiposColumna[colIdx] === 'number') th.classList.add('col-num-header');

    let sortIconHtml = '<span class="sort-idle-icon">⇅</span>';
    if (EstadoApp.indiceColumnaOrden === colIdx) {
      const arrow = EstadoApp.direccionOrden === 'asc' ? '▲' : '▼';
      sortIconHtml = `<span class="sort-icon active">${arrow}</span>`;
    }

    th.innerHTML = `
      <div class="th-content">
        <span>${encabezado}</span>
        ${sortIconHtml}
      </div>
    `;

    th.title = `Ordenar por ${encabezado}`;
    th.addEventListener('click', () => alternarOrdenColumna(colIdx));
    trHead.appendChild(th);
  });

  const thActions = document.createElement('th');
  thActions.className = 'th-actions';
  thActions.textContent = 'Acciones';
  trHead.appendChild(thActions);

  thead.appendChild(trHead);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  EstadoApp.indicesFiltrados.forEach((originalRowIndex: number) => {
    const row = EstadoApp.filasDatos[originalRowIndex];
    const tr = document.createElement('tr');

    if (EstadoApp.indiceFilaSeleccionada === originalRowIndex) {
      tr.classList.add('selected-row');
    }

    EstadoApp.encabezados.forEach((_, colIdx: number) => {
      if (!EstadoApp.indicesColumnasVisibles.includes(colIdx)) return;

      const td = document.createElement('td');
      const val = row[colIdx];
      const tipoCol = EstadoApp.tiposColumna[colIdx];
      const esId = esColumnaIdentificador(EstadoApp.encabezados[colIdx]);

      if (colIdx === 0 || esId) {
        td.classList.add('col-primary');
        if (val) {
          td.title = `ID: ${val} (Clic para copiar)`;
          td.style.cursor = 'pointer';
          td.addEventListener('click', (e: MouseEvent) => {
            e.stopPropagation();
            if (navigator.clipboard) {
              navigator.clipboard.writeText(String(val)).then(() => {
                mostrarNotificacionToast(`📋 ID "${val}" copiado al portapapeles`, 'info');
              });
            }
          });
        }
      }
      if (tipoCol === 'number') td.classList.add('col-number');

      td.innerHTML = formatearValorCelda(val, tipoCol, EstadoApp.consultaBusquedaActiva);
      tr.appendChild(td);
    });

    const tdActions = document.createElement('td');
    tdActions.className = 'td-actions';
    tdActions.innerHTML = `
      <div class="table-action-btns">
        <button class="action-icon-btn view-btn" title="Ver ficha 360°" data-index="${originalRowIndex}">👁️</button>
        <button class="action-icon-btn edit-btn" title="Editar registro" data-index="${originalRowIndex}">✏️</button>
        <button class="action-icon-btn delete-btn" title="Eliminar registro" data-index="${originalRowIndex}">🗑️</button>
      </div>
    `;

    const btnView = tdActions.querySelector('.view-btn') as HTMLButtonElement;
    const btnEdit = tdActions.querySelector('.edit-btn') as HTMLButtonElement;
    const btnDel = tdActions.querySelector('.delete-btn') as HTMLButtonElement;

    btnView?.addEventListener('click', (e: MouseEvent) => {
      e.stopPropagation();
      abrirModalFicha360(originalRowIndex);
    });

    btnEdit?.addEventListener('click', (e: MouseEvent) => {
      e.stopPropagation();
      abrirModalEditarRegistro(originalRowIndex);
    });

    btnDel?.addEventListener('click', (e: MouseEvent) => {
      e.stopPropagation();
      abrirModalConfirmacionEliminar(originalRowIndex);
    });

    tr.appendChild(tdActions);

    tr.addEventListener('click', () => {
      tbody.querySelectorAll('tr').forEach((r: HTMLTableRowElement) => r.classList.remove('selected-row'));
      tr.classList.add('selected-row');
      EstadoApp.indiceFilaSeleccionada = originalRowIndex;
    });

    tr.addEventListener('dblclick', () => {
      abrirModalFicha360(originalRowIndex);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  wrapper.appendChild(table);
  DOM.recordsContainer.appendChild(wrapper);
}

function actualizarContador(): void {
  if (!DOM.counterText) return;
  const showing = EstadoApp.indicesFiltrados.length;
  const total = EstadoApp.filasDatos.length;
  const visibleCols = EstadoApp.indicesColumnasVisibles.length;
  const totalCols = EstadoApp.encabezados.length;
  DOM.counterText.textContent = `Mostrando ${showing} de ${total} registros • ${visibleCols}/${totalCols} cols`;
}

function mostrarEsqueletosDeCarga(): void {
  if (!DOM.recordsContainer) return;
  DOM.recordsContainer.innerHTML = `
    <div class="skeleton-card"></div>
    <div class="skeleton-card"></div>
    <div class="skeleton-card"></div>
  `;
}

function renderizarEstadoVacio(mensaje: string): void {
  if (!DOM.recordsContainer) return;
  DOM.recordsContainer.innerHTML = `
    <div class="empty-state">
      <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
      <p>${mensaje}</p>
    </div>
  `;
}

function mostrarNotificacionToast(mensaje: string, tipo: TipoNotificacion = 'info'): void {
  if (!DOM.toastContainer) return;
  const toast = document.createElement('div');
  toast.className = 'toast';

  const icon = tipo === 'success' ? '✅' : tipo === 'danger' ? '❌' : 'ℹ️';
  toast.innerHTML = `<span>${icon}</span> <span>${mensaje}</span>`;

  DOM.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 2600);
}
