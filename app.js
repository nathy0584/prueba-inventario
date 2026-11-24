/**
 * Control de Inventarios, Salidas y Mermas - Bioclinic
 * Archivo: app.js
 * Arquitectura: 100% Gratuita (GitHub Pages + OneDrive Local Sync / File System Access API + SheetJS + jsPDF)
 */

var FILE_HANDLE = null;
var WORKBOOK = null;
var NOMBRE_ARCHIVO_EXCEL = "Control_Inventarios_Bioclinic.xlsx";

var USUARIO_LOGUEADO = "";
var ROL_LOGUEADO = "";
var SESION_ES_UTM = false;
var MAPA_SUCURSAL_PLANILLA = {};
var LOTE_MANUAL_ITEMS = [];
var PEDIDOS_CARGADOS_CACHE = [];
var INSUMOS_MONITOR_CACHE = [];
var ULTIMO_REPORTE_MENSUAL = null;

// ============================================================================
// 1. CONEXIÓN, LECTURA Y GUARDADO DIRECTO EN ONEDRIVE (.xlsx)
// ============================================================================

async function conectarArchivoOneDrive() {
  if ('showOpenFilePicker' in window) {
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{
          description: 'Archivos Excel (*.xlsx)',
          accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
        }],
        multiple: false
      });
      FILE_HANDLE = handle;
      NOMBRE_ARCHIVO_EXCEL = handle.name;
      const file = await FILE_HANDLE.getFile();
      const buffer = await file.arrayBuffer();
      WORKBOOK = XLSX.read(buffer, { type: 'array', cellDates: true });
      activarFormularioLogin(NOMBRE_ARCHIVO_EXCEL, true);
    } catch (err) {
      if (err.name !== 'AbortError') {
        mostrarNotificacion("No se pudo abrir el archivo: " + err.message, "error");
      }
    }
  } else {
    // Fallback para navegadores sin File System Access API (Firefox / Safari)
    document.getElementById("inputFallbackExcel").click();
  }
}

function cargarExcelFallback(event) {
  const file = event.target.files[0];
  if (!file) return;
  NOMBRE_ARCHIVO_EXCEL = file.name;
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    WORKBOOK = XLSX.read(data, { type: 'array', cellDates: true });
    activarFormularioLogin(NOMBRE_ARCHIVO_EXCEL, false);
  };
  reader.readAsArrayBuffer(file);
}

function activarFormularioLogin(nombreArchivo, escrituraDirecta) {
  const badge = document.getElementById("badgeEstadoArchivo");
  const lbl = document.getElementById("lblNombreArchivo");
  badge.className = "badge bg-success";
  badge.innerText = escrituraDirecta ? "Sincronización Directa" : "Modo Descarga";
  lbl.innerHTML = `<i class="bi bi-file-earmark-excel-fill me-1"></i> ${nombreArchivo}`;
  lbl.classList.remove("d-none");

  document.getElementById("loginUsuario").disabled = false;
  document.getElementById("loginClave").disabled = false;
  document.getElementById("btnIngresarSistema").disabled = false;
  mostrarNotificacion("Archivo Excel conectado correctamente. Ingrese sus credenciales.", "success");
}

async function guardarCambiosEnOneDrive() {
  if (!WORKBOOK) return false;
  const banner = document.getElementById("estadoGuardadoBanner");
  if (banner) {
    banner.className = "badge bg-warning text-dark";
    banner.innerHTML = '<i class="bi bi-arrow-repeat me-1"></i> Guardando en OneDrive...';
  }

  try {
    const wbout = XLSX.write(WORKBOOK, { bookType: 'xlsx', type: 'array' });
    if (FILE_HANDLE) {
      const writable = await FILE_HANDLE.createWritable();
      await writable.write(wbout);
      await writable.close();
      if (banner) {
        banner.className = "badge bg-success";
        banner.innerHTML = '<i class="bi bi-check-circle me-1"></i> Guardado en OneDrive';
      }
      return true;
    } else {
      descargarCopiaExcel();
      if (banner) {
        banner.className = "badge bg-info text-dark";
        banner.innerHTML = '<i class="bi bi-download me-1"></i> Descargado (.xlsx)';
      }
      return true;
    }
  } catch (err) {
    console.error("Error al guardar en Excel:", err);
    if (banner) {
      banner.className = "badge bg-danger";
      banner.innerHTML = '<i class="bi bi-exclamation-triangle me-1"></i> Error al guardar (¿Excel abierto?)';
    }
    mostrarNotificacion("No se pudo escribir en el archivo. Asegúrese de no tener el Excel abierto en otro programa.", "error");
    return false;
  }
}

async function recargarDesdeOneDrive() {
  if (!FILE_HANDLE) {
    mostrarNotificacion("En modo descarga, vuelva a seleccionar el archivo desde el inicio de sesión.", "info");
    return;
  }
  try {
    const file = await FILE_HANDLE.getFile();
    const buffer = await file.arrayBuffer();
    WORKBOOK = XLSX.read(buffer, { type: 'array', cellDates: true });
    cargarProveedoresEnSelect();
    actualizarCatalogoDinamicoEnVivo();
    recargarTablaMonitor();
    mostrarNotificacion("Datos recargados desde OneDrive.", "success");
  } catch (err) {
    mostrarNotificacion("Error al recargar: " + err.message, "error");
  }
}

function descargarCopiaExcel() {
  if (!WORKBOOK) return;
  XLSX.writeFile(WORKBOOK, NOMBRE_ARCHIVO_EXCEL);
}

// ============================================================================
// 2. UTILIDADES DE LECTURA/ESCRITURA DE HOJAS (MATRIZ 2D)
// ============================================================================

function obtenerMatrizHoja(nombreHoja) {
  if (!WORKBOOK) return [];
  const nombreReal = WORKBOOK.SheetNames.find(s => s.trim().toUpperCase() === nombreHoja.trim().toUpperCase());
  if (!nombreReal) return [];
  const sheet = WORKBOOK.Sheets[nombreReal];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
}

function guardarMatrizHoja(nombreHoja, matriz) {
  if (!WORKBOOK) return;
  let nombreReal = WORKBOOK.SheetNames.find(s => s.trim().toUpperCase() === nombreHoja.trim().toUpperCase());
  const nuevaHoja = XLSX.utils.aoa_to_sheet(matriz);
  if (nombreReal) {
    WORKBOOK.Sheets[nombreReal] = nuevaHoja;
  } else {
    XLSX.utils.book_append_sheet(WORKBOOK, nuevaHoja, nombreHoja);
  }
}

function parseNumeroMoneda(val) {
  if (typeof val === "number") return val;
  if (val === undefined || val === null || val === "") return 0;
  let txt = val.toString().trim().replace(/[^0-9.,-]/g, "");
  if (txt.indexOf('.') !== -1 && txt.indexOf(',') === -1) {
    txt = txt.replace(/\./g, "");
  } else if (txt.indexOf(',') !== -1) {
    txt = txt.replace(/\./g, "").replace(',', '.');
  }
  return parseFloat(txt) || 0;
}

function formatearFechaDDMMYYYY(fechaObj) {
  if (!fechaObj) return "";
  const d = (fechaObj instanceof Date) ? fechaObj : new Date(fechaObj);
  if (isNaN(d.getTime())) return fechaObj.toString().substring(0, 10);
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function formatearFechaHora(fechaObj) {
  if (!fechaObj) return "";
  const d = (fechaObj instanceof Date) ? fechaObj : new Date(fechaObj);
  if (isNaN(d.getTime())) return fechaObj.toString();
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${dd}-${mm}-${yyyy} ${hh}:${min}`;
}

function mostrarNotificacion(mensaje, tipo = "success") {
  var container = document.getElementById("toastContainer");
  var toast = document.createElement("div");
  toast.className = `bioclinic-toast toast-${tipo}`;
  var icono = "bi-check-circle-fill text-success";
  if (tipo === "error") icono = "bi-exclamation-triangle-fill text-danger";
  if (tipo === "info") icono = "bi-info-circle-fill text-info";
  if (tipo === "warning") icono = "bi-exclamation-circle-fill text-warning";
  toast.innerHTML = `<i class="bi ${icono} fs-4"></i> <div class="fw-semibold text-dark">${mensaje}</div>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    setTimeout(() => toast.remove(), 300);
  }, 4500);
}

function formatMonto(val) {
  var num = parseFloat(val);
  if (isNaN(num)) return "$0";
  return new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', minimumFractionDigits: 0 }).format(num);
}

// ============================================================================
// 3. AUTENTICACIÓN Y ROLES (LECTURA DE HOJA 'Personal')
// ============================================================================

function ejecutarLogin() {
  if (!WORKBOOK) {
    mostrarNotificacion("Primero conecte el archivo Excel de OneDrive.", "error");
    return;
  }

  var user = document.getElementById("loginUsuario").value.trim();
  var pass = document.getElementById("loginClave").value.trim();
  var errBox = document.getElementById("errorLogin");
  errBox.classList.add("d-none");

  var filasPersonal = obtenerMatrizHoja('Personal');
  if (filasPersonal.length <= 1) {
    errBox.innerText = "No se encontró la pestaña 'Personal' o está vacía en el Excel.";
    errBox.classList.remove("d-none");
    return;
  }

  var usuarioEncontrado = null;
  var filaIndex = -1;

  for (var i = 1; i < filasPersonal.length; i++) {
    var uCelda = filasPersonal[i][0] !== undefined ? filasPersonal[i][0].toString().trim() : "";
    var pCelda = filasPersonal[i][1] !== undefined ? filasPersonal[i][1].toString().trim() : "";
    if (uCelda === user && pCelda === pass) {
      usuarioEncontrado = {
        usuario: uCelda,
        nombre: uCelda,
        rol: filasPersonal[i][2] ? filasPersonal[i][2].toString().trim() : "usuario"
      };
      filaIndex = i;
      break;
    }
  }

  if (!usuarioEncontrado) {
    errBox.innerText = "Credenciales de acceso incorrectas.";
    errBox.classList.remove("d-none");
    return;
  }

  USUARIO_LOGUEADO = usuarioEncontrado.nombre;
  ROL_LOGUEADO = usuarioEncontrado.rol;

  var userUpper = USUARIO_LOGUEADO.toUpperCase();
  var rolUpper = ROL_LOGUEADO.toUpperCase();
  var esDT = (userUpper.includes("DIRECTOR") || userUpper.includes("TECNICO") || userUpper === "DT" || rolUpper.includes("DT"));

  var sucursalesTexto = filasPersonal[filaIndex][3] ? filasPersonal[filaIndex][3].toString().trim() : "";
  var listaSucursales = sucursalesTexto !== ""
    ? sucursalesTexto.split(",").map(s => s.trim()).filter(Boolean)
    : (esDT ? [] : [USUARIO_LOGUEADO]);

  MAPA_SUCURSAL_PLANILLA = {};
  listaSucursales.forEach(function(suc) {
    var sucUpper = suc.toUpperCase();
    if (sucUpper === "ADMINISTRACION") {
      MAPA_SUCURSAL_PLANILLA[suc] = "Solicitud Administracion";
    } else if (sucUpper.startsWith("UTM")) {
      MAPA_SUCURSAL_PLANILLA[suc] = suc;
    } else if (sucUpper === "LABORATORIO") {
      MAPA_SUCURSAL_PLANILLA[suc] = "Solicitud LAB";
    } else if (sucUpper.startsWith("ADM ") || sucUpper.startsWith("ADM_") || sucUpper === "BIOEXPRESS") {
      MAPA_SUCURSAL_PLANILLA[suc] = "Solicitud ADM";
    } else if (sucUpper === "MAHO" || sucUpper === "ESACHS" || sucUpper === "BODEGA") {
      MAPA_SUCURSAL_PLANILLA[suc] = "Solicitud MAHO";
    } else {
      MAPA_SUCURSAL_PLANILLA[suc] = suc;
    }
  });

  var primeraSucursal = listaSucursales[0] || "";
  SESION_ES_UTM = primeraSucursal.toUpperCase().startsWith("UTM");

  document.getElementById("txtNombreUsuario").innerText = USUARIO_LOGUEADO;
  document.getElementById("seccionLogin").classList.add("d-none");
  document.getElementById("seccionApp").classList.remove("d-none");

  // Cargar selectores y catálogos iniciales
  cargarProveedoresEnSelect();
  actualizarCatalogoDinamicoEnVivo();

  var selectSuc = document.getElementById("pedidoSucursalOrigen");
  if (selectSuc) {
    selectSuc.innerHTML = "";
    listaSucursales.forEach(suc => { selectSuc.innerHTML += `<option value="${suc}">${suc}</option>`; });
  }

  poblarSelectorSucursalesReporte(listaSucursales);

  var selectMovimiento = document.getElementById("sucursal");
  if (selectMovimiento) {
    selectMovimiento.innerHTML = '<option value="">Seleccione Destino...</option>';
    listaSucursales.forEach(suc => { selectMovimiento.innerHTML += `<option value="${suc}">${suc}</option>`; });
  }

  // Control de visibilidad de pestañas por Rol
  var listaItemsNav = document.querySelectorAll(".nav-tabs-custom .nav-link");
  var subseccionDT = document.getElementById("subseccionPreAprobacionDT");
  var subseccionBodega = document.getElementById("subseccionAdminPedidos");
  var tabSubbodegas = document.getElementById("tab-subbodegas");
  var tabReporteMensual = document.getElementById("tab-reporte-mensual");

  if (tabSubbodegas) {
    if (userUpper.includes("ADMIN") || rolUpper.includes("ADMIN") ||
        userUpper.includes("FINANZAS") || rolUpper.includes("FINANZAS") ||
        userUpper.includes("ADMINISTRACION") || rolUpper.includes("ADMINISTRACION")) {
      tabSubbodegas.style.setProperty("display", "block", "important");
    } else {
      tabSubbodegas.style.setProperty("display", "none", "important");
    }
  }

  if (userUpper.includes("ADMIN") || rolUpper.includes("ADMIN") || userUpper === "BODEGA") {
    for (var i = 0; i < 6; i++) {
      if (listaItemsNav[i]) listaItemsNav[i].style.setProperty("display", "block", "important");
    }
    if (tabReporteMensual) tabReporteMensual.style.setProperty("display", "block", "important");
    if (subseccionDT) subseccionDT.style.setProperty("display", "none", "important");
    if (subseccionBodega) subseccionBodega.style.setProperty("display", "block", "important");
    cambiarPestaña('panel-manual', document.getElementById("tab-manual"));

  } else if (userUpper.includes("FINANZAS") || rolUpper.includes("FINANZAS")) {
    if (listaItemsNav[0]) listaItemsNav[0].style.setProperty("display", "block", "important");
    if (listaItemsNav[1]) listaItemsNav[1].style.setProperty("display", "none", "important");
    if (listaItemsNav[2]) listaItemsNav[2].style.setProperty("display", "none", "important");
    if (listaItemsNav[3]) listaItemsNav[3].style.setProperty("display", "block", "important");
    if (listaItemsNav[4]) listaItemsNav[4].style.setProperty("display", "block", "important");
    if (listaItemsNav[5]) listaItemsNav[5].style.setProperty("display", "none", "important");
    if (tabReporteMensual) tabReporteMensual.style.setProperty("display", "none", "important");
    if (subseccionDT) subseccionDT.style.setProperty("display", "none", "important");
    if (subseccionBodega) subseccionBodega.style.setProperty("display", "none", "important");
    cambiarPestaña('panel-manual', document.getElementById("tab-manual"));

  } else if (esDT) {
    if (listaItemsNav[0]) listaItemsNav[0].style.setProperty("display", "none", "important");
    if (listaItemsNav[1]) listaItemsNav[1].style.setProperty("display", "block", "important");
    if (listaItemsNav[2]) listaItemsNav[2].style.setProperty("display", "none", "important");
    if (listaItemsNav[3]) listaItemsNav[3].style.setProperty("display", "block", "important");
    if (listaItemsNav[4]) listaItemsNav[4].style.setProperty("display", "block", "important");
    if (listaItemsNav[5]) listaItemsNav[5].style.setProperty("display", "none", "important");
    if (tabReporteMensual) tabReporteMensual.style.setProperty("display", "none", "important");
    if (subseccionDT) subseccionDT.style.setProperty("display", "block", "important");
    if (subseccionBodega) subseccionBodega.style.setProperty("display", "none", "important");
    cambiarPestaña('panel-aprobaciones', document.getElementById("tab-aprobaciones"));

  } else {
    for (var j = 0; j < 5; j++) {
      if (listaItemsNav[j]) listaItemsNav[j].style.setProperty("display", "none", "important");
    }
    if (listaItemsNav[5]) listaItemsNav[5].style.setProperty("display", "block", "important");
    if (tabReporteMensual) tabReporteMensual.style.setProperty("display", "block", "important");
    cambiarPestaña('panel-solicitud', document.getElementById("tab-solicitud"));
  }

  actualizarCatalogoPorSucursalUTM();
  mostrarNotificacion(`Sesión iniciada. Operador: ${USUARIO_LOGUEADO}`, "info");
}

function cargarProveedoresEnSelect() {
  var selectProv = document.getElementById("proveedor");
  if (!selectProv) return;
  selectProv.innerHTML = '<option value="">Seleccione Proveedor...</option>';
  var filas = obtenerMatrizHoja('Proveedores');
  for (var i = 1; i < filas.length; i++) {
    var rut = filas[i][0] ? filas[i][0].toString().trim() : "";
    var nombre = filas[i][1] ? filas[i][1].toString().trim() : "";
    var texto = (rut && nombre) ? `${rut} - ${nombre}` : (rut || nombre);
    if (texto) {
      selectProv.innerHTML += `<option value="${texto}">${texto}</option>`;
    }
  }
}

// ============================================================================
// 4. LECTURA Y CÁLCULO DE INVENTARIO (HOJA 'Inventario_Actual')
// ============================================================================

function obtenerInventarioActual() {
  var datosInsumos = obtenerMatrizHoja('Inventario_Actual');
  var listaInsumos = [];

  for (var i = 1; i < datosInsumos.length; i++) {
    if (datosInsumos[i][0] !== undefined && datosInsumos[i][0] !== "") {
      var stockBulnes = parseFloat(datosInsumos[i][16]) || 0;
      var stockPeninsula = parseFloat(datosInsumos[i][17]) || 0;
      // Si el Excel tiene fórmulas que SheetJS lee en caché, usamos columna I; si Bulnes/Península cambian, recalculamos
      var stockColI = parseFloat(datosInsumos[i][8]);
      var stockNum = !isNaN(stockColI) ? stockColI : (stockBulnes + stockPeninsula);

      var precioUnit = parseNumeroMoneda(datosInsumos[i][13]);
      var valorCeldaJ = parseNumeroMoneda(datosInsumos[i][9]);
      var totalValor = valorCeldaJ > 0 ? valorCeldaJ : (stockNum * precioUnit);

      var estadoVisual = datosInsumos[i][10] ? datosInsumos[i][10].toString() : (stockNum <= 0 ? "SIN STOCK" : "DISPONIBLE");
      var proyAnual = parseFloat(datosInsumos[i][14]) || 0;
      var proyTrimestral = parseFloat(datosInsumos[i][15]) || 0;

      listaInsumos.push({
        filaIndex: i,
        id: datosInsumos[i][0].toString().trim(),
        nombre: datosInsumos[i][1] ? datosInsumos[i][1].toString().trim() : "",
        categoria: datosInsumos[i][2] ? datosInsumos[i][2].toString().trim() : "General",
        um: datosInsumos[i][3] ? datosInsumos[i][3].toString().trim() : "UN",
        stock: stockNum,
        totalValor: totalValor,
        precioUnitario: precioUnit,
        estado: estadoVisual,
        proyeccionAnual: proyAnual,
        proyeccionTrimestral: proyTrimestral,
        stockBulnes: stockBulnes,
        stockPeninsula: stockPeninsula
      });
    }
  }
  return listaInsumos;
}

function actualizarCatalogoDinamicoEnVivo() {
  INSUMOS_MONITOR_CACHE = obtenerInventarioActual();
  var datalist = document.getElementById("listaInsumosDinamica");
  if (datalist) {
    datalist.innerHTML = "";
    INSUMOS_MONITOR_CACHE.forEach(function(insumo) {
      var opt = document.createElement("option");
      opt.value = insumo.id + " - " + insumo.nombre;
      opt.dataset.id = insumo.id;
      opt.dataset.nombre = insumo.nombre;
      opt.dataset.stock = insumo.stock;
      opt.dataset.um = insumo.um;
      opt.dataset.estado = insumo.estado;
      opt.dataset.categoria = insumo.categoria;
      datalist.appendChild(opt);
    });
  }
}

// ============================================================================
// 5. REGISTRO MANUAL DE LOTES (SALIDA, INGRESO, MERMA, AJUSTE)
// ============================================================================

function obtenerOptionInsumoSeleccionado() {
  var val = document.getElementById("txtInputInsumo").value.trim();
  if (!val) return null;
  var datalist = document.getElementById("listaInsumosDinamica");
  for (var i = 0; i < datalist.options.length; i++) {
    var opt = datalist.options[i];
    if (opt.value === val || opt.dataset.id === val) return opt;
  }
  return null;
}

function mostrarMetadataInsumo() {
  var box = document.getElementById("boxMetadata");
  var opt = obtenerOptionInsumoSeleccionado();
  if (!opt) {
    box.classList.add("d-none");
    return;
  }
  box.classList.remove("d-none");
  box.innerHTML = `<strong>Nombre:</strong> ${opt.dataset.nombre} | <strong>Stock:</strong> ${opt.dataset.stock} | <strong>U.M.:</strong> ${opt.dataset.um} | <strong>Estado:</strong> ${opt.dataset.estado}`;
  validarStockDisponible();
}

function validarStockDisponible() {
  var tipo = document.getElementById("tipoMovimiento").value;
  var opt = obtenerOptionInsumoSeleccionado();
  var valorCant = parseFloat(document.getElementById("cantidad").value) || 0;
  var alerta = document.getElementById("alertaStockInsuficiente");

  if ((tipo === "SALIDA" || tipo === "MERMA") && opt) {
    var stockActual = parseFloat(opt.dataset.stock) || 0;
    if (valorCant > stockActual) {
      alerta.classList.remove("d-none");
      return;
    }
  }
  alerta.classList.add("d-none");
}

function alternarCamposDinamicos() {
  var tipo = document.getElementById("tipoMovimiento").value;
  var prov = document.getElementById("bloqueProveedor");
  var sClass = document.getElementById("bloqueSucursal").classList;
  var fact = document.getElementById("bloqueFactura");
  var bodegaClass = document.getElementById("bloqueBodegaOrigen").classList;
  
  prov.classList.add("d-none");
  sClass.add("d-none");
  fact.classList.add("d-none");
  bodegaClass.add("d-none");
  
  if (tipo === "INGRESO") {
    prov.classList.remove("d-none");
    fact.classList.remove("d-none");
    bodegaClass.remove("d-none");
  } else if (tipo === "SALIDA") {
    sClass.remove("d-none");
    bodegaClass.remove("d-none");
  } else if (tipo === "MERMA") {
    sClass.remove("d-none");
    bodegaClass.remove("d-none");
  } else if (tipo === "AJUSTE") {
    bodegaClass.remove("d-none");
  }
  validarStockDisponible();
}

function agregarInsumoAlLoteManual() {
  var opt = obtenerOptionInsumoSeleccionado();
  var cantInput = document.getElementById("cantidad");
  var vencInput = document.getElementById("vencimientoIndividual");
  var cantidad = parseFloat(cantInput.value) || 0;
  var vencimientoVal = vencInput.value ? vencInput.value : "S/V";
  
  if (!opt || cantidad <= 0) {
    mostrarNotificacion("Seleccione un insumo válido de la lista y una cantidad mayor a 0.", "error");
    return;
  }
  
  if (!document.getElementById("alertaStockInsuficiente").classList.contains("d-none")) {
    mostrarNotificacion("No puedes añadir cantidades que superen el stock disponible.", "error");
    return;
  }
  
  LOTE_MANUAL_ITEMS.push({ 
    idInsumo: opt.dataset.id, 
    name: opt.dataset.nombre, 
    categoria: opt.dataset.categoria || "General",
    um: opt.dataset.um || "UN",
    cantidad: cantidad, 
    fechaVenc: vencimientoVal 
  });
  
  document.getElementById("txtInputInsumo").value = "";
  cantInput.value = "1";
  vencInput.value = "";
  document.getElementById("boxMetadata").classList.add("d-none");
  renderizarTablaLoteManual();
}

function eliminarItemLoteManual(index) {
  LOTE_MANUAL_ITEMS.splice(index, 1);
  renderizarTablaLoteManual();
}

function renderizarTablaLoteManual() {
  var tbody = document.getElementById("tbodyLoteManual");
  var btnGuardar = document.getElementById("btnGuardarLoteManual");

  if (LOTE_MANUAL_ITEMS.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-3">No has agregado ningún elemento a este lote.</td></tr>';
    btnGuardar.disabled = true;
    return;
  }

  var html = "";
  LOTE_MANUAL_ITEMS.forEach((item, index) => {
    html += `<tr>
      <td><span class="badge bg-dark">${item.idInsumo}</span></td>
      <td><strong>${item.name}</strong></td>
      <td class="text-center text-primary fw-bold">${item.cantidad}</td>
      <td class="text-center"><span class="badge bg-warning text-dark">${item.fechaVenc}</span></td>
      <td class="text-center"><button type="button" class="btn btn-sm btn-danger" onclick="eliminarItemLoteManual(${index})">❌</button></td>
    </tr>`;
  });
  tbody.innerHTML = html;
  btnGuardar.disabled = false;
}

async function procesarLoteManualCompleto() {
  if (LOTE_MANUAL_ITEMS.length === 0) return;
  
  var btn = document.getElementById("btnGuardarLoteManual");
  btn.disabled = true;
  btn.innerHTML = "⌛ Guardando en OneDrive...";

  var tipoMovimiento = document.getElementById("tipoMovimiento").value.trim().toUpperCase();
  var proveedor = document.getElementById("proveedor").value;
  var sucursal = document.getElementById("sucursal").value;
  var bodegaOrigen = document.getElementById("bodegaOrigen") ? document.getElementById("bodegaOrigen").value : "BULNES";
  var notas = document.getElementById("notas").value;
  var destinoOOrigen = (tipoMovimiento === "INGRESO") ? proveedor : sucursal;

  var matrizHistorial = obtenerMatrizHoja("Historial");
  var matrizInventario = obtenerMatrizHoja("Inventario_Actual");

  // Calcular próximo correlativo por tipo de movimiento
  var maxActual = 0;
  for (var i = 1; i < matrizHistorial.length; i++) {
    var tipoFila = matrizHistorial[i][2] ? matrizHistorial[i][2].toString().trim().toUpperCase() : "";
    if (tipoFila === tipoMovimiento) {
      var num = parseInt(matrizHistorial[i][9]);
      if (!isNaN(num) && num > maxActual) maxActual = num;
    }
  }
  var proximoCorrelativo = maxActual + 1;
  var fechaActual = new Date();

  LOTE_MANUAL_ITEMS.forEach(function(item) {
    var idLimpio = item.idInsumo.toString().trim();
    var cant = parseFloat(item.cantidad) || 0;
    var venc = item.fechaVenc || "S/V";

    for (var k = 1; k < matrizInventario.length; k++) {
      var idMaster = matrizInventario[k][0] ? matrizInventario[k][0].toString().trim() : "";
      if (idMaster === idLimpio) {
        var stockBulnes = parseFloat(matrizInventario[k][16]) || 0;
        var stockPeninsula = parseFloat(matrizInventario[k][17]) || 0;
        var precioUnit = parseNumeroMoneda(matrizInventario[k][13]);

        if (tipoMovimiento === "SALIDA" || tipoMovimiento === "MERMA") {
          if (bodegaOrigen === "PENINSULA") {
            stockPeninsula = Math.max(0, stockPeninsula - cant);
            matrizInventario[k][17] = stockPeninsula;
          } else {
            stockBulnes = Math.max(0, stockBulnes - cant);
            matrizInventario[k][16] = stockBulnes;
          }
          // Actualizar salidas acumuladas (Columna G / índice 6)
          matrizInventario[k][6] = (parseFloat(matrizInventario[k][6]) || 0) + cant;
        } else if (tipoMovimiento === "INGRESO") {
          if (bodegaOrigen === "PENINSULA") {
            stockPeninsula += cant;
            matrizInventario[k][17] = stockPeninsula;
          } else {
            stockBulnes += cant;
            matrizInventario[k][16] = stockBulnes;
          }
          // Actualizar entradas acumuladas (Columna F / índice 5)
          matrizInventario[k][5] = (parseFloat(matrizInventario[k][5]) || 0) + cant;
          if (venc !== "S/V" && venc !== "") matrizInventario[k][12] = venc;
        } else if (tipoMovimiento === "AJUSTE") {
          var stockActualSis = parseFloat(matrizInventario[k][8]) || 0;
          var dif = cant - stockActualSis;
          matrizInventario[k][5] = (parseFloat(matrizInventario[k][5]) || 0) + dif;
          if (bodegaOrigen === "PENINSULA") {
            stockPeninsula = cant;
            matrizInventario[k][17] = stockPeninsula;
          } else {
            stockBulnes = cant;
            matrizInventario[k][16] = stockBulnes;
          }
          if (venc !== "S/V" && venc !== "") matrizInventario[k][12] = venc;
        }

        // Recalcular Stock Final (Columna I / índice 8) y Valor Total (Columna J / índice 9)
        var nuevoStockFinal = stockBulnes + stockPeninsula;
        matrizInventario[k][8] = nuevoStockFinal;
        matrizInventario[k][9] = nuevoStockFinal * precioUnit;
        matrizInventario[k][10] = nuevoStockFinal <= 0 ? "SIN STOCK" : (nuevoStockFinal <= 5 ? "CRÍTICO" : "DISPONIBLE");
        break;
      }
    }

    matrizHistorial.push([
      fechaActual,
      USUARIO_LOGUEADO,
      tipoMovimiento,
      destinoOOrigen,
      "",
      idLimpio,
      item.name,
      cant,
      venc,
      proximoCorrelativo,
      "", "",
      notas,
      bodegaOrigen
    ]);
  });

  guardarMatrizHoja("Inventario_Actual", matrizInventario);
  guardarMatrizHoja("Historial", matrizHistorial);
  const guardadoOk = await guardarCambiosEnOneDrive();

  btn.disabled = false;
  btn.innerHTML = "<i class='bi bi-cloud-arrow-up-fill me-1'></i> Guardar Lote de Movimientos en OneDrive";

  if (guardadoOk) {
    mostrarNotificacion(`¡Éxito! Lote registrado bajo el Comprobante N° ${proximoCorrelativo}`, "success");
    if (tipoMovimiento === "SALIDA") {
      generarPdfComprobanteLocal({
        tipoDocumento: "SALIDA",
        numero: proximoCorrelativo.toString(),
        entregadoA: destinoOOrigen,
        fecha: formatearFechaDDMMYYYY(fechaActual),
        notas: notas,
        filas: LOTE_MANUAL_ITEMS.map(it => ({
          categoria: it.categoria || "General",
          codigo: it.idInsumo,
          descripcion: it.name,
          um: it.um || "UN",
          cantidad: it.cantidad,
          fechaVencimiento: it.fechaVenc || "S/V",
          observacion: ""
        }))
      });
    }
    LOTE_MANUAL_ITEMS = [];
    renderizarTablaLoteManual();
    document.getElementById("notas").value = "";
    document.getElementById("factura").value = "";
    alternarCamposDinamicos();
    actualizarCatalogoDinamicoEnVivo();
  }
}

// ============================================================================
// 6. SOLICITUDES DE PEDIDO Y BANDEJAS DT / BODEGA
// ============================================================================

function actualizarCatalogoPorSucursalUTM() {
  var selectSucursal = document.getElementById("pedidoSucursalOrigen");
  if (!selectSucursal) return;
  var sucursalSeleccionada = selectSucursal.value;
  if (!sucursalSeleccionada) return;

  SESION_ES_UTM = sucursalSeleccionada.toUpperCase().startsWith("UTM");
  var pestañaReal = MAPA_SUCURSAL_PLANILLA[sucursalSeleccionada] || sucursalSeleccionada;

  var inventario = obtenerInventarioActual();
  var mapaStock = {};
  var mapaUM = {};
  inventario.forEach(item => {
    mapaStock[item.id] = item.stock;
    mapaUM[item.id] = item.um;
  });

  var datosConfig = obtenerMatrizHoja(pestañaReal);
  var insumosEspecificos = [];
  for (var i = 1; i < datosConfig.length; i++) {
    var idCell = datosConfig[i][0];
    if (idCell !== undefined && idCell !== null && idCell.toString().trim() !== "") {
      var idLimpio = idCell.toString().trim();
      insumosEspecificos.push({
        id: idLimpio,
        nombre: datosConfig[i][1] ? datosConfig[i][1].toString().trim() : "",
        um: mapaUM[idLimpio] || (datosConfig[i][2] ? datosConfig[i][2].toString().trim() : "UN"),
        sugerido: parseFloat(datosConfig[i][3]) || 0,
        stock: mapaStock[idLimpio] !== undefined ? mapaStock[idLimpio] : 0
      });
    }
  }
  renderizarTablaPedidosDinamica(insumosEspecificos);
}

function renderizarTablaPedidosDinamica(insumosEspecificos) {
  var tbody = document.getElementById("tbodyCarroPedido");
  if (!tbody) return;
  tbody.innerHTML = "";

  var tablaElemento = tbody.closest("table");
  var thead = tablaElemento ? tablaElemento.querySelector("thead") : null;

  if (thead) {
    if (SESION_ES_UTM) {
      thead.innerHTML = `
        <tr>
          <th style="width: 10%;">Código</th>
          <th style="width: 30%; text-align: left;">Descripción</th>
          <th style="width: 8%;">U.M.</th>
          <th style="width: 10%;">Stock</th>
          <th style="width: 10%;">Sugerido</th>
          <th style="width: 12%;">Cantidad a Solicitar</th>
          <th style="width: 12%;">F. Vencimiento</th>
          <th style="width: 8%;">Observaciones</th>
        </tr>`;
    } else {
      thead.innerHTML = `
        <tr>
          <th style="width: 12%;">Código</th>
          <th style="width: 38%; text-align: left;">Descripción</th>
          <th style="width: 10%;">U.M.</th>
          <th style="width: 12%;">Stock</th>
          <th style="width: 14%;">Cantidad a Solicitar</th>
          <th style="width: 10%;">F. Vencimiento</th>
          <th style="width: 4%;">Observaciones</th>
        </tr>`;
    }
  }

  if (insumosEspecificos && insumosEspecificos.length > 0) {
    insumosEspecificos.forEach(function(insumo) {
      var stockFormateado = Math.round(Number(insumo.stock || 0));
      var badgeStockClass = stockFormateado <= 0 ? "bg-danger" : (stockFormateado <= 5 ? "bg-warning text-dark" : "bg-success");
      var sug = Math.round(Number(insumo.sugerido || 0));

      tbody.innerHTML += `
        <tr>
          <td><span class="badge bg-secondary">${insumo.id}</span></td>
          <td style="text-align: left;"><strong>${insumo.nombre}</strong></td>
          <td><span class="badge bg-light text-dark border">${insumo.um}</span></td>
          <td class="text-center fw-bold"><span class="badge ${badgeStockClass}">${stockFormateado}</span></td>
          ${SESION_ES_UTM ? `<td class="text-center fw-bold text-primary">${sug}</td>` : ''}
          <td>
            <input type="number" class="form-control form-control-sm cantidad-solicitar text-center mx-auto" 
                   data-id="${insumo.id}" data-nombre="${insumo.nombre}" data-um="${insumo.um}"
                   min="0" placeholder="0" style="max-width: 100px;" value="0">
          </td>
          <td>
            <input type="date" class="form-control form-control-sm fecha-vencimiento-solicitar" 
                   id="venc_${insumo.id}" style="font-size: 11px; padding: 4px;">
          </td>
          <td>
            <input type="text" class="form-control form-control-sm observacion-solicitar" 
                   id="obs_${insumo.id}" placeholder="Observaciones..." style="min-width: 100px;">
          </td>
        </tr>`;
    });
  } else {
    var totalCols = SESION_ES_UTM ? 8 : 7;
    tbody.innerHTML = `<tr><td colspan="${totalCols}" class="text-center text-muted py-3">No hay insumos configurados para esta sección.</td></tr>`;
  }
}

async function enviarPedidoFijoAlServidor() {
  var sucursalSel = document.getElementById("pedidoSucursalOrigen").value;
  if (!sucursalSel) {
    mostrarNotificacion("Debe seleccionar la sucursal de origen.", "error");
    return;
  }

  var inputs = document.querySelectorAll(".cantidad-solicitar");
  var itemsParaPedir = [];

  inputs.forEach(function(input) {
    var cant = parseFloat(input.value) || 0;
    if (cant > 0) {
      var idInsumo = input.getAttribute("data-id");
      var nombreInsumo = input.getAttribute("data-nombre");
      var umInsumo = input.getAttribute("data-um") || "UN";
      var inputObs = document.getElementById("obs_" + idInsumo);
      var observacionItem = inputObs ? inputObs.value.trim() : "";

      itemsParaPedir.push({
        idInsumo: idInsumo,
        name: nombreInsumo,
        um: umInsumo,
        cantidad: cant,
        observacion: observacionItem
      });
    }
  });

  if (itemsParaPedir.length === 0) {
    mostrarNotificacion("Ingrese una cantidad mayor a 0 en al menos un insumo.", "error");
    return;
  }

  var btn = document.getElementById("btnEnviarPedido");
  btn.disabled = true;
  btn.innerHTML = "⌛ Guardando solicitud en OneDrive...";

  var matrizPedidos = obtenerMatrizHoja("Pedidos");
  var maxFolio = 0;
  for (var i = 1; i < matrizPedidos.length; i++) {
    var num = parseInt(matrizPedidos[i][1]);
    if (!isNaN(num) && num > maxFolio) maxFolio = num;
  }
  var proximoFolio = maxFolio + 1;
  var timestamp = new Date();
  var notasPedido = document.getElementById("pedidoNotas").value;

  var inventario = obtenerInventarioActual();
  var mapaCat = {};
  inventario.forEach(it => { mapaCat[it.id] = it.categoria; });

  itemsParaPedir.forEach(function(it) {
    matrizPedidos.push([
      timestamp,
      proximoFolio,
      USUARIO_LOGUEADO,
      sucursalSel,
      it.idInsumo,
      it.name,
      it.um,
      it.cantidad,
      "PRE-APROBACION",
      notasPedido,
      it.observacion
    ]);
  });

  guardarMatrizHoja("Pedidos", matrizPedidos);
  const ok = await guardarCambiosEnOneDrive();

  btn.disabled = false;
  btn.innerHTML = "<i class='bi bi-send-check-fill me-1'></i> Despachar Solicitud de Pedido a Bodega";

  if (ok) {
    mostrarNotificacion(`¡Solicitud de Pedido N° ${proximoFolio} guardada con éxito!`, "success");
    generarPdfComprobanteLocal({
      tipoDocumento: "SOLICITUD DE PEDIDO",
      numero: proximoFolio.toString(),
      entregadoA: sucursalSel,
      notas: notasPedido,
      fecha: formatearFechaDDMMYYYY(timestamp),
      filas: itemsParaPedir.map(it => ({
        categoria: mapaCat[it.idInsumo] || "General",
        codigo: it.idInsumo,
        descripcion: it.name,
        um: it.um,
        cantidad: it.cantidad,
        fechaVencimiento: "S/V",
        observacion: it.observacion
      }))
    });

    inputs.forEach(input => { input.value = "0"; });
    document.querySelectorAll(".observacion-solicitar").forEach(obs => { obs.value = ""; });
    document.getElementById("pedidoNotas").value = "";
  }
}

// --- BANDEJA DIRECTOR TÉCNICO ---
function cargarPedidosPreAprobacionDT() {
  var tbody = document.getElementById("tbodyPreAprobacionDT");
  if (!tbody) return;
  document.getElementById("checkTodosDT").checked = false;
  actualizarContadorDT();

  var datos = obtenerMatrizHoja("Pedidos");
  var pedidos = [];
  for (var i = 1; i < datos.length; i++) {
    if (!datos[i][0] && !datos[i][1]) continue;
    var estado = datos[i][8] ? datos[i][8].toString().trim().toUpperCase() : "";
    if (estado === "PRE-APROBACION" || estado === "PRE-APROBACIÓN") {
      pedidos.push({
        filaIndex: i,
        fecha: formatearFechaHora(datos[i][0]),
        idPedido: datos[i][1] ? datos[i][1].toString().trim() : "",
        solicitante: datos[i][2] ? datos[i][2].toString().trim() : "",
        sucursal: datos[i][3] ? datos[i][3].toString().trim() : "",
        idInsumo: datos[i][4] ? datos[i][4].toString().trim() : "",
        descripcion: datos[i][5] ? datos[i][5].toString().trim() : "",
        um: datos[i][6] ? datos[i][6].toString().trim() : "UN",
        cantidad: parseFloat(datos[i][7]) || 0,
        observacion: datos[i][10] ? datos[i][10].toString().trim() : ""
      });
    }
  }

  if (pedidos.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="text-center text-muted py-3">No hay solicitudes esperando pre-aprobación en este momento.</td></tr>';
    return;
  }

  var html = "";
  pedidos.forEach(function(p) {
    html += `<tr>
      <td class="text-center"><input type="checkbox" class="check-dt-item" data-fila="${p.filaIndex}" onchange="actualizarContadorDT()"></td>
      <td><small>${p.fecha}</small></td>
      <td><strong>${p.solicitante}</strong></td>
      <td><span class="badge bg-secondary">${p.sucursal}</span></td>
      <td><span class="badge bg-dark">${p.idInsumo}</span></td>
      <td style="text-align: left;"><strong>${p.descripcion}</strong></td>
      <td class="text-center bg-light fw-bold">${p.um}</td>
      <td><input type="number" step="0.01" min="0.01" class="form-control form-control-sm text-center input-dt-cantidad mx-auto" value="${p.cantidad}" style="max-width: 90px; font-weight: bold;"></td>
      <td><small class="text-muted">${p.observacion}</small></td>
    </tr>`;
  });
  tbody.innerHTML = html;
}

function seleccionarTodosDT(master) {
  document.querySelectorAll(".check-dt-item").forEach(cb => cb.checked = master.checked);
  actualizarContadorDT();
}

function actualizarContadorDT() {
  var seleccionados = document.querySelectorAll(".check-dt-item:checked").length;
  var lbl = document.getElementById("lblContadorDT");
  if (lbl) lbl.innerText = seleccionados;
}

function ejecutarLiberacionMasivaDT() {
  var check = document.querySelectorAll(".check-dt-item:checked");
  if (check.length === 0) {
    mostrarNotificacion("Selecciona al menos una solicitud para pre-aprobar.", "warning");
    return;
  }

  document.getElementById("modalTextoCantidad").innerHTML = `¿Desea pre-aprobar los <span class="text-primary fw-bold">${check.length}</span> ítems seleccionados?`;
  var miModal = new bootstrap.Modal(document.getElementById('modalConfirmarDT'));
  miModal.show();

  document.getElementById("btnAceptarModalDT").onclick = async function() {
    miModal.hide();
    var matrizPedidos = obtenerMatrizHoja("Pedidos");
    var procesados = 0;

    check.forEach(function(c) {
      var tr = c.closest("tr");
      var inputCant = tr ? tr.querySelector(".input-dt-cantidad") : null;
      var nuevaCantidad = inputCant ? parseFloat(inputCant.value) || 0 : 0;
      var idx = parseInt(c.dataset.fila);

      if (!isNaN(idx) && idx > 0 && matrizPedidos[idx] && nuevaCantidad > 0) {
        matrizPedidos[idx][8] = "APROBADO_DT";
        matrizPedidos[idx][7] = nuevaCantidad;
        procesados++;
      }
    });

    guardarMatrizHoja("Pedidos", matrizPedidos);
    const ok = await guardarCambiosEnOneDrive();
    if (ok) {
      mostrarNotificacion(`¡Éxito! Se liberaron ${procesados} ítems para Bodega.`, "success");
      cargarPedidosPreAprobacionDT();
    }
  };
}

// --- BANDEJA BODEGA ---
function cargarPedidosPendientesAdmin() {
  var tbody = document.getElementById("tbodyAdminPedidos");
  if (!tbody) return;
  document.getElementById("checkTodosPedidos").checked = false;
  actualizarContadorSeleccionados();

  var inventario = obtenerInventarioActual();
  var mapaStock = {};
  inventario.forEach(it => { mapaStock[it.id] = it.stock; });

  var datos = obtenerMatrizHoja("Pedidos");
  PEDIDOS_CARGADOS_CACHE = [];

  for (var i = 1; i < datos.length; i++) {
    if (!datos[i][0] && !datos[i][1]) continue;
    var estado = datos[i][8] ? datos[i][8].toString().trim().toUpperCase() : "";
    if (estado === "APROBADO_DT" || estado === "APROBADO" || estado === "PENDIENTE") {
      var idInsumoReal = datos[i][4] ? datos[i][4].toString().trim() : "";
      var stockDisp = mapaStock[idInsumoReal] !== undefined ? mapaStock[idInsumoReal] : 0;
      var cantReal = parseFloat(datos[i][7]) || 0;
      var textoObs = datos[i][10] ? datos[i][10].toString().trim() : "";

      PEDIDOS_CARGADOS_CACHE.push({
        filaIndex: i,
        fecha: formatearFechaHora(datos[i][0]),
        fechaRaw: datos[i][0],
        idPedido: datos[i][1] ? datos[i][1].toString().trim() : "",
        solicitante: datos[i][2] ? datos[i][2].toString().trim() : "",
        sucursal: datos[i][3] ? datos[i][3].toString().trim() : "",
        idInsumo: idInsumoReal,
        descripcion: datos[i][5] ? datos[i][5].toString().trim() : "",
        um: datos[i][6] ? datos[i][6].toString().trim() : "UN",
        cantidad: cantReal,
        stockDisponible: stockDisp,
        observacionSolicitud: textoObs
      });
    }
  }

  if (PEDIDOS_CARGADOS_CACHE.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" class="text-center text-muted py-3">No hay solicitudes pendientes en este momento.</td></tr>';
    return;
  }

  var html = "";
  PEDIDOS_CARGADOS_CACHE.forEach((p, index) => {
    var sinStock = p.stockDisponible < p.cantidad;
    var badgeStock = sinStock ? "bg-danger" : "bg-success";
    var contenidoCheck = sinStock
      ? `<i class="bi bi-x-circle-fill text-danger" title="Stock insuficiente (${p.stockDisponible})"></i>`
      : `<input type="checkbox" class="check-pedido-item" data-index="${index}" onchange="actualizarContadorSeleccionados()">`;

    html += `<tr>
      <td class="text-center" id="celda_check_${index}">${contenidoCheck}</td>
      <td><small>${p.fecha}</small></td>
      <td><strong>${p.solicitante}</strong></td>
      <td><span class="badge bg-secondary">${p.sucursal}</span></td>
      <td><span class="badge bg-dark">${p.idInsumo}</span></td>
      <td style="text-align: left;"><strong>${p.descripcion}</strong></td>
      <td class="text-center small text-dark fw-bold bg-light">${p.um}</td>
      <td class="text-center"><input type="number" step="0.01" min="0" class="form-control form-control-sm text-center fw-bold text-primary mx-auto" id="admin_cant_${index}" value="${p.cantidad}" style="max-width: 90px;" oninput="validarCantidadDespachoDinamica(${index})"></td>
      <td class="text-center"><span class="badge ${badgeStock}">${p.stockDisponible}</span></td>
      <td><input type="date" class="form-control form-control-sm" id="admin_venc_${index}" style="font-size: 11px; padding: 2px 4px;"></td>
      <td><input type="text" class="form-control form-control-sm" id="admin_obs_${index}" value="${p.observacionSolicitud}" placeholder="Observaciones..."></td>
    </tr>`;
  });
  tbody.innerHTML = html;
}

function validarCantidadDespachoDinamica(index) {
  var p = PEDIDOS_CARGADOS_CACHE[index];
  if (!p) return;
  var nuevaCantidad = parseFloat(document.getElementById("admin_cant_" + index).value) || 0;
  var stockDisponible = parseFloat(p.stockDisponible) || 0;
  var celdaCheck = document.getElementById("celda_check_" + index);

  if (nuevaCantidad > 0 && nuevaCantidad <= stockDisponible) {
    celdaCheck.innerHTML = `<input type="checkbox" class="check-pedido-item" data-index="${index}" onchange="actualizarContadorSeleccionados()">`;
  } else {
    celdaCheck.innerHTML = `<i class="bi bi-x-circle-fill text-danger" title="Supera stock disponible (${stockDisponible})"></i>`;
  }
  actualizarContadorSeleccionados();
}

function seleccionarTodosLosPedidos(master) {
  document.querySelectorAll(".check-pedido-item").forEach(cb => { cb.checked = master.checked; });
  actualizarContadorSeleccionados();
}

function actualizarContadorSeleccionados() {
  var seleccionados = document.querySelectorAll(".check-pedido-item:checked").length;
  document.getElementById("lblContadorSeleccionados").innerText = seleccionados;
  document.getElementById("btnDespachoGlobal").disabled = (seleccionados === 0);
}

function procesarDespachoMasivoGlobal() {
  var checkboxes = document.querySelectorAll(".check-pedido-item:checked");
  if (checkboxes.length === 0) {
    mostrarNotificacion("Por favor, selecciona al menos un ítem.", "warning");
    return;
  }

  document.getElementById("modalTextoBodegaCantidad").innerHTML = `¿Desea aprobar y despachar los <span class="text-primary fw-bold">${checkboxes.length}</span> ítems seleccionados?`;
  var miModalBodega = new bootstrap.Modal(document.getElementById('modalConfirmarBodega'));
  miModalBodega.show();

  document.getElementById("btnAceptarModalBodega").onclick = async function() {
    miModalBodega.hide();
    var bodegaOrigen = document.getElementById("bodegaDespachoAdmin").value || "BULNES";
    var matrizPedidos = obtenerMatrizHoja("Pedidos");
    var matrizHistorial = obtenerMatrizHoja("Historial");
    var matrizInventario = obtenerMatrizHoja("Inventario_Actual");

    var mapaCat = {};
    var mapaIdxInv = {};
    for (var k = 1; k < matrizInventario.length; k++) {
      var idM = matrizInventario[k][0] ? matrizInventario[k][0].toString().trim() : "";
      if (idM) {
        mapaCat[idM] = matrizInventario[k][2] ? matrizInventario[k][2].toString().trim() : "General";
        mapaIdxInv[idM] = k;
      }
    }

    // Correlativo exclusivo de SALIDA
    var maxSalida = 0;
    for (var f = 1; f < matrizHistorial.length; f++) {
      var tFila = matrizHistorial[f][2] ? matrizHistorial[f][2].toString().trim().toUpperCase() : "";
      if (tFila === "SALIDA") {
        var n = parseInt(matrizHistorial[f][9]);
        if (!isNaN(n) && n > maxSalida) maxSalida = n;
      }
    }
    var proximoCorrelativo = maxSalida + 1;
    var timestamp = new Date();

    var itemsPdfSalida = [];
    var itemsDeclaracion = [];
    var folioSolicitud = "";
    var fechaSolicitud = "";
    var coordinadorSol = "";
    var sucursalDestino = "";

    checkboxes.forEach(function(cb) {
      var index = parseInt(cb.getAttribute("data-index"));
      var p = PEDIDOS_CARGADOS_CACHE[index];
      if (!p) return;

      if (!folioSolicitud) folioSolicitud = p.idPedido;
      if (!fechaSolicitud) fechaSolicitud = formatearFechaDDMMYYYY(p.fechaRaw);
      if (!coordinadorSol) coordinadorSol = p.solicitante;
      if (!sucursalDestino) sucursalDestino = p.sucursal;

      var cantDespachada = parseFloat(document.getElementById("admin_cant_" + index).value) || 0;
      var fechaElegida = document.getElementById("admin_venc_" + index).value || "S/V";
      var observacionItem = document.getElementById("admin_obs_" + index).value.trim();

      // Descontar stock en Inventario_Actual
      if (mapaIdxInv[p.idInsumo] !== undefined && cantDespachada > 0) {
        var idxM = mapaIdxInv[p.idInsumo];
        var sBul = parseFloat(matrizInventario[idxM][16]) || 0;
        var sPen = parseFloat(matrizInventario[idxM][17]) || 0;
        var pUnit = parseNumeroMoneda(matrizInventario[idxM][13]);

        if (bodegaOrigen === "PENINSULA") {
          sPen = Math.max(0, sPen - cantDespachada);
          matrizInventario[idxM][17] = sPen;
        } else {
          sBul = Math.max(0, sBul - cantDespachada);
          matrizInventario[idxM][16] = sBul;
        }
        var nuevoFinal = sBul + sPen;
        matrizInventario[idxM][6] = (parseFloat(matrizInventario[idxM][6]) || 0) + cantDespachada;
        matrizInventario[idxM][8] = nuevoFinal;
        matrizInventario[idxM][9] = nuevoFinal * pUnit;
        matrizInventario[idxM][10] = nuevoFinal <= 0 ? "SIN STOCK" : (nuevoFinal <= 5 ? "CRÍTICO" : "DISPONIBLE");
      }

      matrizHistorial.push([
        timestamp,
        USUARIO_LOGUEADO,
        "SALIDA",
        p.sucursal,
        "Solicitud N° " + p.idPedido,
        p.idInsumo,
        p.descripcion,
        cantDespachada,
        fechaElegida,
        proximoCorrelativo,
        "", "",
        observacionItem,
        bodegaOrigen
      ]);

      if (matrizPedidos[p.filaIndex]) {
        matrizPedidos[p.filaIndex][8] = "DESPACHADO";
      }

      itemsPdfSalida.push({
        categoria: mapaCat[p.idInsumo] || "General",
        codigo: p.idInsumo,
        descripcion: p.descripcion,
        um: p.um,
        cantidad: cantDespachada,
        fechaVencimiento: fechaElegida,
        observacion: observacionItem
      });

      if (cantDespachada < p.cantidad) {
        itemsDeclaracion.push({
          codigo: p.idInsumo,
          descripcion: p.descripcion,
          um: p.um,
          cantSolicitada: p.cantidad,
          cantAprobada: cantDespachada,
          diferencia: p.cantidad - cantDespachada
        });
      }
    });

    guardarMatrizHoja("Inventario_Actual", matrizInventario);
    guardarMatrizHoja("Historial", matrizHistorial);
    guardarMatrizHoja("Pedidos", matrizPedidos);
    const ok = await guardarCambiosEnOneDrive();

    if (ok) {
      mostrarNotificacion(`¡Éxito! Despacho N° ${proximoCorrelativo} registrado en OneDrive.`, "success");
      cargarPedidosPendientesAdmin();
      actualizarCatalogoDinamicoEnVivo();

      if (itemsPdfSalida.length > 0) {
        generarPdfComprobanteLocal({
          tipoDocumento: "SALIDA",
          numero: proximoCorrelativo.toString(),
          solicitudNumero: folioSolicitud,
          solicitudFecha: fechaSolicitud,
          entregadoA: sucursalDestino,
          fecha: formatearFechaDDMMYYYY(timestamp),
          filas: itemsPdfSalida
        });
      }

      if (itemsDeclaracion.length > 0) {
        setTimeout(() => {
          generarPdfDeclaracionLocal({
            numeroSalida: proximoCorrelativo.toString(),
            numeroSolicitud: folioSolicitud,
            sucursal: sucursalDestino,
            fecha: formatearFechaDDMMYYYY(timestamp),
            coordinador: coordinadorSol,
            filas: itemsDeclaracion
          });
        }, 600);
      }
    }
  };
}

// ============================================================================
// 7. MONITORES (SALDOS, SUB-BODEGAS, MERMAS Y GRÁFICO)
// ============================================================================

function recargarTablaMonitor() {
  INSUMOS_MONITOR_CACHE = obtenerInventarioActual();
  table_refresh_view(INSUMOS_MONITOR_CACHE);
}

function filtrarMonitorSaldos() {
  var texto = document.getElementById("txtBuscarMonitor").value.toLowerCase().trim();
  if (texto === "") {
    table_refresh_view(INSUMOS_MONITOR_CACHE);
    return;
  }
  var filtrados = INSUMOS_MONITOR_CACHE.filter(ins =>
    (ins.id && ins.id.toLowerCase().includes(texto)) ||
    (ins.nombre && ins.nombre.toLowerCase().includes(texto)) ||
    (ins.categoria && ins.categoria.toLowerCase().includes(texto))
  );
  table_refresh_view(filtrados);
}

function table_refresh_view(lista) {
  var tbody = document.getElementById("tbodyMonitor");
  if (!lista || lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" class="text-center text-muted py-3">No se encontraron registros de insumos.</td></tr>';
    return;
  }
  var html = "";
  lista.forEach(function(item) {
    var claseStock = item.stock <= 0 ? "bg-danger" : (item.estado.includes("CRÍTICO") ? "bg-warning text-dark" : "bg-success");
    var estiloFila = item.stock < item.proyeccionTrimestral ? "style='background-color: #fff5f5;'" : "";
    var alertaTrimestral = item.stock < item.proyeccionTrimestral
      ? "<span class='badge bg-danger text-white'><i class='bi bi-exclamation-triangle-fill me-1'></i> Insuficiente</span>"
      : "<span class='badge bg-success text-white'><i class='bi bi-check-circle-fill me-1'></i> Cubierto</span>";

    html += `<tr ${estiloFila}>
      <td><span class="badge bg-secondary">${item.id}</span></td>
      <td><strong>${item.nombre}</strong></td>
      <td>${item.categoria}</td>
      <td>${item.um}</td>
      <td class="text-center"><span class="badge badge-stock-size ${claseStock}">${item.stock}</span></td>
      <td class="text-end fw-bold text-success">${formatMonto(item.totalValor)}</td>
      <td class="text-center"><small class="fw-semibold">${item.estado}</small></td>
      <td class="text-center fw-semibold text-secondary">${item.proyeccionTrimestral || 0}</td>
      <td class="text-center text-muted">${item.proyeccionAnual || 0}</td>
      <td class="text-center">${alertaTrimestral}</td>
    </tr>`;
  });
  tbody.innerHTML = html;
}

function recargarTablaSubbodegas() {
  INSUMOS_MONITOR_CACHE = obtenerInventarioActual();
  renderizarTablaSubbodegasHTML(INSUMOS_MONITOR_CACHE);
}

function renderizarTablaSubbodegasHTML(lista) {
  var tbody = document.getElementById("tbodySubbodegas");
  if (!tbody) return;
  if (!lista || lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-3">No hay datos disponibles.</td></tr>';
    return;
  }
  var html = "";
  lista.forEach(function(item) {
    var badgeFinal = item.stock <= 0 ? "bg-danger" : "bg-success";
    var badgeBulnes = item.stockBulnes <= 0 ? "bg-secondary" : "bg-primary";
    var badgePeninsula = item.stockPeninsula <= 0 ? "bg-secondary" : "bg-warning text-dark";
    html += `<tr>
      <td><span class="badge bg-dark">${item.id}</span></td>
      <td><strong>${item.nombre}</strong></td>
      <td class="text-center"><span class="badge ${badgeFinal} fs-6">${item.stock}</span></td>
      <td class="text-center"><span class="badge ${badgeBulnes} fs-6">${item.stockBulnes}</span></td>
      <td class="text-center"><span class="badge ${badgePeninsula} fs-6">${item.stockPeninsula}</span></td>
    </tr>`;
  });
  tbody.innerHTML = html;
}

function filtrarSubbodegas() {
  var texto = document.getElementById("txtBuscarSubbodegas").value.toLowerCase().trim();
  if (texto === "") {
    renderizarTablaSubbodegasHTML(INSUMOS_MONITOR_CACHE);
    return;
  }
  var filtrados = INSUMOS_MONITOR_CACHE.filter(ins =>
    (ins.id && ins.id.toLowerCase().includes(texto)) || (ins.nombre && ins.nombre.toLowerCase().includes(texto))
  );
  renderizarTablaSubbodegasHTML(filtrados);
}

function recargarTablaMermas() {
  var tbody = document.getElementById("tbodyMermas");
  if (!tbody) return;

  var inventario = obtenerInventarioActual();
  var dicPrecios = {};
  var dicNombres = {};
  var dicUM = {};
  inventario.forEach(it => {
    dicPrecios[it.id] = it.precioUnitario;
    dicNombres[it.id] = it.nombre;
    dicUM[it.id] = it.um;
  });

  var datosMovs = obtenerMatrizHoja("Historial");
  var mermasAgrupadas = {};

  for (var j = 1; j < datosMovs.length; j++) {
    var tipo = datosMovs[j][2] ? datosMovs[j][2].toString().trim().toUpperCase() : "";
    if (tipo === "MERMA") {
      var idInsumo = datosMovs[j][5] ? datosMovs[j][5].toString().trim() : "";
      var desc = datosMovs[j][6] ? datosMovs[j][6].toString().trim() : (dicNombres[idInsumo] || "");
      var cant = parseFloat(datosMovs[j][7]) || 0;
      if (idInsumo !== "") {
        if (!mermasAgrupadas[idInsumo]) {
          mermasAgrupadas[idInsumo] = { id: idInsumo, nombre: desc, um: dicUM[idInsumo] || "UN", cantidad: 0, valorTotal: 0 };
        }
        mermasAgrupadas[idInsumo].cantidad += cant;
      }
    }
  }

  var lista = Object.values(mermasAgrupadas).map(m => {
    m.valorTotal = m.cantidad * (dicPrecios[m.id] || 0);
    return m;
  });

  if (lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-4">No se han registrado mermas.</td></tr>';
    return;
  }

  var html = "";
  lista.forEach(m => {
    html += `<tr>
      <td><span class="badge bg-danger">${m.id}</span></td>
      <td><strong>${m.nombre}</strong></td>
      <td class="text-center text-muted">${m.um}</td>
      <td class="text-center fw-bold">${m.cantidad}</td>
      <td class="text-end fw-bold text-danger">${formatMonto(m.valorTotal)}</td>
    </tr>`;
  });
  tbody.innerHTML = html;
}

function abrirModalGraficoMermas() {
  var modalEl = document.getElementById('modalGraficoMermas');
  var myModal = bootstrap.Modal.getInstance(modalEl) || new bootstrap.Modal(modalEl);
  myModal.show();

  var contenedor = document.getElementById('mermasChartContainer');
  var inventario = obtenerInventarioActual();
  var mapaCostos = {};
  inventario.forEach(it => { mapaCostos[it.id] = it.precioUnitario; });

  var datosHistorial = obtenerMatrizHoja("Historial");
  var acumuladoMeses = {};

  for (var i = 1; i < datosHistorial.length; i++) {
    var tipo = datosHistorial[i][2] ? datosHistorial[i][2].toString().trim().toUpperCase() : "";
    if (tipo === "MERMA") {
      var fechaCelda = datosHistorial[i][0];
      var idInsumo = datosHistorial[i][5] ? datosHistorial[i][5].toString().trim() : "";
      var cant = Math.abs(parseFloat(datosHistorial[i][7])) || 0;
      var perdida = cant * (mapaCostos[idInsumo] || 0);

      var d = (fechaCelda instanceof Date) ? fechaCelda : new Date(fechaCelda);
      if (!isNaN(d.getTime())) {
        var clave = String(d.getMonth() + 1).padStart(2, '0') + "-" + d.getFullYear();
        acumuladoMeses[clave] = (acumuladoMeses[clave] || 0) + perdida;
      }
    }
  }

  var llaves = Object.keys(acumuladoMeses).sort((a, b) => {
    var pA = a.split("-"), pB = b.split("-");
    return new Date(pA[1], pA[0] - 1) - new Date(pB[1], pB[0] - 1);
  });

  if (llaves.length === 0) {
    contenedor.innerHTML = '<div class="text-center text-muted py-5">No se registraron mermas valorizadas en el historial.</div>';
    return;
  }

  var nombresMeses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  var dataTable = new google.visualization.DataTable();
  dataTable.addColumn('string', 'Mes');
  dataTable.addColumn('number', 'Total Pérdida ($)');
  dataTable.addColumn({ type: 'string', role: 'annotation' });

  var filas = llaves.map(k => {
    var p = k.split("-");
    var monto = Math.round(acumuladoMeses[k]);
    return [nombresMeses[parseInt(p[0], 10) - 1] + " " + p[1], monto, '$' + monto.toLocaleString('es-CL')];
  });
  dataTable.addRows(filas);

  contenedor.innerHTML = "";
  var chart = new google.visualization.ColumnChart(contenedor);
  chart.draw(dataTable, {
    hAxis: { title: 'Mes-Año' },
    vAxis: { title: 'Pérdida Estimada ($)', minValue: 0, format: '$#,###' },
    colors: ['#dc3545'],
    legend: { position: 'none' },
    chartArea: { width: '85%', height: '70%' }
  });
}

// ============================================================================
// 8. REPORTE MENSUAL POR CICLO OPERATIVO (DÍA 20 MES ANTERIOR AL DÍA 5)
// ============================================================================

function poblarSelectorSucursalesReporte(sucursales) {
  var selectRep = document.getElementById("reporteSucursalSelect");
  if (!selectRep || !sucursales) return;
  selectRep.innerHTML = "";
  sucursales.forEach(s => { selectRep.innerHTML += `<option value="${s}">${s}</option>`; });
}

function consultarReporteMensualBackend() {
  var sucursal = document.getElementById("reporteSucursalSelect").value;
  var mesNum = parseInt(document.getElementById("reporteMesSelect").value, 10);
  var anioNum = parseInt(document.getElementById("reporteAnioInput").value, 10);

  if (!sucursal) {
    mostrarNotificacion("Seleccione una sucursal válida.", "error");
    return;
  }

  var sucBuscada = sucursal.trim().toUpperCase();
  var fechaInicioCiclo = new Date(anioNum, mesNum - 2, 20, 0, 0, 0);
  var fechaFinCiclo = new Date(anioNum, mesNum - 1, 5, 23, 59, 59);

  var inventario = obtenerInventarioActual();
  var mapaUM = {};
  inventario.forEach(it => { mapaUM[it.id] = it.um; });

  var consolidado = {};
  var foliosCiclo = [];

  var datosPedidos = obtenerMatrizHoja("Pedidos");
  for (var i = 1; i < datosPedidos.length; i++) {
    var fechaP = datosPedidos[i][0];
    var sucP = datosPedidos[i][3] ? datosPedidos[i][3].toString().trim().toUpperCase() : "";
    if (fechaP && sucP === sucBuscada) {
      var fObj = (fechaP instanceof Date) ? fechaP : new Date(fechaP);
      if (!isNaN(fObj.getTime()) && fObj >= fechaInicioCiclo && fObj <= fechaFinCiclo) {
        var numSol = datosPedidos[i][1] ? datosPedidos[i][1].toString().trim() : "";
        var cod = datosPedidos[i][4] ? datosPedidos[i][4].toString().trim() : "";
        var desc = datosPedidos[i][5] ? datosPedidos[i][5].toString().trim() : "";
        var cantP = parseFloat(datosPedidos[i][7]) || 0;
        var um = datosPedidos[i][6] ? datosPedidos[i][6].toString().trim() : (mapaUM[cod] || "UN");

        if (cod) {
          if (!consolidado[cod]) {
            consolidado[cod] = { codigo: cod, insumo: desc, um: um, cantPedido: 0, cantDespachada: 0, numSolicitudes: [], numSalidas: [] };
          }
          consolidado[cod].cantPedido += cantP;
          if (numSol) {
            if (!consolidado[cod].numSolicitudes.includes(numSol)) consolidado[cod].numSolicitudes.push(numSol);
            if (!foliosCiclo.includes(numSol)) foliosCiclo.push(numSol);
          }
        }
      }
    }
  }

  var datosHistorial = obtenerMatrizHoja("Historial");
  for (var j = 1; j < datosHistorial.length; j++) {
    var fechaH = datosHistorial[j][0];
    var tipoOp = datosHistorial[j][2] ? datosHistorial[j][2].toString().trim().toUpperCase() : "";
    var sucH = datosHistorial[j][3] ? datosHistorial[j][3].toString().trim().toUpperCase() : "";
    var obsNotas = datosHistorial[j][4] ? datosHistorial[j][4].toString().trim() : "";

    if (tipoOp === "SALIDA" && sucH === sucBuscada) {
      var fObjH = (fechaH instanceof Date) ? fechaH : new Date(fechaH);
      var match = obsNotas.match(/\d+/);
      var numSolVinc = match ? match[0] : "";
      var pertenece = (numSolVinc && foliosCiclo.includes(numSolVinc)) || (!isNaN(fObjH.getTime()) && fObjH >= fechaInicioCiclo && fObjH <= fechaFinCiclo);

      if (pertenece) {
        var codH = datosHistorial[j][5] ? datosHistorial[j][5].toString().trim() : "";
        var descH = datosHistorial[j][6] ? datosHistorial[j][6].toString().trim() : "";
        var cantD = parseFloat(datosHistorial[j][7]) || 0;
        var numSal = datosHistorial[j][9] ? datosHistorial[j][9].toString().trim() : "";

        if (codH) {
          if (!consolidado[codH]) {
            consolidado[codH] = { codigo: codH, insumo: descH, um: mapaUM[codH] || "UN", cantPedido: 0, cantDespachada: 0, numSolicitudes: [], numSalidas: [] };
          }
          consolidado[codH].cantDespachada += cantD;
          if (numSal && !consolidado[codH].numSalidas.includes(numSal)) consolidado[codH].numSalidas.push(numSal);
        }
      }
    }
  }

  var resultadoFinal = Object.values(consolidado).map(item => ({
    codigo: item.codigo,
    insumo: item.insumo,
    um: item.um,
    cantPedido: item.cantPedido,
    cantDespachada: item.cantDespachada,
    solicitudesTexto: item.numSolicitudes.length > 0 ? item.numSolicitudes.join(", ") : "S/N",
    salidasTexto: item.numSalidas.length > 0 ? item.numSalidas.join(", ") : "S/N"
  }));

  ULTIMO_REPORTE_MENSUAL = { sucursal: sucursal, mes: mesNum, anio: anioNum, datos: resultadoFinal };
  document.getElementById("lblSucursalReporte").innerText = sucursal;
  document.getElementById("lblMesAnioReporte").innerText = mesNum + " / " + anioNum;

  var tbody = document.getElementById("tbodyReporteMensual");
  var btnPdf = document.getElementById("btnDescargarPdfReporte");

  if (resultadoFinal.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted py-4">No se encontraron movimientos registrados para esta sucursal en el mes seleccionado.</td></tr>';
    btnPdf.disabled = true;
    return;
  }

  var html = "";
  resultadoFinal.forEach(item => {
    html += `<tr>
      <td><strong>${item.insumo}</strong> <small class="text-muted">(${item.codigo})</small></td>
      <td class="text-center"><span class="badge bg-light text-dark border">${item.um}</span></td>
      <td class="text-center fw-bold text-primary fs-6">${item.cantPedido}</td>
      <td class="text-center"><span class="badge bg-secondary">${item.solicitudesTexto}</span></td>
      <td class="text-center fw-bold text-success fs-6">${item.cantDespachada}</td>
      <td class="text-center"><span class="badge bg-dark">${item.salidasTexto}</span></td>
    </tr>`;
  });
  tbody.innerHTML = html;
  btnPdf.disabled = false;
}

// ============================================================================
// 9. GENERACIÓN DE COMPROBANTES PDF EN EL NAVEGADOR (jsPDF)
// ============================================================================

function generarPdfComprobanteLocal(datos) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });
  const esSolicitud = datos.tipoDocumento.toUpperCase().includes("SOLICITUD");

  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.setTextColor(26, 74, 117);
  doc.text("LABORATORIO BIOCLINIC", 14, 18);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(85, 85, 85);
  doc.text("Juan Antonio Ríos #2889 | Tel: 572-277483", 14, 23);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  if (esSolicitud) {
    doc.text("Nº Registro: PO-ADM-02-R1 | Versión 02 (Solicitud)", 14, 28);
  } else {
    doc.text("N° Registro: PO-ADM-02-R2 | Versión 01 (Salida)", 14, 28);
    if (datos.solicitudNumero) {
      doc.setTextColor(26, 74, 117);
      doc.text(`Correspondiente a Solicitud N° ${datos.solicitudNumero} ${datos.solicitudFecha ? 'de fecha ' + datos.solicitudFecha : ''}`, 14, 33);
    }
  }

  // Caja de Título y Número
  doc.setDrawColor(26, 74, 117);
  doc.setLineWidth(0.6);
  doc.roundedRect(145, 12, 56, 20, 2, 2);
  doc.setFontSize(10);
  doc.setTextColor(26, 74, 117);
  doc.text(datos.tipoDocumento, 173, 19, { align: "center" });
  doc.setFontSize(13);
  doc.setTextColor(220, 38, 38);
  doc.text(`N° ${datos.numero}`, 173, 27, { align: "center" });

  // Tabla de Metadatos
  doc.autoTable({
    startY: 38,
    body: [[
      { content: 'SUCURSAL / DESTINO:', styles: { fillColor: [143, 170, 220], fontStyle: 'bold' } },
      { content: datos.entregadoA || 'General' },
      { content: 'FECHA:', styles: { fillColor: [143, 170, 220], fontStyle: 'bold' } },
      { content: datos.fecha || '', styles: { halign: 'center' } }
    ]],
    theme: 'grid',
    styles: { fontSize: 9, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 }
  });

  let nextY = doc.lastAutoTable.finalY + 4;
  if (datos.notas) {
    doc.setFontSize(9);
    doc.setTextColor(113, 63, 18);
    doc.text(`Observación General: ${datos.notas}`, 14, nextY + 4);
    nextY += 8;
  }

  const filasTabla = (datos.filas || []).map(f => [
    f.categoria || "General",
    f.codigo || "",
    f.descripcion || "",
    f.um || "UN",
    f.cantidad || 0,
    f.fechaVencimiento || "S/V",
    f.observacion || ""
  ]);

  while (filasTabla.length < 5) {
    filasTabla.push(["", "", "", "", "", "", ""]);
  }

  doc.autoTable({
    startY: nextY + 2,
    head: [['CATEGORÍA', 'CÓDIGO', 'DESCRIPCIÓN', 'U.M.', 'CANT.', 'F. VENC.', 'OBSERVACIONES']],
    body: filasTabla,
    theme: 'grid',
    headStyles: { fillColor: [143, 170, 220], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' },
    styles: { fontSize: 8.5, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 },
    columnStyles: {
      3: { halign: 'center' },
      4: { halign: 'center' },
      5: { halign: 'center' }
    }
  });

  const firmaY = doc.lastAutoTable.finalY + 35;
  doc.setDrawColor(0, 0, 0);
  doc.line(25, firmaY, 90, firmaY);
  doc.line(125, firmaY, 190, firmaY);

  doc.setFontSize(9);
  doc.setTextColor(0, 0, 0);
  doc.setFont("helvetica", "bold");
  doc.text("QUIEN ENTREGA", 57.5, firmaY + 5, { align: "center" });
  doc.text("QUIEN RECIBE / SOLICITA", 157.5, firmaY + 5, { align: "center" });

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text("Área Administración - Bioclinic Ltda.", 57.5, firmaY + 10, { align: "center" });
  doc.text(datos.entregadoA || "", 157.5, firmaY + 10, { align: "center" });

  const nombrePdf = `${datos.tipoDocumento.replace(/[^a-zA-Z0-9]/g, "_")}_N_${datos.numero}.pdf`;
  doc.save(nombrePdf);
}

function generarPdfDeclaracionLocal(datos) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });

  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(26, 74, 117);
  doc.text("LABORATORIO BIOCLINIC", 108, 18, { align: "center" });

  doc.setFontSize(10);
  doc.setTextColor(80, 80, 80);
  doc.text("DECLARACIÓN DE ADECUACIÓN Y MODIFICACIÓN DE SOLICITUD DE PEDIDO", 108, 24, { align: "center" });

  doc.setFontSize(9.5);
  doc.setTextColor(0, 0, 0);
  doc.text(`Correspondiente a Solicitud N°: ${datos.numeroSolicitud || 'S/N'}`, 14, 36);
  doc.text(`Generado en Salida N°: ${datos.numeroSalida || 'S/N'}`, 130, 36);
  doc.text(`Sucursal Destino: ${datos.sucursal || 'General'}`, 14, 42);
  doc.text(`Fecha de Emisión: ${datos.fecha || ''}`, 130, 42);

  doc.setFont("helvetica", "normal");
  const textoDecl = `En base a las estadísticas revisadas y la disponibilidad de inventario en bodega, se modifica la Solicitud de Pedido N° ${datos.numeroSolicitud} y se adecua formalmente a las cantidades despachadas en la Salida N° ${datos.numeroSalida}.`;
  const lineas = doc.splitTextToSize(textoDecl, 185);
  doc.text(lineas, 14, 52);

  doc.autoTable({
    startY: 66,
    head: [['CÓDIGO', 'DESCRIPCIÓN MATERIAL', 'U.M.', 'SOLICITADA', 'DESPACHADA', 'DIFERENCIA']],
    body: datos.filas.map(it => [it.codigo, it.descripcion, it.um, it.cantSolicitada, it.cantAprobada, `-${it.diferencia}`]),
    theme: 'grid',
    headStyles: { fillColor: [26, 74, 117], textColor: [255, 255, 255], halign: 'center' },
    styles: { fontSize: 9, halign: 'center' },
    columnStyles: { 1: { halign: 'left' } }
  });

  const firmaY = doc.lastAutoTable.finalY + 35;
  doc.line(25, firmaY, 90, firmaY);
  doc.line(125, firmaY, 190, firmaY);
  doc.setFont("helvetica", "bold");
  doc.text(datos.coordinador || "Coordinador Solicitante", 57.5, firmaY + 5, { align: "center" });
  doc.text("Director Técnico (DT)", 157.5, firmaY + 5, { align: "center" });

  doc.save(`DECLARACION_AJUSTE_SOL_${datos.numeroSolicitud}_SAL_${datos.numeroSalida}.pdf`);
}

function descargarPdfReporteMensual() {
  if (!ULTIMO_REPORTE_MENSUAL || !ULTIMO_REPORTE_MENSUAL.datos.length) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'letter' });
  const r = ULTIMO_REPORTE_MENSUAL;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.setTextColor(26, 74, 117);
  doc.text("LABORATORIO BIOCLINIC - REPORTE MENSUAL", 14, 18);

  doc.setFontSize(10);
  doc.setTextColor(0, 0, 0);
  doc.text(`Sucursal: ${r.sucursal}   |   Período: ${r.mes} / ${r.anio}`, 14, 26);

  doc.autoTable({
    startY: 32,
    head: [['INSUMO', 'U.M.', 'CANT. PEDIDO', 'N° SOLICITUD', 'CANT. DESP.', 'N° SALIDA']],
    body: r.datos.map(it => [`${it.insumo} (${it.codigo})`, it.um, it.cantPedido, it.solicitudesTexto, it.cantDespachada, it.salidasTexto]),
    theme: 'grid',
    headStyles: { fillColor: [26, 74, 117], textColor: [255, 255, 255], halign: 'center' },
    styles: { fontSize: 8.5, halign: 'center' },
    columnStyles: { 0: { halign: 'left' } }
  });

  doc.save(`REPORTE_MENSUAL_${r.sucursal.replace(/[^a-zA-Z0-9]/g, "_")}_${r.mes}_${r.anio}.pdf`);
}

// ============================================================================
// 10. NAVEGACIÓN DE PESTAÑAS Y CIERRE DE SESIÓN
// ============================================================================

function cambiarPestaña(id, btn) {
  document.querySelectorAll('.tab-panel-content').forEach(p => {
    p.classList.remove('active');
    p.style.setProperty("display", "none", "important");
  });
  document.querySelectorAll('.nav-tabs-custom .nav-link').forEach(b => b.classList.remove('active'));
  
  var panelActivo = document.getElementById(id);
  if (panelActivo) {
    panelActivo.classList.add('active');
    panelActivo.style.setProperty("display", "block", "important");
  }
  if (btn) btn.classList.add('active');
  
  if (id === 'panel-manual') alternarCamposDinamicos();
  if (id === 'panel-aprobaciones') {
    var u = USUARIO_LOGUEADO.toUpperCase();
    var r = ROL_LOGUEADO.toUpperCase();
    if (u.includes("ADMIN") || r.includes("ADMIN") || u === "BODEGA") {
      cargarPedidosPendientesAdmin();
    } else {
      cargarPedidosPreAprobacionDT();
    }
  }
  if (id === 'panel-monitor') recargarTablaMonitor();
  if (id === 'panel-mermas') recargarTablaMermas();
  if (id === 'panel-subbodegas') recargarTablaSubbodegas();
}

function procesarArchivoCSV() {
  const input = document.getElementById("archivoCSV");
  if (!input.files.length) {
    mostrarNotificacion("Seleccione un archivo .CSV primero.", "warning");
    return;
  }
  const reader = new FileReader();
  reader.onload = async function(e) {
    const wbCsv = XLSX.read(e.target.result, { type: 'string' });
    const hojaCsv = wbCsv.Sheets[wbCsv.SheetNames[0]];
    const filas = XLSX.utils.sheet_to_json(hojaCsv, { header: 1 });
    if (filas.length <= 1) {
      mostrarNotificacion("El archivo CSV está vacío.", "error");
      return;
    }
    const matrizHistorial = obtenerMatrizHoja("Historial");
    const fechaActual = new Date();
    let importados = 0;

    for (let i = 1; i < filas.length; i++) {
      if (!filas[i][0] || !filas[i][2]) continue;
      matrizHistorial.push([
        fechaActual,
        USUARIO_LOGUEADO,
        filas[i][0].toString().toUpperCase().trim(),
        filas[i][1] || "",
        "",
        filas[i][2].toString().trim(),
        filas[i][3] || "",
        parseFloat(filas[i][4]) || 0,
        filas[i][5] || "S/V",
        "", "", "",
        "Importación CSV",
        filas[i][6] || "BULNES"
      ]);
      importados++;
    }
    guardarMatrizHoja("Historial", matrizHistorial);
    const ok = await guardarCambiosEnOneDrive();
    if (ok) {
      mostrarNotificacion(`Se importaron ${importados} filas al Historial.`, "success");
      input.value = "";
    }
  };
  reader.readAsText(input.files[0]);
}

function cerrarSesion() {
  document.getElementById("seccionApp").classList.add("d-none");
  document.getElementById("seccionLogin").classList.remove("d-none");
  document.getElementById("loginUsuario").value = "";
  document.getElementById("loginClave").value = "";
  document.getElementById("errorLogin").classList.add("d-none");
  
  LOTE_MANUAL_ITEMS = [];
  PEDIDOS_CARGADOS_CACHE = [];
  renderizarTablaLoteManual();
  USUARIO_LOGUEADO = "";
  ROL_LOGUEADO = "";
}
