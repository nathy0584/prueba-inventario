// Configuración de conexión a Supabase - Bioclinic
const SUPABASE_URL = 'https://inbjanhqikoqozsxcwnk.supabase.co';
const SUPABASE_KEY = 'sb_publishable_RnxQUIpE2Ak-2Q2VnXPFHA_xHwAdW59';
const db = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

let currentUser = null;
let cacheInventario = [];
let cacheSolicitudes = [];
let cacheMovimientos = [];
let cachePersonal = [];

// Navegación por pestañas
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
    btn.classList.add('active');
    const target = document.getElementById(btn.dataset.tab);
    if (target) target.classList.remove('hidden');
  });
});

function mostrarMensaje(texto, tipo = 'success') {
  const el = document.getElementById('global-msg');
  if (!el) return;
  el.textContent = texto;
  el.className = `alert alert-${tipo}`;
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 4500);
}

// INICIO DE SESIÓN
document.getElementById('login-form').addEventListener('submit', async (e) => {
  e.preventDefault();
  const usuario = document.getElementById('login-user').value.trim();
  const password = document.getElementById('login-pass').value.trim();
  const errBox = document.getElementById('login-error');
  const btn = document.getElementById('btn-login');

  errBox.classList.add('hidden');
  btn.disabled = true;
  btn.textContent = 'Verificando...';

  const { data, error } = await db
    .from('personal')
    .select('*')
    .eq('usuario', usuario)
    .eq('clave', password)
    .maybeSingle();

  btn.disabled = false;
  btn.textContent = 'Ingresar al Sistema';

  if (error || !data) {
    errBox.textContent = 'Usuario o contraseña incorrectos.';
    errBox.classList.remove('hidden');
    return;
  }

  currentUser = data;
  iniciarSesionUI();
});

function iniciarSesionUI() {
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('app-screen').classList.remove('hidden');

  const sucursalTexto = currentUser.sucursales || currentUser.sucursal || '-';
  document.getElementById('user-banner').textContent =
    `Usuario: ${currentUser.usuario} | Rol: ${currentUser.rol} | Sucursal: ${sucursalTexto}`;

  const rolUpper = (currentUser.rol || '').toUpperCase();
  const isAdmin = rolUpper.includes('ADMIN');

  const navInv = document.getElementById('nav-inventario');
  const navMov = document.getElementById('nav-movimientos');
  const navPer = document.getElementById('nav-personal');
  if (navInv) navInv.classList.toggle('hidden', !isAdmin);
  if (navMov) navMov.classList.toggle('hidden', !isAdmin);
  if (navPer) navPer.classList.toggle('hidden', !isAdmin);

  cargarTodo();
}

function cerrarSesion() {
  currentUser = null;
  document.getElementById('login-form').reset();
  document.getElementById('app-screen').classList.add('hidden');
  document.getElementById('login-screen').classList.remove('hidden');
}

// CARGA CENTRAL DE DATOS
async function cargarTodo() {
  const [invRes, solRes, movRes, perRes] = await Promise.all([
    db.from('inventario_actual').select('*').order('nombre', { ascending: true }),
    db.from('pedidos').select('*').order('created_at', { ascending: false }),
    db.from('historial').select('*').order('created_at', { ascending: false }),
    db.from('personal').select('*').order('usuario', { ascending: true })
  ]);

  cacheInventario = (invRes.data || []).map(p => ({
    ...p,
    codigo: p.id,
    descripcion: p.nombre,
    unidad: p.um,
    stock_actual: Number(p.stock) || 0,
    stock_minimo: Number(p.proyeccion_trimestral) || 0
  }));

  cacheSolicitudes = (solRes.data || []).map(s => ({
    ...s,
    fecha: s.created_at,
    codigo_producto: s.id_insumo
  }));

  cacheMovimientos = (movRes.data || []).map(m => ({
    ...m,
    fecha: m.created_at,
    codigo_producto: m.id_insumo,
    destino_motivo: m.destino_origen || m.notas || '',
    responsable: m.operador
  }));

  cachePersonal = (perRes.data || []).map(u => ({
    ...u,
    nombre: u.usuario,
    sucursal: u.sucursales || ''
  }));

  renderSelectoresProductos();
  renderInventario();
  renderSolicitudes();
  renderMovimientos();
  renderPersonal();
}

function renderSelectoresProductos() {
  const opciones = cacheInventario.map(p =>
    `<option value="${p.codigo}">${p.codigo} - ${p.descripcion} (Disp: ${p.stock_actual} ${p.unidad || ''})</option>`
  ).join('');

  const placeholder = '<option value="">-- Seleccione un producto --</option>';
  const selSol = document.getElementById('sol-producto');
  const selMov = document.getElementById('mov-producto');
  if (selSol) selSol.innerHTML = placeholder + opciones;
  if (selMov) selMov.innerHTML = placeholder + opciones;
}

// 1. GESTIÓN DE SOLICITUDES
document.getElementById('form-solicitud').addEventListener('submit', async (e) => {
  e.preventDefault();
  const codigo = document.getElementById('sol-producto').value;
  const cantidad = Number(document.getElementById('sol-cantidad').value);
  const observacion = document.getElementById('sol-obs').value.trim();
  const prod = cacheInventario.find(p => p.codigo === codigo);

  if (!prod || cantidad <= 0) return;

  const { data: ult } = await db.from('pedidos').select('folio').order('folio', { ascending: false }).limit(1);
  const folio = (ult && ult.length ? Number(ult[0].folio) || 0 : 0) + 1;
  const sucursalOrigen = (currentUser.sucursales || currentUser.sucursal || '').split(',')[0].trim() || currentUser.usuario;

  const { error } = await db.from('pedidos').insert([{
    folio,
    sucursal: sucursalOrigen,
    solicitante: currentUser.usuario,
    id_insumo: prod.codigo,
    descripcion: prod.descripcion,
    um: prod.unidad || 'UN',
    cantidad,
    estado: 'PRE-APROBACION',
    observacion
  }]);

  if (error) {
    mostrarMensaje('Error al enviar solicitud: ' + error.message, 'error');
  } else {
    document.getElementById('form-solicitud').reset();
    mostrarMensaje('Solicitud enviada correctamente a Coordinación/DT.');
    cargarTodo();
  }
});

function renderSolicitudes() {
  const tbody = document.getElementById('tbody-solicitudes');
  if (!tbody) return;
  let lista = cacheSolicitudes;
  const rolUpper = (currentUser.rol || '').toUpperCase();

  if (rolUpper === 'USUARIO' || rolUpper === 'SUCURSAL') {
    const misSucursales = (currentUser.sucursales || currentUser.sucursal || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
    if (misSucursales.length) {
      lista = lista.filter(s => misSucursales.includes((s.sucursal || '').toUpperCase()));
    }
  }

  tbody.innerHTML = lista.map(s => {
    const fecha = s.fecha ? new Date(s.fecha).toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' }) : '-';
    let badgeClass = 'badge-pendiente';
    if (s.estado === 'APROBADO_DT' || s.estado === 'APROBADA') badgeClass = 'badge-aprobada';
    if (s.estado === 'DESPACHADA' || s.estado === 'DESPACHADO') badgeClass = 'badge-despachada';
    if (s.estado === 'RECHAZADA') badgeClass = 'badge-rechazada';

    let acciones = '-';
    const esDT = rolUpper.includes('DIRECTOR') || rolUpper.includes('TECNICO') || rolUpper === 'DT';
    const esAdmin = rolUpper.includes('ADMIN');

    if ((esDT || esAdmin) && (s.estado === 'PRE-APROBACION' || s.estado === 'PENDIENTE_DT')) {
      acciones = `
        <button class="btn btn-sm btn-success" onclick="cambiarEstadoSolicitud(${s.id}, 'APROBADO_DT')">Aprobar</button>
        <button class="btn btn-sm btn-danger" onclick="cambiarEstadoSolicitud(${s.id}, 'RECHAZADA')">Rechazar</button>
      `;
    } else if (esAdmin && (s.estado === 'APROBADO_DT' || s.estado === 'APROBADA')) {
      acciones = `
        <button class="btn btn-sm btn-primary" onclick="despacharSolicitud(${s.id})">Despachar y Descontar Stock</button>
      `;
    }

    return `
      <tr>
        <td>${fecha}</td>
        <td><strong>${s.sucursal}</strong></td>
        <td>${s.solicitante}</td>
        <td>${s.codigo_producto} - ${s.descripcion}</td>
        <td><strong>${s.cantidad}</strong></td>
        <td><span class="badge ${badgeClass}">${s.estado}</span></td>
        <td>${s.observacion || ''}</td>
        <td>${acciones}</td>
      </tr>
    `;
  }).join('');
}

async function cambiarEstadoSolicitud(id, nuevoEstado) {
  const { error } = await db.from('pedidos').update({ estado: nuevoEstado }).eq('id', id);
  if (error) return mostrarMensaje('Error al actualizar estado', 'error');
  mostrarMensaje(`Solicitud marcada como ${nuevoEstado}.`);
  cargarTodo();
}

async function despacharSolicitud(id) {
  const sol = cacheSolicitudes.find(s => s.id === id);
  const prod = cacheInventario.find(p => p.codigo === sol.codigo_producto);
  if (!sol || !prod) return;

  if (Number(prod.stock_actual) < Number(sol.cantidad)) {
    return mostrarMensaje(`Stock insuficiente en bodega (Disponible: ${prod.stock_actual}).`, 'error');
  }

  const nuevoStock = Number(prod.stock_actual) - Number(sol.cantidad);
  const { data: ult } = await db.from('historial').select('correlativo').eq('tipo', 'SALIDA').order('correlativo', { ascending: false }).limit(1);
  const correlativo = (ult && ult.length ? Number(ult[0].correlativo) || 0 : 0) + 1;

  await db.from('inventario_actual').update({ stock: nuevoStock }).eq('id', prod.codigo);
  await db.from('pedidos').update({ estado: 'DESPACHADO' }).eq('id', id);
  await db.from('historial').insert([{
    tipo: 'SALIDA',
    id_insumo: prod.codigo,
    descripcion: prod.descripcion,
    cantidad: sol.cantidad,
    destino_origen: sol.sucursal,
    referencia_nota: `Solicitud #${sol.folio || sol.id}`,
    correlativo,
    operador: currentUser.usuario
  }]);

  mostrarMensaje('Solicitud despachada y stock descontado automáticamente.');
  cargarTodo();
}

// 2. INVENTARIO CENTRAL (ADMIN)
document.getElementById('form-producto').addEventListener('submit', async (e) => {
  e.preventDefault();
  const payload = {
    id: document.getElementById('inv-codigo').value.trim().toUpperCase(),
    nombre: document.getElementById('inv-desc').value.trim(),
    categoria: document.getElementById('inv-cat').value.trim(),
    um: document.getElementById('inv-unidad').value.trim(),
    stock: Number(document.getElementById('inv-stock').value),
    proyeccion_trimestral: Number(document.getElementById('inv-min').value)
  };

  const { error } = await db.from('inventario_actual').upsert([payload]);
  if (error) return mostrarMensaje('Error al guardar producto: ' + error.message, 'error');

  document.getElementById('form-producto').reset();
  mostrarMensaje('Producto guardado en inventario central.');
  cargarTodo();
});

function renderInventario() {
  const tbody = document.getElementById('tbody-inventario');
  if (!tbody) return;
  tbody.innerHTML = cacheInventario.map(p => {
    const bajo = Number(p.stock_actual) <= Number(p.stock_minimo);
    return `
      <tr>
        <td><strong>${p.codigo}</strong></td>
        <td>${p.descripcion}</td>
        <td>${p.categoria || '-'}</td>
        <td>${p.unidad || '-'}</td>
        <td><strong>${p.stock_actual}</strong></td>
        <td>${p.stock_minimo}</td>
        <td><span class="badge ${bajo ? 'badge-bajo' : 'badge-ok'}">${bajo ? 'STOCK BAJO' : 'ÓPTIMO'}</span></td>
      </tr>
    `;
  }).join('');
}

// 3. SALIDAS Y MERMAS (ADMIN)
document.getElementById('form-movimiento').addEventListener('submit', async (e) => {
  e.preventDefault();
  const tipo = document.getElementById('mov-tipo').value;
  const codigo = document.getElementById('mov-producto').value;
  const cantidad = Number(document.getElementById('mov-cantidad').value);
  const destino_motivo = document.getElementById('mov-motivo').value.trim();
  const prod = cacheInventario.find(p => p.codigo === codigo);

  if (!prod) return;
  if (Number(prod.stock_actual) < cantidad) {
    return mostrarMensaje(`Stock insuficiente (Disponible: ${prod.stock_actual}).`, 'error');
  }

  const nuevoStock = Number(prod.stock_actual) - cantidad;
  const { data: ult } = await db.from('historial').select('correlativo').eq('tipo', tipo).order('correlativo', { ascending: false }).limit(1);
  const correlativo = (ult && ult.length ? Number(ult[0].correlativo) || 0 : 0) + 1;

  await db.from('inventario_actual').update({ stock: nuevoStock }).eq('id', prod.codigo);
  await db.from('historial').insert([{
    tipo,
    id_insumo: prod.codigo,
    descripcion: prod.descripcion,
    cantidad,
    destino_origen: destino_motivo,
    notas: destino_motivo,
    correlativo,
    operador: currentUser.usuario
  }]);

  document.getElementById('form-movimiento').reset();
  mostrarMensaje(`${tipo} registrada y descontada del stock.`);
  cargarTodo();
});

function renderMovimientos() {
  const tbody = document.getElementById('tbody-movimientos');
  if (!tbody) return;
  tbody.innerHTML = cacheMovimientos.map(m => {
    const fecha = m.fecha ? new Date(m.fecha).toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' }) : '-';
    return `
      <tr>
        <td>${fecha}</td>
        <td><span class="badge ${m.tipo === 'MERMA' ? 'badge-rechazada' : 'badge-aprobada'}">${m.tipo}</span></td>
        <td>${m.codigo_producto}</td>
        <td>${m.descripcion}</td>
        <td><strong>${m.cantidad}</strong></td>
        <td>${m.destino_motivo}</td>
        <td>${m.responsable}</td>
      </tr>
    `;
  }).join('');
}

// 4. USUARIOS Y ROLES (ADMIN)
document.getElementById('form-usuario').addEventListener('submit', async (e) => {
  e.preventDefault();
  const payload = {
    usuario: document.getElementById('usr-user').value.trim(),
    clave: document.getElementById('usr-pass').value.trim(),
    rol: document.getElementById('usr-rol').value,
    sucursales: document.getElementById('usr-sucursal').value.trim().toUpperCase()
  };

  const { error } = await db.from('personal').insert([payload]);
  if (error) return mostrarMensaje('Error al crear usuario: ' + error.message, 'error');

  document.getElementById('form-usuario').reset();
  mostrarMensaje('Usuario creado correctamente.');
  cargarTodo();
});

function renderPersonal() {
  const tbody = document.getElementById('tbody-personal');
  if (!tbody) return;
  tbody.innerHTML = cachePersonal.map(u => `
    <tr>
      <td><strong>${u.usuario}</strong></td>
      <td>${u.nombre}</td>
      <td>${u.rol}</td>
      <td>${u.sucursal}</td>
      <td>
        ${u.usuario.toLowerCase() !== 'admin'
          ? `<button class="btn btn-sm btn-danger" onclick="eliminarUsuario(${u.id})">Eliminar</button>`
          : 'Principal'}
      </td>
    </tr>
  `).join('');
}

async function eliminarUsuario(id) {
  await db.from('personal').delete().eq('id', id);
  mostrarMensaje('Usuario eliminado.');
  cargarTodo();
}

// EXPORTAR RESPALDO A EXCEL (ADMIN)
function exportarExcel() {
  if (typeof XLSX === 'undefined') {
    return mostrarMensaje('Librería XLSX no cargada en la página.', 'error');
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheInventario), 'Inventario');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheSolicitudes), 'Solicitudes');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheMovimientos), 'Salidas_y_Mermas');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cachePersonal), 'Personal');
  XLSX.writeFile(wb, `Bioclinic_Inventario_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
