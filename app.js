// Configuración de conexión a Supabase - Bioclinic
const SUPABASE_URL = 'https://xekvvmtfwobwttluknvn.supabase.co';
const SUPABASE_KEY = 'sb_publishable_-qdaMaDIAFsF9MCRZ9RjgQ_eM-AjhRE';
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
    document.getElementById(btn.dataset.tab).classList.remove('hidden');
  });
});

function mostrarMensaje(texto, tipo = 'success') {
  const el = document.getElementById('global-msg');
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
    .eq('password', password)
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
  document.getElementById('user-banner').textContent =
    `Usuario: ${currentUser.nombre} | Rol: ${currentUser.rol} | Sucursal: ${currentUser.sucursal}`;

  // Permisos visuales según Rol
  const isAdmin = currentUser.rol === 'ADMIN';
  document.getElementById('nav-inventario').classList.toggle('hidden', !isAdmin);
  document.getElementById('nav-movimientos').classList.toggle('hidden', !isAdmin);
  document.getElementById('nav-personal').classList.toggle('hidden', !isAdmin);

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
    db.from('inventario').select('*').order('descripcion', { ascending: true }),
    db.from('solicitudes').select('*').order('fecha', { ascending: false }),
    db.from('movimientos').select('*').order('fecha', { ascending: false }),
    db.from('personal').select('*').order('nombre', { ascending: true })
  ]);

  cacheInventario = invRes.data || [];
  cacheSolicitudes = solRes.data || [];
  cacheMovimientos = movRes.data || [];
  cachePersonal = perRes.data || [];

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
  document.getElementById('sol-producto').innerHTML = placeholder + opciones;
  document.getElementById('mov-producto').innerHTML = placeholder + opciones;
}

// 1. GESTIÓN DE SOLICITUDES
document.getElementById('form-solicitud').addEventListener('submit', async (e) => {
  e.preventDefault();
  const codigo = document.getElementById('sol-producto').value;
  const cantidad = Number(document.getElementById('sol-cantidad').value);
  const observacion = document.getElementById('sol-obs').value.trim();
  const prod = cacheInventario.find(p => p.codigo === codigo);

  if (!prod) return;

  const { error } = await db.from('solicitudes').insert([{
    sucursal: currentUser.sucursal,
    solicitante: currentUser.nombre,
    codigo_producto: prod.codigo,
    descripcion: prod.descripcion,
    cantidad,
    estado: 'PENDIENTE_DT',
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
  let lista = cacheSolicitudes;

  // Si es SUCURSAL, solo ve las solicitudes de su propia sucursal
  if (currentUser.rol === 'SUCURSAL') {
    lista = lista.filter(s => s.sucursal === currentUser.sucursal);
  }

  tbody.innerHTML = lista.map(s => {
    const fecha = new Date(s.fecha).toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' });
    let badgeClass = 'badge-pendiente';
    if (s.estado === 'APROBADA') badgeClass = 'badge-aprobada';
    if (s.estado === 'DESPACHADA') badgeClass = 'badge-despachada';
    if (s.estado === 'RECHAZADA') badgeClass = 'badge-rechazada';

    let acciones = '-';
    if ((currentUser.rol === 'DT' || currentUser.rol === 'ADMIN') && s.estado === 'PENDIENTE_DT') {
      acciones = `
        <button class="btn btn-sm btn-success" onclick="cambiarEstadoSolicitud(${s.id}, 'APROBADA')">Aprobar</button>
        <button class="btn btn-sm btn-danger" onclick="cambiarEstadoSolicitud(${s.id}, 'RECHAZADA')">Rechazar</button>
      `;
    } else if (currentUser.rol === 'ADMIN' && s.estado === 'APROBADA') {
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
  const { error } = await db.from('solicitudes').update({ estado: nuevoEstado }).eq('id', id);
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

  await db.from('inventario').update({ stock_actual: nuevoStock }).eq('codigo', prod.codigo);
  await db.from('solicitudes').update({ estado: 'DESPACHADA' }).eq('id', id);
  await db.from('movimientos').insert([{
    tipo: 'SALIDA',
    codigo_producto: prod.codigo,
    descripcion: prod.descripcion,
    cantidad: sol.cantidad,
    destino_motivo: `Despacho Solicitud #${sol.id} - ${sol.sucursal}`,
    responsable: currentUser.nombre
  }]);

  mostrarMensaje('Solicitud despachada y stock descontado automáticamente.');
  cargarTodo();
}

// 2. INVENTARIO CENTRAL (ADMIN)
document.getElementById('form-producto').addEventListener('submit', async (e) => {
  e.preventDefault();
  const payload = {
    codigo: document.getElementById('inv-codigo').value.trim().toUpperCase(),
    descripcion: document.getElementById('inv-desc').value.trim(),
    categoria: document.getElementById('inv-cat').value.trim(),
    unidad: document.getElementById('inv-unidad').value.trim(),
    stock_actual: Number(document.getElementById('inv-stock').value),
    stock_minimo: Number(document.getElementById('inv-min').value)
  };

  const { error } = await db.from('inventario').upsert([payload]);
  if (error) return mostrarMensaje('Error al guardar producto: ' + error.message, 'error');

  document.getElementById('form-producto').reset();
  mostrarMensaje('Producto guardado en inventario central.');
  cargarTodo();
});

function renderInventario() {
  const tbody = document.getElementById('tbody-inventario');
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

  await db.from('inventario').update({ stock_actual: nuevoStock }).eq('codigo', prod.codigo);
  await db.from('movimientos').insert([{
    tipo,
    codigo_producto: prod.codigo,
    descripcion: prod.descripcion,
    cantidad,
    destino_motivo,
    responsable: currentUser.nombre
  }]);

  document.getElementById('form-movimiento').reset();
  mostrarMensaje(`${tipo} registrada y descontada del stock.`);
  cargarTodo();
});

function renderMovimientos() {
  const tbody = document.getElementById('tbody-movimientos');
  tbody.innerHTML = cacheMovimientos.map(m => {
    const fecha = new Date(m.fecha).toLocaleString('es-CL', { dateStyle: 'short', timeStyle: 'short' });
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
    password: document.getElementById('usr-pass').value.trim(),
    nombre: document.getElementById('usr-nombre').value.trim(),
    rol: document.getElementById('usr-rol').value,
    sucursal: document.getElementById('usr-sucursal').value.trim().toUpperCase()
  };

  const { error } = await db.from('personal').insert([payload]);
  if (error) return mostrarMensaje('Error al crear usuario: ' + error.message, 'error');

  document.getElementById('form-usuario').reset();
  mostrarMensaje('Usuario creado correctamente.');
  cargarTodo();
});

function renderPersonal() {
  const tbody = document.getElementById('tbody-personal');
  tbody.innerHTML = cachePersonal.map(u => `
    <tr>
      <td><strong>${u.usuario}</strong></td>
      <td>${u.nombre}</td>
      <td>${u.rol}</td>
      <td>${u.sucursal}</td>
      <td>
        ${u.usuario !== 'admin'
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
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheInventario), 'Inventario');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheSolicitudes), 'Solicitudes');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cacheMovimientos), 'Salidas_y_Mermas');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cachePersonal), 'Personal');
  XLSX.writeFile(wb, `Bioclinic_Inventario_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
