// panel.js - Panel administrativo completo
const API_URL = 'https://backend-control-operarios.onrender.com';

// Credenciales
const USUARIO_PERMITIDO = 'admin';
const PASSWORD_PERMITIDO = 'empuvilla2025';

// 💰 SALARIOS BASE POR OPERARIO (ajústalos según corresponda)
const SALARIOS = {
  'Hernan Aragon': 1423500,
  'Rossa Viafara': 1423500,
  'Jaiver Casaran': 1423500,
  'Arnoldo Camacho': 1423500,
};

// Constantes laborales de Colombia 2025
const HORAS_MENSUALES_BASE = 240; // 8 horas/día * 30 días
const RECARGO_HORA_EXTRA = 1.25; // 25% de recargo
const RECARGO_DOMINICAL = 1.75; // 75% de recargo

// Elementos del DOM
const seccionLogin = document.getElementById('seccionLogin');
const seccionReporte = document.getElementById('seccionReporte');
const seccionPermisos = document.getElementById('seccionPermisos');
const mensajeLogin = document.getElementById('mensajeLogin');

const btnLogin = document.getElementById('btnLogin');
const btnConsultar = document.getElementById('btnConsultar');
const btnEliminar = document.getElementById('btnEliminar');
const btnMostrarPermisos = document.getElementById('btnMostrarPermisos');
const btnRegistrarPermiso = document.getElementById('btnRegistrarPermiso');

const operarioReporte = document.getElementById('operarioReporte');
const fechaDesde = document.getElementById('fechaDesde');
const fechaHasta = document.getElementById('fechaHasta');

const resumenDiv = document.getElementById('resumen');
const tablaDetalleDiv = document.getElementById('tablaDetalle');

const operarioPermiso = document.getElementById('operarioPermiso');
const fechaPermiso = document.getElementById('fechaPermiso');
const horasPermiso = document.getElementById('horasPermiso');
const motivoPermiso = document.getElementById('motivoPermiso');
const mensajePermiso = document.getElementById('mensajePermiso');
const listaPermisosDiv = document.getElementById('listaPermisos');

// 1. LOGIN
btnLogin.addEventListener('click', () => {
  const usuario = document.getElementById('usuario').value.trim();
  const password = document.getElementById('password').value.trim();

  if (usuario === USUARIO_PERMITIDO && password === PASSWORD_PERMITIDO) {
    mensajeLogin.textContent = 'Ingreso exitoso.';
    mensajeLogin.style.color = '#22c55e';

    seccionLogin.classList.add('oculto');
    seccionReporte.classList.remove('oculto');
  } else {
    mensajeLogin.textContent = 'Usuario o contraseña incorrectos.';
    mensajeLogin.style.color = '#f97316';
  }
});

// 2. CONSULTAR REPORTE CON CÁLCULO DE PAGO
btnConsultar.addEventListener('click', async () => {
  const operario = operarioReporte.value;
  const desde = fechaDesde.value;
  const hasta = fechaHasta.value;

  if (!operario || !desde || !hasta) {
    resumenDiv.textContent = 'Debes seleccionar operario, fecha desde y hasta.';
    resumenDiv.style.color = '#f97316';
    return;
  }

  try {
    const url = `${API_URL}/api/reporte-horas?operario=${encodeURIComponent(
      operario
    )}&desde=${desde}&hasta=${hasta}`;

    const resp = await fetch(url);
    const data = await resp.json();

    if (!resp.ok) {
      resumenDiv.textContent = data.mensaje || 'Error al obtener el reporte.';
      resumenDiv.style.color = '#f97316';
      return;
    }

    // Calcular valores monetarios
    const salarioBase = SALARIOS[operario] || 0;
    const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
    const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
    const valorHoraDominical = valorHoraNormal * RECARGO_DOMINICAL;

    const pagoNormal = data.horasNormales * valorHoraNormal;
    const pagoExtra = data.horasExtra * valorHoraExtra;
    const pagoDominical = data.horasDominicales * valorHoraDominical;
    const pagoTotal = pagoNormal + pagoExtra + pagoDominical;

    // Guardar datos para exportar
    datosReporteActual = data;

    // Calcular horas netas (trabajadas - permisos)
    const horasNetas = data.totalHoras - data.horasPermiso;

    // Mostrar resumen
    resumenDiv.style.color = '#e5e7eb';
    resumenDiv.innerHTML = `
      <div style="text-align: right; margin-bottom: 15px;">
        <button onclick="exportarExcel()" style="background: #059669; padding: 10px 20px; border-radius: 8px; border: none; color: white; font-weight: 600; cursor: pointer; font-size: 14px;">
          📥 Exportar a Excel
        </button>
      </div>
    `;
      <div style="background: #1f2937; padding: 20px; border-radius: 12px; margin-bottom: 15px;">
        <h3 style="margin: 0 0 15px 0; color: #60a5fa; font-size: 18px;">📊 Resumen de Horas</h3>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; font-size: 14px;">
          <div><strong>Operario:</strong> ${data.operario}</div>
          <div><strong>Período:</strong> ${data.desde} a ${data.hasta}</div>
          <div><strong>Salario base:</strong> $${salarioBase.toLocaleString('es-CO')}</div>
          <div><strong>Total trabajadas:</strong> ${data.totalHoras} h</div>
          <div><strong>Horas normales:</strong> ${data.horasNormales} h</div>
          <div><strong>Horas extra:</strong> ${data.horasExtra} h</div>
          <div><strong>Horas dominicales:</strong> ${data.horasDominicales} h</div>
          <div style="color: #fbbf24;"><strong>⚠️ Horas permiso:</strong> ${data.horasPermiso} h</div>
          <div style="color: #22c55e;"><strong>✅ Horas netas:</strong> ${horasNetas.toFixed(2)} h</div>
        </div>
      </div>

      <div style="background: #1f2937; padding: 20px; border-radius: 12px;">
        <h3 style="margin: 0 0 15px 0; color: #22c55e; font-size: 18px;">💰 Cálculo de Pago</h3>
        <div style="display: grid; gap: 10px; font-size: 14px;">
          <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
            <span>Pago horas normales (${data.horasNormales} h × $${valorHoraNormal.toFixed(0)}):</span>
            <strong style="color: #93c5fd;">$${pagoNormal.toLocaleString('es-CO', {maximumFractionDigits: 0})}</strong>
          </div>
          <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
            <span>Pago horas extra (${data.horasExtra} h × $${valorHoraExtra.toFixed(0)}):</span>
            <strong style="color: #fbbf24;">$${pagoExtra.toLocaleString('es-CO', {maximumFractionDigits: 0})}</strong>
          </div>
          <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
            <span>Pago horas dominicales (${data.horasDominicales} h × $${valorHoraDominical.toFixed(0)}):</span>
            <strong style="color: #c084fc;">$${pagoDominical.toLocaleString('es-CO', {maximumFractionDigits: 0})}</strong>
          </div>
          <div style="display: flex; justify-content: space-between; padding: 12px; background: #22c55e; border-radius: 6px; margin-top: 8px;">
            <span style="color: #052e16; font-weight: 700; font-size: 16px;">TOTAL A PAGAR:</span>
            <strong style="color: #052e16; font-size: 18px;">$${pagoTotal.toLocaleString('es-CO', {maximumFractionDigits: 0})}</strong>
          </div>
        </div>
      </div>
    `;

    // Mostrar detalle en tabla
    if (Array.isArray(data.detalle) && data.detalle.length > 0) {
      let html = `
        <div style="margin-top: 15px;">
          <h3 style="color: #60a5fa; margin-bottom: 10px;">📅 Detalle por Día</h3>
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Domingo</th>
                <th>Horas normales</th>
                <th>Horas extra</th>
                <th>Horas dominicales</th>
              </tr>
            </thead>
            <tbody>
      `;

      for (const fila of data.detalle) {
        html += `
          <tr>
            <td>${fila.fecha}</td>
            <td>${fila.domingo ? 'Sí' : 'No'}</td>
            <td>${fila.horasNormalesDia.toFixed(2)}</td>
            <td>${fila.horasExtraDia.toFixed(2)}</td>
            <td>${fila.horasDominicalesDia.toFixed(2)}</td>
          </tr>
        `;
      }

      html += '</tbody></table></div>';
      tablaDetalleDiv.innerHTML = html;
    } else {
      tablaDetalleDiv.innerHTML = '<p>No hay detalle para este rango.</p>';
    }
  } catch (err) {
    console.error(err);
    resumenDiv.textContent = 'No se pudo conectar con el servidor. Revisa tu conexión.';
    resumenDiv.style.color = '#f97316';
  }
});

// 3. ELIMINAR REGISTROS
btnEliminar.addEventListener('click', async () => {
  const operario = operarioReporte.value;
  const desde = fechaDesde.value;
  const hasta = fechaHasta.value;

  if (!operario || !desde || !hasta) {
    resumenDiv.textContent = 'Debes seleccionar operario, fecha desde y hasta antes de eliminar.';
    resumenDiv.style.color = '#f97316';
    return;
  }

  const seguro = window.confirm(
    `Vas a eliminar TODOS los registros de ${operario}\nentre ${desde} y ${hasta}.\n\nEsta acción no se puede deshacer.\n\n¿Estás seguro?`
  );

  if (!seguro) return;

  try {
    const resp = await fetch(`${API_URL}/api/registros-rango`, {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ operario, desde, hasta }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      resumenDiv.textContent = data.mensaje || 'Error al eliminar registros.';
      resumenDiv.style.color = '#f97316';
      return;
    }

    resumenDiv.style.color = '#22c55e';
    resumenDiv.textContent = `Se eliminaron ${data.eliminados} registros de ${operario} entre ${desde} y ${hasta}.`;
    tablaDetalleDiv.innerHTML = '';
  } catch (err) {
    console.error(err);
    resumenDiv.textContent = 'No se pudo conectar con el servidor al intentar eliminar.';
    resumenDiv.style.color = '#f97316';
  }
});

// 4. MOSTRAR SECCIÓN DE PERMISOS
btnMostrarPermisos.addEventListener('click', () => {
  seccionReporte.classList.add('oculto');
  seccionPermisos.classList.remove('oculto');
  cargarPermisos();
});

// 5. REGISTRAR PERMISO
btnRegistrarPermiso.addEventListener('click', async () => {
  const nombreOperario = operarioPermiso.value;
  const fechaPerm = fechaPermiso.value;
  const horas = horasPermiso.value;
  const motivo = motivoPermiso.value.trim();

  if (!nombreOperario || !fechaPerm || !horas || !motivo) {
    mensajePermiso.textContent = 'Debes completar todos los campos.';
    mensajePermiso.style.color = '#f97316';
    return;
  }

  try {
    const resp = await fetch(`${API_URL}/api/permisos`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        nombreOperario,
        fechaPermiso: fechaPerm,
        horasPermiso: parseFloat(horas),
        motivo,
      }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      mensajePermiso.textContent = data.mensaje || 'Error al registrar permiso.';
      mensajePermiso.style.color = '#f97316';
      return;
    }

    mensajePermiso.textContent = '✅ Permiso registrado correctamente.';
    mensajePermiso.style.color = '#22c55e';

    // Limpiar formulario
    horasPermiso.value = '';
    motivoPermiso.value = '';

    // Recargar lista
    cargarPermisos();
  } catch (err) {
    console.error(err);
    mensajePermiso.textContent = 'No se pudo conectar con el servidor.';
    mensajePermiso.style.color = '#f97316';
  }
});

// 6. CARGAR LISTA DE PERMISOS
async function cargarPermisos() {
  const operario = operarioPermiso.value;
  if (!operario) {
    listaPermisosDiv.innerHTML = '<p style="color: #9ca3af;">Selecciona un operario para ver sus permisos.</p>';
    return;
  }

  // Cargar permisos del mes actual
  const ahora = new Date();
  const año = ahora.getFullYear();
  const mes = String(ahora.getMonth() + 1).padStart(2, '0');
  const desde = `${año}-${mes}-01`;
  const hasta = `${año}-${mes}-31`;

  try {
    const url = `${API_URL}/api/permisos?operario=${encodeURIComponent(
      operario
    )}&desde=${desde}&hasta=${hasta}`;

    const resp = await fetch(url);
    const permisos = await resp.json();

    if (!resp.ok) {
      listaPermisosDiv.innerHTML = '<p style="color: #f97316;">Error al cargar permisos.</p>';
      return;
    }

    if (!Array.isArray(permisos) || permisos.length === 0) {
      listaPermisosDiv.innerHTML = '<p style="color: #9ca3af;">No hay permisos registrados este mes.</p>';
      return;
    }

    const totalHoras = permisos.reduce((sum, p) => sum + p.horasPermiso, 0);

    let html = `
      <div style="background: #1f2937; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <strong style="color: #60a5fa;">Total horas de permiso este mes: ${totalHoras.toFixed(2)} h</strong>
      </div>
      <table>
        <thead>
          <tr>
            <th>Fecha</th>
            <th>Horas</th>
            <th>Motivo</th>
            <th>Acción</th>
          </tr>
        </thead>
        <tbody>
    `;

    for (const permiso of permisos) {
      html += `
        <tr>
          <td>${permiso.fechaPermiso}</td>
          <td>${permiso.horasPermiso}</td>
          <td>${permiso.motivo}</td>
          <td>
            <button onclick="eliminarPermiso('${permiso._id}')" 
                    style="padding: 4px 10px; background: #b91c1c; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 12px;">
              Eliminar
            </button>
          </td>
        </tr>
      `;
    }

    html += '</tbody></table>';
    listaPermisosDiv.innerHTML = html;
  } catch (err) {
    console.error(err);
    listaPermisosDiv.innerHTML = '<p style="color: #f97316;">No se pudo conectar con el servidor.</p>';
  }
}

// 7. ELIMINAR PERMISO
window.eliminarPermiso = async function (id) {
  if (!confirm('¿Estás seguro de eliminar este permiso?')) return;

  try {
    const resp = await fetch(`${API_URL}/api/permisos/${id}`, {
      method: 'DELETE',
    });

    const data = await resp.json();

    if (!resp.ok) {
      alert(data.mensaje || 'Error al eliminar permiso.');
      return;
    }

    mensajePermiso.textContent = '✅ Permiso eliminado correctamente.';
    mensajePermiso.style.color = '#22c55e';
    cargarPermisos();
  } catch (err) {
    console.error(err);
    alert('No se pudo conectar con el servidor.');
  }
};

// Cargar permisos cuando cambie el operario
operarioPermiso.addEventListener('change', cargarPermisos);

// ========== EXPORTAR A EXCEL ==========

let datosReporteActual = null; // Guardar datos del último reporte

// Función para exportar a Excel
window.exportarExcel = function() {
  if (!datosReporteActual) {
    alert('Primero debes consultar un reporte para poder exportarlo.');
    return;
  }

  const data = datosReporteActual;
  const operario = data.operario;
  const desde = data.desde;
  const hasta = data.hasta;

  // Calcular valores monetarios
  const salarioBase = SALARIOS[operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraDominical = valorHoraNormal * RECARGO_DOMINICAL;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoDominical = data.horasDominicales * valorHoraDominical;
  const pagoTotal = pagoNormal + pagoExtra + pagoDominical;
  const horasNetas = data.totalHoras - data.horasPermiso;

  // Crear el contenido del Excel en formato CSV
  let csv = '';
  
  // ENCABEZADO
  csv += 'EMPUVILLA S.A. E.S.P.\n';
  csv += 'REPORTE DE HORAS TRABAJADAS Y LIQUIDACIÓN\n';
  csv += '\n';
  
  // INFORMACIÓN GENERAL
  csv += 'INFORMACIÓN GENERAL\n';
  csv += `Operario:,${operario}\n`;
  csv += `Período:,${desde} a ${hasta}\n`;
  csv += `Fecha de generación:,${new Date().toLocaleDateString('es-CO')}\n`;
  csv += `Salario base mensual:,${salarioBase.toLocaleString('es-CO')}\n`;
  csv += '\n';

  // RESUMEN DE HORAS
  csv += 'RESUMEN DE HORAS\n';
  csv += 'Concepto,Cantidad (horas)\n';
  csv += `Total trabajadas,${data.totalHoras}\n`;
  csv += `Horas normales,${data.horasNormales}\n`;
  csv += `Horas extra,${data.horasExtra}\n`;
  csv += `Horas dominicales,${data.horasDominicales}\n`;
  csv += `Horas permiso,${data.horasPermiso}\n`;
  csv += `Horas netas (trabajadas - permisos),${horasNetas.toFixed(2)}\n`;
  csv += '\n';

  // CÁLCULO DE LIQUIDACIÓN
  csv += 'CÁLCULO DE LIQUIDACIÓN\n';
  csv += 'Concepto,Horas,Valor por hora,Subtotal\n';
  csv += `Horas normales,${data.horasNormales},${valorHoraNormal.toFixed(0)},${Math.round(pagoNormal).toLocaleString('es-CO')}\n`;
  csv += `Horas extra (+25%),${data.horasExtra},${valorHoraExtra.toFixed(0)},${Math.round(pagoExtra).toLocaleString('es-CO')}\n`;
  csv += `Horas dominicales (+75%),${data.horasDominicales},${valorHoraDominical.toFixed(0)},${Math.round(pagoDominical).toLocaleString('es-CO')}\n`;
  csv += `TOTAL A PAGAR,,,${Math.round(pagoTotal).toLocaleString('es-CO')}\n`;
  csv += '\n';

  // DETALLE DIARIO
  if (Array.isArray(data.detalle) && data.detalle.length > 0) {
    csv += 'DETALLE DIARIO\n';
    csv += 'Fecha,Domingo,Horas normales,Horas extra,Horas dominicales\n';
    
    for (const fila of data.detalle) {
      csv += `${fila.fecha},${fila.domingo ? 'Sí' : 'No'},${fila.horasNormalesDia.toFixed(2)},${fila.horasExtraDia.toFixed(2)},${fila.horasDominicalesDia.toFixed(2)}\n`;
    }
  }

  // Crear archivo y descargar
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  const nombreArchivo = `Reporte_${operario.replace(/\s+/g, '_')}_${desde}_${hasta}.csv`;
  
  link.setAttribute('href', url);
  link.setAttribute('download', nombreArchivo);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
