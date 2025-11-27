// panel.js - Panel administrativo con soporte PTAP/PTAR y exportación Excel
const API_URL = 'https://backend-control-operarios.onrender.com';

// Credenciales
const USUARIO_PERMITIDO = 'admin';
const PASSWORD_PERMITIDO = 'empuvilla2025';

// 💰 SALARIOS POR OPERARIO Y PLANTA
const SALARIOS = {
  PTAP: {
    'Hernan Aragon': 1423500,
    'Rossa Viafara': 1423500,
    'Jaiver Casaran': 1423500,
    'Arnoldo Camacho': 1423500,
  },
  PTAR: {
    'Bladimir Cifuentes': 1423500,
    'Willington Granja': 1423500,
    'Rigoberto Arrechea': 1423500,
  }
};

// Constantes laborales Colombia 2025
const HORAS_MENSUALES_BASE = 240;
const HORAS_SEMANALES_PTAR = 45;
const RECARGO_HORA_EXTRA = 1.25;
const RECARGO_DOMINICAL = 1.75;
const RECARGO_NOCTURNO = 1.35; // 35% adicional

// Operarios por planta
const OPERARIOS = {
  PTAP: ['Hernan Aragon', 'Rossa Viafara', 'Jaiver Casaran', 'Arnoldo Camacho'],
  PTAR: ['Bladimir Cifuentes', 'Willington Granja', 'Rigoberto Arrechea']
};

// Elementos DOM
const seccionLogin = document.getElementById('seccionLogin');
const seccionReporte = document.getElementById('seccionReporte');
const seccionPermisos = document.getElementById('seccionPermisos');
const mensajeLogin = document.getElementById('mensajeLogin');

const btnLogin = document.getElementById('btnLogin');
const btnConsultar = document.getElementById('btnConsultar');
const btnEliminar = document.getElementById('btnEliminar');
const btnMostrarPermisos = document.getElementById('btnMostrarPermisos');
const btnRegistrarPermiso = document.getElementById('btnRegistrarPermiso');

const plantaReporte = document.getElementById('plantaReporte');
const operarioReporte = document.getElementById('operarioReporte');
const fechaDesde = document.getElementById('fechaDesde');
const fechaHasta = document.getElementById('fechaHasta');

const resumenDiv = document.getElementById('resumen');
const tablaDetalleDiv = document.getElementById('tablaDetalle');

const plantaPermiso = document.getElementById('plantaPermiso');
const operarioPermiso = document.getElementById('operarioPermiso');
const fechaPermiso = document.getElementById('fechaPermiso');
const horasPermiso = document.getElementById('horasPermiso');
const motivoPermiso = document.getElementById('motivoPermiso');
const mensajePermiso = document.getElementById('mensajePermiso');
const listaPermisosDiv = document.getElementById('listaPermisos');

let datosReporteActual = null;

// ==================== LOGIN ====================

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

// ==================== SELECTOR DE PLANTA EN REPORTE ====================

plantaReporte.addEventListener('change', () => {
  const planta = plantaReporte.value;
  if (!planta) {
    operarioReporte.innerHTML = '<option value="">Primero selecciona planta...</option>';
    return;
  }

  operarioReporte.innerHTML = '<option value="">Seleccione...</option>';
  OPERARIOS[planta].forEach(op => {
    operarioReporte.innerHTML += `<option value="${op}">${op}</option>`;
  });
});

plantaPermiso.addEventListener('change', () => {
  const planta = plantaPermiso.value;
  if (!planta) {
    operarioPermiso.innerHTML = '<option value="">Primero selecciona planta...</option>';
    return;
  }

  operarioPermiso.innerHTML = '<option value="">Seleccione...</option>';
  OPERARIOS[planta].forEach(op => {
    operarioPermiso.innerHTML += `<option value="${op}">${op}</option>`;
  });
});

// ==================== CONSULTAR REPORTE ====================

btnConsultar.addEventListener('click', async () => {
  const planta = plantaReporte.value;
  const operario = operarioReporte.value;
  const desde = fechaDesde.value;
  const hasta = fechaHasta.value;

  if (!planta || !operario || !desde || !hasta) {
    resumenDiv.textContent = 'Completa todos los campos';
    resumenDiv.style.color = '#f97316';
    return;
  }

  try {
    const url = `${API_URL}/api/reporte-horas?operario=${encodeURIComponent(operario)}&planta=${planta}&desde=${desde}&hasta=${hasta}`;
    const resp = await fetch(url);
    const data = await resp.json();

    if (!resp.ok) {
      resumenDiv.textContent = data.mensaje || 'Error al obtener reporte';
      resumenDiv.style.color = '#f97316';
      return;
    }

    datosReporteActual = data;

    if (planta === 'PTAP') {
      mostrarReportePTAP(data);
    } else {
      mostrarReportePTAR(data);
    }
  } catch (err) {
    console.error(err);
    resumenDiv.textContent = 'Error de conexión';
    resumenDiv.style.color = '#f97316';
  }
});

// ==================== MOSTRAR REPORTE PTAP ====================

function mostrarReportePTAP(data) {
  const salarioBase = SALARIOS.PTAP[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraDominical = valorHoraNormal * RECARGO_DOMINICAL;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoDominical = data.horasDominicales * valorHoraDominical;
  const pagoTotal = pagoNormal + pagoExtra + pagoDominical;

  const horasNetas = data.totalHoras - data.horasPermiso;

  resumenDiv.style.color = '#e5e7eb';
  resumenDiv.innerHTML = `
    <div style="text-align: right; margin-bottom: 15px;">
      <button onclick="exportarExcel()" style="background: #059669; padding: 10px 20px; border-radius: 8px; border: none; color: white; font-weight: 600; cursor: pointer;">
        📥 Exportar a Excel
      </button>
    </div>
    
    <div style="background: #1f2937; padding: 20px; border-radius: 12px; margin-bottom: 15px;">
      <h3 style="margin: 0 0 15px 0; color: #60a5fa;">📊 PTAP - Resumen de Horas</h3>
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
      <h3 style="margin: 0 0 15px 0; color: #22c55e;">💰 Cálculo de Pago</h3>
      <div style="display: grid; gap: 10px;">
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas normales (${data.horasNormales} h × $${valorHoraNormal.toFixed(0)}):</span>
          <strong style="color: #93c5fd;">$${Math.round(pagoNormal).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas extra (${data.horasExtra} h × $${valorHoraExtra.toFixed(0)}):</span>
          <strong style="color: #fbbf24;">$${Math.round(pagoExtra).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas dominicales (${data.horasDominicales} h × $${valorHoraDominical.toFixed(0)}):</span>
          <strong style="color: #c084fc;">$${Math.round(pagoDominical).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 12px; background: #22c55e; border-radius: 6px; margin-top: 8px;">
          <span style="color: #052e16; font-weight: 700; font-size: 16px;">TOTAL A PAGAR:</span>
          <strong style="color: #052e16; font-size: 18px;">$${Math.round(pagoTotal).toLocaleString('es-CO')}</strong>
        </div>
      </div>
    </div>
  `;

  // Detalle diario
  if (data.detalle && data.detalle.length > 0) {
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
  }
}

// ==================== MOSTRAR REPORTE PTAR ====================

function mostrarReportePTAR(data) {
  const salarioBase = SALARIOS.PTAR[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraNocturna = valorHoraNormal * RECARGO_NOCTURNO;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoNocturno = data.horasNocturnas * valorHoraNocturna;
  const pagoTotal = pagoNormal + pagoExtra + pagoNocturno;

  const horasNetas = data.totalHoras - data.horasPermiso;

  resumenDiv.style.color = '#e5e7eb';
  resumenDiv.innerHTML = `
    <div style="text-align: right; margin-bottom: 15px;">
      <button onclick="exportarExcel()" style="background: #059669; padding: 10px 20px; border-radius: 8px; border: none; color: white; font-weight: 600; cursor: pointer;">
        📥 Exportar a Excel
      </button>
    </div>
    
    <div style="background: #1f2937; padding: 20px; border-radius: 12px; margin-bottom: 15px;">
      <h3 style="margin: 0 0 15px 0; color: #60a5fa;">🌙 PTAR - Resumen de Horas</h3>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; font-size: 14px;">
        <div><strong>Operario:</strong> ${data.operario}</div>
        <div><strong>Período:</strong> ${data.desde} a ${data.hasta}</div>
        <div><strong>Salario base:</strong> $${salarioBase.toLocaleString('es-CO')}</div>
        <div><strong>Total trabajadas:</strong> ${data.totalHoras} h</div>
        <div><strong>Horas normales:</strong> ${data.horasNormales} h (hasta 45h/semana)</div>
        <div><strong>Horas extra:</strong> ${data.horasExtra} h (sobre 45h/semana)</div>
        <div style="color: #a78bfa;"><strong>🌙 Horas nocturnas:</strong> ${data.horasNocturnas} h (19:00-06:00)</div>
        <div style="color: #fbbf24;"><strong>⚠️ Horas permiso:</strong> ${data.horasPermiso} h</div>
        <div style="color: #22c55e;"><strong>✅ Horas netas:</strong> ${horasNetas.toFixed(2)} h</div>
      </div>
    </div>

    <div style="background: #1f2937; padding: 20px; border-radius: 12px;">
      <h3 style="margin: 0 0 15px 0; color: #22c55e;">💰 Cálculo de Pago</h3>
      <div style="display: grid; gap: 10px;">
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas normales (${data.horasNormales} h × $${valorHoraNormal.toFixed(0)}):</span>
          <strong style="color: #93c5fd;">$${Math.round(pagoNormal).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas extra +25% (${data.horasExtra} h × $${valorHoraExtra.toFixed(0)}):</span>
          <strong style="color: #fbbf24;">$${Math.round(pagoExtra).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 8px; background: #111827; border-radius: 6px;">
          <span>Horas nocturnas +35% (${data.horasNocturnas} h × $${valorHoraNocturna.toFixed(0)}):</span>
          <strong style="color: #a78bfa;">$${Math.round(pagoNocturno).toLocaleString('es-CO')}</strong>
        </div>
        <div style="display: flex; justify-content: space-between; padding: 12px; background: #22c55e; border-radius: 6px; margin-top: 8px;">
          <span style="color: #052e16; font-weight: 700; font-size: 16px;">TOTAL A PAGAR:</span>
          <strong style="color: #052e16; font-size: 18px;">$${Math.round(pagoTotal).toLocaleString('es-CO')}</strong>
        </div>
      </div>
    </div>
  `;

  // Detalle semanal
  if (data.detalleSemanas && data.detalleSemanas.length > 0) {
    let html = `
      <div style="margin-top: 15px;">
        <h3 style="color: #60a5fa; margin-bottom: 10px;">📆 Detalle por Semana</h3>
        <table>
          <thead>
            <tr>
              <th>Semana</th>
              <th>Total horas</th>
              <th>Normales</th>
              <th>Extra</th>
              <th>Nocturnas</th>
            </tr>
          </thead>
          <tbody>
    `;

    for (const sem of data.detalleSemanas) {
      html += `
        <tr>
          <td>${sem.inicio} a ${sem.fin}</td>
          <td>${sem.horasTotales.toFixed(2)}</td>
          <td>${sem.horasNormales.toFixed(2)}</td>
          <td>${sem.horasExtra.toFixed(2)}</td>
          <td>${sem.horasNocturnas.toFixed(2)}</td>
        </tr>
      `;
    }

    html += '</tbody></table></div>';
    tablaDetalleDiv.innerHTML = html;
  }
}

// ==================== ELIMINAR REGISTROS ====================

btnEliminar.addEventListener('click', async () => {
  const planta = plantaReporte.value;
  const operario = operarioReporte.value;
  const desde = fechaDesde.value;
  const hasta = fechaHasta.value;

  if (!planta || !operario || !desde || !hasta) {
    resumenDiv.textContent = 'Completa todos los campos';
    resumenDiv.style.color = '#f97316';
    return;
  }

  if (!confirm(`¿Eliminar TODOS los registros de ${operario} (${planta}) entre ${desde} y ${hasta}?`)) return;

  try {
    const resp = await fetch(`${API_URL}/api/registros-rango`, {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ operario, planta, desde, hasta }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      resumenDiv.textContent = data.mensaje || 'Error al eliminar';
      resumenDiv.style.color = '#f97316';
      return;
    }

    resumenDiv.style.color = '#22c55e';
    resumenDiv.textContent = `Se eliminaron ${data.eliminados} registros`;
    tablaDetalleDiv.innerHTML = '';
  } catch (err) {
    console.error(err);
    resumenDiv.textContent = 'Error de conexión';
    resumenDiv.style.color = '#f97316';
  }
});

// ==================== GESTIÓN DE PERMISOS ====================

btnMostrarPermisos.addEventListener('click', () => {
  seccionReporte.classList.add('oculto');
  seccionPermisos.classList.remove('oculto');
  cargarPermisos();
});

btnRegistrarPermiso.addEventListener('click', async () => {
  const planta = plantaPermiso.value;
  const operario = operarioPermiso.value;
  const fecha = fechaPermiso.value;
  const horas = horasPermiso.value;
  const motivo = motivoPermiso.value.trim();

  if (!planta || !operario || !fecha || !horas || !motivo) {
    mensajePermiso.textContent = 'Completa todos los campos';
    mensajePermiso.style.color = '#f97316';
    return;
  }

  try {
    const resp = await fetch(`${API_URL}/api/permisos`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        nombreOperario: operario,
        planta,
        fechaPermiso: fecha,
        horasPermiso: parseFloat(horas),
        motivo,
      }),
    });

    const data = await resp.json();

    if (!resp.ok) {
      mensajePermiso.textContent = data.mensaje || 'Error';
      mensajePermiso.style.color = '#f97316';
      return;
    }

    mensajePermiso.textContent = '✅ Permiso registrado';
    mensajePermiso.style.color = '#22c55e';

    horasPermiso.value = '';
    motivoPermiso.value = '';
    cargarPermisos();
  } catch (err) {
    console.error(err);
    mensajePermiso.textContent = 'Error de conexión';
    mensajePermiso.style.color = '#f97316';
  }
});

async function cargarPermisos() {
  const planta = plantaPermiso.value;
  const operario = operarioPermiso.value;
  
  if (!planta || !operario) {
    listaPermisosDiv.innerHTML = '<p style="color: #9ca3af;">Selecciona planta y operario</p>';
    return;
  }

  const ahora = new Date();
  const año = ahora.getFullYear();
  const mes = String(ahora.getMonth() + 1).padStart(2, '0');
  const desde = `${año}-${mes}-01`;
  const hasta = `${año}-${mes}-31`;

  try {
    const url = `${API_URL}/api/permisos?operario=${encodeURIComponent(operario)}&planta=${planta}&desde=${desde}&hasta=${hasta}`;
    const resp = await fetch(url);
    const permisos = await resp.json();

    if (!resp.ok) {
      listaPermisosDiv.innerHTML = '<p style="color: #f97316;">Error al cargar</p>';
      return;
    }

    if (!Array.isArray(permisos) || permisos.length === 0) {
      listaPermisosDiv.innerHTML = '<p style="color: #9ca3af;">Sin permisos este mes</p>';
      return;
    }

    const totalHoras = permisos.reduce((sum, p) => sum + p.horasPermiso, 0);

    let html = `
      <div style="background: #1f2937; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <strong style="color: #60a5fa;">Total horas permiso: ${totalHoras.toFixed(2)} h</strong>
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
                    style="padding: 4px 10px; background: #b91c1c; color: white; border: none; border-radius: 6px; cursor: pointer;">
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
    listaPermisosDiv.innerHTML = '<p style="color: #f97316;">Error de conexión</p>';
  }
}

window.eliminarPermiso = async function(id) {
  if (!confirm('¿Eliminar este permiso?')) return;

  try {
    const resp = await fetch(`${API_URL}/api/permisos/${id}`, {
      method: 'DELETE',
    });

    if (!resp.ok) {
      alert('Error al eliminar');
      return;
    }

    mensajePermiso.textContent = '✅ Permiso eliminado';
    mensajePermiso.style.color = '#22c55e';
    cargarPermisos();
  } catch (err) {
    console.error(err);
    alert('Error de conexión');
  }
};

operarioPermiso.addEventListener('change', cargarPermisos);

// ==================== EXPORTAR A EXCEL ====================

window.exportarExcel = async function() {
  if (!datosReporteActual) {
    alert('Primero consulta un reporte');
    return;
  }

  const data = datosReporteActual;
  const wb = XLSX.utils.book_new();

  if (data.planta === 'PTAP') {
    generarExcelPTAP(wb, data);
  } else {
    generarExcelPTAR(wb, data);
  }

  const nombreArchivo = `Reporte_${data.planta}_${data.operario.replace(/\s+/g, '_')}_${data.desde}_${data.hasta}.xlsx`;
  XLSX.writeFile(wb, nombreArchivo);

  mostrarMensajeExito('✅ Excel descargado correctamente');
};

function generarExcelPTAP(wb, data) {
  const salarioBase = SALARIOS.PTAP[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraDominical = valorHoraNormal * RECARGO_DOMINICAL;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoDominical = data.horasDominicales * valorHoraDominical;
  const pagoTotal = pagoNormal + pagoExtra + pagoDominical;
  const horasNetas = (data.totalHoras - data.horasPermiso).toFixed(2);

  // ================= HOJA 1: RESUMEN =================
  const resumen = [
    ['EMPUVILLA S.A. E.S.P. - PTAP'],
    ['REPORTE DE HORAS TRABAJADAS Y LIQUIDACIÓN'],
    [],
    ['INFORMACIÓN GENERAL'],
    ['Operario:', data.operario],
    ['Planta:', 'PTAP'],
    ['Período:', `${data.desde} a ${data.hasta}`],
    ['Fecha generación:', new Date().toLocaleDateString('es-CO')],
    ['Salario base:', `${salarioBase.toLocaleString('es-CO')}`],
    [],
    ['RESUMEN DE HORAS'],
    ['Concepto', 'Cantidad (h)'],
    ['Total trabajadas', data.totalHoras],
    ['Horas normales', data.horasNormales],
    ['Horas extra', data.horasExtra],
    ['Horas dominicales', data.horasDominicales],
    ['Horas permiso', data.horasPermiso],
    ['Horas netas', horasNetas],
    [],
    ['LIQUIDACIÓN'],
    ['Concepto', 'Horas', 'Valor/hora', 'Subtotal'],
    ['Horas normales', data.horasNormales, `${Math.round(valorHoraNormal).toLocaleString('es-CO')}`, `${Math.round(pagoNormal).toLocaleString('es-CO')}`],
    ['Horas extra (+25%)', data.horasExtra, `${Math.round(valorHoraExtra).toLocaleString('es-CO')}`, `${Math.round(pagoExtra).toLocaleString('es-CO')}`],
    ['Horas dominicales (+75%)', data.horasDominicales, `${Math.round(valorHoraDominical).toLocaleString('es-CO')}`, `${Math.round(pagoDominical).toLocaleString('es-CO')}`],
    [],
    ['TOTAL A PAGAR', '', '', `${Math.round(pagoTotal).toLocaleString('es-CO')}`],
  ];

  const wsResumen = XLSX.utils.aoa_to_sheet(resumen);

  // Ancho de columnas
  wsResumen['!cols'] = [
    { wch: 35 },
    { wch: 20 },
    { wch: 18 },
    { wch: 18 },
  ];

  // Unir celdas para títulos y secciones
  wsResumen['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }, // A1:D1
    { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } }, // A2:D2
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } }, // "INFORMACIÓN GENERAL"
    { s: { r: 10, c: 0 }, e: { r: 10, c: 3 } }, // "RESUMEN DE HORAS"
    { s: { r: 19, c: 0 }, e: { r: 19, c: 3 } }, // "LIQUIDACIÓN"
  ];

  // Helper para aplicar estilo a una celda
  function setCellStyle(ws, addr, style) {
    if (!ws[addr]) return;
    ws[addr].s = Object.assign({}, ws[addr].s || {}, style);
  }

  // Estilos base
  const tituloPrincipal = {
    font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1D4ED8' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  };

  const tituloSecundario = {
    font: { bold: true, sz: 12, color: { rgb: 'E5E7EB' } },
    fill: { patternType: 'solid', fgColor: { rgb: '111827' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  };

  const encabezadoTabla = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
    alignment: { horizontal: 'center' },
  };

  const filaTotal = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '16A34A' } },
  };

  // Aplicar estilos títulos
  setCellStyle(wsResumen, 'A1', tituloPrincipal);
  setCellStyle(wsResumen, 'A2', tituloSecundario);
  setCellStyle(wsResumen, 'A4', tituloSecundario);  // INFORMACIÓN GENERAL
  setCellStyle(wsResumen, 'A11', tituloSecundario); // RESUMEN DE HORAS
  setCellStyle(wsResumen, 'A20', tituloSecundario); // LIQUIDACIÓN

  // Encabezados de tablas
  setCellStyle(wsResumen, 'A12', encabezadoTabla);
  setCellStyle(wsResumen, 'B12', encabezadoTabla);
  setCellStyle(wsResumen, 'A21', encabezadoTabla);
  setCellStyle(wsResumen, 'B21', encabezadoTabla);
  setCellStyle(wsResumen, 'C21', encabezadoTabla);
  setCellStyle(wsResumen, 'D21', encabezadoTabla);

  // Fila TOTAL A PAGAR (A26:D26 → índice fila 25)
  setCellStyle(wsResumen, 'A26', filaTotal);
  setCellStyle(wsResumen, 'D26', filaTotal);

  XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

  // ================= HOJA 2: DETALLE DIARIO =================
  if (data.detalle && data.detalle.length > 0) {
    const detalle = [
      ['DETALLE DIARIO - PTAP'],
      ['Operario:', data.operario],
      [],
      ['Fecha', 'Domingo', 'H. Normales', 'H. Extra', 'H. Dominicales', 'Total'],
    ];

    for (const d of data.detalle) {
      const total = d.horasNormalesDia + d.horasExtraDia + d.horasDominicalesDia;
      detalle.push([
        d.fecha,
        d.domingo ? 'Sí' : 'No',
        d.horasNormalesDia.toFixed(2),
        d.horasExtraDia.toFixed(2),
        d.horasDominicalesDia.toFixed(2),
        total.toFixed(2),
      ]);
    }

    detalle.push([]);
    detalle.push([
      'TOTALES',
      '',
      data.horasNormales.toFixed(2),
      data.horasExtra.toFixed(2),
      data.horasDominicales.toFixed(2),
      data.totalHoras.toFixed(2),
    ]);

    const wsDetalle = XLSX.utils.aoa_to_sheet(detalle);
    wsDetalle['!cols'] = [
      { wch: 15 },
      { wch: 10 },
      { wch: 12 },
      { wch: 12 },
      { wch: 15 },
      { wch: 12 },
    ];

    // Merges título
    wsDetalle['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }, // A1:F1
    ];

    // Estilos título y encabezado
    setCellStyle(wsDetalle, 'A1', tituloPrincipal);
    setCellStyle(wsDetalle, 'A4', encabezadoTabla);
    setCellStyle(wsDetalle, 'B4', encabezadoTabla);
    setCellStyle(wsDetalle, 'C4', encabezadoTabla);
    setCellStyle(wsDetalle, 'D4', encabezadoTabla);
    setCellStyle(wsDetalle, 'E4', encabezadoTabla);
    setCellStyle(wsDetalle, 'F4', encabezadoTabla);

    // Fila de totales (última)
    const ultimaFila = 4 + data.detalle.length + 2; // 0-based: fila de "TOTALES"
    const totalRowAddr = ['A', 'B', 'C', 'D', 'E', 'F'].map(
      (col) => `${col}${ultimaFila + 1}`,
    );
    totalRowAddr.forEach((addr) => setCellStyle(wsDetalle, addr, filaTotal));

    XLSX.utils.book_append_sheet(wb, wsDetalle, 'Detalle Diario');
  }
}


function generarExcelPTAR(wb, data) {
  const salarioBase = SALARIOS.PTAR[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraNocturna = valorHoraNormal * RECARGO_NOCTURNO;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoNocturno = data.horasNocturnas * valorHoraNocturna;
  const pagoTotal = pagoNormal + pagoExtra + pagoNocturno;
  const horasNetas = (data.totalHoras - data.horasPermiso).toFixed(2);

  const resumen = [
    ['EMPUVILLA S.A. E.S.P. - PTAR'],
    ['REPORTE DE HORAS TRABAJADAS Y LIQUIDACIÓN'],
    [],
    ['INFORMACIÓN GENERAL'],
    ['Operario:', data.operario],
    ['Planta:', 'PTAR'],
    ['Período:', `${data.desde} a ${data.hasta}`],
    ['Fecha generación:', new Date().toLocaleDateString('es-CO')],
    ['Salario base:', `${salarioBase.toLocaleString('es-CO')}`],
    [],
    ['RESUMEN DE HORAS'],
    ['Concepto', 'Cantidad (h)'],
    ['Total trabajadas', data.totalHoras],
    ['Horas normales (hasta 45h/sem)', data.horasNormales],
    ['Horas extra (sobre 45h/sem)', data.horasExtra],
    ['Horas nocturnas (19:00-06:00)', data.horasNocturnas],
    ['Horas permiso', data.horasPermiso],
    ['Horas netas', horasNetas],
    [],
    ['LIQUIDACIÓN'],
    ['Concepto', 'Horas', 'Valor/hora', 'Subtotal'],
    ['Horas normales', data.horasNormales, `${Math.round(valorHoraNormal).toLocaleString('es-CO')}`, `${Math.round(pagoNormal).toLocaleString('es-CO')}`],
    ['Horas extra (+25%)', data.horasExtra, `${Math.round(valorHoraExtra).toLocaleString('es-CO')}`, `${Math.round(pagoExtra).toLocaleString('es-CO')}`],
    ['Horas nocturnas (+35%)', data.horasNocturnas, `${Math.round(valorHoraNocturna).toLocaleString('es-CO')}`, `${Math.round(pagoNocturno).toLocaleString('es-CO')}`],
    [],
    ['TOTAL A PAGAR', '', '', `${Math.round(pagoTotal).toLocaleString('es-CO')}`],
  ];

  const wsResumen = XLSX.utils.aoa_to_sheet(resumen);
  wsResumen['!cols'] = [
    { wch: 35 },
    { wch: 20 },
    { wch: 20 },
    { wch: 20 },
  ];

  // Merges para títulos y secciones
  wsResumen['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }, // A1:D1
    { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } }, // A2:D2
    { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } }, // INFORMACIÓN GENERAL
    { s: { r: 10, c: 0 }, e: { r: 10, c: 3 } }, // RESUMEN DE HORAS
    { s: { r: 19, c: 0 }, e: { r: 19, c: 3 } }, // LIQUIDACIÓN
  ];

  function setCellStyle(ws, addr, style) {
    if (!ws[addr]) return;
    ws[addr].s = Object.assign({}, ws[addr].s || {}, style);
  }

  const tituloPrincipal = {
    font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1E293B' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  };

  const tituloSecundario = {
    font: { bold: true, sz: 12, color: { rgb: 'E5E7EB' } },
    fill: { patternType: 'solid', fgColor: { rgb: '020617' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  };

  const encabezadoTabla = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '0F172A' } },
    alignment: { horizontal: 'center' },
  };

  const filaTotal = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '16A34A' } },
  };

  // Estilos de títulos y encabezados
  setCellStyle(wsResumen, 'A1', tituloPrincipal);
  setCellStyle(wsResumen, 'A2', tituloSecundario);
  setCellStyle(wsResumen, 'A4', tituloSecundario);
  setCellStyle(wsResumen, 'A11', tituloSecundario);
  setCellStyle(wsResumen, 'A20', tituloSecundario);

  setCellStyle(wsResumen, 'A12', encabezadoTabla);
  setCellStyle(wsResumen, 'B12', encabezadoTabla);
  setCellStyle(wsResumen, 'A21', encabezadoTabla);
  setCellStyle(wsResumen, 'B21', encabezadoTabla);
  setCellStyle(wsResumen, 'C21', encabezadoTabla);
  setCellStyle(wsResumen, 'D21', encabezadoTabla);

  // Total a pagar
  setCellStyle(wsResumen, 'A26', filaTotal);
  setCellStyle(wsResumen, 'D26', filaTotal);

  XLSX.utils.book_append_sheet(wb, wsResumen, 'Resumen');

  // ================= HOJA 2: DETALLE SEMANAL =================
  if (data.detalleSemanas && data.detalleSemanas.length > 0) {
    const detalle = [
      ['DETALLE POR SEMANA - PTAR'],
      ['Operario:', data.operario],
      ['Límite semanal: 45 horas'],
      [],
      ['Semana', 'Total', 'Normales', 'Extra', 'Nocturnas'],
    ];

    for (const sem of data.detalleSemanas) {
      detalle.push([
        `${sem.inicio} a ${sem.fin}`,
        sem.horasTotales.toFixed(2),
        sem.horasNormales.toFixed(2),
        sem.horasExtra.toFixed(2),
        sem.horasNocturnas.toFixed(2),
      ]);
    }

    detalle.push([]);
    detalle.push([
      'TOTALES',
      data.totalHoras.toFixed(2),
      data.horasNormales.toFixed(2),
      data.horasExtra.toFixed(2),
      data.horasNocturnas.toFixed(2),
    ]);

    const wsDetalle = XLSX.utils.aoa_to_sheet(detalle);
    wsDetalle['!cols'] = [
      { wch: 25 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
      { wch: 12 },
    ];

    wsDetalle['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // A1:E1
    ];

    // Reutilizamos estilos
    setCellStyle(wsDetalle, 'A1', tituloPrincipal);
    setCellStyle(wsDetalle, 'A5', encabezadoTabla);
    setCellStyle(wsDetalle, 'B5', encabezadoTabla);
    setCellStyle(wsDetalle, 'C5', encabezadoTabla);
    setCellStyle(wsDetalle, 'D5', encabezadoTabla);
    setCellStyle(wsDetalle, 'E5', encabezadoTabla);

    // Fila de totales (al final)
    const ultimaFila = 5 + data.detalleSemanas.length + 2; // línea "TOTALES"
    ['A', 'B', 'C', 'D', 'E'].forEach((col) =>
      setCellStyle(wsDetalle, `${col}${ultimaFila + 1}`, filaTotal),
    );

    XLSX.utils.book_append_sheet(wb, wsDetalle, 'Detalle Semanal');
  }
}

function mostrarMensajeExito(texto) {
  const mensaje = document.createElement('div');
  mensaje.textContent = texto;
  mensaje.style.cssText = 'position: fixed; top: 20px; right: 20px; background: #22c55e; color: white; padding: 15px 25px; border-radius: 8px; font-weight: 600; box-shadow: 0 4px 12px rgba(0,0,0,0.3); z-index: 9999;';
  document.body.appendChild(mensaje);
  
  setTimeout(() => mensaje.remove(), 3000);
}

