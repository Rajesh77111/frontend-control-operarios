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

// NUEVA EXPORTACIÓN A EXCEL CON ESTILOS (ExcelJS)
window.exportarExcel = async function () {
  if (!datosReporteActual) {
    alert('Primero consulta un reporte');
    return;
  }

  const data = datosReporteActual;
  const workbook = new ExcelJS.Workbook();

  if (data.planta === 'PTAP') {
    crearExcelPTAP(workbook, data);
  } else {
    crearExcelPTAR(workbook, data);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const nombreArchivo = `Reporte_${data.planta}_${data.operario.replace(/\s+/g, '_')}_${data.desde}_${data.hasta}.xlsx`;
  saveAs(blob, nombreArchivo);

  mostrarMensajeExito('✅ Excel descargado correctamente');
};

// NUEVA EXPORTACIÓN PTAP
function crearExcelPTAP(workbook, data) {
  const salarioBase = SALARIOS.PTAP[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraDominical = valorHoraNormal * RECARGO_DOMINICAL;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoDominical = data.horasDominicales * valorHoraDominical;
  const pagoTotal = pagoNormal + pagoExtra + pagoDominical;
  const horasNetas = (data.totalHoras - data.horasPermiso).toFixed(2);

  // ===== Hoja 1: Resumen =====
  const ws = workbook.addWorksheet('Resumen');

  ws.columns = [
    { header: '', key: 'col1', width: 35 },
    { header: '', key: 'col2', width: 20 },
    { header: '', key: 'col3', width: 18 },
    { header: '', key: 'col4', width: 18 },
  ];

  // Título principal
  ws.mergeCells('A1:D1');
  const titulo1 = ws.getCell('A1');
  titulo1.value = 'EMPUVILLA S.A. E.S.P. - PTAP';
  titulo1.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  titulo1.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };

  // Subtítulo
  ws.mergeCells('A2:D2');
  const titulo2 = ws.getCell('A2');
  titulo2.value = 'REPORTE DE HORAS TRABAJADAS Y LIQUIDACIÓN';
  titulo2.font = { bold: true, size: 12, color: { argb: 'FFE5E7EB' } };
  titulo2.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };

  ws.addRow([]);
  ws.addRow(['INFORMACIÓN GENERAL']);
  const infoGeneralRow = ws.getRow(ws.lastRow.number);
  infoGeneralRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  infoGeneralRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
  ws.mergeCells(`A${infoGeneralRow.number}:D${infoGeneralRow.number}`);

  ws.addRow(['Operario:', data.operario]);
  ws.addRow(['Planta:', 'PTAP']);
  ws.addRow(['Período:', `${data.desde} a ${data.hasta}`]);
  ws.addRow(['Fecha generación:', new Date().toLocaleDateString('es-CO')]);
  ws.addRow(['Salario base:', salarioBase.toLocaleString('es-CO')]);
  ws.addRow([]);

  // Resumen de horas
  ws.addRow(['RESUMEN DE HORAS']);
  const resumenHorasRow = ws.getRow(ws.lastRow.number);
  resumenHorasRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  resumenHorasRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
  ws.mergeCells(`A${resumenHorasRow.number}:D${resumenHorasRow.number}`);

  const headerHoras = ws.addRow(['Concepto', 'Cantidad (h)']);
  headerHoras.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerHoras.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

  ws.addRow(['Total trabajadas', data.totalHoras]);
  ws.addRow(['Horas normales', data.horasNormales]);
  ws.addRow(['Horas extra', data.horasExtra]);
  ws.addRow(['Horas dominicales', data.horasDominicales]);
  ws.addRow(['Horas permiso', data.horasPermiso]);
  ws.addRow(['Horas netas', horasNetas]);
  ws.addRow([]);

  // Liquidación
  ws.addRow(['LIQUIDACIÓN']);
  const liqRow = ws.getRow(ws.lastRow.number);
  liqRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  liqRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
  ws.mergeCells(`A${liqRow.number}:D${liqRow.number}`);

  const headerLiq = ws.addRow(['Concepto', 'Horas', 'Valor/hora', 'Subtotal']);
  headerLiq.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerLiq.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

  ws.addRow(['Horas normales', data.horasNormales, Math.round(valorHoraNormal).toLocaleString('es-CO'), Math.round(pagoNormal).toLocaleString('es-CO')]);
  ws.addRow(['Horas extra (+25%)', data.horasExtra, Math.round(valorHoraExtra).toLocaleString('es-CO'), Math.round(pagoExtra).toLocaleString('es-CO')]);
  ws.addRow(['Horas dominicales (+75%)', data.horasDominicales, Math.round(valorHoraDominical).toLocaleString('es-CO'), Math.round(pagoDominical).toLocaleString('es-CO')]);
  ws.addRow([]);

  const totalRow = ws.addRow(['TOTAL A PAGAR', '', '', Math.round(pagoTotal).toLocaleString('es-CO')]);
  totalRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };

  // ===== Hoja 2: Detalle Diario =====
  if (data.detalle && data.detalle.length > 0) {
    const wsDet = workbook.addWorksheet('Detalle Diario');
    wsDet.columns = [
      { header: 'Fecha', key: 'fecha', width: 15 },
      { header: 'Domingo', key: 'dom', width: 10 },
      { header: 'H. Normales', key: 'hn', width: 15 },
      { header: 'H. Extra', key: 'he', width: 12 },
      { header: 'H. Dominicales', key: 'hd', width: 18 },
      { header: 'Total', key: 'tot', width: 12 },
    ];

    const tituloDet = wsDet.addRow(['DETALLE DIARIO - PTAP']);
    wsDet.mergeCells('A1:F1');
    tituloDet.alignment = { horizontal: 'center' };
    tituloDet.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    tituloDet.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };

    wsDet.addRow(['Operario:', data.operario]);
    wsDet.addRow([]);
    const header = wsDet.getRow(4);
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

    for (const d of data.detalle) {
      const totalDia = d.horasNormalesDia + d.horasExtraDia + d.horasDominicalesDia;
      wsDet.addRow([
        d.fecha,
        d.domingo ? 'Sí' : 'No',
        d.horasNormalesDia.toFixed(2),
        d.horasExtraDia.toFixed(2),
        d.horasDominicalesDia.toFixed(2),
        totalDia.toFixed(2),
      ]);
    }

    wsDet.addRow([]);
    const totales = wsDet.addRow([
      'TOTALES',
      '',
      data.horasNormales.toFixed(2),
      data.horasExtra.toFixed(2),
      data.horasDominicales.toFixed(2),
      data.totalHoras.toFixed(2),
    ]);
    totales.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    totales.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };
  }
}

// NUEVA EXPORTACIÓN PTAR
function crearExcelPTAR(workbook, data) {
  const salarioBase = SALARIOS.PTAR[data.operario] || 0;
  const valorHoraNormal = salarioBase / HORAS_MENSUALES_BASE;
  const valorHoraExtra = valorHoraNormal * RECARGO_HORA_EXTRA;
  const valorHoraNocturna = valorHoraNormal * RECARGO_NOCTURNO;

  const pagoNormal = data.horasNormales * valorHoraNormal;
  const pagoExtra = data.horasExtra * valorHoraExtra;
  const pagoNocturno = data.horasNocturnas * valorHoraNocturna;
  const pagoTotal = pagoNormal + pagoExtra + pagoNocturno;
  const horasNetas = (data.totalHoras - data.horasPermiso).toFixed(2);

  // ===== Hoja 1: Resumen =====
  const ws = workbook.addWorksheet('Resumen');

  ws.columns = [
    { header: '', key: 'col1', width: 35 },
    { header: '', key: 'col2', width: 20 },
    { header: '', key: 'col3', width: 20 },
    { header: '', key: 'col4', width: 20 },
  ];

  ws.mergeCells('A1:D1');
  const titulo1 = ws.getCell('A1');
  titulo1.value = 'EMPUVILLA S.A. E.S.P. - PTAR';
  titulo1.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
  titulo1.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };

  ws.mergeCells('A2:D2');
  const titulo2 = ws.getCell('A2');
  titulo2.value = 'REPORTE DE HORAS TRABAJADAS Y LIQUIDACIÓN';
  titulo2.font = { bold: true, size: 12, color: { argb: 'FFE5E7EB' } };
  titulo2.alignment = { horizontal: 'center', vertical: 'middle' };
  titulo2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF020617' } };

  ws.addRow([]);
  ws.addRow(['INFORMACIÓN GENERAL']);
  const infoGeneralRow = ws.getRow(ws.lastRow.number);
  infoGeneralRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  infoGeneralRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF020617' } };
  ws.mergeCells(`A${infoGeneralRow.number}:D${infoGeneralRow.number}`);

  ws.addRow(['Operario:', data.operario]);
  ws.addRow(['Planta:', 'PTAR']);
  ws.addRow(['Período:', `${data.desde} a ${data.hasta}`]);
  ws.addRow(['Fecha generación:', new Date().toLocaleDateString('es-CO')]);
  ws.addRow(['Salario base:', salarioBase.toLocaleString('es-CO')]);
  ws.addRow([]);

  ws.addRow(['RESUMEN DE HORAS']);
  const resumenHorasRow = ws.getRow(ws.lastRow.number);
  resumenHorasRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  resumenHorasRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF020617' } };
  ws.mergeCells(`A${resumenHorasRow.number}:D${resumenHorasRow.number}`);

  const headerHoras = ws.addRow(['Concepto', 'Cantidad (h)']);
  headerHoras.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerHoras.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

  ws.addRow(['Total trabajadas', data.totalHoras]);
  ws.addRow(['Horas normales (hasta 45h/sem)', data.horasNormales]);
  ws.addRow(['Horas extra (sobre 45h/sem)', data.horasExtra]);
  ws.addRow(['Horas nocturnas (19:00-06:00)', data.horasNocturnas]);
  ws.addRow(['Horas permiso', data.horasPermiso]);
  ws.addRow(['Horas netas', horasNetas]);
  ws.addRow([]);

  ws.addRow(['LIQUIDACIÓN']);
  const liqRow = ws.getRow(ws.lastRow.number);
  liqRow.font = { bold: true, color: { argb: 'FFE5E7EB' } };
  liqRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF020617' } };
  ws.mergeCells(`A${liqRow.number}:D${liqRow.number}`);

  const headerLiq = ws.addRow(['Concepto', 'Horas', 'Valor/hora', 'Subtotal']);
  headerLiq.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerLiq.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

  ws.addRow(['Horas normales', data.horasNormales, Math.round(valorHoraNormal).toLocaleString('es-CO'), Math.round(pagoNormal).toLocaleString('es-CO')]);
  ws.addRow(['Horas extra (+25%)', data.horasExtra, Math.round(valorHoraExtra).toLocaleString('es-CO'), Math.round(pagoExtra).toLocaleString('es-CO')]);
  ws.addRow(['Horas nocturnas (+35%)', data.horasNocturnas, Math.round(valorHoraNocturna).toLocaleString('es-CO'), Math.round(pagoNocturno).toLocaleString('es-CO')]);
  ws.addRow([]);

  const totalRow = ws.addRow(['TOTAL A PAGAR', '', '', Math.round(pagoTotal).toLocaleString('es-CO')]);
  totalRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };

  // ===== Hoja 2: Detalle Semanal =====
  if (data.detalleSemanas && data.detalleSemanas.length > 0) {
    const wsDet = workbook.addWorksheet('Detalle Semanal');
    wsDet.columns = [
      { header: 'Semana', key: 'sem', width: 25 },
      { header: 'Total horas', key: 'tot', width: 15 },
      { header: 'Normales', key: 'hn', width: 12 },
      { header: 'Extra', key: 'he', width: 12 },
      { header: 'Nocturnas', key: 'hnoc', width: 12 },
    ];

    const tituloDet = wsDet.addRow(['DETALLE POR SEMANA - PTAR']);
    wsDet.mergeCells('A1:E1');
    tituloDet.alignment = { horizontal: 'center' };
    tituloDet.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    tituloDet.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };

    wsDet.addRow(['Operario:', data.operario]);
    wsDet.addRow(['Límite semanal: 45 horas']);
    wsDet.addRow([]);

    const header = wsDet.getRow(5);
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };

    for (const sem of data.detalleSemanas) {
      wsDet.addRow([
        `${sem.inicio} a ${sem.fin}`,
        sem.horasTotales.toFixed(2),
        sem.horasNormales.toFixed(2),
        sem.horasExtra.toFixed(2),
        sem.horasNocturnas.toFixed(2),
      ]);
    }

    wsDet.addRow([]);
    const totales = wsDet.addRow([
      'TOTALES',
      data.totalHoras.toFixed(2),
      data.horasNormales.toFixed(2),
      data.horasExtra.toFixed(2),
      data.horasNocturnas.toFixed(2),
    ]);
    totales.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    totales.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };
  }
}




function mostrarMensajeExito(texto) {
  const mensaje = document.createElement('div');
  mensaje.textContent = texto;
  mensaje.style.cssText = 'position: fixed; top: 20px; right: 20px; background: #22c55e; color: white; padding: 15px 25px; border-radius: 8px; font-weight: 600; box-shadow: 0 4px 12px rgba(0,0,0,0.3); z-index: 9999;';
  document.body.appendChild(mensaje);
  
  setTimeout(() => mensaje.remove(), 3000);
}


