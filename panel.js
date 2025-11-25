// Misma URL del backend
const API_URL = 'https://backend-control-operarios.onrender.com';

// Credenciales simples (para uso interno)
// OJO: esto NO es seguridad fuerte, solo sirve para uso interno.
const USUARIO_PERMITIDO = 'admin';
const PASSWORD_PERMITIDO = 'empuvilla2025';

// Elementos del DOM
const seccionLogin = document.getElementById('seccionLogin');
const seccionReporte = document.getElementById('seccionReporte');
const mensajeLogin = document.getElementById('mensajeLogin');

const btnLogin = document.getElementById('btnLogin');
const btnConsultar = document.getElementById('btnConsultar');

const operarioReporte = document.getElementById('operarioReporte');
const fechaDesde = document.getElementById('fechaDesde');
const fechaHasta = document.getElementById('fechaHasta');

const resumenDiv = document.getElementById('resumen');
const tablaDetalleDiv = document.getElementById('tablaDetalle');

// 1. LOGIN básico
btnLogin.addEventListener('click', () => {
  const usuario = document.getElementById('usuario').value.trim();
  const password = document.getElementById('password').value.trim();

  if (
    usuario === USUARIO_PERMITIDO &&
    password === PASSWORD_PERMITIDO
  ) {
    mensajeLogin.textContent = 'Ingreso exitoso.';
    mensajeLogin.style.color = '#22c55e';

    // Mostrar sección de reporte
    seccionLogin.classList.add('oculto');
    seccionReporte.classList.remove('oculto');
  } else {
    mensajeLogin.textContent = 'Usuario o contraseña incorrectos.';
    mensajeLogin.style.color = '#f97316';
  }
});

// 2. Consultar reporte
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

    // Mostrar resumen
    resumenDiv.style.color = '#e5e7eb';
    resumenDiv.innerHTML = `
      <strong>Operario:</strong> ${data.operario}<br>
      <strong>Rango:</strong> ${data.desde} a ${data.hasta}<br>
      <strong>Total horas:</strong> ${data.totalHoras} h<br>
      <strong>Horas normales:</strong> ${data.horasNormales} h<br>
      <strong>Horas extra:</strong> ${data.horasExtra} h<br>
      <strong>Horas dominicales:</strong> ${data.horasDominicales} h
    `;

    // Mostrar detalle en tabla
    if (Array.isArray(data.detalle) && data.detalle.length > 0) {
      let html = `
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

      html += '</tbody></table>';
      tablaDetalleDiv.innerHTML = html;
    } else {
      tablaDetalleDiv.innerHTML = '<p>No hay detalle para este rango.</p>';
    }
  } catch (err) {
    console.error(err);
    resumenDiv.textContent =
      'No se pudo conectar con el servidor. Revisa tu conexión.';
    resumenDiv.style.color = '#f97316';
  }
});
