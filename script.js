// script.js - Control de registro con soporte PTAP/PTAR
const API_URL = 'https://backend-control-operarios.onrender.com';

// Operarios por planta
const OPERARIOS = {
  PTAP: ['Hernan Aragon', 'Rossa Viafara', 'Jaiver Casaran', 'Arnoldo Camacho'],
  PTAR: ['Bladimir Cifuentes', 'Willington Granja', 'Rigoberto Arrechea']
};

// Elementos DOM
const plantaSelect = document.getElementById('planta');
const operarioSelect = document.getElementById('nombreOperario');
const tipoSelect = document.getElementById('tipo');
const labelJustificacion = document.getElementById('labelJustificacion');
const justificacionTextarea = document.getElementById('justificacionExtra');
const horaActualInput = document.getElementById('horaActual');
const estadoGpsDiv = document.getElementById('estadoGps');
const latInput = document.getElementById('lat');
const lngInput = document.getElementById('lng');
const formRegistro = document.getElementById('formRegistro');
const mensajeDiv = document.getElementById('mensajeRespuesta');

let configPlanta = null;

// Actualizar hora
function actualizarHora() {
  horaActualInput.value = new Date().toLocaleTimeString('es-CO');
}
setInterval(actualizarHora, 1000);
actualizarHora();

// Cambio de planta
plantaSelect.addEventListener('change', async () => {
  const planta = plantaSelect.value;
  
  if (!planta) {
    operarioSelect.innerHTML = '<option value="">Primero selecciona una planta...</option>';
    operarioSelect.disabled = true;
    return;
  }

  // Cargar operarios
  operarioSelect.innerHTML = '<option value="">Seleccione...</option>';
  OPERARIOS[planta].forEach(op => {
    operarioSelect.innerHTML += `<option value="${op}">${op}</option>`;
  });
  operarioSelect.disabled = false;

  // Cargar configuración de geocerca
  try {
    const resp = await fetch(`${API_URL}/api/config-planta/${planta}`);
    configPlanta = await resp.json();
    obtenerUbicacion();
  } catch (err) {
    console.error('Error al cargar config de planta:', err);
  }
});

// Mostrar justificación si es PTAP y salida después de 17:00
tipoSelect.addEventListener('change', () => {
  const planta = plantaSelect.value;
  const tipo = tipoSelect.value;
  const hora = new Date().getHours();

  if (planta === 'PTAP' && tipo === 'salida' && hora >= 17) {
    labelJustificacion.classList.remove('oculto');
    justificacionTextarea.classList.remove('oculto');
  } else {
    labelJustificacion.classList.add('oculto');
    justificacionTextarea.classList.add('oculto');
    justificacionTextarea.value = '';
  }
});

// Calcular distancia
function distanciaMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = g => g * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat/2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng/2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

// Obtener GPS
function obtenerUbicacion() {
  if (!navigator.geolocation || !configPlanta) {
    estadoGpsDiv.textContent = 'GPS no disponible';
    estadoGpsDiv.className = 'info error';
    return;
  }

  navigator.geolocation.getCurrentPosition(
    pos => {
      latInput.value = pos.coords.latitude;
      lngInput.value = pos.coords.longitude;

      const dist = distanciaMetros(pos.coords.latitude, pos.coords.longitude, configPlanta.lat, configPlanta.lng);

      if (dist <= configPlanta.radio) {
        estadoGpsDiv.textContent = `✓ Dentro de la zona (${plantaSelect.value})`;
        estadoGpsDiv.className = 'info ok';
      } else {
        estadoGpsDiv.textContent = `⚠️ Fuera de zona: ${dist.toFixed(0)}m (máx: ${configPlanta.radio}m)`;
        estadoGpsDiv.className = 'info error';
      }
    },
    err => {
      estadoGpsDiv.textContent = 'Error al obtener ubicación';
      estadoGpsDiv.className = 'info error';
    },
    { enableHighAccuracy: true, timeout: 15000 }
  );
}

// Enviar registro
formRegistro.addEventListener('submit', async e => {
  e.preventDefault();

  const datos = {
    planta: plantaSelect.value,
    nombreOperario: operarioSelect.value,
    tipo: tipoSelect.value,
    lat: parseFloat(latInput.value),
    lng: parseFloat(lngInput.value),
    justificacionExtra: justificacionTextarea.value.trim()
  };

  if (!datos.planta || !datos.nombreOperario || !datos.tipo || !datos.lat || !datos.lng) {
    mensajeDiv.textContent = '⚠️ Completa todos los campos';
    mensajeDiv.className = 'info error';
    mensajeDiv.style.display = 'block';
    return;
  }

  mensajeDiv.textContent = '⏳ Registrando...';
  mensajeDiv.className = 'info';
  mensajeDiv.style.display = 'block';

  try {
    const resp = await fetch(`${API_URL}/api/registro`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(datos)
    });

    const json = await resp.json();

    if (!resp.ok) {
      mensajeDiv.textContent = '❌ ' + json.mensaje;
      mensajeDiv.className = 'info error';
      return;
    }

    mensajeDiv.textContent = `✅ ¡REGISTRO EXITOSO!${json.turno ? ` (Turno: ${json.turno})` : ''}`;
    mensajeDiv.className = 'info ok';
    mensajeDiv.style.fontSize = '18px';
    mensajeDiv.style.fontWeight = 'bold';

    setTimeout(() => {
      tipoSelect.value = '';
      mensajeDiv.style.display = 'none';
      mensajeDiv.style.fontSize = '14px';
    }, 3000);

  } catch (err) {
    mensajeDiv.textContent = '❌ Error de conexión';
    mensajeDiv.className = 'info error';
  }
});





