// URL de tu backend en Render
const API_URL = 'https://backend-control-operarios.onrender.com';

// ⚠️ IMPORTANTE: Cambia estas coordenadas por las de tu planta EMPUVILLA
const PLANTA_LAT = 3.17253; // Latitud de tu planta
const PLANTA_LNG = -76.4588; // Longitud de tu planta
const RADIO_PERMITIDO_METROS = 100; // Radio en metros (debe coincidir con tu backend)

const horaActualInput = document.getElementById('horaActual');
const estadoGpsDiv = document.getElementById('estadoGps');
const latInput = document.getElementById('lat');
const lngInput = document.getElementById('lng');
const formRegistro = document.getElementById('formRegistro');
const mensajeRespuestaDiv = document.getElementById('mensajeRespuesta');

// 1. Actualizar hora en el formulario cada segundo
function actualizarHora() {
  const ahora = new Date();
  horaActualInput.value = ahora.toLocaleTimeString('es-CO', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

setInterval(actualizarHora, 1000);
actualizarHora();

// 2. Función para calcular distancia entre dos puntos (en metros)
function distanciaMetros(lat1, lng1, lat2, lng2) {
  const R = 6371000; // Radio de la Tierra en metros
  const toRad = (grado) => (grado * Math.PI) / 180;

  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);

  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLng / 2) *
      Math.sin(dLng / 2);

  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c; // Retorna distancia en metros
}

// 3. Pedir ubicación por GPS
function obtenerUbicacion() {
  if (!navigator.geolocation) {
    estadoGpsDiv.textContent = 'Tu dispositivo no soporta GPS en el navegador.';
    estadoGpsDiv.className = 'info error';
    return;
  }

  navigator.geolocation.getCurrentPosition(
    (posicion) => {
      const { latitude, longitude } = posicion.coords;
      latInput.value = latitude;
      lngInput.value = longitude;

      // Calcular distancia a la planta
      const distancia = distanciaMetros(latitude, longitude, PLANTA_LAT, PLANTA_LNG);
      
      if (distancia <= RADIO_PERMITIDO_METROS) {
        estadoGpsDiv.textContent = `Ubicación obtenida ✓ - Dentro de la zona de trabajo`;
        estadoGpsDiv.className = 'info ok';
      } else {
        estadoGpsDiv.textContent = `⚠️ FUERA DE LA ZONA: Estás a ${distancia.toFixed(0)} metros de la planta`;
        estadoGpsDiv.className = 'info error';
      }
    },
    (error) => {
      console.error(error);
      estadoGpsDiv.textContent = 'No se pudo obtener la ubicación. Activa el GPS y permite el acceso.';
      estadoGpsDiv.className = 'info error';
    },
    {
      enableHighAccuracy: true,
      timeout: 15000,
      maximumAge: 0,
    }
  );
}

obtenerUbicacion();

// 4. Enviar datos al backend al enviar el formulario
formRegistro.addEventListener('submit', async (e) => {
  e.preventDefault();

  const nombreOperario = document.getElementById('nombreOperario').value;
  const tipo = document.getElementById('tipo').value;
  const lat = latInput.value ? parseFloat(latInput.value) : null;
  const lng = lngInput.value ? parseFloat(lngInput.value) : null;

  // Mostrar el div de mensaje (por si está oculto)
  mensajeRespuestaDiv.style.display = 'block';

  // Validar campos básicos
  if (!nombreOperario || !tipo) {
    mensajeRespuestaDiv.textContent = '⚠️ Debes seleccionar el operario y el tipo de registro.';
    mensajeRespuestaDiv.className = 'info error';
    return;
  }

  // Validar que tenga ubicación GPS
  if (!lat || !lng) {
    mensajeRespuestaDiv.textContent = '⚠️ No se pudo obtener tu ubicación GPS. Activa el GPS e intenta nuevamente.';
    mensajeRespuestaDiv.className = 'info error';
    return;
  }

  // VALIDACIÓN FRONTEND: Verificar que esté dentro de la zona
  const distancia = distanciaMetros(lat, lng, PLANTA_LAT, PLANTA_LNG);
  
  if (distancia > RADIO_PERMITIDO_METROS) {
    mensajeRespuestaDiv.textContent = `🚫 REGISTRO DENEGADO: Debes estar en la zona de trabajo para registrar tu ${tipo}. Estás a ${distancia.toFixed(0)} metros de la planta (máximo permitido: ${RADIO_PERMITIDO_METROS}m).`;
    mensajeRespuestaDiv.className = 'info error';
    mensajeRespuestaDiv.style.fontSize = '16px';
    mensajeRespuestaDiv.style.fontWeight = 'bold';
    mensajeRespuestaDiv.style.padding = '20px';
    return;
  }

  const datos = { nombreOperario, tipo, lat, lng };

  // Mostrar mensaje de carga
  mensajeRespuestaDiv.textContent = '⏳ Registrando...';
  mensajeRespuestaDiv.className = 'info';
  mensajeRespuestaDiv.style.fontSize = '14px';

  try {
    const respuesta = await fetch(`${API_URL}/api/registro`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(datos),
    });

    const json = await respuesta.json();

    if (!respuesta.ok) {
      mensajeRespuestaDiv.textContent = '❌ ' + (json.mensaje || 'Error al registrar. Intenta de nuevo.');
      mensajeRespuestaDiv.className = 'info error';
      mensajeRespuestaDiv.style.fontSize = '16px';
      mensajeRespuestaDiv.style.fontWeight = 'bold';
      return;
    }

    // ✅ ÉXITO - MENSAJE GRANDE Y VISIBLE
    const emoji = tipo === 'ingreso' ? '🟢' : '🔴';
    mensajeRespuestaDiv.style.display = 'block';
    mensajeRespuestaDiv.textContent = `${emoji} ¡REGISTRO GUARDADO EXITOSAMENTE!`;
    mensajeRespuestaDiv.className = 'info ok';
    mensajeRespuestaDiv.style.fontSize = '20px';
    mensajeRespuestaDiv.style.fontWeight = 'bold';
    mensajeRespuestaDiv.style.textAlign = 'center';
    mensajeRespuestaDiv.style.padding = '25px';

    // Limpiar el tipo después de 3 segundos
    setTimeout(() => {
      document.getElementById('tipo').value = '';
      mensajeRespuestaDiv.style.display = 'none';
      mensajeRespuestaDiv.style.fontSize = '14px';
      mensajeRespuestaDiv.style.padding = '12px';
    }, 3000);

  } catch (err) {
    console.error(err);
    mensajeRespuestaDiv.textContent = '❌ No se pudo conectar con el servidor. Revisa tu conexión a internet.';
    mensajeRespuestaDiv.className = 'info error';
    mensajeRespuestaDiv.style.fontSize = '16px';
    mensajeRespuestaDiv.style.fontWeight = 'bold';
  }
});



