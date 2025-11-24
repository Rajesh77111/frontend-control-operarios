// URL de tu backend en Render (CÁMBIALA por la real)
const API_URL = 'https://backend-control-operarios.onrender.com';

const horaActualInput = document.getElementById('horaActual');
const estadoGpsDiv = document.getElementById('estadoGps');
const latInput = document.getElementById('lat');
const lngInput = document.getElementById('lng');
const formRegistro = document.getElementById('formRegistro');
const mensajeRespuestaDiv = document.getElementById('mensajeRespuesta');

// 1. Actualizar hora en el formulario cada segundo
function actualizarHora() {
  const ahora = new Date();
  // Muestra la hora en formato HH:MM:SS
  horaActualInput.value = ahora.toLocaleTimeString('es-CO', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
}

setInterval(actualizarHora, 1000);
actualizarHora(); // Llamada inicial

// 2. Pedir ubicación por GPS
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

      estadoGpsDiv.textContent = `Ubicación obtenida ✓ (lat: ${latitude.toFixed(
        5
      )}, lng: ${longitude.toFixed(5)})`;
      estadoGpsDiv.className = 'info ok';
    },
    (error) => {
      console.error(error);
      estadoGpsDiv.textContent =
        'No se pudo obtener la ubicación. Activa el GPS y permite el acceso.';
      estadoGpsDiv.className = 'info error';
    },
    {
      enableHighAccuracy: true,
      timeout: 15000,
      maximumAge: 0,
    }
  );
}

// Llamamos a la función al cargar la página
obtenerUbicacion();

// 3. Enviar datos al backend al enviar el formulario
formRegistro.addEventListener('submit', async (e) => {
  e.preventDefault(); // Evita que recargue la página

  const nombreOperario = document.getElementById('nombreOperario').value;
  const tipo = document.getElementById('tipo').value;
  const lat = latInput.value ? parseFloat(latInput.value) : null;
  const lng = lngInput.value ? parseFloat(lngInput.value) : null;

  if (!nombreOperario || !tipo) {
    mensajeRespuestaDiv.textContent =
      'Debes seleccionar el operario y el tipo de registro.';
    mensajeRespuestaDiv.className = 'info error';
    return;
  }

  const datos = { nombreOperario, tipo, lat, lng };

  try {
    const respuesta = await fetch(`${API_URL}/api/registro`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(datos),
    });

    const json = await respuesta.json();

    if (!respuesta.ok) {
      mensajeRespuestaDiv.textContent =
        json.mensaje || 'Error al registrar. Intenta de nuevo.';
      mensajeRespuestaDiv.className = 'info error';
      return;
    }

    // Éxito
   mensajeRespuestaDiv.textContent = '✅ REGISTRO GUARDADO EXITOSAMENTE';
mensajeRespuestaDiv.className = 'info ok';
mensajeRespuestaDiv.style.fontSize = '18px';
mensajeRespuestaDiv.style.fontWeight = 'bold';
mensajeRespuestaDiv.style.textAlign = 'center';
mensajeRespuestaDiv.style.padding = '20px';

// Opcional: limpiar el formulario después de 3 segundos
setTimeout(() => {
  document.getElementById('tipo').value = '';
  mensajeRespuestaDiv.style.display = 'none';
}, 3000);

    // Opcional: limpiar solo el tipo
    document.getElementById('tipo').value = '';
  } catch (err) {
    console.error(err);
    mensajeRespuestaDiv.textContent =
      'No se pudo conectar con el servidor. Revisa tu conexión a internet.';
    mensajeRespuestaDiv.className = 'info error';
  }
});

