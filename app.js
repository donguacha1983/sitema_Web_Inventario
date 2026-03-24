const express = require('express');
const path = require('path');
const {
  registrarPedido,
  actualizarPedido,
  eliminarPedido,
  registrarCustodia,
  registrarRuta,
  listarCustodios,
  crearCustodio,
  actualizarCustodio,
  eliminarCustodio,
  listarVehiculos,
  crearVehiculo,
  actualizarVehiculo,
  eliminarVehiculo,
  obtenerDatosCustodia,
  obtenerDatosRutas,
  buscarPedido,
  obtenerListas,
  obtenerPedidos,
  inicializarExcel
} = require('./inventario');
const {
  listUsers,
  authenticateUser,
  createUser,
  updateUser,
  deleteUser
} = require('./usuarios');

// Inicializa servidor HTTP y puerto configurable por variable de entorno.
const app = express();
const PORT = process.env.PORT || 3000;

// Habilita parseo de JSON y formularios URL encoded.
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Permite llamadas al API desde el navegador incluso si el HTML se abre desde file://
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type,X-User-Role,X-User-Name');
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});

app.use(express.static(path.join(__dirname)));

// Valida por cabecera que la operación fue solicitada por un administrador.
function requireAdminRole(req, res) {
  const role = (req.header('X-User-Role') || '').toString().trim().toLowerCase();
  if (role !== 'admin') {
    res.status(403).json({ error: 'Solo un administrador puede realizar esta acción' });
    return false;
  }

  return true;
}

// --- AUTENTICACION ---
// Valida usuario/clave y devuelve solo datos seguros para sesión.
app.post('/api/login', async (req, res) => {
  try {
    const user = await authenticateUser(req.body.username, req.body.password);

    if (!user) {
      return res.status(401).json({ error: 'Usuario o contraseña incorrectos' });
    }

    return res.json({ user });
  } catch (error) {
    return res.status(error.status || 500).json({ error: error.message || 'No se pudo validar credenciales' });
  }
});

// --- GESTION DE USUARIOS ---
// Lista usuarios sin exponer contraseñas.
app.get('/api/usuarios', async (req, res) => {
  try {
    const users = await listUsers();
    res.json(users);
  } catch (error) {
    res.status(error.status || 500).json({ error: error.message || 'No se pudo listar usuarios' });
  }
});

// Crea un usuario nuevo.
app.post('/api/usuarios', async (req, res) => {
  try {
    const created = await createUser(req.body || {});
    res.status(201).json({ usuario: created });
  } catch (error) {
    res.status(error.status || 500).json({ error: error.message || 'No se pudo crear usuario' });
  }
});

// Actualiza usuario por username.
app.put('/api/usuarios/:username', async (req, res) => {
  try {
    const updated = await updateUser(req.params.username, req.body || {});
    res.json({ usuario: updated });
  } catch (error) {
    res.status(error.status || 500).json({ error: error.message || 'No se pudo actualizar usuario' });
  }
});

// Elimina usuario por username.
app.delete('/api/usuarios/:username', async (req, res) => {
  if (!requireAdminRole(req, res)) {
    return;
  }

  try {
    await deleteUser(req.params.username);
    res.json({ mensaje: 'Usuario eliminado' });
  } catch (error) {
    res.status(error.status || 500).json({ error: error.message || 'No se pudo eliminar usuario' });
  }
});

// --- CATALOGOS Y PEDIDOS ---
// Lista custodios para el panel de configuracion.
app.get('/api/custodios', async (req, res) => {
  try {
    const custodios = await listarCustodios();
    res.json(custodios);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo listar custodios: ' + error.message });
  }
});

app.get('/api/vehiculos', async (req, res) => {
  try {
    const vehiculos = await listarVehiculos();
    res.json(vehiculos);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo listar vehículos: ' + error.message });
  }
});

app.post('/api/vehiculos', async (req, res) => {
  try {
    const mensaje = await crearVehiculo(req.body || {});
    const status = mensaje.includes('correctamente') ? 201 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo crear vehículo: ' + error.message });
  }
});

app.put('/api/vehiculos/:id', async (req, res) => {
  try {
    const mensaje = await actualizarVehiculo(req.params.id || '', req.body || {});
    const status = mensaje.includes('correctamente') ? 200 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo actualizar vehículo: ' + error.message });
  }
});

app.delete('/api/vehiculos/:id', async (req, res) => {
  if (!requireAdminRole(req, res)) {
    return;
  }

  try {
    const mensaje = await eliminarVehiculo(req.params.id || '');
    const status = mensaje.includes('correctamente') ? 200 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo eliminar vehículo: ' + error.message });
  }
});

// Crea un custodio nuevo.
app.post('/api/custodios', async (req, res) => {
  try {
    const mensaje = await crearCustodio(req.body || {});
    const status = mensaje.includes('correctamente') ? 201 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo crear custodio: ' + error.message });
  }
});

// Actualiza un custodio por CI.
app.put('/api/custodios/:ci', async (req, res) => {
  try {
    const mensaje = await actualizarCustodio(req.params.ci || '', req.body || {});
    const status = mensaje.includes('correctamente') ? 200 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo actualizar custodio: ' + error.message });
  }
});

// Elimina custodio por CI (solo admin).
app.delete('/api/custodios/:ci', async (req, res) => {
  if (!requireAdminRole(req, res)) {
    return;
  }

  try {
    const mensaje = await eliminarCustodio(req.params.ci || '');
    const status = mensaje.includes('correctamente') ? 200 : 400;
    res.status(status).json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo eliminar custodio: ' + error.message });
  }
});

// Devuelve listas de apoyo para selects (unidades y grupos).
app.get('/api/listas', async (req, res) => {
  try {
    const listas = await obtenerListas();
    res.json(listas);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo obtener listas: ' + error.message });
  }
});

// Devuelve listas para formulario de custodia desde hojas CUSTODIOS y vehiculos.
app.get('/api/custodias/listas', async (req, res) => {
  try {
    const datos = await obtenerDatosCustodia();
    res.json(datos);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo obtener listas de custodia: ' + error.message });
  }
});

async function crearPedidoHandler(req, res) {
  try {
    const pedido = req.body;
    if (!pedido.codigo || !pedido.nombre) {
      return res.status(400).json({ error: 'Código y nombre son obligatorios' });
    }

    const resultado = await registrarPedido(pedido);

    if (typeof resultado === 'string') {
      return res.json({ mensaje: resultado });
    }

    return res.json(resultado);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo registrar pedido: ' + error.message });
  }
}

async function actualizarPedidoHandler(req, res) {
  try {
    const codigo = req.params.codigo || '';
    const pedido = req.body || {};

    if (!codigo) {
      return res.status(400).json({ error: 'Identificador de pedido inválido' });
    }

    const mensaje = await actualizarPedido(codigo, pedido);
    const status = mensaje.includes('correctamente') ? 200 : 400;
    return res.status(status).json({ mensaje });
  } catch (error) {
    return res.status(500).json({ error: 'No se pudo actualizar pedido: ' + error.message });
  }
}

async function eliminarPedidoHandler(req, res) {
  if (!requireAdminRole(req, res)) {
    return;
  }

  try {
    const codigo = req.params.codigo || '';

    if (!codigo) {
      return res.status(400).json({ error: 'Identificador de pedido inválido' });
    }

    const mensaje = await eliminarPedido(codigo);
    const status = mensaje.includes('correctamente') ? 200 : 400;
    return res.status(status).json({ mensaje });
  } catch (error) {
    return res.status(500).json({ error: 'No se pudo eliminar pedido: ' + error.message });
  }
}

// Registra, actualiza y elimina pedidos.
app.post('/api/pedidos', crearPedidoHandler);
app.put('/api/pedidos/:codigo', actualizarPedidoHandler);
app.delete('/api/pedidos/:codigo', eliminarPedidoHandler);

// Alias legacy para compatibilidad con clientes anteriores.
app.post('/api/productos', crearPedidoHandler);
app.put('/api/productos/:codigo', actualizarPedidoHandler);
app.delete('/api/productos/:codigo', eliminarPedidoHandler);

// Busca pedidos por texto libre.
app.get('/api/productos', async (req, res) => {
  try {
    const texto = req.query.q || '';
    const resultados = await buscarPedido(texto);
    res.json(resultados);
  } catch (error) {
    res.status(500).json({ error: 'No se pudo buscar pedido: ' + error.message });
  }
});

// --- CUSTODIA Y REPORTES ---
// Registra los datos de custodia asociados a un pedido.
app.post('/api/custodias', async (req, res) => {
  try {
    const custodia = req.body;
    const pedidoId = (custodia.pedidoId || custodia.codigo || '').toString().trim();

    if (!pedidoId || !custodia.fecha || !custodia.tipo || !custodia.cantidad) {
      return res.status(400).json({ error: 'Datos de custodia incompletos' });
    }

    const mensaje = await registrarCustodia(custodia);
    res.json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'No se pudo registrar custodia: ' + error.message });
  }
});

// Registra una ruta asociada a un pedido.
app.post('/api/rutas', async (req, res) => {
  try {
    const ruta = req.body || {};
    const pedidoId = (ruta.pedidoId || ruta.codigo || '').toString().trim();

    if (!pedidoId || !ruta.ruta) {
      return res.status(400).json({ error: 'ID Pedido y nombre de ruta son obligatorios' });
    }

    const mensaje = await registrarRuta(ruta);
    const status = mensaje.includes('correctamente') ? 200 : 400;
    return res.status(status).json({ mensaje });
  } catch (error) {
    return res.status(500).json({ error: 'No se pudo registrar ruta: ' + error.message });
  }
});

// Devuelve listas para formulario de rutas (ID Pedido).
app.get('/api/rutas/listas', async (req, res) => {
  try {
    const datos = await obtenerDatosRutas();
    return res.json(datos);
  } catch (error) {
    return res.status(500).json({ error: 'No se pudo obtener listas de rutas: ' + error.message });
  }
});

// Obtiene el listado consolidado de pedidos.
app.get('/api/pedidos', async (req, res) => {
  try {
    const texto = (req.query.q || '').toString().trim();
    if (texto) {
      const resultados = await buscarPedido(texto);
      return res.json(resultados);
    }

    const pedidos = await obtenerPedidos();
    return res.json(pedidos);
  } catch (error) {
    return res.status(500).json({ error: 'No se pudo obtener pedidos: ' + error.message });
  }
});

// Inicializa o repara las hojas y columnas del Excel (accion del boton Configuracion BD).
app.post('/api/configurar-bd', async (req, res) => {
  try {
    const mensaje = await inicializarExcel();
    res.json({ mensaje });
  } catch (error) {
    res.status(500).json({ error: 'Error al configurar la BD: ' + error.message });
  }
});

// Arranca el servidor y muestra la URL local.
app.listen(PORT, () => {
  console.log(`Servidor iniciado en http://localhost:${PORT}`);
});
