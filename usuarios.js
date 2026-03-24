const fs = require('fs/promises');
const path = require('path');

const USERS_FILE = path.join(__dirname, 'data', 'users.json');
const VALID_ROLES = new Set(['admin', 'user']);

let usersCache = null;

// Crea errores con codigo HTTP para manejo uniforme en rutas.
function createHttpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

// Normaliza username para evitar duplicados por mayusculas/espacios.
function normalizeUsername(value) {
  return (value || '').toString().trim().toLowerCase();
}

// Valida y normaliza el rol recibido.
function normalizeRole(value) {
  const role = (value || '').toString().trim().toLowerCase();
  return VALID_ROLES.has(role) ? role : null;
}

// Elimina campos sensibles al responder al cliente.
function sanitizeUser(user) {
  return {
    username: user.username,
    role: user.role,
    displayName: user.displayName
  };
}

// Persiste el cache de usuarios en data/users.json.
async function persistUsers() {
  await fs.writeFile(USERS_FILE, JSON.stringify(usersCache, null, 2), 'utf8');
}

// Carga usuarios desde disco una sola vez y cachea en memoria.
async function loadUsers() {
  if (usersCache) {
    return usersCache;
  }

  try {
    const raw = await fs.readFile(USERS_FILE, 'utf8');
    const parsed = JSON.parse(raw);

    if (!Array.isArray(parsed)) {
      throw createHttpError(500, 'Formato de usuarios invalido en data/users.json');
    }

    usersCache = parsed
      .map((item) => ({
        username: normalizeUsername(item.username),
        password: (item.password || '').toString(),
        role: normalizeRole(item.role) || 'user',
        displayName: (item.displayName || '').toString().trim()
      }))
      .filter((item) => item.username && item.password && item.displayName);

    return usersCache;
  } catch (error) {
    if (error.code === 'ENOENT') {
      usersCache = [];
      await persistUsers();
      return usersCache;
    }

    throw error;
  }
}

// Lista usuarios sin contraseñas.
async function listUsers() {
  const users = await loadUsers();
  return users.map(sanitizeUser);
}

// Valida credenciales para login.
async function authenticateUser(usernameInput, passwordInput) {
  const users = await loadUsers();
  const username = normalizeUsername(usernameInput);
  const password = (passwordInput || '').toString();

  const user = users.find((item) => item.username === username && item.password === password);
  return user ? sanitizeUser(user) : null;
}

// Crea usuario nuevo con validaciones basicas y unicidad.
async function createUser(payload) {
  const users = await loadUsers();

  const username = normalizeUsername(payload.username);
  const password = (payload.password || '').toString();
  const displayName = (payload.displayName || '').toString().trim();
  const role = normalizeRole(payload.role) || 'user';

  if (!username || !password || !displayName) {
    throw createHttpError(400, 'username, password y displayName son obligatorios');
  }

  if (users.some((item) => item.username === username)) {
    throw createHttpError(409, 'Ya existe un usuario con ese username');
  }

  const newUser = { username, password, role, displayName };
  users.push(newUser);
  await persistUsers();

  return sanitizeUser(newUser);
}

// Actualiza usuario existente, conservando reglas de integridad.
async function updateUser(currentUsernameInput, payload) {
  const users = await loadUsers();
  const currentUsername = normalizeUsername(currentUsernameInput);
  const userIndex = users.findIndex((item) => item.username === currentUsername);

  if (userIndex === -1) {
    throw createHttpError(404, 'Usuario no encontrado');
  }

  const nextUsername = payload.username !== undefined
    ? normalizeUsername(payload.username)
    : users[userIndex].username;

  const nextPassword = payload.password !== undefined
    ? (payload.password || '').toString()
    : users[userIndex].password;

  const nextDisplayName = payload.displayName !== undefined
    ? (payload.displayName || '').toString().trim()
    : users[userIndex].displayName;

  const nextRole = payload.role !== undefined
    ? normalizeRole(payload.role)
    : users[userIndex].role;

  if (!nextUsername || !nextPassword || !nextDisplayName || !nextRole) {
    throw createHttpError(400, 'Datos invalidos para actualizar usuario');
  }

  const duplicate = users.find((item, index) => index !== userIndex && item.username === nextUsername);
  if (duplicate) {
    throw createHttpError(409, 'El username nuevo ya esta en uso');
  }

  const updated = {
    username: nextUsername,
    password: nextPassword,
    role: nextRole,
    displayName: nextDisplayName
  };

  users[userIndex] = updated;

  const adminCount = users.filter((item) => item.role === 'admin').length;
  if (adminCount === 0) {
    throw createHttpError(400, 'Debe existir al menos un usuario admin');
  }

  await persistUsers();
  return sanitizeUser(updated);
}

// Elimina usuario validando que no se elimine el ultimo admin.
async function deleteUser(usernameInput) {
  const users = await loadUsers();
  const username = normalizeUsername(usernameInput);
  const userIndex = users.findIndex((item) => item.username === username);

  if (userIndex === -1) {
    throw createHttpError(404, 'Usuario no encontrado');
  }

  const candidate = users[userIndex];
  const adminCount = users.filter((item) => item.role === 'admin').length;

  if (candidate.role === 'admin' && adminCount <= 1) {
    throw createHttpError(400, 'No se puede eliminar el ultimo usuario admin');
  }

  users.splice(userIndex, 1);
  await persistUsers();
}

module.exports = {
  listUsers,
  authenticateUser,
  createUser,
  updateUser,
  deleteUser
};
