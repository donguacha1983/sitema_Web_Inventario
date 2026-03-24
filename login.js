let currentUser = null;

const LOGIN_VIEW_FALLBACK_HTML = `
<div id="loginOverlay" class="login-overlay">
  <div class="login-card">
    <h2>Ingreso al Sistema</h2>
    <p class="login-subtitle">Inicie sesión para acceder al panel de inventario.</p>
    <form id="loginForm">
      <div class="form-group">
        <label for="loginUser">Usuario</label>
        <input id="loginUser" type="text" placeholder="Ingrese su usuario" required>
      </div>
      <div class="form-group" style="margin-top: 12px;">
        <label for="loginPassword">Contraseña</label>
        <div class="password-input-row">
          <input id="loginPassword" type="password" placeholder="Ingrese su contraseña" required>
          <button id="loginTogglePassword" type="button" class="btn btn-secondary">Mostrar</button>
        </div>
      </div>
      <div style="margin-top: 16px;">
        <button type="submit" class="btn btn-primary" style="width: 100%;">Ingresar</button>
      </div>
    </form>
    <div id="loginError" class="message error hidden" style="margin-top: 12px;"></div>
  </div>
</div>
`;

// Genera candidatos de URL para encontrar el API local en distintos escenarios.
function getApiBaseCandidates() {
  const candidates = [];

  if (window.location.origin && window.location.origin.startsWith('http')) {
    candidates.push(window.location.origin);
  }

  // Fallbacks comunes cuando el HTML se abre desde file:// o en otro host.
  candidates.push('http://localhost:3000');
  candidates.push('http://127.0.0.1:3000');
  candidates.push('http://localhost:3001');
  candidates.push('http://127.0.0.1:3001');
  candidates.push('http://localhost:3002');
  candidates.push('http://127.0.0.1:3002');

  if (window.location.hostname && window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
    candidates.push(`http://${window.location.hostname}:3000`);
    candidates.push(`http://${window.location.hostname}:3001`);
  }

  return [...new Set(candidates)];
}

// Realiza fetch probando varias bases de API hasta encontrar una valida.
async function fetchWithApiFallback(path, options) {
  const bases = getApiBaseCandidates();
  let lastError;

  for (const base of bases) {
    try {
      const response = await fetch(`${base}${path}`, options);

      // Si este servidor responde pero no conoce /api/login, intentamos el siguiente.
      if (response.status === 404) {
        const preview = await response.clone().text();
        if (preview.includes('Cannot POST /api/login') || preview.includes('Cannot GET /api/login')) {
          continue;
        }
      }

      return response;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('No se pudo conectar con el API');
}

// Monta el login desde archivo externo para mantener index.html más limpio.
async function mountLoginView() {
  const mount = document.getElementById('loginMount');
  if (!mount) return;

  if (document.getElementById('loginOverlay')) {
    return;
  }

  try {
    const response = await fetchWithApiFallback('/login-view.html', { method: 'GET' });
    if (!response.ok) {
      throw new Error('No se pudo cargar login-view.html');
    }

    const html = await response.text();
    mount.innerHTML = html;
  } catch (error) {
    // Fallback local para no bloquear acceso si falla la plantilla externa.
    mount.innerHTML = LOGIN_VIEW_FALLBACK_HTML;
  }
}

// Vincula eventos de login una sola vez luego de montar la vista.
function bindLoginEvents() {
  const loginForm = document.getElementById('loginForm');
  const togglePasswordButton = document.getElementById('loginTogglePassword');

  if (loginForm && !loginForm.dataset.bound) {
    loginForm.addEventListener('submit', handleLogin);
    loginForm.dataset.bound = 'true';
  }

  if (togglePasswordButton && !togglePasswordButton.dataset.bound) {
    togglePasswordButton.addEventListener('click', togglePasswordVisibility);
    togglePasswordButton.dataset.bound = 'true';
  }
}

// Alterna visibilidad de la contraseña en login.
function togglePasswordVisibility() {
  const passwordInput = document.getElementById('loginPassword');
  const toggleButton = document.getElementById('loginTogglePassword');
  if (!passwordInput || !toggleButton) return;

  const isPassword = passwordInput.type === 'password';
  passwordInput.type = isPassword ? 'text' : 'password';
  toggleButton.textContent = isPassword ? 'Ocultar' : 'Mostrar';
}

// Procesa envío de login contra /api/login.
async function handleLogin(event) {
  event.preventDefault();
  const username = document.getElementById('loginUser').value.trim();
  const password = document.getElementById('loginPassword').value;
  const loginError = document.getElementById('loginError');

  try {
    const response = await fetchWithApiFallback('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password })
    });

    const data = await response.json();
    if (!response.ok || data.error || !data.user) {
      loginError.textContent = data.error || 'No se pudo iniciar sesión';
      loginError.classList.remove('hidden');
      return;
    }

    loginError.classList.add('hidden');
    startUserSession(data.user);
  } catch (error) {
    loginError.textContent = 'Error de conexión al validar credenciales. Verifique que el servidor Node esté activo.';
    loginError.classList.remove('hidden');
  }
}

// Restaura sesión guardada en sessionStorage.
function restoreSession() {
  const saved = sessionStorage.getItem('inventarioSessionUser');
  if (!saved) {
    showLoginOverlay();
    return;
  }

  try {
    const parsed = JSON.parse(saved);
    if (!parsed || !parsed.username || !parsed.role || !parsed.displayName) {
      showLoginOverlay();
      return;
    }
    startUserSession(parsed, true);
  } catch (error) {
    showLoginOverlay();
  }
}

// Inicializa estado visual y de sesión del usuario autenticado.
function startUserSession(user, isRestore = false) {
  currentUser = {
    username: user.username,
    role: user.role,
    displayName: user.displayName
  };

  sessionStorage.setItem('inventarioSessionUser', JSON.stringify({
    username: currentUser.username,
    role: currentUser.role,
    displayName: currentUser.displayName
  }));

  document.body.classList.add('logged-in');
  document.getElementById('loginOverlay').classList.add('hidden');
  document.getElementById('currentUserText').textContent = `${currentUser.displayName} (${currentUser.role === 'admin' ? 'Administrador' : 'Usuario'})`;

  applyRolePermissions();

  if (!isRestore) {
    document.getElementById('loginUser').value = '';
    document.getElementById('loginPassword').value = '';
    const toggleButton = document.getElementById('loginTogglePassword');
    if (toggleButton) toggleButton.textContent = 'Mostrar';
    document.getElementById('loginPassword').type = 'password';
  }

  if (typeof initializeAfterLogin === 'function') {
    initializeAfterLogin();
  }
}

// Muestra el overlay de login y limpia estado de sesión actual.
function showLoginOverlay() {
  currentUser = null;
  document.body.classList.remove('logged-in');
  const loginOverlay = document.getElementById('loginOverlay');
  if (loginOverlay) {
    loginOverlay.classList.remove('hidden');
  }
}

// Indica si el usuario actual es administrador.
function isAdminUser() {
  return currentUser && currentUser.role === 'admin';
}

// Aplica permisos visuales según rol de usuario.
function applyRolePermissions() {
  const navConfigItem = document.getElementById('navConfigItem');
  if (!navConfigItem) return;

  if (isAdminUser()) {
    navConfigItem.classList.remove('hidden');
  } else {
    navConfigItem.classList.add('hidden');
    if (typeof currentTab !== 'undefined' && currentTab === 'configuracion') {
      showTab('dashboard');
    }
  }
}

// Cierra sesión local y vuelve a pantalla de login.
function logout() {
  sessionStorage.removeItem('inventarioSessionUser');
  currentUser = null;
  document.getElementById('currentUserText').textContent = 'Sin sesión';
  showLoginOverlay();
}

// Punto de entrada de la app en frontend.
async function initializeApp() {
  await mountLoginView();
  bindLoginEvents();
  restoreSession();
}
