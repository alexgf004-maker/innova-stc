# INNOVA STC — Guía de configuración

## Fase 0 completada: Base y autenticación

---

## 1. Configurar Firebase

1. Ve a [console.firebase.google.com](https://console.firebase.google.com)
2. Crea un proyecto nuevo llamado `innova-stc`
3. Activa estos servicios:
   - **Authentication** → Sign-in method → **Email/Password** → Habilitar
   - **Firestore Database** → Crear en modo **producción**
4. En Configuración del proyecto → Tus apps → Agrega app Web
5. Copia el objeto `firebaseConfig`

---

## 2. Configurar el proyecto

Edita `js/firebase-config.js` y reemplaza:
- Los valores de `FIREBASE_CONFIG` con los de tu proyecto
- El valor de `SEED` con una cadena aleatoria propia (mínimo 16 caracteres)

**⚠️ No cambies el SEED después de crear usuarios. Si lo cambias, nadie podrá entrar.**

---

## 3. Aplicar Security Rules

1. Ve a Firebase Console → Firestore → Reglas
2. Copia el contenido de `firestore.rules`
3. Pega y publica las reglas

---

## 4. Crear el primer usuario administrador

Como no hay registro libre, el primer admin se crea manualmente.

### Opción A: Desde Firebase Console

1. **Authentication → Agregar usuario:**
   - Email: `admin@innova-stc.internal`
   - Password: (cualquiera, se sobrescribirá)
   - Copia el **UID** generado

2. **Generar el hash del PIN inicial:**
   Abre `login.html` en el navegador, abre la consola del navegador (F12) y ejecuta:
   ```javascript
   // Ejemplo para PIN "1234" con salt aleatorio
   const salt = Array.from(crypto.getRandomValues(new Uint8Array(32)))
     .map(b => b.toString(16).padStart(2,'0')).join('');

   const enc = new TextEncoder().encode(salt + '1234');
   const buf = await crypto.subtle.digest('SHA-256', enc);
   const hash = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');

   console.log('Salt:', salt);
   console.log('Hash:', hash);
   ```
   Copia el `salt` y el `hash` resultantes.

3. **Firestore → Crear colección `users` → Nuevo documento:**
   - ID del documento: el UID del paso 1
   - Campos:
     ```
     uid:           <UID copiado>
     username:      "admin"
     displayName:   "Administrador"
     role:          "admin"
     internalEmail: "admin@innova-stc.internal"
     pinHash:       <hash del paso 2>
     pinSalt:       <salt del paso 2>
     active:        true (boolean)
     createdAt:     (timestamp — click en el icono de calendario)
     createdBy:     <mismo UID>
     ```

4. **Actualizar la contraseña en Firebase Auth:**
   Con el UID del admin, ve a Authentication → ese usuario → editar → cambia la contraseña.
   La contraseña debe ser `SHA-256(uid + SEED)` de tu `firebase-config.js`.
   
   Calcula así en consola:
   ```javascript
   const uid  = 'EL_UID_DEL_ADMIN';
   const seed = 'TU_SEED_DE_FIREBASE_CONFIG';
   const enc  = new TextEncoder().encode(uid + seed);
   const buf  = await crypto.subtle.digest('SHA-256', enc);
   const pass = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
   console.log('Password:', pass);
   ```

---

## 5. Subir a GitHub Pages

1. Crea un repositorio en GitHub (puede ser privado)
2. Sube todos los archivos del proyecto
3. Ve a Settings → Pages → Source: main branch → / (root)
4. La app estará disponible en `https://tu-usuario.github.io/innova-stc/login.html`

---

## Estructura del proyecto

```
innova-stc/
├── index.html              ← Shell principal de la app
├── login.html              ← Pantalla de login
├── css/
│   ├── base.css            ← Variables y reset
│   ├── layout.css          ← Sidebar, topbar, bottom nav
│   ├── components.css      ← Badges, cards, botones, tablas
│   └── pages/
│       └── login.css       ← Estilos del login
├── js/
│   ├── firebase-config.js  ← ⚠️ Configurar antes de usar
│   ├── crypto.js           ← SHA-256 nativo (sin dependencias)
│   ├── router.js           ← Navegación SPA
│   ├── ui.js               ← Toast, modal, badges, helpers
│   └── modules/
│       ├── admin.js        ← Gestión de usuarios
│       ├── kardex.js       ← Fase 1 (stub preparado)
│       ├── cambios.js      ← Fase 2 (pendiente)
│       ├── otc.js          ← Fase 3 (pendiente)
│       └── dashboard.js    ← Fase 4 (pendiente)
├── views/
│   ├── home.html           ← Vista de inicio
│   ├── admin-usuarios.html ← Panel de usuarios
│   ├── kardex.html         ← Fase 1 (preparado)
│   └── perfil.html         ← Cambio de PIN
├── assets/                 ← Logos e imágenes
└── firestore.rules         ← Copiar a Firebase Console
```

---

## Roles disponibles

| Rol | Acceso |
|-----|--------|
| `admin` | Todo: usuarios, kardex, módulos futuros |
| `coordinadora` | Operativo: kardex, módulos operativos. Sin gestión de usuarios |
| `campo` | Limitado: solo lo que le corresponde por módulo |

---

## Para agregar un módulo nuevo (Fase 2+)

1. Crear `js/modules/cambios.js` con `export async function initCambios(session) {...}`
2. Crear `views/cambios.html` con el HTML del módulo
3. En `js/router.js`, agregar en `ROUTES` y `NAV_ITEMS`
4. En `firestore.rules`, agregar las reglas de la colección del módulo

Eso es todo. El shell, la navegación y la autenticación no se tocan.
