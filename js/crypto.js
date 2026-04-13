/**
 * crypto.js
 * Utilidades criptográficas usando la Web Crypto API nativa del navegador.
 * Sin dependencias externas.
 */

/**
 * Convierte un ArrayBuffer a string hexadecimal.
 */
function bufferToHex(buffer) {
  return Array.from(new Uint8Array(buffer))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
}

/**
 * Calcula SHA-256 de un string y retorna hex string.
 */
async function sha256(text) {
  const encoded = new TextEncoder().encode(text);
  const hashBuf = await crypto.subtle.digest('SHA-256', encoded);
  return bufferToHex(hashBuf);
}

/**
 * Genera un salt aleatorio de 32 bytes en formato hex.
 * Se llama al crear un usuario o cambiar el PIN.
 */
export function generateSalt() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return bufferToHex(bytes.buffer);
}

/**
 * Hashea un PIN con su salt.
 * Formato: SHA-256(salt + pin)
 *
 * @param {string} salt  - Salt hex del usuario (guardado en Firestore)
 * @param {string} pin   - PIN ingresado por el usuario (solo dígitos)
 * @returns {Promise<string>} hash hex
 */
export async function hashPin(salt, pin) {
  return sha256(salt + pin);
}

/**
 * Deriva la contraseña interna de Firebase Auth para un usuario.
 * Produce una contraseña única por usuario a partir de su UID y el SEED.
 * El SEED está en firebase-config.js y nunca se sube a repos públicos.
 *
 * @param {string} uid   - Firebase Auth UID del usuario
 * @param {string} seed  - SEED del proyecto (de firebase-config.js)
 * @returns {Promise<string>} contraseña interna hex (64 chars)
 */
export async function derivePassword(uid, seed) {
  return sha256(uid + seed);
}
