import { BACKUP_FORMAT, ENCRYPTED_BACKUP_FORMAT, SIGNED_BACKUP_FORMAT } from '../types/public.js';
import { canonicalJsonStringify } from '../utils/canonical_json.js';

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function toB64(bytes) {
  if (typeof btoa === 'function') {
    let bin = '';
    for (let i = 0; i < bytes.length; i += 1) bin += String.fromCharCode(bytes[i]);
    return btoa(bin);
  }
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
  throw new Error('No base64 encoder available');
}
function fromB64(text) {
  if (typeof atob === 'function') {
    const bin = atob(text);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
    return out;
  }
  if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(text, 'base64'));
  throw new Error('No base64 decoder available');
}

export async function createIntegrity(bundle) {
  const withoutIntegrity = { ...bundle };
  delete withoutIntegrity.integrity;
  const bytes = encoder.encode(canonicalJsonStringify(withoutIntegrity));
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  return { hash: 'SHA-256', payloadHashB64: toB64(digest) };
}

export async function encryptBackup(bundle, password, { iterations = 310000 } = {}) {
  if (!password || password.length < 8) throw new Error('Password must be at least 8 chars');
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const baseKey = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', hash: 'SHA-256', salt, iterations },
    baseKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );

  const aad = { app: bundle.app, scope: bundle.scope };
  const payload = encoder.encode(canonicalJsonStringify(bundle));
  const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv, additionalData: encoder.encode(JSON.stringify(aad)), tagLength: 128 }, key, payload));

  return {
    format: ENCRYPTED_BACKUP_FORMAT,
    formatVersion: 1,
    createdAt: new Date().toISOString(),
    kdf: { name: 'PBKDF2', hash: 'SHA-256', saltB64: toB64(salt), iterations },
    cipher: { name: 'AES-GCM', ivB64: toB64(iv), tagBits: 128 },
    payloadB64: toB64(cipher),
    aad
  };
}

export async function decryptBackup(envelope, password) {
  if (envelope.format !== ENCRYPTED_BACKUP_FORMAT) throw new Error('Expected encrypted backup envelope');
  const salt = fromB64(envelope.kdf.saltB64);
  const iv = fromB64(envelope.cipher.ivB64);
  const baseKey = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', hash: 'SHA-256', salt, iterations: envelope.kdf.iterations },
    baseKey,
    { name: 'AES-GCM', length: 256 },
    false,
    ['decrypt']
  );
  const plain = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv, additionalData: encoder.encode(JSON.stringify(envelope.aad)), tagLength: envelope.cipher.tagBits },
    key,
    fromB64(envelope.payloadB64)
  );
  return JSON.parse(decoder.decode(plain));
}

export async function signBackup(payload, { privateKey, publicKey }) {
  const bytes = encoder.encode(canonicalJsonStringify(payload));
  const signature = new Uint8Array(await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, privateKey, bytes));
  return {
    format: SIGNED_BACKUP_FORMAT,
    formatVersion: 1,
    payload,
    signature: {
      alg: 'ECDSA-P256-SHA256',
      sigB64: toB64(signature),
      publicKeyJwk: await crypto.subtle.exportKey('jwk', publicKey)
    }
  };
}

export async function verifyBackup(envelope) {
  if (envelope.format !== SIGNED_BACKUP_FORMAT) return { ok: false, reason: 'not-signed' };
  const pub = await crypto.subtle.importKey('jwk', envelope.signature.publicKeyJwk, { name: 'ECDSA', namedCurve: 'P-256' }, true, ['verify']);
  const ok = await crypto.subtle.verify(
    { name: 'ECDSA', hash: 'SHA-256' },
    pub,
    fromB64(envelope.signature.sigB64),
    encoder.encode(canonicalJsonStringify(envelope.payload))
  );
  return { ok, reason: ok ? 'ok' : 'invalid-signature' };
}


export async function verifyIntegrity(bundle) {
  if (!bundle?.integrity?.payloadHashB64) return false;
  const actual = await createIntegrity(bundle);
  return actual.payloadHashB64 === bundle.integrity.payloadHashB64;
}

export function isBackupBundle(value) {
  return value?.format === BACKUP_FORMAT;
}
