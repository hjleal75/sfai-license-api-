// lib/redis.js — cliente Upstash compartido
// Upstash usa HTTP — funciona perfecto en Vercel serverless
import { Redis } from '@upstash/redis';

let client;

export function getRedis() {
  if (!client) {
    client = new Redis({
      url:   process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
  }
  return client;
}

// ── Helpers de licencias ──────────────────────────────────────
export function licKey(email) {
  return `lic:${email.toLowerCase().trim()}`;
}

export async function getLicense(email) {
  return await getRedis().get(licKey(email));
}

export async function setLicense(email, data) {
  return await getRedis().set(licKey(email), data);
}

export async function deleteLicense(email) {
  return await getRedis().del(licKey(email));
}

export async function listLicenses() {
  const redis = getRedis();
  const keys  = await redis.keys('lic:*');
  if (!keys.length) return [];
  const values = await redis.mget(...keys);
  return values.filter(Boolean);
}
