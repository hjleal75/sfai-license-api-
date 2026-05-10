# SFAI License API — Vercel + Upstash
### Sistema de licenciamiento para OneXcel Dashboard Builder

---

## Arquitectura

```
Excel (VBA) → POST /api/validate → Vercel Function → Upstash Redis
                                                          ↑
Tú (curl)  → POST /api/admin/license ──────────────────→ ┘
```

**Upstash** reemplaza al KV de Cloudflare — misma API, mismos comandos.
**Vercel** reemplaza al Worker — el código es casi idéntico.

---

## PASO 1 — Crear cuenta en Upstash (base de datos)

1. Ve a → https://console.upstash.com
2. Regístrate con **GitHub** (sin email de confirmación)
3. Crea una base de datos: **+ Create Database**
   - Name: `sfai-licenses`
   - Type: **Redis**
   - Region: `US-East-1` o `US-West-1`
   - Plan: **Free** ✓
4. En la pantalla de tu DB, copia:
   - `UPSTASH_REDIS_REST_URL`
   - `UPSTASH_REDIS_REST_TOKEN`

---

## PASO 2 — Subir el código a GitHub

```bash
# En tu terminal
git init
git add .
git commit -m "SFAI License API v1.0"
gh repo create sfai-license-api --private --push
# (o sube manualmente desde github.com/new)
```

---

## PASO 3 — Desplegar en Vercel

1. Ve a → https://vercel.com/new
2. Importa el repo `sfai-license-api`
3. Framework Preset: **Other**
4. En **Environment Variables**, agrega:

| Variable | Valor |
|---|---|
| `EXCEL_TOKEN` | `sfai_xl_2026_prod_XXXXXXXX` (inventa uno) |
| `ADMIN_TOKEN` | `sfai_admin_XXXXXXXXXXXXXXXX` (inventa uno) |
| `UPSTASH_REDIS_REST_URL` | (del Paso 1) |
| `UPSTASH_REDIS_REST_TOKEN` | (del Paso 1) |

5. Clic en **Deploy** → en 60 segundos tienes tu URL:
   `https://sfai-license-api.vercel.app`

---

## PASO 4 — Verificar que funciona

```bash
# Health check
curl https://sfai-license-api.vercel.app/api/health
# → {"status":"ok","product":"sfai-license-api"}

# Crear primera licencia
curl -X POST https://sfai-license-api.vercel.app/api/admin/license \
  -H "X-Admin-Token: sfai_admin_TU_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{
    "action":  "create",
    "email":   "cliente@empresa.com",
    "key":     "SFAI-48291-23",
    "tier":    "PRO",
    "expira":  "2026-12-31",
    "cliente": "Empresa Demo SA de CV"
  }'
# → {"ok":true,"action":"create","record":{...}}
```

---

## PASO 5 — Actualizar el VBA del Excel

1. Abre el archivo `mod_Licencia.bas` de este paquete
2. Edita la línea:
   ```vba
   Private Const API_URL = "https://sfai-license-api.vercel.app"
   ```
   Cambia `sfai-license-api` por el nombre real de tu proyecto en Vercel.
3. Edita también:
   ```vba
   Private Const EXCEL_TOKEN = "sfai_xl_2026_prod_TU_TOKEN_AQUI"
   ```
   Pon el mismo token que configuraste en el Paso 3.
4. En el Editor VBA del OneXcel: reemplaza el `mod_Licencia` existente.

---

## Comandos de operación diaria

### Crear licencia nueva
```bash
curl -X POST https://sfai-license-api.vercel.app/api/admin/license \
  -H "X-Admin-Token: TU_ADMIN_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"action":"create","email":"EMAIL","key":"KEY","tier":"PRO","expira":"2026-12-31","cliente":"NOMBRE"}'
```

### Revocar licencia (el Excel se cierra al siguiente intento de abrir)
```bash
curl -X POST https://sfai-license-api.vercel.app/api/admin/license \
  -H "X-Admin-Token: TU_ADMIN_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"action":"revoke","email":"EMAIL"}'
```

### Reactivar licencia
```bash
curl -X POST https://sfai-license-api.vercel.app/api/admin/license \
  -H "X-Admin-Token: TU_ADMIN_TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"action":"reactivate","email":"EMAIL"}'
```

### Ver todas las licencias
```bash
curl https://sfai-license-api.vercel.app/api/admin/licenses \
  -H "X-Admin-Token: TU_ADMIN_TOKEN"
```

---

## Límites del plan Free

| Recurso | Vercel Hobby | Upstash Free |
|---|---|---|
| Invocaciones/mes | 100,000 | 10,000/día |
| Duración máxima | 10 segundos | — |
| KV requests | — | 10,000/día |
| Costo | $0 | $0 |

Para el licenciamiento de un producto Excel con < 50 clientes:
- Cada cliente abre el Excel ~5 veces/día = 250 requests/día
- Límite: 10,000/día → aguanta hasta ~200 clientes en plan gratuito

---

## Estructura del proyecto

```
sfai-license-api/
├── api/
│   ├── health.js           ← GET  /api/health
│   ├── validate.js         ← POST /api/validate  (VBA)
│   └── admin/
│       ├── license.js      ← POST /api/admin/license
│       └── licenses.js     ← GET  /api/admin/licenses
├── lib/
│   └── redis.js            ← Cliente Upstash compartido
├── mod_Licencia.bas        ← Módulo VBA actualizado
├── .env.local              ← Variables locales (NO subir a GitHub)
├── .gitignore
├── vercel.json
└── package.json
```
