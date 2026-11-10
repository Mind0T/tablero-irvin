// Función serverless de Vercel: puente entre el tablero y la API de Apps Script.
// Variables de entorno (Vercel > Settings > Environment Variables):
//   APPS_SCRIPT_URL  URL de la implementación web de Apps Script (termina en /exec)
//   API_SECRET       La clave que generó generarSecreto() en Apps Script
//   TEAM_CODE        (opcional) Código de acceso que escribes una vez para entrar al tablero

import crypto from "node:crypto";

const ACTIONS = new Set(["save", "move", "delete", "hecho", "corte", "corteDel", "modulo", "pagoJandy"]);

function sameText(a, b) {
  const ha = crypto.createHash("sha256").update(String(a)).digest();
  const hb = crypto.createHash("sha256").update(String(b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string" && req.body) return JSON.parse(req.body);
  return {};
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const { APPS_SCRIPT_URL, API_SECRET, TEAM_CODE } = process.env;

  if (!APPS_SCRIPT_URL || !API_SECRET) {
    return res.status(500).json({ ok: false, error: "Falta configurar APPS_SCRIPT_URL o API_SECRET en Vercel." });
  }
  if (TEAM_CODE && !sameText(req.headers["x-team-code"] || "", TEAM_CODE)) {
    return res.status(401).json({ ok: false, error: "Código de acceso incorrecto." });
  }

  let upstream;
  try {
    if (req.method === "GET") {
      const url = `${APPS_SCRIPT_URL}?api=1&secret=${encodeURIComponent(API_SECRET)}&t=${Date.now()}`;
      upstream = await fetch(url, { redirect: "follow" });
    } else if (req.method === "POST") {
      let body;
      try { body = readBody(req); } catch { return res.status(400).json({ ok: false, error: "El cuerpo no es JSON válido." }); }
      if (!ACTIONS.has(body.action)) return res.status(400).json({ ok: false, error: "Acción no válida." });
      const payload = {
        action: body.action, id: body.id, estado: body.estado, task: body.task, user: body.user, secret: API_SECRET,
        clave: body.clave, fecha: body.fecha, done: body.done,              // hecho (marcar una ocurrencia)
        corte: body.corte,                                                   // corte con evidencia (nota, link, imagen)
        camino: body.camino, modulo: body.modulo,                            // modulo del temario
        no: body.no,                                                          // número de pago (Pagos Jandy)
      };
      upstream = await fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload),
        redirect: "follow",
      });
    } else {
      res.setHeader("Allow", "GET, POST");
      return res.status(405).json({ ok: false, error: "Método no permitido." });
    }
  } catch (err) {
    return res.status(502).json({ ok: false, error: "No pude contactar a Apps Script: " + err.message });
  }

  const text = await upstream.text();
  let data;
  try { data = JSON.parse(text); } catch {
    return res.status(502).json({
      ok: false,
      error: "Apps Script no respondió JSON. Revisa que la implementación tenga acceso «Cualquier persona» y que la URL termine en /exec.",
    });
  }
  if (data.ok === false) {
    const status = data.error === "No autorizado" || /API_SECRET/.test(data.error) ? 502 : 400;
    if (status === 502 && data.error === "No autorizado") data.error = "Apps Script rechazó la clave: revisa que API_SECRET en Vercel sea la misma.";
    return res.status(status).json(data);
  }
  return res.status(200).json(data);
}
