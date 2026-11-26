// Función serverless de Vercel: puente entre el tablero y la API de Apps Script.
// Variables de entorno (Vercel > Settings > Environment Variables):
//   APPS_SCRIPT_URL      URL de la implementación web de Apps Script (termina en /exec)
//   API_SECRET           La clave que generó generarSecreto() en Apps Script
//   EDITOR_CODE          Contraseña de editor: ve todo (incluido lo personal) y puede editar.
//   READER_CODE          Contraseña de lector: ve el tablero sin lo personal y NO puede editar.
//   TEAM_CODE            (heredado) se usa como EDITOR_CODE si no defines EDITOR_CODE.
//   PERSONAL_FIJA_TIPOS  (opcional) tipos de actividad del calendario que son personales,
//                        separados por coma. Por defecto: "Medicamento,Salud,Cita,Médico".
//
// Seguridad: la censura de lo personal y el bloqueo de escrituras ocurren AQUÍ, en el
// servidor. Al lector nunca se le envía el contenido personal ni se le acepta un cambio,
// así que el candado no depende del navegador. El API_SECRET solo vive en el servidor.

import crypto from "node:crypto";

const ACTIONS = new Set([
  "save", "move", "delete", "hecho", "corte", "corteDel", "modulo", "pagoJandy",
  // Notas (vista tipo Keep)
  "notaSave", "notaDelete", "notaOrden", "notaItemSave", "notaItemEstado", "notaItemDelete", "notaItemOrden",
]);

// Tipos de "Fijas" (calendario) que se consideran personales y se ocultan al lector.
const PERSONAL_FIJA_TIPOS = (process.env.PERSONAL_FIJA_TIPOS || "Medicamento,Salud,Cita,Médico,Medico")
  .split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);

// Campos de texto de una tarea que se vacían cuando la tarea es personal y el rol es lector.
const TASK_BLANK = ["categoria", "tipo", "titulo", "descripcion", "responsable", "instituciones", "notas", "enlace", "materiales"];

function sameText(a, b) {
  const ha = crypto.createHash("sha256").update(String(a)).digest();
  const hb = crypto.createHash("sha256").update(String(b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

// Decide el rol a partir del código recibido. Devuelve "editor", "reader" o null (inválido).
function roleFor(code) {
  const editor = process.env.EDITOR_CODE || process.env.TEAM_CODE || "";
  const reader = process.env.READER_CODE || "";
  if (!editor && !reader) return "editor"; // sin códigos configurados: acceso abierto (como antes)
  if (editor && sameText(code, editor)) return "editor";
  if (reader && sameText(code, reader)) return "reader";
  return null;
}

function isPersonalTask(t) {
  return !!t && (t.ambito === "Personal" || /^per-/i.test(String(t.id || "")));
}

// Quita del tablero todo lo personal antes de enviarlo a un lector.
function redactForReader(data) {
  if (!data || data.ok === false) return data;
  const d = { ...data };

  if (Array.isArray(d.tasks)) {
    d.tasks = d.tasks.map((t) => {
      if (!isPersonalTask(t)) return t;
      const o = { ...t, ambito: "Personal", locked: true };
      for (const k of TASK_BLANK) if (k in o) o[k] = "";
      return o; // conserva id, estado, prioridad, fecha, hora, inicio para ubicar un bloque bloqueado
    });
  }
  if (Array.isArray(d.recursos)) {
    d.recursos = d.recursos.filter((r) => String(r && r.ambito) !== "Personal");
  }
  if (Array.isArray(d.fijas)) {
    d.fijas = d.fijas.map((f) => {
      const tipo = String((f && f.tipo) || "").toLowerCase();
      if (!PERSONAL_FIJA_TIPOS.includes(tipo)) return f;
      return { ...f, actividad: "", notas: "", lugar: "", locked: true };
    });
  }
  // "Jandy" es personal; las evidencias (cortes) pueden contener contenido personal.
  d.pagosJandy = Array.isArray(d.pagosJandy) ? [] : {};
  d.cortes = Array.isArray(d.cortes) ? [] : {};
  return d;
}

function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string" && req.body) return JSON.parse(req.body);
  return {};
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const { APPS_SCRIPT_URL, API_SECRET } = process.env;

  if (!APPS_SCRIPT_URL || !API_SECRET) {
    return res.status(500).json({ ok: false, error: "Falta configurar APPS_SCRIPT_URL o API_SECRET en Vercel." });
  }

  const role = roleFor(req.headers["x-team-code"] || "");
  if (role === null) {
    return res.status(401).json({ ok: false, error: "Código de acceso incorrecto." });
  }

  let upstream;
  try {
    if (req.method === "GET") {
      const url = `${APPS_SCRIPT_URL}?api=1&secret=${encodeURIComponent(API_SECRET)}&t=${Date.now()}`;
      upstream = await fetch(url, { redirect: "follow" });
    } else if (req.method === "POST") {
      // Cualquier cambio necesita la contraseña de editor.
      if (role !== "editor") {
        return res.status(403).json({ ok: false, error: "Se necesita la contraseña de editor para hacer cambios.", needEditor: true });
      }
      let body;
      try { body = readBody(req); } catch { return res.status(400).json({ ok: false, error: "El cuerpo no es JSON válido." }); }
      if (!ACTIONS.has(body.action)) return res.status(400).json({ ok: false, error: "Acción no válida." });
      const payload = {
        action: body.action, id: body.id, estado: body.estado, task: body.task, user: body.user, secret: API_SECRET,
        clave: body.clave, fecha: body.fecha, done: body.done,              // hecho (marcar una ocurrencia)
        corte: body.corte,                                                   // corte con evidencia (nota, link, imagen)
        camino: body.camino, modulo: body.modulo,                            // modulo del temario
        no: body.no,                                                          // número de pago (Pagos Jandy)
        nota: body.nota, item: body.item, ids: body.ids,                     // notas (vista tipo Keep)
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

  // GET de lector: censura lo personal antes de responder. El rol viaja al frontend.
  if (req.method === "GET" && role === "reader") data = redactForReader(data);
  if (data && typeof data === "object" && !Array.isArray(data)) data.role = role;

  return res.status(200).json(data);
}
