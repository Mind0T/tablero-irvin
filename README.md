# Tablero personal · Irvin

Kanban para **escuela** y **vida**, hecho a partir del tablero de Creativity del SCD CDMX 2026.
Los datos viven en un Google Sheet; el tablero corre en Vercel.

```
Navegador ──► Vercel /api/board ──► Apps Script (/exec) ──► Google Sheets
 index.html    agrega API_SECRET      valida la clave         pestaña "Tareas"
               valida TEAM_CODE
```

La URL de Apps Script, la clave y tu número **nunca están en este repo**.

## Cómo está organizado

| Ámbito   | Materias / categorías | Tipos | Escuela / plataforma (hasta 2) | IDs |
|----------|-----------------------|-------|--------------------------------|-----|
| Escuela  | Algoritmos y Estructura de Datos · Programación para Ciencia de Datos · Introducción a la Ciencia de Datos · Ética y Legalidad · Nube · Framework · UX/UI · Python · C · FrontEnd · BackEnd | Tarea · Práctica · Estudio · Curso · Workshop · Certificación · Recordatorio · Examen · Nota | ESCOM · EBAC · Google · AWS · Oracle · FinOps · Udemy · Domestika · Código Facilito · Platzi · DataCamp | `ESC-001` |
| Personal | IMSS · Jandy · Familia · Pagos | Recordatorio · Evento · Cita · Encargo | — | `PER-001` |

Cada asignación lleva **ámbito + materia + tipo + urgencia** (Urgente / Importante / Normal), hasta dos escuelas o plataformas, fecha límite, hora y fecha de inicio opcionales.
Ejemplo: *Certificación* de *Nube* en *AWS + Código Facilito*, *Urgente*, para el 10 de diciembre.

Las listas se cambian en `apps-script/Code.gs` (constante `AMBITOS`) y el tablero las toma solas.

## 1 · Google Sheet + Apps Script

1. Abre el Google Sheet **«Tablero personal · Irvin»** (ya tiene todas las pestañas con encabezados y menús) → **Extensiones → Apps Script**.
2. Reemplaza `Código.gs` con `apps-script/Code.gs`.
3. (Solo si partes de un Sheet vacío) ejecuta **`crearHoja`** (▶) para crear las pestañas.
4. Ejecuta **`generarSecreto`** → copia el `API_SECRET` del registro de ejecución.
5. **⚙️ Configuración del proyecto → Propiedades de la secuencia de comandos → Agregar**:
   - `MI_WHATSAPP` = tu número a 10 dígitos (el botón «Mandármelo por WhatsApp» lo usa).
6. **Implementar → Nueva implementación → Aplicación web** · Ejecutar como: **Yo** · Acceso: **Cualquier persona** → copia la URL `/exec`.

## 2 · Vercel

Importa el repo (Framework: **Other**, sin build) y agrega:

| Variable          | Valor                                   |
|-------------------|-----------------------------------------|
| `APPS_SCRIPT_URL` | La URL `/exec`                          |
| `API_SECRET`      | La clave de `generarSecreto()`          |
| `TEAM_CODE`       | Tu código de acceso (lo pide una vez)   |

Si cambias `Code.gs`: **Implementar → Gestionar implementaciones → ✏️ → Nueva versión** (la URL no cambia).

## La hoja

- **Tareas**: ID · Ámbito · Categoría · Tipo · Tarea · Descripción · Responsable · Estado · Prioridad · Fecha · Hora · Notas · Entregable · Recursos · Creado · Actualizado · Actualizado por · Inicio · Instituciones
  - **Prioridad** es la urgencia (Urgente / Importante / Normal). **Inicio** es desde cuándo arranca algo que no tiene fecha límite. **Instituciones**: hasta dos, separadas por coma (`AWS, Udemy`).
  - **Entregable**: link del resultado final (repo, notebook, documento, comprobante).
  - **Recursos**: un link por línea, `Nombre | https://…`.
  - Las columnas se ubican por el texto del encabezado: puedes reordenarlas.
- **Equipo**: Nombre · Color. Solo estás tú (Irvin).
- **Recursos**: Etiqueta · Enlace · Notas · Ámbito · Categoría. Salen en el botón **Recursos** y se precargan en cada asignación nueva del mismo ámbito y categoría. Ámbito o categoría vacíos = aplica a todo.
- **Horario**: tu horario de clases tal cual (columna `HORA` con `07:00 - 08:30` y una columna por día). El calendario reconoce la materia y el salón.
- **Plan**: bloques de estudio. ID (de la tarea) · Actividad · Días · Inicio · Fin · Desde · Hasta · Notas. Días acepta `Diario`, `L-V` o letras `L,M,X,J,V,S,D`. El color sale de la urgencia de la tarea ligada.
- **Calendario IPN**: Evento · Inicio · Fin · Tipo · Notas. Tipo `Clases` marca el rango en que se muestra el horario; `Descanso` y `Vacaciones` ocultan las clases ese día.
- **Plan UX EBAC**: módulos del curso con su fecha meta y una casilla de Hecho (solo vive en la hoja).

## Funciones

- **Escuela / Personal**: el selector de arriba cambia de tablero; se recuerda en tu navegador.
- **Filtros plegables** por materia, tipo, escuela/plataforma y urgencia, más buscador.
- **Calendario** (Mes / Semana / Día): clases del horario, bloques del Plan, fechas límite y calendario IPN. Toca un bloque para abrir su asignación; toca un día para verlo completo.
- **Próximo pendiente** con días restantes y conteo de vencidas.
- **Google Calendar**: con fecha, botón para agregarla (con hora si la pusiste) y `.ics` para otros calendarios, con recordatorio un día antes (y una hora antes si tiene hora).
- **Mandármelo por WhatsApp**: abre tu chat contigo mismo con todo el detalle y el link a la asignación. También aparece al crear una nueva.
- **Link directo**: `https://…vercel.app/#ESC-003` abre esa asignación.

## Aviso automático diario (opcional)

Para que te llegue un WhatsApp cada mañana (7:00) con lo vencido, lo de hoy, lo de mañana y tus bloques de estudio del día, sin abrir el tablero:

1. Saca tu apikey gratis en [callmebot.com](https://www.callmebot.com/blog/free-api-whatsapp-messages/) (sección WhatsApp: le mandas un mensaje de activación y te responde con la clave).
2. Agrégala en las propiedades del script como `CALLMEBOT_APIKEY`.
3. Ejecuta **`probarAviso`** para recibir uno de prueba y luego **`instalarAvisoDiario`**. Para apagarlo: **`quitarAvisoDiario`**.

## Desarrollo local

```bash
npm i -g vercel
cp .env.example .env.local
vercel dev
```

Si abres `index.html` sin backend entra en **modo demo** con datos de ejemplo.
