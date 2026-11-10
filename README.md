# Dame la Vida · tablero de Irvin

*Fake it till you make it.* Kanban y calendario para **desarrollo profesional** y **vida**, hecho a partir del tablero de Creativity del SCD CDMX 2026.
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
| Desarrollo profesional (DP) | Algoritmos y Estructura de Datos · Programación para Ciencia de Datos · Introducción a la Ciencia de Datos · Ética y Legalidad · Matemáticas · Inglés · Nube · Framework · UX/UI · Python · C · FrontEnd · BackEnd | Tarea · Práctica · Estudio · Curso · Workshop · Certificación · Evento · Recordatorio · Examen · Nota | ESCOM · EBAC · Google · AWS · Oracle · FinOps · Udemy · Domestika · Código Facilito · Platzi · DataCamp | `DP-001` |
| Personal | IMSS · Salud · Jandy · Familia · Pagos | Recordatorio · Evento · Cita · Encargo | — | `PER-001` |

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
- **Fijas**: tu rutina. Actividad · Tipo · Días · Día del mes · Inicio · Fin · Desde · Hasta · Excepto · Lugar · Notas · Aviso (min antes). Sin *Inicio* = todo el día (pagos); sin *Fin* = marca (comidas, medicamento). *Día del mes* acepta números y `último` (`15, último`). Tipo `Clase` se oculta en días sin clases.
- **Plan**: bloques de estudio. ID (de la tarea) · Actividad · Días · Inicio · Fin · Desde · Hasta · Notas · Excepto (fechas en que no aplica, separadas por coma). Días acepta `Diario`, `L-V` o letras `L,M,X,J,V,S,D`. El color sale de la urgencia de la tarea ligada.
- **Calendario IPN**: Evento · Inicio · Fin · Tipo · Notas. Tipo `Clases` marca el rango en que se muestra el horario; `Descanso` y `Vacaciones` ocultan las clases ese día.
- **Plan UX EBAC**: módulos del curso con su fecha meta y una casilla de Hecho (solo vive en la hoja).

## Funciones

- **Escuela / Personal**: el selector de arriba cambia de tablero; se recuerda en tu navegador.
- **Filtros plegables** por materia, tipo, escuela/plataforma y urgencia, más buscador.
- **Calendario** (Mes / Semana / Día): clases y deporte en azul, comidas como marcas verdes, rutina, pagos, bloques del Plan (color por urgencia: rojo, ámbar, verde), fechas límite y calendario IPN. Toca una asignación para editarla; toca una clase o actividad fija para ver su ventanita con detalles, WhatsApp y Google Calendar (repetido).
- **Reloj** con la hora de Ciudad de México en el encabezado.
- **Próximo pendiente** con días restantes y conteo de vencidas.
- **Google Calendar**: con fecha, botón para agregarla (con hora si la pusiste) y `.ics` para otros calendarios, con recordatorio un día antes (y una hora antes si tiene hora).
- **Mandármelo por WhatsApp**: abre tu chat contigo mismo con todo el detalle y el link a la asignación. También aparece al crear una nueva.
- **Link directo**: `https://…vercel.app/#ESC-003` abre esa asignación.

## Recordatorios a lo largo del día

El Apps Script revisa cada 5 minutos qué empieza pronto y te manda un mensaje:

- **Al despertar**: el resumen del día (pagos, fechas límite, vencidas, clases y plan de estudio).
- **Clases**: 10 min antes, para salir a tiempo.
- **Actividades fijas**: según la columna *Aviso (min antes)* de la pestaña Fijas (vacía = sin aviso).
- **Bloques del Plan**: al empezar. **Tareas con hora**: 1 h antes.
- Lo que cae a la misma hora llega en un solo mensaje.

Canales (propiedades del script):

| Propiedad | Para qué |
|-----------|----------|
| `FORMSPREE_ENDPOINT` | Tu formulario, ej. `https://formspree.io/f/xxxx`. En Formspree activa el plugin **WhatsApp** en la pestaña *Workflow* del formulario. |
| `FORMSPREE_LIMITE` | Envíos al mes antes de pausar Formspree. Por defecto 50 (plan gratis). Súbelo si tienes un plan de pago. |
| `CALLMEBOT_APIKEY` | Opcional y gratis. Se usa si no hay Formspree o si ya se acabó el cupo del mes. |

Un día normal son unos 15 mensajes (unos 450 al mes): el plan gratis de Formspree no alcanza, así que conviene CallMeBot como respaldo o un plan de pago.

Funciones (▶ en el editor): **`probarRecordatorio`** manda un mensaje de prueba, **`verAvisosDeHoy`** lista en el registro lo que te llegaría hoy, **`instalarRecordatorios`** los activa y **`quitarRecordatorios`** los apaga.

## Desarrollo local

```bash
npm i -g vercel
cp .env.example .env.local
vercel dev
```

Si abres `index.html` sin backend entra en **modo demo** con datos de ejemplo.
