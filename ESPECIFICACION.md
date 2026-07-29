# Especificación única — Generación de PPT de Seguimiento Operativo

> **Este archivo es la fuente de verdad del proyecto.** Reemplaza y unifica a
> `contexto.txt` (prompt del equipo) y `requerimientoPPT.txt` (sintaxis previa),
> que quedan en la carpeta `archivos/` solo como referencia histórica.
> Está escrito para leerse como usuario y para que el asistente lo use como
> especificación mantenible. Los bloques ```yaml``` contienen la configuración
> operativa; el texto explica el porqué.

**Versión:** 1.0 · **Actualizado:** 29-07-2026

---

## 1. Objetivo

Generar **dos** presentaciones PowerPoint (`.pptx`) editables, una por torre, a
partir de datos de Excel, replicando el formato de las presentaciones modelo.

```yaml
entregables:
  - "Seguimiento Semanal Torre Comercial y Técnico.pptx"
  - "Seguimiento Semanal Torre Corporativa.pptx"
```

---

## 2. Ubicación de archivos (lectura SIEMPRE desde /modelo-archivos)

Desde ahora, toda la base de datos y los recursos se leen desde `modelo-archivos/`.

```yaml
rutas:
  raiz_proyecto: "/PPT"
  base_estatica: "/PPT/modelo-archivos"           # recursos fijos (modelos, imágenes)
  # --- DATOS: cambian cada semana, se SUBEN a la carpeta de la ejecución ---
  datos_semana:  "/PPT/Ejecuciones/<dd-mm-yyyy>/datos/"   # aquí se sube tickets + desarrollos
    # tickets     -> archivo con hoja "Page 1"      (agrupa por "Grupo Asignado")
    # desarrollos -> archivo con hoja "Desarrollos" (agrupa por "Frente")
  # --- Recursos estáticos ---
  pdf_modelo:
    corporativa: "/PPT/modelo-archivos/pdf-modelo/MTGS - Seguimiento Operativo Sistemas Corporativos_23-07-2026.pdf"
    comercial:   "/PPT/modelo-archivos/pdf-modelo/Presentación de Comité Operativo 23.07.2026.pdf"
  imagenes:
    logo_metrogas:      "/PPT/modelo-archivos/imagenes/LOGO_metrogas.png"
    logo_nttdata_azul:  "/PPT/modelo-archivos/imagenes/LOGO_nttdata_azul.png"
    logo_nttdata_blanco:"/PPT/modelo-archivos/imagenes/LOGO_nttdata_blanco.png"
    imagen_corporativa: "/PPT/modelo-archivos/imagenes/IMG_corporativa.png"
  ejecucion:
    carpeta: "/PPT/Ejecuciones/<dd-mm-yyyy>/"
    datos:   "/PPT/Ejecuciones/<dd-mm-yyyy>/datos/"   # entrada (se sube semanalmente)
    log:     "/PPT/Ejecuciones/<dd-mm-yyyy>/logs.txt" # salida-registro
  test:
    carpeta: "/PPT/test/<dd-mm-yyyy>/"                # misma estructura que una ejecución
    proposito: "corridas de prueba; NO se consideran en el flujo real"
    regla: "la 'reunión anterior' se lee SOLO de /Ejecuciones/, nunca de /test/"
```

> Importante: `/test/` replica la estructura de una ejecución (datos/, logs.txt,
> .pptx) pero está fuera de `/Ejecuciones/` a propósito, para que una corrida de
> prueba no sea tomada como "reunión anterior" del flujo real.

> Flujo semanal: cada semana se crea la carpeta `Ejecuciones/<dd-mm-yyyy>/` y se
> **sube dentro de `datos/`** el export de tickets y el de desarrollos de esa
> semana. El proceso lee de ahí, escribe el `logs.txt` en la misma carpeta y usa
> los recursos estáticos de `modelo-archivos/`. Los nombres de los `.xlsx` pueden
> variar; se toma el de tickets (hoja "Page 1") y el de desarrollos (hoja
> "Desarrollos") más recientes dentro de `datos/`.

---

## 3. Fuentes y agrupación

- **Tickets** → `datos/tickets`, se agrupa por la columna **`Grupo Asignado`**.
- **Desarrollos** → `datos/desarrollos`, se agrupa por la columna **`Frente`**.
- El nombre del módulo **no** se usa para filtrar; se usa el valor real de la columna.

```yaml
match:
  tipo: tolerante            # ignora mayús/minús, espacios, guiones,
                             # sufijo " - Metrogas" y prefijo "SAP "
  salvaguarda: "no confundir tokens distintos (ej. Fico != Fica)"
```

---

## 3.1. Modelado de datos (esquema de los Excel)

Dos fuentes independientes, relacionadas por el concepto **Módulo**:
`Tickets."Grupo Asignado"` ↔ `Desarrollos."Frente"` (vía la tabla de mapeo de la
sección 4, con match tolerante).

![Modelo de datos](diagramas/modelo_datos.png)

**TICKETS** — `task (4).xlsx`, hoja `Page 1` (13 columnas; 8 en uso):

| Col | Campo | Tipo | Uso en la PPT |
|---|---|---|---|
| A | Número | texto | Columna "Número" |
| B | Empresa | texto | Columna "Sociedad" |
| C | Estado | texto | Columna "Estado" + **filtro** (excluir Cerrado) |
| D | Resumen | texto | Columna "Resumen" |
| E | Abierto por | texto | Columna "Usuario MTGS/GS00" |
| G | Creado | datetime | "Fecha de apertura" + cálculo "Días Tranc" |
| H | Tipo de tarea | texto | **filtro** (Incidente / Elemento solicitado) |
| I | Grupo Asignado | texto | **CLAVE**: agrupa por módulo |
| F, J, K, L, M | Creado por, Prioridad, Usuario Asignado, Cliente [Incidente], Descripción | texto | no usadas |

**DESARROLLOS** — `SHP01072026.xlsx`, hoja `Desarrollos` (37 columnas; 9 en uso + Torre de referencia):

| Col | Campo | Tipo | Uso en la PPT |
|---|---|---|---|
| A | Codigo | texto | Columna "Código" |
| B | Tipo de requerimiento | texto | Columna "Tipo de requerimiento" |
| C | Título | texto | Columna "Título" |
| G | Status | texto | Columna "Status" + cuadro resumen |
| H | Frente | texto | **CLAVE**: agrupa por módulo |
| K | Creado | datetime | Columna "Creado" |
| L | Solicitante | texto | Columna "Solicitante" |
| N | Torre | texto | referencia (Corporativa / Comercial / Técnica) |
| T | Fecha plan entrega QA | datetime | Columna "Fecha plan entrega QA" |
| W | Fecha entrega a QA | datetime | Columna "Fecha entrega a QA" |
| (otras 27) | campos operativos varios | — | no usadas |

---

## 4. Mapeo por módulo

```yaml
modulos:
  # ---- Deck COMERCIAL Y TÉCNICA ----
  - nombre: NET
    deck: comercial
    tickets_grupo: "Net - Metrogas"
    desarrollos_frente: "Portales y autoatencion"
  - nombre: Sharepoint
    deck: comercial
    tickets_grupo: "Sharepoint - Metrogas"
    desarrollos_frente: null            # sin frente -> tabla vacía
  - nombre: "B&I"
    deck: comercial
    tickets_grupo: "B&I - Metrogas"
    desarrollos_frente: "SAP ISU"
  - nombre: FICA
    deck: comercial
    tickets_grupo: "Fica - Metrogas"
    desarrollos_frente: "SAP FICA"
  - nombre: CRM
    deck: comercial
    tickets_grupo: "Crm - Metrogas"
    desarrollos_frente: "SAP CRM"
  - nombre: PM
    deck: comercial
    tickets_grupo: "Pm - Metrogas"
    desarrollos_frente: "SAP PM"
  - nombre: PS
    deck: comercial
    tickets_grupo: "Ps - Metrogas"
    desarrollos_frente: "SAP PS"
  # ---- Deck CORPORATIVA ----
  - nombre: HCM
    deck: corporativa
    tickets_grupo: "Hcm - Metrogas"
    desarrollos_frente: "SAP HCM"
  - nombre: SSFF
    deck: corporativa
    tickets_grupo: "Ssff - Metrogas"
    desarrollos_frente: null            # sin frente -> tabla vacía
  - nombre: "R&P"
    deck: corporativa
    tickets_grupo: "R&P - Metrogas"
    desarrollos_frente: null            # sin frente -> tabla vacía
  - nombre: SD
    deck: corporativa
    tickets_grupo: "Sd - Metrogas"
    desarrollos_frente: "SAP SD"
  - nombre: BW
    deck: corporativa
    tickets_grupo: "Bw-Hana - Metrogas"
    desarrollos_frente: "BI"
  - nombre: FICO
    deck: corporativa
    tickets_grupo: "Fico - Metrogas"
    desarrollos_frente: "SAP FI"
  - nombre: MM
    deck: corporativa
    tickets_grupo: "Mm - Metrogas"
    desarrollos_frente: "SAP MM"
```

---

## 5. Filtros de datos

```yaml
filtros_tickets:
  estado_excluir: ["Cerrado", "Cancelado"]           # confirmado 29-07 (¿+ Resuelto/Terminado? por revisar)
  tipo_tarea: ["Incidente", "Elemento solicitado"]   # PROVISIONAL: ¿ambos o solo Incidente?
  empresa: "todas"                     # Metrogas SA, Gas Sur, Innergy
filtros_desarrollos:
  # solo se filtra por Frente (ver mapeo). Sin filtro de status para la tabla.
  modulo_sin_frente: "mostrar tabla solo con cabecera"
```

---

## 6. Tabla de Tickets (por módulo)

Columnas y origen (mandan las del requerimiento previo):

```yaml
tabla_tickets:
  columnas:
    - {titulo: "Estado",            origen: "Estado"}
    - {titulo: "Sociedad",          origen: "Empresa"}
    - {titulo: "Número",            origen: "Número"}
    - {titulo: "Resumen",           origen: "Resumen"}
    - {titulo: "Fecha de apertura", origen: "Creado"}
    - {titulo: "Usuario MTGS/GS00", origen: "Abierto por"}
    - {titulo: "Días Tranc",        origen: "calculado = fecha_actual - Creado (días, positivo)"}
  orden: ["Estado", "Fecha de creación DESC"]
```

---

## 7. Tabla de Desarrollos (por módulo)

```yaml
tabla_desarrollos:
  columnas:
    - "Status"
    - "Creado"
    - "Código"
    - "Tipo de requerimiento"
    - "Título"
    - "Fecha plan entrega QA"
    - "Fecha entrega a QA"
    - "Solicitante"
  orden: ["Estado", "Fecha de creación DESC"]
  cuadro_resumen: "conteo de desarrollos por Status (recuadro inferior)"
```

---

## 8. Volumetría de Tickets (slide 3) y logs

```yaml
volumetria:
  tabla: "por módulo -> Reunión Anterior | Reunión Actual | Diferencia + fila total"
  actual: "conteo desde datos/tickets con los filtros de la sección 5"
  anterior: "se lee del logs.txt de la ejecución previa"
  primera_ejecucion: "los valores 'Anterior' se ingresan a mano"
  diferencia: "Actual - Anterior"

logs:
  ruta: "/PPT/Ejecuciones/<dd-mm-yyyy>/logs.txt"
  contenido_minimo:
    - "Cabecera: ID, fecha/hora, responsable, tipo, estado"
    - "Fuentes: archivos, fecha del dato, n° registros"
    - "Parámetros aplicados (sección 5)"
    - "Volumetría por módulo y deck (anterior/actual/diferencia) + totales"
    - "Resumen de Desarrollos por módulo (conteo por Status) + totales"
    - "Observaciones (módulos vacíos, provisionales, alertas)"
    - "Entregables generados"
  regla: "los 'Actual' de esta ejecución son los 'Anterior' de la siguiente"
  ejemplo: "/PPT/Ejecuciones/29-07-2026/logs.txt"
```

---

## 9. Estilo visual

Colores tomados del PDF modelo (para que las slides sean coherentes con él):

```yaml
estilo:
  colores:
    verde_marca:     "#A2AD00"   # título portada + bloque barra inferior
    azul_subtitulo:  "#284973"   # subtítulo de la portada
    titulo_texto:    "#3B5A9E"   # texto del título de cada slide (dentro del recuadro)
    titulo_borde:    "#6B88C0"   # borde del recuadro redondeado del título
    navy_tabla:      "#002060"   # cabecera de tabla + 1ª columna + fila "Total general"
    zebra_clara:     "#E3E7F0"   # filas alternas (zebra) del cuerpo de tabla
    barra_navy:      "#002060"   # barra inferior de marca (junto al verde)
  titulo_slide:
    formato: "recuadro redondeado, fondo blanco, borde azul (titulo_borde), texto centrado en titulo_texto, negrita"
  tablas:
    cabecera:    "fila navy_tabla, texto blanco negrita"
    col_categoria: "Tickets: 1ª y 2ª columna (Estado, Sociedad) en navy_tabla texto blanco; Desarrollos: 1ª columna (Status) en navy_tabla texto blanco"
    cuerpo:      "resto de columnas con zebra (blanco / zebra_clara), texto oscuro"
    total:       "fila 'Total general' del resumen en navy_tabla texto blanco negrita"
  resumen_desarrollos: "el cuadro 'Resumen por Status' lista TODOS los estados presentes (orden por cantidad desc) + Total general, de modo que la suma cuadre"
  grafico: "en cada slide de Desarrollos: gráfico de dona con % por Status + leyenda de colores (paleta por estado). Se omite si el módulo no tiene desarrollos"
  barra_inferior: "en TODAS las slides de contenido: bloque verde_marca a la izquierda + barra barra_navy a lo ancho"
  tipografias:
    titulos: "FS EMERIC"        # si no está instalada -> reemplazo (ej. Aptos/Calibri)
    tablas:  "APTOS NARROW 11"
  logos:
    metrogas: "sup. izq. de todas las diapositivas"
    nttdata:  "sup. der. de todas las diapositivas"
  cierre: "última diapositiva: logo Metrogas sobre fondo azul"
  reglas:
    - "si una tabla no cabe en una diapositiva, crear diapositivas adicionales con el mismo formato"
    - "si un módulo no tiene tickets/desarrollos, mostrar la tabla solo con la cabecera"
```

---

## 10. Estructura de diapositivas (ambos decks)

```yaml
estructura_slides:
  1: "Portada — 'Seguimiento Semanal' + subtítulo de la torre + logo Metrogas + fecha (formato largo) + imagen corporativa"
  2: "Agenda — 1) Volumetría tickets  2) Detalle tickets por módulo  3) Detalle desarrollos por módulo"
  3: "Volumetría Tickets — tabla total por módulo (anterior/actual/diferencia) + total general"
  "4..N": "por cada módulo: (a) Tickets módulo X  (b) Desarrollos módulo X + cuadro resumen"
  ultima: "logo Metrogas sobre fondo azul"
decks:
  comercial:  [NET, Sharepoint, "B&I", FICA, CRM, PM, PS]
  corporativa:[HCM, SSFF, "R&P", SD, BW, FICO, MM]
```

---

## 11. Diagramas de funcionamiento

**Diagrama de flujo (proceso de generación):**

![Diagrama de flujo](diagramas/diagrama_flujo.png)

**Diagrama de secuencia (interacción semanal):**

![Diagrama de secuencia](diagramas/diagrama_secuencia.png)

---

## 12. Pendientes (ver `decisiones_y_pendientes.txt`)

1. **[Mañana]** Estados a excluir en Tickets (¿solo `Cerrado` o + `Cancelado`/`Resuelto`/`Terminado`?).
2. **[Mañana]** Tipo de tarea (¿ambos INC+RITM o solo Incidente?).
3. No hay `.pptx` modelo (solo PDF): el layout se aproxima desde el PDF.
4. Confirmar tipografías FS EMERIC / APTOS NARROW o usar reemplazo.
5. Validar que ambas fuentes correspondan a la misma reunión (tickets 24-07 vs desarrollos 01-07).
