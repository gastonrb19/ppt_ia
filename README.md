# ppt_ia — Generación automática de PPT de Seguimiento Operativo (Metrogas)

Automatiza la creación de **dos presentaciones PowerPoint semanales** —Torre
Comercial y Técnica, y Torre Corporativa— a partir de los export de **Tickets**
(ServiceNow) y **Desarrollos** (SharePoint), replicando el formato de las
presentaciones modelo.

Cada semana se suben los Excel a la carpeta de la ejecución y el proceso arma los
`.pptx` (portada, agenda, volumetría, y por módulo: tabla de Tickets + tabla de
Desarrollos con cuadro resumen y gráfico de dona) y deja un log de seguimiento.

---

## 📄 Documentos de definición (empezar por aquí)

| Archivo | Qué es |
|---|---|
| **[ESPECIFICACION.md](ESPECIFICACION.md)** | **Fuente de verdad.** Rutas, fuentes de datos, mapeo por módulo, filtros, columnas, estilo/colores y estructura de slides. |
| **[RESUMEN_PROYECTO.md](RESUMEN_PROYECTO.md)** | Visión general del proyecto y de la estructura de carpetas. |
| **[decisiones_y_pendientes.txt](decisiones_y_pendientes.txt)** | Bitácora de decisiones tomadas y lo que queda pendiente. |
| **[proceso/README.md](proceso/README.md)** | Cómo ejecutar el proceso (comandos, test vs. flujo real). |

## 🖼️ Diagramas

- [Modelo de datos](diagramas/modelo_datos.png) — esquema de los Excel y su relación.
- [Diagrama de flujo](diagramas/diagrama_flujo.png) — proceso de generación.
- [Diagrama de secuencia](diagramas/diagrama_secuencia.png) — interacción semanal.

---

## 📁 Estructura del repositorio

```
├── ESPECIFICACION.md            fuente de verdad (spec completa)
├── RESUMEN_PROYECTO.md          visión general
├── decisiones_y_pendientes.txt  bitácora de decisiones/pendientes
├── modelo-archivos/             recursos estáticos
│   ├── pdf-modelo/              PPT modelo de referencia (PDF)
│   └── imagenes/                logos Metrogas / NTT DATA + imagen corporativa
├── proceso/                     pipeline reproducible
│   ├── extraer_datos.py         filtra y arma _data.json
│   ├── generar_ppt.js           construye las 2 .pptx (pptxgenjs)
│   └── README.md                instrucciones de ejecución
├── Ejecuciones/<dd-mm-yyyy>/    corridas REALES (datos/ + logs.txt + .pptx)
├── test/<dd-mm-yyyy>/           corridas de PRUEBA (no cuentan como reunión anterior)
├── diagramas/                   imágenes de los diagramas
└── archivos/                    versiones históricas (contexto.txt, requerimientoPPT.txt)
```

---

## ⚙️ Uso rápido

```bash
# 1) subir el Excel de tickets y el de desarrollos a Ejecuciones/<fecha>/datos/
# 2) extraer y filtrar
python3 proceso/extraer_datos.py Ejecuciones/<dd-mm-yyyy>/datos <dd-mm-yyyy>
# 3) generar las 2 presentaciones
npm install pptxgenjs        # solo la primera vez
node proceso/generar_ppt.js Ejecuciones/<dd-mm-yyyy>
```

Detalle completo en [proceso/README.md](proceso/README.md).

---

## 🔎 Reglas clave (resumen)

- **Tickets** se agrupan por `Grupo Asignado`; **Desarrollos** por `Frente` (ver mapeo en la spec).
- Filtro de tickets: se excluyen `Cerrado` y `Cancelado`; se incluyen INC + RITM; todas las sociedades.
- La **volumetría** compara Reunión Anterior / Actual / Diferencia; el "anterior" se lee del log previo en `/Ejecuciones/` (nunca de `/test/`).
- Estilo alineado al PDF modelo (título en recuadro, 1ª columna navy, filas zebra, barra inferior, dona por Status).

> Estado: **en progreso** (rama `inprogress`). Pendientes menores documentados en `decisiones_y_pendientes.txt`.
