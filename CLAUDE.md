# CLAUDE.md - Panel Keep It

## Instrucciones del Proyecto (Permanentes)

1. **Documentación**: Usar siempre este único archivo `CLAUDE.md` para documentar.
2. **Registro de cambios**: Documentar todas las mejoras y cambios realizados en este archivo (con fecha).
3. **Registro de peticiones**: Mantener `prompt.txt` con cada petición del usuario.
4. **Git commits**: Siempre hacer commit con detalle de mejoras y cambios.
5. **Context7 MCP**: Verificar documentación actualizada de librerías/frameworks vía Context7 antes de responder con código.
6. **Entorno virtual**: Siempre usar `venv` con Python.
7. **Lenguaje**: Python.

## Descripción del Proyecto

> Panel "Keep It" — procesar exports/Excel y generar hojas planas listas para
> Looker Studio (sin blends) y/o un panel web Flask propio.
>
> _(Completar: qué datos entran, qué transformaciones se hacen y qué gráficas/hojas
> salen. Ir llenando a medida que se construye.)_

### Objetivo
- Tomar los archivos de origen (exports/Excel) y producir hojas `G_*` planas, una
  por gráfica, para conectar cada visualización a UNA sola fuente en Looker Studio.

### Convención de módulos (igual que dashboard_voa_sodimac)
- Cada bloque de procesamiento es una carpeta autocontenida con:
  - `input/` — archivos de origen que deja el usuario
  - `output/` — resultados con timestamp (`<algo>_YYYY-MM-DD_HH-MM.xlsx`)
  - `procesar_*.py` — script del bloque
  - `procesados.json` — registro anti-duplicación (no reprocesar archivos ya corridos)
- `graficas_looker/` — pre-calcula las hojas `G_*` y genera un `G_para_looker.xlsx`
  chico para subir a Drive/Looker (evita el límite de 10M celdas del Sheet).
- Rutas resueltas con `Path(__file__).parent` para poder ejecutar desde cualquier cwd.

### Entornos
- **WSL2**: `python3 -m venv venv && source venv/bin/activate`
- **Windows**: `python -m venv venv && venv\Scripts\activate`
- Instalar dependencias: `pip install -r requirements.txt`

---

## Modelo de Datos (análisis de los archivos de origen)

### Negocio: ¿qué es "Keep It"?
Programa de **retención por cupón**. Ante un reclamo (producto en mal estado, pieza
faltante, incompleto, no corresponde, etc.), en vez de hacer **logística inversa**
(retirar el producto + asumir la pérdida de su costo) se entrega un **cupón de
descuento** para que el cliente *se quede con el producto*. El ahorro es la diferencia
entre el costo evitado (producto + flete inverso) y el costo del cupón regalado.

### Archivos de origen (carpeta `Consolidadores/`)
**1. `Cupones BOT - Keep It 2023.xlsx`** — generador/fuente de cupones (un BOT emite
los códigos). Una hoja por año: `2023` (881), `2024` (2.025), `2025` (2.422),
`2026` (1.250) + auxiliares (`solicitudes cupones`, `Cupones Especiales`, `Hoja1`,
`carga oct`). Columnas: `Codigo`, `Monto_Bot` (timestamp emisión), `Monto`,
`Compra Min`, `Caso SF`, `Ejecutivo`, `Equipo`, `Rut Cliente`, `Reserva`.
⚠️ El formato de código cambió en 2025: `KEEPIT20XX_##M-xxxx` → `CL_C_KEEP##M###-xxxx`.

**2. `DATA KEEP IT 2024.xlsx`** — consolidador analítico (alimenta el panel). 4 hojas.
Aunque se llama "2024", los datos llegan hasta **abril 2026** (sigue creciendo).

| Hoja | Filas | Cols | Rol |
|---|---|---|---|
| `Cupones Keep it` | 5.039 | 35 | Cupones entregados (consolidado del BOT) + enriquecido |
| `Casos Keep it`   | 33.225 | 43 | Todos los casos/reclamos de Salesforce (con o sin cupón) |
| `SKU`             | 234.997 | 6 | Maestro de productos (cod SKU→nombre) + tabla mes nº→nombre (E:F) |
| `Ahorro`          | 4.649 | 30 | Cálculo financiero del ahorro por cupón |

### Cruce de información (llaves) — verificado con las fórmulas reales del Excel
- **Llave principal = Número de caso**: `Casos.Número del caso` ↔ `Cupones.Caso SF` ↔ `Ahorro.Caso`.
  - `Ahorro.Caso ∩ Cupones.Caso SF` = 99% · `Cupones.Caso SF ∩ Casos.Número del caso` = 2.550 (conversión Keep It: no todo caso recibe cupón).
- **Maestro SKU** se usa vía VLOOKUP para el nombre del producto en las 3 hojas:
  - Cupones `AG` = `VLOOKUP(SKU; SKU!A:B; 2)` ; Casos `S/AQ` = `VLOOKUP(SKU; SKU!A:B; 2)`.
- Casos trae el cupón: `AG/AP` = `VLOOKUP(Nº caso; 'Cupones Keep it'!J:K; 2)` → `Cúpon`, `tiene keep it`.
- Casos `Nombre mes` = `VLOOKUP(mes; SKU!E:F; 2) & " - " & YEAR(fecha)`.
- `Año`/`Mes` se derivan con `YEAR()`/`MONTH()` sobre la fecha de entrega/apertura.
- Reserva es llave secundaria débil (`Cupones.Número de reserva ∩ Casos.Número de reserva 2` ≈ 389).

### Fórmulas de cálculo del ahorro (reversadas; columnas pegadas como valores)
- `Perdida costo  = Precio Costo / 2`                                  (100% exacto)
- `Ahorro         = Perdida costo − Costo cupón + Logística inversa`   (~95% exacto)
- `Costo cupón    ≈ Monto cupón × 0,7` (proporción típica)
- Totales del consolidado: Ahorro ≈ $263.288.880 · Monto cupones $122.925.000 · Costo cupones $91.868.000.

### Dimensiones para gráficas (hojas `G_*` futuras)
`Tipo Reclamo` (Producto en mal estado, Pieza faltante, Incompleto, No corresponde…),
`Familia_1/Subfamilia` producto, `Equipo` (BO domina), `Origen del caso` (Portal,
Teléfono, Chat…), `Estado` del caso, `Nivel 2` (Cambios/Devoluciones/Entrega),
`Rango precio`, `Portable/No portable`, `Mes/Año`, `Monto` del cupón.

---

## Fuentes de actualización del panel (insumos por hoja)

El panel ya existe (`DATA KEEP IT 2024.xlsx`); el objetivo es **mantenerlo actualizado**
alimentándolo desde estos insumos.

### Origen de cada hoja
- **Casos Keep it** ← `Informe Productos-Reserva Keep-it-*.xlsx` (export crudo de
  Salesforce, MENSUAL). Cabecera en la **fila 15**; filas 1-14 son metadatos/filtros.
  Misma estructura de 33 columnas que la hoja. El export trae "Todos los casos".
  - **FILTRO de la hoja (confirmado con datos):** `Nombre Tipificación` ∈
    {Producto en mal estado, Producto incompleto, Pieza faltante, Producto no
    corresponde, Producto con falla de funcionamiento/técnica, Producto Pieza Rota,
    Producto con empaque deteriorado}. (NO se filtra por el flag==1.)
  - **DEDUPE:** por `Número del caso` → ~1 fila por caso (99,6%; solo 116 casos con
    2-4 filas por multi-producto).
  - `Caso Keep It` (numérico **1/0**): se auto-marca con esas tipificaciones; es un
    SUBCONJUNTO de la hoja (18.617 con 1, 14.608 con 0 — todos con tipificación Keep It).
  - `Acción de Cierre` = decisión del cliente: `Si Acepta` (se queda → cupón) /
    `No Acepta, <razón>` (rechaza). Conversión histórica ≈ **33%**.
- **Cupones Keep it** ← `Cupones BOT - Keep It 2023.xlsx`, una hoja por año. Aporta el
  `Monto` del cupón.
- **Ahorro** ← receta validada (94%):
  `Ahorro = Precio Costo/2 − Costo cupón + Logística inversa`, donde:
  - `Precio Costo` = `Precio Unitario ($)` del archivo de costos `c240543e-...xlsx`
    (lookup por SKU corto; ratio 1.008 = idéntico).
  - `Costo cupón` = `Monto` (del BOT) × **0,7** (factor estándar).
  - `Logística inversa` = tarifa MENSUAL fija (en 2026 = **15.847**).
  - ⚠️ El archivo de costos es **snapshot de precios actual** (no histórico) → reconstruye
    meses pasados solo ~66% dentro de ±20%; para meses nuevos sirve bien.
  - ⚠️ Cobertura del archivo de costos = **33%** de los SKU históricos de Ahorro
    (27.579 SKU). Falta un extracto de costos más completo para el resto.
- **Familia / Sub Familia** (en las 3 hojas) ← `2279dfef-...xlsx` = **maestro de
  jerarquía** (136.660 filas: `SKU | SKU F.COM | Familia | Sub Familia | Grupo`).
  (El archivo de costos `c240543e` trae las mismas 5 columnas + costo.)
  Resuelve los `#N/D` que antes venían de un VLOOKUP externo perdido.
  - Llaves: **Cupones/Casos se unen por `SKU F.COM`** (numérico largo, ~80% cobertura);
    **Ahorro por `SKU`** (código corto con X, 99%).
  - Los valores traen prefijo `"NN - "` (ej `"23 - Jardin"`) → quitar el prefijo.

### Estado de actualización (al 2026-06-23)
- Cupones: hasta 30-abr-2026 → falta **may + jun 2026** (el BOT 2026 llega al 22-jun ✓).
- Casos: hasta 30-abr-2026 → falta **may + jun**; export disponible solo **mayo**
  → falta el **export de junio** (usuario descargando).
- Ahorro: hasta 07-mar-2026 → falta **mar(resto)+abr+may+jun**; ya hay receta + costos
  (con caveats de snapshot y cobertura 33%).

## Script de actualización — `procesar_keepit.py`

Automatiza el proceso manual del panel. Lee los insumos de `Consolidadores/` y genera
un archivo NUEVO `DATA KEEP IT 2024_actualizado_<fecha>.xlsx` (no toca el original),
agregando a las 3 hojas las filas de los meses nuevos.

- **Casos**: combina los exports `Informe Productos-Reserva*` (may+jun 2026), filtra por
  tipificación: los **6 motivos** Keep It SIEMPRE
  (Producto en mal estado, incompleto, Pieza faltante, falla técnica, Pieza Rota,
  empaque deteriorado) + **"Producto no corresponde" SOLO si `Caso Keep It == 1`**
  (regla del usuario; sin esto inflaba ~1.500 casos/2 meses). Deduplica por `Número del caso`,
  excluye casos ya en el panel, convierte fechas (texto d/m/a → datetime) y agrega
  Familia/Subfamilia (jerarquía por SKU F.COM), Portable, Mes, Nombre mes.
- **Cupones**: del BOT hoja `2026`, los códigos no presentes en el panel. ⚠️ El BOT 2026
  NO trae SKU ni Monto → el **SKU se saca del caso** (`Caso SF`→export/Casos) y el
  **Monto se parsea del código** (`...10M...` → 10.000).
- **Ahorro**: una fila por cupón nuevo con la receta validada
  `Ahorro = Precio Costo/2 − Monto×0,7 + Logística inversa`. `Precio Costo` =
  `Precio Unitario ($)` del costo `ad4a57bf` por SKU F.COM + mes. `Num Cupon` continúa
  la secuencia. Parámetros arriba del script (`FACTOR_PERDIDA`, `FACTOR_COSTO_CUPON`,
  `LOGISTICA_INVERSA`) — **confirmar con negocio**.

Cobertura: ~80% de cupones nuevos quedan con costo→ahorro; el resto sin costo en Qlik
(quedan `#N/D`, igual que en el proceso manual original).

### Análisis exhaustivo de columnas de `Ahorro` (2026-06-24)
Las 30 columnas se clasifican así (por qué algunas quedan vacías en filas nuevas):
- **Calculadas (núcleo):** Precio Costo, Perdida costo, Costo cupón, Logística inversa,
  Ahorro, Monto, Año, Mes → llenas al 80-100% (el 20% sin Precio Costo es límite de datos).
- **Derivables (se llenan):** Tipo Reclamo (=tipificación del caso), Categoria
  (="Liquidación"), NCR (="cupón"), Margen (=0,30 def.), tipo producto (="Otras"),
  Categoria/Sub familia producto (jerarquía).
- Donde falla un lookup (sin costo/familia/SKU), las columnas afectadas muestran
  `#N/D` (texto), igual que el original — no se dejan vacías. Columnas: SKU, Costo Prod.,
  Precio Costo, Perdida costo, Tipo Reclamo, Categoria/Sub familia producto, Ahorro.
- **NO reproducibles (van como `#N/D`, honesto):**
  - `Costo Prod.` — costo independiente de fuente Qlik desconocida; NO entra en la fórmula.
  - `Rango precio`, `Rango precio piso`, `Rango precio linea blanca` — **están ROTOS en
    el original** (las etiquetas de rango no corresponden al Precio Costo; se solapan).
  - `Nombre`, `RSVA`, `Documento`, `Equipo` — vacías también en el original (0%).
- **Límite de cobertura de costo (80%):** de los cupones sin costo, ~⅔ son casos sin
  SKU/producto y ~⅓ son SKU que no existen en ningún extracto de costo de Qlik
  (sellers/marketplace). No recuperable; idéntico a la brecha del proceso manual.

## Registro de Cambios

### 2026-06-23 — Script de actualización + primera corrida
- Creado `procesar_keepit.py`. Corrida hasta hoy:
  Casos 33.225→37.805 (+4.580 may+jun), Cupones 5.039→5.373 (+334),
  Ahorro 4.649→4.983 (+334; ahorro nuevo ≈ $17,26M). Filas viejas intactas; fórmula
  de ahorro validada (94,9% sobre filas viejas; aceptación 33% igual al histórico).
  Regla afinada: "Producto no corresponde" solo entra con flag Keep It = 1.

### 2026-06-23 — Análisis de fuentes y plan de actualización
- Analizados `DATA KEEP IT 2024.xlsx`, `Cupones BOT - Keep It 2023.xlsx`,
  `Informe Productos-Reserva Keep-it-*.xlsx` (export Salesforce mensual, mayo 2026) y
  `2279dfef-...xlsx` (maestro de jerarquía de productos).
- Documentado modelo de datos, llaves de cruce (Nº de caso + SKU/SKU F.COM vía VLOOKUP),
  fórmula de ahorro (`Perdida costo − Costo cupón + Logística inversa`), origen de cada
  hoja y gaps de fechas.
- Pendiente del usuario: regla exacta del filtro/dedupe de `Casos Keep it`, fuente de
  costos para `Ahorro`, y export de junio 2026.

### 2026-06-22 — Setup inicial
- Creado el repositorio con el andamiaje genérico (`.gitignore`, `requirements.txt`,
  `CLAUDE.md`, `README.md`, `prompt.txt`, carpeta `graficas_looker/`).
- Entorno virtual `venv` con pandas, openpyxl, flask, reportlab.
- Pendiente: definir las fuentes de datos y el primer bloque de procesamiento.
