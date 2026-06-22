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

## Registro de Cambios

### 2026-06-22 — Setup inicial
- Creado el repositorio con el andamiaje genérico (`.gitignore`, `requirements.txt`,
  `CLAUDE.md`, `README.md`, `prompt.txt`, carpeta `graficas_looker/`).
- Entorno virtual `venv` con pandas, openpyxl, flask, reportlab.
- Pendiente: definir las fuentes de datos y el primer bloque de procesamiento.
