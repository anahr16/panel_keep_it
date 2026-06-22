# Panel Keep It

Panel de datos: procesa exports/Excel y genera hojas planas para Looker Studio
(sin blends) y/o un panel web Flask propio.

## Requisitos
- Python 3.10+

## Instalación

### WSL2 / Linux
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Windows
```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Estructura
- `<bloque>/input/` · `<bloque>/output/` · `<bloque>/procesar_*.py` — módulos de procesamiento
- `graficas_looker/` — genera las hojas `G_*` y el `G_para_looker.xlsx` para subir a Drive/Looker

## Documentación
Todos los cambios y decisiones se registran en `CLAUDE.md`.
