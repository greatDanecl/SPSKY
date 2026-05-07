# FDS Portal · Flight Data System

Dashboard de productividad de tripulación para pilotos. Se actualiza automáticamente cada vez que se sube un nuevo archivo xlsx a la carpeta `/data`.

## Configuración inicial (una sola vez)

### 1. Activar GitHub Pages

1. Ve a **Settings → Pages**
2. En **Source** selecciona **GitHub Actions**
3. Guarda los cambios

### 2. Estructura de carpetas

```
fds-pilotos/
├── data/              ← Aquí van los xlsx mensualmente
├── parser.py          ← Script que genera el dashboard
├── requirements.txt
├── .github/
│   └── workflows/
│       └── update.yml
└── index.html         ← Se genera automáticamente (no editar)
```

## Flujo mensual

1. Recibes el xlsx del mes (programado y/o efectuado)
2. Lo arrastras a la carpeta `data/` en GitHub (botón **Add file → Upload files**)
3. Haces commit (botón verde **Commit changes**)
4. GitHub Actions ejecuta el parser automáticamente (~2 minutos)
5. El dashboard en `https://TU-USUARIO.github.io/fds-pilotos` queda actualizado

## Convención de nombres para los archivos xlsx

El parser detecta el período desde el contenido del archivo, no del nombre. Sin embargo, para mantener orden se recomienda:

```
YYYY-MM_tipo_base.xlsx

Ejemplos:
2026-04_programado_SCL.xlsx
2026-04_efectuado_SCL.xlsx
2026-04_efectuado_Horas_PMC.xlsx
```

## Criterio de exclusión del promedio comparativo

Un piloto queda excluido del cálculo del promedio de su cargo en un mes determinado si:
- Más del 35% de sus días registrados son vacaciones o licencia médica
- Sus horas de bloque efectuadas son menores a 5h

Los datos del piloto **se siguen mostrando** en el gráfico (con marcador triangular), pero no afectan el promedio de referencia del cargo.

## Límites regulatorios monitoreados (DAN 121 / Ley 20.321)

| Indicador | Umbral alerta | Límite máximo |
|---|---|---|
| Horas bloque mensual | 85h | 100h |
| Horas bloque anual | 750h acum. | 1.000h |
| Días libres mensual | < 10d | Mínimo 8d |
| Horas deber mensual | > 105h | Según FDP |

## Correr localmente (opcional)

```bash
pip install -r requirements.txt
python parser.py
# Abre index.html en tu navegador
```
