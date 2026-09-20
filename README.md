# Sistema de Automatización de Reportes e Incidencias

App en Python con Streamlit para armar reportes. La armé para no perder tiempo formateando y corrigiendo texto a mano. Escribes, la IA te lo pule, y sale el reporte listo para copiar o descargar en Excel.

## Stack

| Parte | Qué usa |
|---|---|
| Interfaz | Streamlit |
| IA principal | DeepSeek (`deepseek-chat`) |
| IA respaldo | OpenRouter (`inclusionai/ling-3.0-flash`) |
| Clima | Windy API, endpoint `point-forecast/v2` |
| Excel | OpenPyXL |
| HTTP | Requests |
| Persistencia | JSON local |

La IA se llama con el mismo cliente de OpenAI. Solo cambia `base_url` y la key según el proveedor. Si DeepSeek falla, cae a OpenRouter, y si los dos fallan, devuelve el texto original con un warning.

## Módulos

**Reporte Diario** — encabezado, responsable, área y detalle. El botón de IA reescribe el texto sin quitarle información. Descarga a Excel o TXT.

**Reporte de Incidentes** — ubicación jerárquica (región → ciudad → zona → barrio), coordenadas, y consulta de clima a Windy. Parámetros que pide a Windy: `wind-surface`, `temp-surface`, `rh-surface`, `pressure-surface`, `precip-surface`. Aparte, tiene un botón para generar un reporte ejecutivo (resumen corto hecho con IA a partir de la reseña y las acciones).

**Reportes Varios** — tres pestañas: notas informativas, reporte meteorológico y un resumen que lee directo de `memoria.json`.

## Instalación

Python 3.10 o superior.

```bash
git clone https://github.com/rahudy03/sistema-automatizacion-reportes-incidencias.git
cd sistema-automatizacion-reportes-incidencias

python -m venv venv

# Windows
venv\Scripts\activate

# Linux / Mac
source venv/bin/activate

pip install -r requirements.txt
```

## Configuración

Crear `.streamlit/secrets.toml`:

```toml
WINDY_API_KEY = "..."
DEEPSEEK_API_KEY = "..."
OPENROUTER_API_KEY = "..."
```

Claves:
- Windy → https://api.windy.com/
- DeepSeek → https://platform.deepseek.com/
- OpenRouter → https://openrouter.ai/

El archivo está en `.gitignore`, no se sube. Hay un `secrets.toml.example` de plantilla.

## Correr

```bash
streamlit run app.py
```

`http://localhost:8501`

## Estructura

```
.
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
├── LICENSE
└── .streamlit/
    ├── secrets.toml          # ignorado por git
    └── secrets.toml.example  # plantilla
```

## Persistencia

Sin base de datos. Dos archivos JSON que se crean solos:

- `memoria.json` — guarda el último valor de cada campo del formulario
- `ubicaciones.json` — estructura de regiones, ciudades, zonas y barrios

Si se borran, se regeneran con los valores por defecto.

## Helpers

Para que la app no se caiga si un JSON queda mal:

- `int_seguro(valor, default)` — convierte a int, devuelve default si falla
- `float_seguro(valor, default)` — igual pero para float
- `bool_seguro(valor, default)` — acepta bool, números, y strings tipo `"true"` o `"si"`
- `normalizar_ubicaciones(ubicaciones)` — revisa que cada ciudad tenga sus claves (`zonas`, `barrios`, `sub_barrios`) y las crea si faltan

## Excel

Una función por módulo:

- `excel_reporte_diario(datos)`
- `excel_reporte_incidente(datos)`
- `excel_reporte_mixto(titulo, secciones)`

Devuelven un `BytesIO` que se pasa directo al `st.download_button`. El formato se arma desde código con OpenPyXL: encabezados con fondo, filas alternadas, bordes finos, anchos de columna calculados.

## Licencia

Todos los derechos reservados. Para uso comercial, escribir al correo de abajo.

## Contacto

Rahudy Osorio
rahudy03@gmail.com
GitHub: https://github.com/rahudy03
