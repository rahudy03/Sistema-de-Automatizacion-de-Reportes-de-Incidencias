# Sistema de Automatización de Reportes e Incidencias

Aplicación web desarrollada en Python sobre Streamlit para la generación asistida de reportes operativos. El sistema integra modelos de lenguaje para corrección y reescritura de texto, consulta meteorológica en tiempo real y exportación a formatos estructurados (XLSX, TXT).

## Arquitectura

La app es single-file (`app.py`) con persistencia en archivos JSON planos. No usa base de datos. El estado entre sesiones se conserva mediante un archivo de memoria (`memoria.json`) que guarda el último valor introducido en cada campo del formulario, más un archivo de estructura de ubicaciones (`ubicaciones.json`).

El flujo de ejecución sigue el patrón estándar de Streamlit: el script se re-ejecuta completo en cada interacción, y el estado se mantiene en `st.session_state` y en los JSON de persistencia.

## Stack

| Capa | Implementación |
|---|---|
| UI | Streamlit (widgets nativos + CSS inyectado con `st.markdown`) |
| LLM primario | DeepSeek API — modelo `deepseek-chat` |
| LLM fallback | OpenRouter API — modelo `inclusionai/ling-3.0-flash` |
| Cliente HTTP LLM | SDK oficial de OpenAI (`openai>=1.0.0`), con `base_url` custom |
| Clima | Windy Point Forecast API v2 (`POST /api/point-forecast/v2`) |
| Excel | OpenPyXL — generación programática sin templates |
| HTTP | Requests |
| Runtime | Python 3.10+ |

## Integración con IA

La función `mejorar_redaccion_ia(texto, tipo_texto)` implementa un patrón de cascada sobre dos proveedores:

1. Intenta con DeepSeek (`https://api.deepseek.com`, modelo `deepseek-chat`).
2. Si falla, intenta con OpenRouter (`https://openrouter.ai/api/v1`, modelo `inclusionai/ling-3.0-flash`).
3. Si ambos fallan, retorna el texto original capitalizado con un warning.

El `max_tokens` se ajusta según el tipo de texto (`reporte`, `resumen` y `nota` usan 1000; el resto 500). La temperatura está fijada en 0.3 para reducir variabilidad.

Cada tipo de texto tiene su propio prompt de sistema. Por ejemplo, `acciones realizadas` espera formato `HH:MM Hrs descripción`, mientras que `resumen` fusiona reseña y acciones en un solo párrafo narrativo en pasado y tercera persona.

## Integración con Windy

`consultar_clima_windy(lat, lon)` hace un POST al endpoint `point-forecast/v2` con los parámetros `wind`, `temp`, `rh`, `pressure` y `precip` a nivel `surface`. La respuesta se parsea con un `safe_get` interno que valida que cada clave exista, sea lista, no esté vacía y que el primer elemento no sea `None`.

La velocidad del viento se calcula primero intentando `wind-surface`; si no está disponible, se hace la magnitud del vector a partir de `wind_u-surface` y `wind_v-surface`. La temperatura se convierte de Kelvin a Celsius. La presión se pasa de Pascales a hectopascales.

## Módulos

**Reporte Diario** — Formulario con encabezado, responsable, área y detalle. Reescritura opcional vía IA. Exportación a Excel (`excel_reporte_diario`) y TXT.

**Reporte de Incidentes** — Formulario extendido:
- Selectores jerárquicos encadenados: Región → Ciudad → Zona → Barrio, con fallback a texto libre cuando la estructura no está definida
- Entrada de coordenadas (`float`, formato `%.6f`)
- Consulta meteorológica vía Windy con persistencia del resultado en `st.session_state`
- Generación de reporte ejecutivo: aplica regex para eliminar marcas de hora (`\d{2}:\d{2}\s*Hrs`) y limpiar frases repetidas, luego pasa el texto combinado por la IA

**Reportes Varios** — Tres submódulos independientes en tabs:
- Notas informativas con reescritura IA
- Reporte meteorológico (independiente de Windy, entrada manual)
- Resumen general que lee `memoria.json` directo y lo muestra en `st.dataframe`

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

## Instalación

```bash
git clone https://github.com/rahudy03/sistema-automatizacion-reportes-incidencias.git
cd sistema-automatizacion-reportes-incidencias

python -m venv venv

# Windows
venv\Scripts\activate

# Linux / macOS
source venv/bin/activate

pip install -r requirements.txt
```

## Configuración

Archivo `.streamlit/secrets.toml`:

```toml
WINDY_API_KEY = "..."
DEEPSEEK_API_KEY = "..."
OPENROUTER_API_KEY = "..."
```

Credenciales:
- Windy → https://api.windy.com/
- DeepSeek → https://platform.deepseek.com/
- OpenRouter → https://openrouter.ai/

El archivo está listado en `.gitignore`. Se incluye `secrets.toml.example` como plantilla.

## Ejecución

```bash
streamlit run app.py
```

Puerto por defecto: `8501`.

## Persistencia

Sin motor de base de datos. Dos archivos JSON en la raíz del proyecto:

- `memoria.json` — diccionario plano `clave → valor` con el último estado de cada widget. Lectura/escritura mediante `guardar_memoria()` y `cargar_memoria()`, ambos con manejo de excepciones y fallback a default.
- `ubicaciones.json` — árbol de tres niveles (`región → ciudad → {zonas, barrios, sub_barrios}`). Se normaliza al cargar con `normalizar_ubicaciones()`.

Ambos se regeneran desde defaults si no existen.

## Helpers de conversión

Los datos que vienen de JSON pueden estar corruptos o con tipos incorrectos. Para evitar excepciones:

- `int_seguro(valor, default)` — `int()` con try/except, devuelve default en caso de fallo
- `float_seguro(valor, default)` — ídem para `float`
- `bool_seguro(valor, default)` — acepta `bool`, `str` (`"true"`, `"1"`, `"si"`, `"sí"`, `"yes"`) e `int`/`float`
- `normalizar_ubicaciones(ubicaciones)` — recorre el árbol y garantiza que cada ciudad tenga las claves `zonas`, `barrios` y `sub_barrios`, y que cada zona exista como clave en `barrios`

## Exportación a Excel

Una función por módulo, cada una devuelve un `BytesIO` listo para `st.download_button`:

- `excel_reporte_diario(datos)`
- `excel_reporte_incidente(datos)`
- `excel_reporte_mixto(titulo, secciones)`

El estilo se aplica con OpenPyXL desde código:

- Paleta definida en constantes (`COLOR_HEADER`, `COLOR_ROW_ALT`, `COLOR_TITLE`, `COLOR_BORDER`, `COLOR_LABEL`)
- `_escribir_titulo()` — fila combinada con `merge_cells`, fuente bold 16, fondo verde, altura 30
- `_escribir_encabezados()` — fondo oscuro, texto blanco, alineación centrada, `wrap_text=True`
- `_escribir_fila_datos()` — bordes finos por celda, alternado opcional para efecto zebra
- `_escribir_pares_clave_valor()` — columna izquierda con fondo claro y bold, columna derecha combinada para el valor
- `_escribir_bloque_texto()` — bloque multilínea con `merge_cells` vertical y `vertical="top"`
- `_ajustar_anchos()` — anchos de columna vía `get_column_letter`

## Limitaciones conocidas

- La app es síncrona. Las llamadas a las APIs de IA y Windy bloquean la UI; se mitiga parcialmente con `st.spinner`.
- `memoria.json` solo guarda el último valor de cada campo. No hay historial de reportes generados.
- El árbol de ubicaciones se guarda completo en cada actualización. Para árboles muy grandes convendría mover a SQLite.
- No hay autenticación. Si se expone públicamente, cualquiera con la URL puede generar reportes.

## Licencia

Todos los derechos reservados. Uso comercial requiere autorización escrita.

## Contacto

Rahudy Osorio
rahudy03@gmail.com
GitHub: https://github.com/rahudy03
