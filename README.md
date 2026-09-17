# Clasificador de clases de idiomas con Groq

Prototipo que transcribe clases de idiomas (audio, video o texto), detecta qué
idioma se enseña y extrae los temas de cada clase, con los resultados en un Excel.
Es una prueba de concepto para, más adelante, clasificar otro tipo de documentos,
como entrevistas.

Coloca audios, videos, TXT o Markdown en `entrada` y haz doble clic en
`Clasificar clases de idiomas.cmd`.

> **Nota:** el repositorio no incluye audios de ejemplo ni resultados. Usa tus
> propios archivos; la carpeta `entrada` se crea sola en la primera ejecución.

El programa distingue dos cosas diferentes:

- **Idioma hablado:** el idioma usado por el profesor para explicar.
- **Idioma enseñado:** el idioma que está aprendiendo el estudiante.

También identifica varios temas concretos por clase. Por ejemplo, un video puede
estar hablado en inglés, enseñar turco y contener los temas «verbo estar»,
«adjetivos posesivos» y «presentaciones».

## Qué hace con cada archivo

1. Si es audio o video, Groq ejecuta `whisper-large-v3-turbo` para transcribirlo.
2. Si es TXT o MD, Python lee el contenido sin transcribirlo.
3. Groq ejecuta `openai/gpt-oss-20b` para identificar idiomas, temas y resumen.
4. Python copia el original a una carpeta según el **idioma enseñado**.
5. Python actualiza `salida/clases.xlsx`.

Los originales de `entrada` nunca se mueven ni se eliminan. Si un archivo no ha
cambiado, no se procesa de nuevo. Las copias idénticas se reconocen aunque tengan
nombres distintos.

## ¿Qué sucede si aparece un idioma nuevo?

No hace falta configurar una categoría ni crear la carpeta manualmente. El modelo
devuelve el código internacional ISO-639-1 del idioma enseñado y Python crea la
carpeta durante esa misma ejecución.

El código se valida antes de usarlo. Errores habituales del modelo se corrigen,
por ejemplo `gr` → `el` para griego, `jp` → `ja` para japonés y `sp` → `es` para
español. Un código inventado no crea una carpeta equivocada: queda como
`sin_identificar` para revisión.

Por ejemplo, si aparece una clase de sueco, se crea `salida/clases/sueco`. El
programa ya tiene nombres en español para numerosos idiomas habituales. Si recibe
un código válido que todavía no está en esa lista, no falla: utiliza temporalmente
el código como nombre de carpeta —por ejemplo, `sw`— y lo muestra igual en el
Excel. Después podemos agregar su nombre en español sin volver a transcribir.

## Resultado

```text
salida/
  clases.xlsx
  clases/
    aleman/
    frances/
    griego/
    italiano/
    turco/
  _datos/
    transcripciones/
    detalles/
    estado.json
```

`clases.xlsx` contiene:

- **Resumen:** cantidad de clases y temas consolidados por idioma enseñado.
- **Clases:** una fila por archivo, con idioma hablado, idioma enseñado, todos los
  temas, resumen y cobertura del texto analizado.
- **Temas por idioma:** una fila por cada tema detectado, útil para filtrar.

Las carpetas `_datos` son el registro interno del programa. Los temas no dependen
de un archivo de categorías: se extraen libremente de cada clase.

La columna **Cobertura del texto** indica `Completo` o informa cuántos caracteres
se analizaron cuando una transcripción fue demasiado larga. La transcripción
completa siempre se conserva en `_datos/transcripciones`.

## Instalación en otro PC con Windows

### 1. Requisitos

El nuevo computador necesita:

- **Python 3.12**. Durante su instalación marca la opción **Add Python to PATH**.
- Conexión a Internet para comunicarse con Groq.
- Una cuenta y una clave de API de Groq.
- Excel, LibreOffice o Google Sheets únicamente si quieres abrir el resultado
  `clases.xlsx`. Microsoft Excel no es obligatorio.

No necesitas instalar Whisper, Ollama ni FFmpeg: los modelos se ejecutan en los
servidores de Groq.

### 2. Descargar el proyecto

En GitHub, usa **Code > Download ZIP** y descomprímelo, o clónalo con Git:

```powershell
git clone <URL-de-este-repositorio>
```

Si copias el proyecto desde otro computador, no copies `.venv` (es propio de cada
PC), `.env` (contiene la clave privada) ni `salida` (resultados anteriores).

### 3. Crear el entorno virtual

Abre PowerShell y entra en la carpeta donde copiaste el proyecto. Sustituye la ruta
del ejemplo por la ruta real del nuevo computador:

```powershell
Set-Location -LiteralPath "C:\ruta\al\clasificador-clases-idiomas"
python --version
python -m venv .venv
Set-ExecutionPolicy -Scope Process Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Al activarse correctamente aparece `(.venv)` al comienzo de la línea. El archivo
`requirements.txt` instala automáticamente:

- `groq`, para comunicarse con los modelos de Groq;
- `python-dotenv`, para leer la clave desde `.env`;
- `openpyxl`, para crear el Excel consolidado.

### 4. Configurar la clave de Groq

Crea una cuenta gratuita en <https://console.groq.com> y genera una clave en
<https://console.groq.com/keys>. Luego crea el archivo privado `.env` a partir del ejemplo:

```powershell
Copy-Item .env.example .env
notepad .env
```

Dentro de `.env` guarda:

```text
GROQ_API_KEY=tu_clave_real
```

Es preferible crear una clave diferente para cada computador. Para comprobar que
Python puede leerla:

```powershell
python -c "from dotenv import load_dotenv; import os; load_dotenv(); print('Clave cargada:', bool(os.getenv('GROQ_API_KEY')))"
```

Debe aparecer `Clave cargada: True`. Esta prueba solo comprueba que existe una
clave; no consume una transcripción.

### 5. Primera ejecución

Coloca un archivo corto y no sensible en `entrada` y haz doble clic en
`Clasificar clases de idiomas.cmd`. La carpeta `salida` se crea automáticamente.

No publiques ni compartas `.env`. Los audios y textos se procesan temporalmente en
Groq. Para material sensible, revisa y activa **Zero Data Retention** en los
controles de datos de la cuenta.

### Problemas de instalación

- Si PowerShell dice que `python` no existe, cierra y vuelve a abrir PowerShell.
  Si continúa, reinstala Python marcando **Add Python to PATH**.
- Si no aparece `(.venv)`, vuelve a ejecutar la línea de activación.
- Si falta una biblioteca, activa `.venv` y repite
  `python -m pip install -r requirements.txt`.
- Si aparece `Invalid API Key`, revisa `.env` o crea una clave nueva en Groq.

## Uso

La forma normal es hacer doble clic en `Clasificar clases de idiomas.cmd`. El lanzador usa
el Python aislado de `.venv` y abre el Excel al terminar.

Desde PowerShell también puedes ejecutar:

```powershell
python clasificar_clases_idiomas.py
```

Opciones:

```powershell
python clasificar_clases_idiomas.py --reprocesar
python clasificar_clases_idiomas.py --sin-copiar
python clasificar_clases_idiomas.py --idioma en
```

`--idioma` indica el idioma hablado esperado para Whisper; no indica el idioma
enseñado. En uso normal déjalo en detección automática.

Los límites temporales y errores de conexión se reintentan tanto durante la
transcripción con Whisper como durante el análisis. Cuando Groq indica cuánto
esperar, el programa usa ese tiempo en lugar de una pausa fija. Si el límite tarda
más en restablecerse, vuelve a ejecutar el icono: los archivos terminados se
omitirán y solo se intentarán los pendientes.
