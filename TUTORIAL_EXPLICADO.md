# Tutorial explicado: qué hace el programa

## El recorrido completo

```text
archivo de entrada
  ├─ audio o video → Whisper produce texto
  └─ TXT o MD      → Python lee el texto directamente
                         ↓
              GPT-OSS analiza la clase
                         ↓
       idioma hablado + idioma enseñado + varios temas
                         ↓
          carpetas por idioma + Excel consolidado
```

## 1. La carpeta `entrada`

El programa busca MP3, WAV, M4A, MP4, FLAC, OGG, AAC, MPEG, WEBM, TXT y MD.
Reconoce el tipo por la extensión. No funciona solo por copiar un archivo: debes
hacer doble clic en `Clasificar clases de idiomas.cmd` para iniciar la revisión.

## 2. Transcripción: `transcribir()`

Solo se ejecuta para audio o video. Python abre el archivo y lo envía mediante la
biblioteca `groq` al modelo `whisper-large-v3-turbo`. Whisper devuelve el texto y
una estimación del idioma hablado. No hay un prompt largo en este paso: la orden
está en la llamada `cliente.audio.transcriptions.create(...)`.

Para TXT o MD se usa `leer_texto()`, por lo que no se llama a Whisper.

## 3. Análisis: `analizar_clase()`

Aquí sí existe un prompt. Le pide a `openai/gpt-oss-20b` que devuelva JSON con:

```json
{
  "idiomas_hablados": ["en"],
  "idiomas_ensenados": ["tr"],
  "temas": [
    {"idioma": "tr", "tema": "verbo estar"},
    {"idioma": "tr", "tema": "adjetivos posesivos"}
  ],
  "resumen": "Resumen breve en español"
}
```

La distinción evita el error anterior de confundir una explicación en inglés con
una clase de inglés. Los temas no están limitados a una lista fija: el modelo los
extrae del contenido y puede devolver varios.

Los idiomas tampoco están limitados a una lista cerrada. El modelo devuelve un
código ISO-639-1. Python lo convierte a un nombre en español cuando lo conoce y
crea automáticamente la carpeta correspondiente. Si el código aún no tiene un
nombre configurado, lo utiliza directamente como nombre de carpeta y en el Excel;
el procesamiento continúa normalmente.

`normalizar_idioma()` valida los códigos y corrige errores frecuentes del modelo:
`gr` se convierte en `el`, `jp` en `ja` y `sp` en `es`. Si recibe un código que no
existe, devuelve `desconocido` en vez de crear una carpeta incorrecta.

`llamar_groq_con_reintentos()` protege tanto la transcripción como el análisis. Si
Groq responde con un límite `429`, un error temporal de servidor o un problema de
conexión, lee el tiempo de espera indicado por Groq y reintenta hasta tres veces.
El análisis también reintenta si el modelo genera JSON inválido.

Si la transcripción supera 24.000 caracteres, `abreviar_clase()` conserva el
principio, el centro y el final para el análisis y muestra un aviso visible. El TXT
completo no se recorta. El Excel registra si la cobertura fue completa o recortada.

## 4. Control de repetidos

`firma()` combina ruta, tamaño y fecha de modificación. `hash_contenido()` calcula
una huella del contenido. Con ambas, `main()` puede:

- omitir un archivo ya procesado que no cambió;
- detectar una copia idéntica con otro nombre;
- reutilizar una transcripción existente cuando solo cambia la forma de analizarla.

El registro vive en `salida/_datos/estado.json`.

## 5. Carpetas y Excel

`copiar_por_idioma()` conserva una copia del original en
`salida/clases/<idioma enseñado>`. Si una clase enseña dos idiomas, aparece en las
dos carpetas. Los temas no crean más subcarpetas, porque un video puede tener
muchos; quedan consolidados en `clases.xlsx`.

`crear_excel()` genera tres hojas:

1. **Resumen:** un idioma por fila, cantidad de clases, lista de temas y gráfico.
2. **Clases:** una fila por clase, incluida la cobertura del texto analizado.
3. **Temas por idioma:** una fila por tema, ideal para filtros y búsquedas.

La versión anterior mostraba una “confianza” calculada por el mismo modelo. Se
eliminó porque no era una medida validada y casi siempre devolvía valores muy
parecidos.

El diseño usa encabezados oscuros, acentos turquesa, tablas con filtros y columnas
ajustadas para lectura.

## 6. Dónde está la clave y qué sale del equipo

`.env` contiene `GROQ_API_KEY`. `python-dotenv` la carga y `groq` la usa para
autenticar las solicitudes. La clave no está escrita dentro del programa.

| Elemento | Permanece local | Se envía a Groq |
|---|---:|---:|
| Original en `entrada` | Sí | Sí, solo si requiere transcripción |
| TXT/MD | Sí | Su contenido para analizar |
| Transcripción | Sí | Sí, para analizar |
| Excel, JSON y estado | Sí | No |

Revoca la clave desde la consola de Groq cuando termines la prueba. Si OneDrive
sincroniza esta carpeta, `.env` también puede subirse a tu OneDrive; no compartas
la carpeta y elimina o revoca la clave al terminar.

## 7. Ejecución normal

Haz doble clic en `Clasificar clases de idiomas.cmd`. El lanzador:

1. entra en la carpeta correcta;
2. ejecuta `.venv\Scripts\python.exe`;
3. muestra el progreso;
4. abre `salida\clases.xlsx` si todo terminó correctamente.

Para forzar un análisis nuevo de todos los archivos:

```powershell
python clasificar_clases_idiomas.py --reprocesar
```

Normalmente no necesitas esa opción: agrega archivos nuevos y vuelve a usar el
icono; los anteriores se omiten.
