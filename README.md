# Evaluador de recuperación de información con validación humana

Proyecto de la Actividad 5 del módulo de IA Generativa del programa Chief Digital Officer – MIT Professional Education.

**Autor:** Danny Salazar

## Qué hace

Le hice preguntas sobre tenis a ChatGPT con tres manuales como fuente, y quise comprobar si sus respuestas salían de esos documentos o de su conocimiento general. Este script revisa cada respuesta contra la fuente de verdad en tres pasos:

1. **Sugerir**: busca en los PDF el pasaje que sustenta cada respuesta, me dice en qué manual y página está, y verifica que la cita exista textual en el documento.
2. **Validar**: yo reviso cada sugerencia contra el PDF y la marco como verdadera. La validación humana es la que convierte la sugerencia en fuente de verdad.
3. **Evaluar**: un modelo juez (Claude) califica de 1 a 5 cada respuesta de ChatGPT:
   - **Recuperación** (`retrieval_score`): ¿el fragmento sirve para responder la pregunta?
   - **Sustento** (`groundedness`): ¿lo que dijo ChatGPT está respaldado por el documento?
   - **Relevancia** (`relevance`): ¿cubre lo esencial frente a una respuesta de referencia construida solo con el documento?

Inspirado en el flujo [eval-qna-rag-metrics](https://github.com/johntango/PromptFlowEval) del profesor Williams, implementado a menor escala en Python sin Prompt Flow ni Azure.

## Archivos

| Archivo | Contenido |
|---|---|
| `eval_rag.py` | El script |
| `tenis_eval.xlsx` | Hoja **Fuentes** (los manuales) y hoja **Preguntas** (preguntas, respuestas de ChatGPT y validación) |
| `resultados.xlsx` | Puntajes del juez, resumen por fuente y razones de cada calificación |

## Fuentes de verdad

Los PDF no se incluyen en el repositorio porque son documentos de terceros.

| Documento | Autor |
|---|---|
| Reglas del Tenis 2020 | International Tennis Federation (ITF) |
| Manual de Tenis | Alejo D. Russell |
| Guía de Entrenamiento de Tenis | Special Olympics |

## Cómo correrlo

```
py -m pip install anthropic python-dotenv openpyxl pypdf
```

Crear un archivo `.env` con la clave de la API de Claude:

```
ANTHROPIC_API_KEY=sk-ant-...
JUDGE_MODEL=claude-haiku-5-5
```

Poner los PDF en la misma carpeta y correr:

```
py eval_rag.py sugerir      # busca el sustento en los PDF
                            # (validar en el Excel: SI en la columna "validado")
py eval_rag.py evaluar      # califica las filas validadas -> resultados.xlsx
```

Opciones útiles: `--dry-run` prueba el flujo sin llamar a la API, `--ids q1,q3` procesa solo esas preguntas y `--debug` guarda los fragmentos que vio el modelo.

## Hallazgos principales

- Una respuesta correcta no es lo mismo que una respuesta sustentada: ChatGPT acertó el nombre del golpe (revés con slice), pero no salió de los documentos.
- La automatización también se equivoca: el buscador no encontró al inicio un caso del reglamento por diferencias de vocabulario ("tierra" vs. "terreno"), y la primera versión del juez dejó calificaciones vacías.
- La calidad de la validación humana pesa: un fragmento copiado incompleto bajó el puntaje de sustento de una respuesta correcta.
