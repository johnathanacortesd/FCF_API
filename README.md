# Analisis FCF

App alternativa para clasificar noticias de la Federacion Colombiana de Futbol.

## Ejecutar

```powershell
cd FCF-API
streamlit run app.py
```

La app espera un archivo `.xlsx` con el formato FCF:

`ID`, `FECHA`, `HORA`, `TIPO DE MEDIO`, `NOMBRE DE MEDIO`, `REGION`, `SECCIÓN`, `TÍTULO`, ` TAMAÑO.-.CARACTERES `, `AUDIENCIA`, ` PAGINA `, `ROI`, `Impacto`, `TEMA`, `SUBTEMA`, `VOCERO`, `LINK`, `RESUMEN`, `WEB`, `MARCA`, `MES`, `VERI`.

El texto que se analiza sale de `TÍTULO` y `RESUMEN`.

La app busca `Configuracion.xlsx` automaticamente dentro del repo y actualiza `REGION` con un buscarv usando `NOMBRE DE MEDIO` como llave. La configuracion puede tener columnas `NOMBRE DE MEDIO` y `REGION`, o una hoja `Regiones` donde la primera columna sea el medio y la segunda la region. No se debe subir este archivo desde la interfaz.

## Salidas

- `Impacto`: `Positivo`, `Negativo` o `Neutro`.
- `Tema`: uno de los temas prestablecidos de FCF.
- `Subtema`: etiqueta especifica de la noticia o grupo.
- `VOCERO`: `Ramón Jesurun` si el resumen menciona a Ramón Jesurun, Ramón Jesurun Franco, Jesurun o Ramon Jesurun; de lo contrario `Sin vocero`.

Si el resumen contiene `Foto: FCF`, `Foto:FCF`, `Foto FCF`, `Tomado de FCF` o una variante equivalente, la fila queda como `Neutro`, `Institucional` y `Foto`.

Si ni `TÍTULO` ni `RESUMEN` mencionan `FCF`, `Federación Colombiana de Fútbol`, `Federación Colombiana` ni `Ramón Jesurun`/`Jesurun`, la fila no se envia al modelo y queda como `Neutro`, `Institucional` y `Logo`. La palabra `Federación` sola no cuenta como mención válida.

La descarga conserva los hipervinculos incrustados en la palabra `Link` para las columnas `LINK` y `WEB`.

## Duplicados

La app marca como `Duplicada` en `Impacto` y `-` en `TEMA`, `SUBTEMA` y `VOCERO` cuando:

- `NOMBRE DE MEDIO` es igual y `TÍTULO` es igual o muy similar, incluyendo variaciones donde solo cambian comillas.
- El hipervinculo incrustado en `WEB` es igual en dos o mas filas.

## Subtemas de partidos

Para noticias de partidos de Colombia, Mundial, Copa del Mundo, Copa Mundo, Mundial 2026 o Eliminatorias, la app normaliza el subtema despues de la respuesta del modelo:

- Si el texto habla de Mundial, Copa del Mundo, Copa Mundo o Mundial 2026, usa `Partidos del Mundial` sin equipos.
- Si no detecta rival y el texto habla de Eliminatorias, usa `Partido Eliminatorias`.
- Variantes como `Partido Mundial`, `Partido Copa del Mundo`, `Partido Copa Mundo` y `Partido Mundial 2026` se agrupan como `Partidos del Mundial`.
