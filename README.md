# Cotizador Sellado de Fluidos — Streamlit

Primera versión de la app web basada en el cotizador Excel validado.

## Qué incluye

- Empaquetaduras planas:
  - Material
  - Tipo de corte
  - Cantidad
  - Factor/margen
  - Cantidad por plancha
  - Costo de corte
  - Costo de material por unidad
  - Precio unitario y total

- Empaquetaduras trenzadas:
  - Estilo
  - Sección dependiente del estilo
  - Kg a cotizar
  - Costo real por kg
  - Factor/margen
  - Precio de venta/kg y total
  - Código EMPTRE y stock registrado

- Cotización:
  - Cliente, número y fecha
  - Agregar múltiples productos
  - Total general
  - Descargar cotización en Excel

## Ejecutar en tu PC

1. Instala Python.
2. Abre una terminal en esta carpeta.
3. Ejecuta:

   pip install -r requirements.txt

4. Luego:

   streamlit run app.py

Streamlit abrirá la app en el navegador.

## Publicarla en internet

La carpeta está preparada para subirse a un repositorio de GitHub y desplegarse con Streamlit Community Cloud.

Documentación oficial:
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app

## Actualizar costos

La app utiliza:

Cotizador_Empaques_2026_VERSION_FINAL.xlsx

Para actualizar los costos, reemplaza ese archivo por la versión actualizada manteniendo el mismo nombre y la misma estructura de las hojas LISTADO y MAYOR AUX G.


## Corrección importante: factor vs. margen

El cotizador original no calcula igual un **factor multiplicador** que un **margen porcentual**.

- Factor multiplicador: `ROUND(costo × factor, -2)`
- Margen porcentual: `ROUNDUP(costo / (1 - margen), 0)`

Ejemplo: para un 50% de margen selecciona **Margen %** e ingresa `50`.
No ingreses `0.5` como factor, porque eso reduciría el precio a la mitad.


## Corrección V3 — coincidencia exacta con Excel

El archivo LISTADO contiene algunos materiales repetidos. Excel `VLOOKUP(...,FALSE)` utiliza la **primera coincidencia exacta**.
La app ahora hace exactamente lo mismo y ya no reemplaza ese costo con registros duplicados posteriores.

Ejemplo validado:
- FF150 24" + DURLON 9000 1/8
- Costo plancha usado por Excel: $1.057.463,75
- Costo corte: $41.522,73
- Cantidad por plancha: 1
- Costo unitario: $1.098.986,49


## Generación automática de cotización Word

La app incluye `Plantilla.docx`, basada en la plantilla corporativa suministrada.
En la sección **Cotización**, completa los datos comerciales y presiona
**Descargar cotización Word**. Los productos agregados desde la app se insertan
en la tabla de la plantilla junto con cantidad, unidad, precio unitario y total.

Para GitHub debes subir/reemplazar:
- `app.py`
- `requirements.txt`
- `Plantilla.docx`

El Excel de costos se mantiene igual.
