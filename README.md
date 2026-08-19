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
