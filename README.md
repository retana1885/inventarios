# Conteo fisico de inventario (sin base de datos)

Aplicacion web simple para agilizar el conteo por filas (F1 a F8) usando un archivo Excel de existencias.

## Funciones principales

- Carga un archivo Excel (`.xlsx` o `.xls`).
- Soporta arrastrar y soltar archivo en la zona de carga.
- Oculta automaticamente la seccion de carga 3s despues de importar (se puede mostrar de nuevo con una pestana flotante).
- Detecta automaticamente columnas: `CODIGO`, `NOMBRE`, `POSICION`, `EXISTENCIA`.
- Calcula estadisticas iniciales (productos, piezas, existencias en cero, etc.).
- Filtra por fila (`F1` a `F8`) o muestra todas.
- Permite busqueda por `codigo`, `nombre` o `posicion`.
- Ordena resultados de busqueda priorizando coincidencia exacta.
- Permite capturar `CONTEO FISICO` desde la web.
- Atajo de captura: `Enter` enfoca el siguiente pendiente.
- Incluye botones de operacion rapida: `Siguiente pendiente` y `Solo diferencias`.
- Autoguarda avance cada 20 segundos e indica "guardado hace..." en pantalla.
- Calcula automaticamente `SOBRANTE` y `FALTANTE` con base en (`conteo - existencia`).
- Destaca visualmente estatus por renglon (`Pendiente`, `OK`, `Faltante`, `Sobrante`).
- Guarda/restaura avance en el navegador (`localStorage`).
- Exporta a Excel usando la plantilla `Plantilla_exportar_datos.xlsx`.

## Como usar

1. Abre `index.html` en tu navegador.
2. Selecciona el archivo de existencias.
3. Elige la fila a contar y captura el conteo fisico.
4. Usa **Guardar avance local** si quieres continuar despues.
5. Al finalizar, usa **Exportar resultados**.

## Publicar en GitHub Pages

1. Sube estos archivos a un repositorio:
   - `index.html`
   - `styles.css`
   - `app.js`
2. En GitHub, activa Pages en la rama principal (`Settings > Pages`).
3. Comparte la URL para abrirla desde celular.

## Nota

Esta version no usa base de datos ni multiusuario en tiempo real. Todo ocurre localmente en el navegador.
