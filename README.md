# Alquería × Eficacia — Dashboard de Evaluación en Campo

## Qué contiene
- `index.html`: dashboard completo, listo para publicar como Static Site.
- `modelo_evaluacion_mercaderista_original.html`: copia de referencia.
- `assets/`: reservado para futuras imágenes externas.

## Publicarlo en Render desde cero

### Opción recomendada: GitHub + Render

1. Entra a GitHub: https://github.com/
2. Inicia sesión.
3. Pulsa **New repository**.
4. Pon como nombre:
   `alqueria-eficacia-dashboard`
5. Puedes dejarlo **Private**.
6. Crea el repositorio.
7. Dentro del repositorio pulsa **Add file → Upload files**.
8. Sube `index.html` y, si quieres, los demás archivos de esta carpeta.
9. Pulsa **Commit changes**.

### Crear el sitio en Render

1. Entra a https://dashboard.render.com/
2. Inicia sesión.
3. Pulsa **New +**.
4. Selecciona **Static Site**.
5. Conecta GitHub si Render lo solicita.
6. Selecciona:
   `alqueria-eficacia-dashboard`
7. Configura:

   Name:
   `alqueria-eficacia`

   Branch:
   `main`

   Root Directory:
   dejar vacío

   Build Command:
   `echo "Static HTML"`

   Publish Directory:
   `.`

8. Pulsa **Create Static Site** / **Deploy Static Site**.
9. Espera a que termine el deploy.
10. Render te dará una dirección similar a:
    `https://alqueria-eficacia.onrender.com`

## Actualizar la página

Cada vez que cambies `index.html`:
1. Sube el archivo nuevo a GitHub.
2. Haz commit.
3. Render detectará el cambio y volverá a desplegarlo.

## Importante sobre la seguridad

Este proyecto es un dashboard HTML/JavaScript del lado del navegador. La pantalla de resultados tiene acceso con usuario y contraseña, pero esas credenciales forman parte del código de la aplicación.

Para una demo o piloto está bien. Para producción con información sensible, se recomienda posteriormente implementar autenticación real y una base de datos/backend.

## Importante sobre Excel

El dashboard usa la librería XLSX desde CDN para generar archivos Excel desde el navegador. Por eso, al usar las funciones de exportación, el navegador necesita acceso a Internet para cargar esa librería.

## Estructura

alqueria-eficacia-dashboard/
├── index.html
├── modelo_evaluacion_mercaderista_original.html
└── assets/
