# factpalmera

Tablero de márgenes de Palmera (vista Márgenes: costo del producto, facturación total y margen retail por cadena y SKU).
La vista Ventas está oculta; se reactiva con `SHOW_VENTAS=true` en el script de `index.html`.

## Cómo se actualiza en Netlify

`index.html` llama a la función `netlify/functions/data.js`, que lee las hojas CONSOLIDADO y FACTURA de
`Facturación_Palmera_SEP.xlsx` con Microsoft Graph. La página vuelve a leer cada 5 minutos y con el botón Actualizar.
Si la función no está configurada, el botón **Cargar .xlsx** sigue funcionando.

### Configuración (una sola vez, la hace TI de Microsoft 365)

1. Azure (Entra ID) > Registros de aplicaciones > Nuevo registro. Anotar *Id. de directorio (inquilino)* e *Id. de aplicación (cliente)*.
2. Certificados y secretos > Nuevo secreto de cliente. Anotar el valor.
3. Permisos de API > Microsoft Graph > Permisos de aplicación > `Files.Read.All` > Conceder consentimiento de administrador.
4. Netlify > Site configuration > Environment variables:

| Variable | Valor |
|---|---|
| `TENANT_ID` | Id. de directorio (inquilino) |
| `CLIENT_ID` | Id. de aplicación (cliente) |
| `CLIENT_SECRET` | Valor del secreto |
| `ACCESS_KEY` | Clave que pedirá el tablero al abrirlo |
| `DRIVE_ID`, `ITEM_ID`, `FILE_NAME` | Opcionales. Por defecto apuntan a `Facturación_Palmera_SEP.xlsx` |

5. Netlify > Deploys > Trigger deploy.

`Files.Read.All` da lectura sobre todos los archivos del inquilino. Para limitarlo, mover el Excel a un sitio de SharePoint y usar `Sites.Selected`.
El secreto del cliente vence (máximo 24 meses): renovarlo y actualizar `CLIENT_SECRET`.
