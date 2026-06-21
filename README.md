# ProdePAS

Aplicacion web de solo lectura para mostrar la tabla general del prode, los resultados oficiales y el detalle por usuario.

## Componentes

- [server.js](server.js): backend Node.js sin dependencias externas.
- [public/](public/): interfaz HTML5 + Bootstrap 5.
- [data/](data/): TXT exportados desde Excel en formato JSON.
- [excel/ExportProdeTxt.bas](excel/ExportProdeTxt.bas): macro VBA para exportar los TXT automaticamente.

## Ejecutar

```bash
node server.js
```

Luego abrir `http://localhost:3000`.

## Formato TXT

El export de Excel genera estos archivos:

- `data/tabla_general.txt`
- `data/resultados_oficiales.txt`
- `data/users/<usuario_tab>.txt`

### Tabla general

```json
{
  "updatedAt": "2026-06-20T18:30:00-03:00",
  "rows": [
    {
      "position": 1,
      "usuario_tab": "Marcos",
      "nombre_fantasia": "Marquinhos",
      "points": 37,
      "exactos": 5,
      "acertados": 16,
      "played": 72
    }
  ]
}
```

### Resultados oficiales

```json
{
  "updatedAt": "2026-06-20T18:30:00-03:00",
  "matches": [
    {
      "number": 1,
      "fecha": "jue 11 Jun",
      "grupo": "A",
      "local": "MÉXICO",
      "goles_local": 2,
      "goles_visitante": 0,
      "visitante": "SUDÁFRICA",
      "score": "2 - 0",
      "status": "✔ Jugado",
      "stadium": "Ciudad de México"
    }
  ]
}
```

### Detalle de usuario

```json
{
  "usuario_tab": "Marcos",
  "nombre_fantasia": "Marquinhos",
  "updatedAt": "2026-06-20T18:30:00-03:00",
  "matches": [
    {
      "number": 1,
      "match": "MÉXICO vs SUDÁFRICA",
      "prediction": "2 - 2",
      "official": "2 - 0",
      "points": 1
    }
  ]
}
```

## Excel

El macro [excel/ExportProdeTxt.bas](excel/ExportProdeTxt.bas) se ejecuta con `Auto_Open` y vuelve a generar los TXT en la carpeta `data/` del libro. Si el libro se abre en Excel con macros habilitadas, la exportacion se hace automaticamente.

### Instructivo rapido para pegar el modulo

1. Abrir `ProdeMundialBoleta_PAS.xlsm` en Excel.
2. Habilitar macros y, si Excel lo pide, activar el acceso al proyecto VBA.
3. Presionar `Alt + F11` para abrir el editor de VBA.
4. Ir a `File > Import File` y elegir [excel/ExportProdeTxt.bas](excel/ExportProdeTxt.bas).
5. Guardar el libro como `.xlsm`.
6. Cerrar y volver a abrir el archivo para que `Auto_Open` exporte los TXT automaticamente.

### Resultado esperado

Al abrir el libro, Excel debe generar o actualizar:

- `data/tabla_general.txt`
- `data/resultados_oficiales.txt`
- `data/users/<usuario_tab>.txt`