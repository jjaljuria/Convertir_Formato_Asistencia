# Convertir Formato de Asistencia

Herramienta simple en Node.js para convertir registros de asistencia (XLS/XLSX) a un reporte de asistencia en formato Excel más legible.

Características principales

- Lee un archivo de entrada con registros de asistencia (columna `Fecha/Hora`).
- Agrupa por usuario y día, determina hora de entrada y salida.
- Marca días de fin de semana y faltas.
- Genera un archivo Excel `.xlsx` con columnas claras: ID, Nombre, Fecha, Entrada, Salida, Departamento, Estado.

Requisitos

- Node.js 16+ instalado.
- Recomendado: `pnpm` (se incluye `packageManager`), también funciona con `npm`.

Instalación

1. Clona o descarga el repositorio.
2. Desde la carpeta del proyecto instala dependencias:

```bash
pnpm install
# o, si usas npm:
npm install
```

Uso

1. Formato de entrada esperado:
   - Archivo Excel (`.xls` / `.xlsx`) cuya primera hoja contiene registros con, las columnas:
     - `ID de usuario`
     - `Nombre`
     - `Departamento`
     - `Fecha/Hora` (un valor de fecha/hora que ExcelJS pueda interpretar correctamente. Se recomienda un formato consistente como `YYYY-MM-DD HH:MM:SS` o `DD/MM/YYYY HH:MM:SS` para asegurar una correcta lectura.

2. Ejecuta el script indicando ruta de entrada y ruta de salida:

```bash
cofa <ruta_entrada> <ruta_salida>
# Ejemplo:
cofa ./asistencia.xls ./reporte_abril.xlsx
```

### Objeto de Configuración (`usuarioConfig`)

La función `generarReporte` acepta un tercer argumento opcional, `usuarioConfig`, que es un objeto de configuración. Este objeto es para uso programático y permite personalizar los nombres de las columnas que el script busca en el archivo de entrada y los nombres de las columnas que genera en el archivo de salida.

#### `usuarioConfig.entrada`

Este objeto se utiliza para mapear los nombres de las columnas del archivo de entrada a las claves internas del script. Si los nombres de tus columnas de entrada difieren de los predeterminados, puedes especificarlos aquí.

**Valores predeterminados para `entrada`:**

- `id`: 'ID de usuario'
- `nombre`: 'Nombre'
- `fechaHora`: 'Fecha/Hora'
- `departamento`: 'Departamento'

#### `usuarioConfig.salida`

Este objeto se utiliza para definir los nombres de los encabezados de las columnas en el archivo de reporte de salida. Puedes cambiar estos nombres para que se ajusten a tus preferencias.

**Valores predeterminados para `salida`:**

- `id`: 'iD'
- `nombre`: 'Nombre'
- `fecha`: 'Fecha'
- `entrada`: 'Entrada'
- `salida`: 'Salida'
- `departamento`: 'Departamento'
- `estado`: 'Estado'

**Ejemplo de uso de `usuarioConfig`:**

```javascript
generarReporte("./asistencia.xls", "./reporte_personalizado.xlsx", {
  entrada: {
    id: "Número de Empleado",
    fechaHora: "Registro de Horario",
  },
  salida: {
    id: "ID Empleado",
    entrada: "Hora de Entrada",
    salida: "Hora de Salida",
  },
});
```

Este ejemplo le dice al script que en el archivo de entrada, la columna para el ID del usuario se llama 'Número de Empleado' y la de fecha/hora es 'Registro de Horario'. Para el archivo de salida, los encabezados de las columnas de ID, entrada y salida se renombrarán a 'ID Empleado', 'Hora de Entrada' y 'Hora de Salida' respectivamente.

Salida

- Se genera un archivo `.xlsx` con una hoja llamada `Asistencia` y las columnas:
  `ID`, `Nombre`, `Fecha`, `Entrada`, `Salida`, `Departamento`, `Estado`.
- El estado puede ser `PRESENTE`, `FALTÓ`, `FIN DE SEMANA` o `TRABAJÓ FIN DE SEMANA`.

Notas y consejos

- Si no pasas argumentos, el script imprimirá el uso correcto.
- El script determina la hora de entrada como la primera marca del día y la hora de salida como la última.
- Asegúrate de que la columna `Fecha/Hora` esté en un formato de fecha válido en el archivo de Excel para que la librería pueda parsearla correctamente.

Contribuciones

- Si quieres mejorar la herramienta, abre un issue o un pull request con cambios concretos.

[Repositorio en GitHub](https://github.com/jjaljuria/Convertir_Formato_Asistencia)

Licencia

- Revisa el archivo `LICENSE` incluido en el repositorio.
