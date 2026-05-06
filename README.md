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
   - Archivo Excel (`.xls` / `.xlsx`) cuya primera hoja contiene registros con, como mínimo, las columnas:
     - `ID de usuario`
     - `Nombre`
     - `Departamento`
     - `Fecha/Hora` (valor de fecha/hora reconocible por Excel)

2. Ejecuta el script indicando ruta de entrada y ruta de salida:

```bash
cofa <ruta_entrada> <ruta_salida>
# Ejemplo:
cofa ./asistencia.xls ./reporte_abril.xlsx
```

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
