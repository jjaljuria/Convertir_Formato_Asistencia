# Convertir Formato de Asistencia [![Node.js CI](https://github.com/jjaljuria/Convertir_Formato_Asistencia/actions/workflows/ci.yml/badge.svg)](https://github.com/jjaljuria/Convertir_Formato_Asistencia/actions/workflows/ci.yml)

Herramienta simple en Node.js para convertir registros de asistencia (XLS/XLSX) a un reporte de asistencia en formato Excel más legible.

Características principales

- Lee un archivo de entrada con registros de asistencia (columna `Fecha/Hora`).
- Agrupa por usuario, determina hora de entrada y salida.
- Marca días de fin de semana y faltas.
- Genera un archivo Excel `.xlsx` con columnas claras: ID, Nombre, Fecha, Entrada, Salida, Departamento, Estado.

Requisitos

- Node.js 20+ instalado.
- Recomendado: `pnpm` (se incluye `packageManager`), también funciona con `npm` o `yarn`.

### Instalación Global (para uso CLI)

Para usar la herramienta como un comando global (`cofa`) desde cualquier directorio en tu terminal, puedes instalarla globalmente:

```bash
pnpm install -g convertir-formato-asistencia
# o, si usas npm:
npm install -g convertir-formato-asistencia
# o, si usas yarn:
yarn global add convertir-formato-asistencia
```

Asegúrate de ejecutar este comando dentro del directorio raíz del proyecto.

### 1. Formato de entrada esperado:

- Archivo Excel (`.xls` / `.xlsx`) cuya primera hoja contiene registros con, las columnas:
  - `ID de usuario`
  - `Nombre`
  - `Departamento`
  - `Fecha/Hora` (un valor de fecha/hora que ExcelJS pueda interpretar correctamente. Se recomienda un formato consistente como `YYYY-MM-DD HH:MM:SS` o `DD/MM/YYYY HH:MM:SS` para asegurar una correcta lectura.

### 2. Uso CLI (Command Line Interface)

Una vez que la herramienta esté instalada globalmente, puedes ejecutarla directamente desde tu terminal.

```bash
cofa <ruta_entrada> <ruta_salida>
# Ejemplo:
cofa ./asistencia.xls ./reporte_abril.xlsx
```

- `<ruta_entrada>`: La ruta al archivo de Excel (`.xls` o `.xlsx`) que contiene los registros de asistencia.
- `<ruta_salida>`: La ruta donde se guardará el reporte de asistencia formateado generado `.xlsx`.

### 3. Uso Programático

Instalación

    ```bash
    pnpm install convertir-formato-asistencia
    # o, si usas npm:
    npm install convertir-formato-asistencia
    # o, si usas yarn:
    yarn install convertir-formato-asistencia
    ```

Puedes importar y utilizar la función `generarReporte` directamente en tus proyectos Node.js.

```javascript
import { generarReporte } from "convertir-formato-asistencia";

async function ejecutarReporte() {
  const rutaEntrada = "./asistencia.xls";
  const rutaSalida = "./reporte_personalizado.xlsx";

  // Configuración opcional para personalizar las columnas de entrada y salida
  const usuarioConfig = {
    entrada: {
      id: "Número de Empleado",
      fechaHora: "Registro de Horario",
      // Puedes añadir más mapeos de columnas aquí
    },
    salida: {
      id: "ID Empleado",
      entrada: "Hora de Entrada",
      salida: "Hora de Salida",
      // Puedes añadir más nombres de columnas de salida aquí
    },
  };

  try {
    await generarReporte(rutaEntrada, rutaSalida, usuarioConfig);
    console.log("Reporte generado exitosamente de forma programática.");
  } catch (error) {
    console.error("Error al generar el reporte de forma programática:", error);
  }
}

ejecutarReporte();
```

### Objeto de Configuración (`usuarioConfig`)

La función `generarReporte` acepta un tercer argumento opcional, `usuarioConfig`, que es un objeto de configuración. Permite personalizar los nombres de las columnas que el script busca en el archivo de entrada y los nombres de las columnas que genera en el archivo de salida.

#### `usuarioConfig.entrada`

Este objeto se utiliza para mapear los nombres de las columnas del archivo de entrada.

**Valores predeterminados para `entrada`:**

- `id`: 'ID de usuario'
- `nombre`: 'Nombre'
- `fechaHora`: 'Fecha/Hora'
- `departamento`: 'Departamento'

#### `usuarioConfig.salida`

Este objeto se utiliza para definir los nombres de los encabezados de las columnas en el archivo de reporte de salida.

**Valores predeterminados para `salida`:**

- `id`: 'iD'
- `nombre`: 'Nombre'
- `fecha`: 'Fecha'
- `entrada`: 'Entrada'
- `salida`: 'Salida'
- `departamento`: 'Departamento'
- `estado`: 'Estado'

Notas y consejos

- Si no pasas argumentos, el script imprimirá el uso correcto.
- El script determina la hora de entrada como la primera marca del día y la hora de salida como la última.
- Asegúrate de que la columna `Fecha/Hora` esté en un formato de fecha válido en el archivo de Excel para que la librería pueda parsearla correctamente.

Contribuciones

- Si quieres mejorar la herramienta, abre un issue o un pull request con cambios concretos.

[Repositorio en GitHub](https://github.com/jjaljuria/Convertir_Formato_Asistencia)

Licencia

- Revisa el archivo `LICENSE` incluido en el repositorio.
