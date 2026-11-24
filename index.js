#!/usr/bin/env node
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url'
import path from 'node:path';
import xlsx from 'xlsx';
import fs from 'node:fs';


// --- FUNCIONES DE AYUDA ---
const formatearFecha = (f) => {
    const dd = String(f.getDate()).padStart(2, '0');
    const mm = String(f.getMonth() + 1).padStart(2, '0');
    const yy = String(f.getFullYear()).slice(-2);
    return `${dd}/${mm}/${yy}`;
};

const generarCalendario = (data) => {
    const columnas = data.columnas;
    // --- PROCESAMIENTO DE FECHAS ---
    const fechas = data.map(reg => new Date(reg[columnas.fechaHora]));

    const minFecha = new Date(Math.min(...fechas));
    const maxFecha = new Date(Math.max(...fechas));
    minFecha.setHours(0, 0, 0, 0);
    maxFecha.setHours(0, 0, 0, 0);

    let fechaAux = new Date(minFecha);
    const calendario = [];

    while (fechaAux <= maxFecha) {
        calendario.push(formatearFecha(fechaAux));
        fechaAux.setDate(fechaAux.getDate() + 1);
    }

    return calendario;
}

const normalizarId = (id) => String(id ?? '').trim();

// --- AGRUPAR TRABAJADORES ---
const obtenerTrabajadores = (data, personal = []) => {
    const columnas = data.columnas
    const trabajadoresPorId = new Map();
    data.forEach(reg => {
        const id = reg[columnas.id];
        const llave = normalizarId(id);
        if (!trabajadoresPorId.has(llave)) {
            const departamento = reg[columnas.departamento] || "";
            trabajadoresPorId.set(llave, { id, nombre: reg[columnas.nombre], depto: departamento });
        }
    });

    personal.forEach(trabajador => {
        const llave = normalizarId(trabajador.id);
        if (!trabajadoresPorId.has(llave)) {
            trabajadoresPorId.set(llave, trabajador);
        }
    });

    return [...trabajadoresPorId.values()];
}

// -- MAPEA LOS REGISTROS PARA SU FACIL ACCESO
const generarMapaAsistencia = (data) => {
    const mapaAsistencia = {};
    const columnas = data.columnas
    data.forEach(reg => {
        const fLimpia = formatearFecha(new Date(reg[columnas.fechaHora]));
        const llave = `${reg[columnas.id]}-${fLimpia}`;
        if (!mapaAsistencia[llave]) mapaAsistencia[llave] = [];
        mapaAsistencia[llave].push(reg);
    });

    return mapaAsistencia
}

const obtenerDepartamentos = (trabajadores) => {
    return [...new Set(trabajadores.map(trabajador => trabajador.depto).filter(depto => depto !== undefined && depto !== null))]
}

const normalizarEncabezado = (encabezado) => String(encabezado ?? '')
    .normalize('NFD')
    .replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

const leerPersonal = (rutaPersonal) => {
    const bookPersonal = xlsx.readFile(rutaPersonal, { cellDates: true });
    const nombreHoja = bookPersonal.SheetNames?.[0];
    if (!nombreHoja) {
        throw new Error('El archivo de personal no contiene hojas.');
    }
    const hojaPersonal = bookPersonal.Sheets[nombreHoja];
    const filas = xlsx.utils.sheet_to_json(hojaPersonal, { header: 1, defval: null });
    const encabezados = filas[0] || [];
    const columnasRequeridas = {
        id: 'Usuario Nro.',
        nombre: 'Nombre',
        departamento: 'Departamento'
    };
    const indices = {};

    Object.entries(columnasRequeridas).forEach(([campo, encabezadoEsperado]) => {
        const indice = encabezados.findIndex(encabezado =>
            normalizarEncabezado(encabezado) === normalizarEncabezado(encabezadoEsperado)
        );
        if (indice === -1) {
            throw new Error(`El archivo de personal no contiene la columna requerida "${encabezadoEsperado}".`);
        }
        indices[campo] = indice;
    });

    return filas.slice(1).flatMap((fila, indice) => {
        const id = fila[indices.id];
        const nombre = fila[indices.nombre];
        const departamento = fila[indices.departamento];
        if ([id, nombre, departamento].every(valor => valor === null || valor === undefined || valor === '')) {
            return [];
        }
        if (id === null || id === undefined || id === '') {
            throw new Error(`Falta "Usuario Nro" en la fila ${indice + 2} del archivo de personal.`);
        }
        if (nombre === null || nombre === undefined || nombre === '') {
            throw new Error(`Falta "Nombre" en la fila ${indice + 2} del archivo de personal.`);
        }
        return [{ id, nombre, depto: departamento || '' }];
    });
}

async function generarReporte(rutaEntrada, rutaSalida, usuarioConfig = {}, rutaPersonal) {
    const defaultConfig = {
        entrada: {
            id: 'ID de usuario',
            nombre: 'Nombre',
            fechaHora: 'Fecha/Hora',
            departamento: 'Departamento'
        },
        salida: {
            id: 'iD',
            nombre: 'Nombre',
            fecha: 'Fecha',
            entrada: 'Entrada',
            salida: 'Salida',
            departamento: 'Departamento',
            estado: "Estado"
        },
        mensajes: {
            presente: 'PRESENTE',
            falto: 'SIN ASISTENCIA',
            finDeSemana: 'FIN DE SEMANA',
            vacacionesPermiso: ''
        }
    };

    const config = {
        entrada: { ...defaultConfig.entrada, ...usuarioConfig.entrada },
        salida: { ...defaultConfig.salida, ...usuarioConfig.salida },
        mensajes: { ...defaultConfig.mensajes, ...usuarioConfig.mensajes }
    };
    // Validar que se pasaron los argumentos necesarios
    if (!rutaEntrada || !rutaSalida) {
        console.error("❌ ERROR: Faltan argumentos de entrada y/o salida.");
        console.log("Para ver el uso correcto, ejecuta: cofa --help");
        return;
    }

    // Validar si el archivo de entrada realmente existe
    if (!fs.existsSync(rutaEntrada)) {
        console.error(`❌ ERROR: El archivo de entrada no existe en: ${rutaEntrada}`);
        return;
    }

    if (rutaPersonal && !fs.existsSync(rutaPersonal)) {
        console.error(`❌ ERROR: El archivo de personal no existe en: ${rutaPersonal}`);
        return;
    }


    console.log(`🚀 Procesando: ${path.basename(rutaEntrada)}...`);

    try {
        // 1. LEER DATOS
        const book = xlsx.readFile(rutaEntrada, { cellDates: true });
        const data = xlsx.utils.sheet_to_json(book.Sheets[book.SheetNames[0]]);

        if (data.length === 0) {
            throw new Error("El archivo de entrada está vacío o no se leyó correctamente.");
        }
        console.log(`✅ Leídos ${data.length} registros del archivo original.`);

        data.columnas = config.entrada

        const calendario = generarCalendario(data)
        const personal = rutaPersonal ? leerPersonal(rutaPersonal) : [];
        const trabajadores = obtenerTrabajadores(data, personal)
        const departamentos = obtenerDepartamentos(trabajadores)
        const mapaAsistencia = generarMapaAsistencia(data)
        const idsConAsistencia = new Set(data.map(reg => normalizarId(reg[config.entrada.id])))


        // 2. CONFIGURAR EXCELJS
        const workbook = new ExcelJS.Workbook();


        const sheetColumns = [
            { header: config.salida.id, key: 'id', width: 10 },
            { header: config.salida.nombre, key: 'nombre', width: 30 },
            { header: config.salida.fecha, key: 'fecha', width: 15 },
            { header: config.salida.entrada, key: 'entrada', width: 12 },
            { header: config.salida.salida, key: 'salida', width: 12 },
            { header: config.salida.departamento, key: 'depto', width: 25 },
            { header: config.salida.estado, key: 'estado', width: 25 }
        ];

        console.log("⏳ Generando hojas por departamento y aplicando lógica de entrada/salida...");
        departamentos.forEach(depto => {
            const sheetName = depto;
            const worksheet = workbook.addWorksheet(sheetName);
            worksheet.columns = sheetColumns;

            // FORMATO DE ENCABEZADO
            worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
            worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };

            //GENERAR FILAS
            console.log(`⏳ Procesando los registros de ${depto}...`);
            const trabajadoresDepto = trabajadores.filter(t => t.depto === depto)

            trabajadoresDepto.forEach(trabajador => {
                calendario.forEach(dia => {
                    const llave = `${trabajador.id}-${dia}`;
                    const registros = mapaAsistencia[llave];
                    const noTieneAsistencia = !idsConAsistencia.has(normalizarId(trabajador.id));

                    const [d, m, y] = dia.split('/');
                    const fEval = new Date(2000 + parseInt(y), parseInt(m) - 1, parseInt(d));
                    const esFinde = (fEval.getDay() === 0 || fEval.getDay() === 6);

                    let rowData;
                    if (registros && registros.length > 0) {
                        registros.sort((a, b) => new Date(a[config.entrada.fechaHora]) - new Date(b[config.entrada.fechaHora]));
                        const hEntrada = new Date(registros[0][config.entrada.fechaHora]).toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit', hour12: false });
                        const hSalida = registros.length > 1
                            ? new Date(registros[registros.length - 1][config.entrada.fechaHora]).toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit', hour12: false })
                            : "--:--";

                        rowData = {
                            id: trabajador.id, nombre: trabajador.nombre, fecha: dia,
                            entrada: hEntrada, salida: hSalida, depto: trabajador.depto,
                            estado: esFinde ? config.mensajes.finDeSemana : config.mensajes.presente
                        };
                    } else if (noTieneAsistencia) {
                        rowData = {
                            id: trabajador.id, nombre: trabajador.nombre, fecha: dia,
                            entrada: '--:--', salida: '--:--', depto: trabajador.depto,
                            estado: esFinde ? config.mensajes.finDeSemana : config.mensajes.vacacionesPermiso
                        };
                    } else {
                        rowData = {
                            id: trabajador.id, nombre: trabajador.nombre, fecha: dia,
                            entrada: '--:--', salida: '--:--', depto: trabajador.depto,
                            estado: esFinde ? config.mensajes.finDeSemana : config.mensajes.falto
                        };
                    }

                    const row = worksheet.addRow(rowData);
                    // Color rojo si faltó
                    if (rowData.estado === config.mensajes.falto) {
                        row.getCell('estado').font = { color: { argb: 'FFFF0000' }, bold: true };
                    }
                });
            });
        })


        // GUARDAR (CON AWAIT)
        console.log(`💾 Intentando guardar archivo en: ${rutaSalida}`);
        await workbook.xlsx.writeFile(rutaSalida);

        console.log("✨ ¡PROCESO COMPLETADO! El archivo se generó correctamente.");

    } catch (error) {
        console.error("❌ ERROR CRÍTICO:", error.message);
    }
}

function mostrarAyuda() {
    console.log("\nUso: cofa [--personal <ruta_personal>] <ruta_entrada> <ruta_salida>\n");
    console.log("Descripción: Herramienta para convertir registros de asistencia de archivos XLS/XLSX a un reporte de Excel legible, agrupando por usuario y marcando ausencias.\n");
    console.log("Argumentos:\n");
    console.log("  <ruta_entrada>     Ruta al archivo de asistencia de entrada (XLS/XLSX).");
    console.log("  <ruta_salida>      Ruta donde se guardará el reporte de Excel generado (XLSX).");
    console.log("\nOpciones:\n");
    console.log("  --personal <ruta>  Excel de personal con Usuario Nro, Nombre y Departamento.");
    console.log("  --help, -h         Muestra esta información de ayuda.");
    console.log("  --version, -v      Muestra la versión de la aplicación.");
    console.log("\nEjemplo:\n");
    console.log("  cofa ./asistencia.xls ./reporte_abril.xlsx");
    console.log("  cofa --personal ./personal.xlsx ./asistencia.xls ./reporte_abril.xlsx\n");
    process.exit(0);
}

function resolverArgumentosCLI(args) {
    let rutaPersonal;
    const argsPosicionales = [];

    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--personal') {
            if (!args[i + 1] || args[i + 1].startsWith('-')) {
                throw new Error('--personal requiere la ruta del archivo de personal.');
            }
            rutaPersonal = args[++i];
        } else {
            argsPosicionales.push(args[i]);
        }
    }

    return {
        rutaEntrada: argsPosicionales[0],
        rutaSalida: argsPosicionales[1],
        rutaPersonal
    };
}

// convertir import.meta.url a path de archivo
const __filename = fileURLToPath(import.meta.url);

const invoked = path.basename(process.argv[1] || '');
const scriptName = path.basename(__filename);

if (invoked === scriptName || invoked === 'cofa') {
    const args = process.argv.slice(2);

    const packageJsonPath = path.join(path.dirname(__filename), "package.json");
    const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf8"));
    const appVersion = packageJson.version;

    if (args.includes("--version") || args.includes("-v")) {
        console.log(`cofa v${appVersion}`);
        process.exit(0);
    }

    if (args.includes("--help") || args.includes("-h")) {
        mostrarAyuda();
    }

    let rutaEntrada;
    let rutaSalida;
    let rutaPersonal;

    try {
        ({ rutaEntrada, rutaSalida, rutaPersonal } = resolverArgumentosCLI(args));
    } catch (error) {
        console.error(`❌ ERROR: ${error.message}`);
        process.exit(1);
    }

    (async () => {
        await generarReporte(rutaEntrada, rutaSalida, {}, rutaPersonal);
    })();
}



export { generarReporte, resolverArgumentosCLI };
