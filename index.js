#!/usr/bin/env node
const ExcelJS = require('exceljs');
const path = require('node:path');
const xlsx = require('xlsx');
const fs = require('node:fs')


// --- FUNCIONES DE AYUDA ---
const formatearFecha = (f) => {
    const dd = String(f.getDate()).padStart(2, '0');
    const mm = String(f.getMonth() + 1).padStart(2, '0');
    const yy = String(f.getFullYear()).slice(-2);
    return `${dd}/${mm}/${yy}`;
};

const generarCalendario = (data) =>{
    // --- PROCESAMIENTO DE FECHAS ---
    const fechas = data.map(reg => new Date(reg['Fecha/Hora']));

    const minFecha = new Date(Math.min(...fechas));
    const maxFecha = new Date(Math.max(...fechas));
    minFecha.setHours(0,0,0,0);
    maxFecha.setHours(0,0,0,0);

    let fechaAux = new Date(minFecha);
    const calendario = [];

    while (fechaAux <= maxFecha) {
        calendario.push(formatearFecha(fechaAux));
        fechaAux.setDate(fechaAux.getDate() + 1);
    }

    return calendario;
}

// --- AGRUPAR TRABAJADORES ---
const obtenerTrabajadores = (data) =>{
    return data.reduce((acc, reg) => {
        if (!acc.find(t => t.id === reg["ID de usuario"])) {
            acc.push({ id: reg["ID de usuario"], nombre: reg["Nombre"], depto: reg["Departamento"] });
        }
        return acc;
    }, []);
}

// -- MAPEA LOS REGISTROS PARA SU FACIL ACCESO
const generarMapaAsistencia = (data) =>{
    const mapaAsistencia = {};
    data.forEach(reg => {
        const fLimpia = formatearFecha(new Date(reg["Fecha/Hora"]));
        const llave = `${reg["ID de usuario"]}-${fLimpia}`;
        if (!mapaAsistencia[llave]) mapaAsistencia[llave] = [];
        mapaAsistencia[llave].push(reg);
    });

    return mapaAsistencia
}

const obtenerDepartamentos = (data) =>{
    return [...new Set(data.map(reg => reg["Departamento"]))]
}

async function generarReporte(rutaEntrada, rutaSalida) {
// Validar que se pasaron los argumentos necesarios
    if (!rutaEntrada || !rutaSalida) {
        console.error("❌ ERROR: Faltan argumentos.");
        console.log("Uso correcto: node index.js <ruta_entrada> <ruta_salida>");
        console.log('Ejemplo: node index.js ./asistencia.xls ./reporte_abril.xlsx');
        return;
    }

    // Validar si el archivo de entrada realmente existe
    if (!fs.existsSync(rutaEntrada)) {
        console.error(`❌ ERROR: El archivo de entrada no existe en: ${rutaEntrada}`);
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
    
        const departamentos = obtenerDepartamentos(data)
        

        const calendario = generarCalendario(data)
        const trabajadores = obtenerTrabajadores(data)
        const mapaAsistencia = generarMapaAsistencia(data)
        

        // 2. CONFIGURAR EXCELJS
        const workbook = new ExcelJS.Workbook();
        

        const sheetColumns = [
            { header: 'ID', key: 'id', width: 10 },
            { header: 'Nombre', key: 'nombre', width: 30 },
            { header: 'Fecha', key: 'fecha', width: 15 },
            { header: 'Entrada', key: 'entrada', width: 12 },
            { header: 'Salida', key: 'salida', width: 12 },
            { header: 'Departamento', key: 'depto', width: 25 },
            { header: 'Estado', key: 'estado', width: 25 }
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
            console.log(`⏳ Procesando los registros de ${depto}...` );
            const trabajadoresDepto = trabajadores.filter(t => t.depto === depto)

            trabajadoresDepto.forEach(trabajador => {
                calendario.forEach(dia => {
                    const llave = `${trabajador.id}-${dia}`;
                    const registros = mapaAsistencia[llave];
                    
                    const [d, m, y] = dia.split('/');
                    const fEval = new Date(2000 + parseInt(y), parseInt(m)-1, parseInt(d));
                    const esFinde = (fEval.getDay() === 0 || fEval.getDay() === 6);

                    let rowData;
                    if (registros && registros.length > 0) {
                        registros.sort((a, b) => new Date(a["Fecha/Hora"]) - new Date(b["Fecha/Hora"]));
                        const hEntrada = new Date(registros[0]["Fecha/Hora"]).toLocaleTimeString('es-VE', {hour:'2-digit', minute:'2-digit', hour12:false});
                        const hSalida = registros.length > 1 
                            ? new Date(registros[registros.length-1]["Fecha/Hora"]).toLocaleTimeString('es-VE', {hour:'2-digit', minute:'2-digit', hour12:false})
                            : "--:--";

                        rowData = {
                            id: trabajador.id, nombre: trabajador.nombre, fecha: dia,
                            entrada: hEntrada, salida: hSalida, depto: trabajador.depto,
                            estado: esFinde ? 'TRABAJÓ FIN DE SEMANA' : 'PRESENTE'
                        };
                    } else {
                        rowData = {
                            id: trabajador.id, nombre: trabajador.nombre, fecha: dia,
                            entrada: '--:--', salida: '--:--', depto: trabajador.depto,
                            estado: esFinde ? 'FIN DE SEMANA' : 'FALTÓ'
                        };
                    }

                    const row = worksheet.addRow(rowData);
                    // Color rojo si faltó
                    if (rowData.estado === 'FALTÓ') {
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

if (require.main === module) {
    // 1. Capturar argumentos de la terminal
    // process.argv[0] es 'node', process.argv[1] es el script.
    // El [2] es el input y el [3] es el output.
    const rutaEntrada = process.argv[2];
    const rutaSalida = process.argv[3];

    (async () => {
        await generarReporte(rutaEntrada, rutaSalida);
    })();
}

module.exports = {
    generarReporte
}