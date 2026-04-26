const xlsx = require('xlsx')
const path = require('path')
const { worker } = require('cluster')


const filePath = path.join(__dirname, 'asistencia-01-04-2026_24-04-2026.xls')

const book = xlsx.readFile(filePath, {cellDates: true})
const sheetName = book.SheetNames[0]

// Leer la hoja manteniendo los valores en crudo para preservar objetos Date
const data = xlsx.utils.sheet_to_json(book.Sheets[sheetName], { raw: true })

const fechas = data.map(reg => new Date(reg['Fecha/Hora']))

const formatearFecha = (f) => {
    return f.toLocaleDateString('es-VE',{
        day: '2-digit',
        month: '2-digit',
        year: '2-digit'
    })
};

const minFecha = new Date(Math.min(...fechas))
minFecha.setHours(0, 0, 0, 0)

const maxFecha = new Date(Math.max(...fechas))
maxFecha.setHours(0, 0, 0, 0)

let fechaAuxiliar = new Date(minFecha)
fechaAuxiliar.setHours(0, 0, 0, 0)
const calendario = []

while(fechaAuxiliar <= maxFecha){
    const dia =  formatearFecha(fechaAuxiliar)

    calendario.push(dia)
    fechaAuxiliar.setDate(fechaAuxiliar.getDate() + 1)
}

// 'datos' es lo que obtuviste de xlsx.utils.sheet_to_json(hoja)
const trabajadoresUnicos = data.reduce((acc, registro) => {
    const id = registro["ID de usuario"];
    // Si aún no hemos agregado este ID al acumulador...
    if (!acc.find(t => t.id === id)) {
        acc.push({
            id: id,
            nombre: registro["Nombre"],
            departamento: registro["Departamento"]
        });
    }
    return acc;
}, []);

const mapaAsistencia = {};
data.forEach(reg => {
    const fechaObj = reg["Fecha/Hora"] instanceof Date ? reg["Fecha/Hora"] : new Date(reg["Fecha/Hora"]);
    const fechaDia = formatearFecha(fechaObj);
    const llave = `${reg["ID de usuario"]}-${fechaDia}`;
    if (!mapaAsistencia[llave]) mapaAsistencia[llave] = [];
    mapaAsistencia[llave].push(fechaObj);
});

// Helper: formatear hora a HH:mm
const formatHora = (d) => {
    if (!d) return '';
    const date = d instanceof Date ? d : new Date(d);
    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    return `${hh}:${mm}`;
};

// Helper: obtener entrada/salida (primera/ultima) desde un array de Date
const obtenerEntradaSalida = (fechasArray) => {
    if (!fechasArray || fechasArray.length === 0) return { entrada: '', salida: '' };
    const sorted = fechasArray.slice().sort((a, b) => a - b);
    const entrada = sorted[0];
    const salida = sorted[sorted.length - 1];
    return { entrada: formatHora(entrada), salida: formatHora(salida) };
};

const reporteConFaltas = [];

for (const trabajador of trabajadoresUnicos) {
    for (const dia of calendario) {
        const llaveBusqueda = `${trabajador.id}-${dia}`;
        const marcas = mapaAsistencia[llaveBusqueda];

        if (marcas && marcas.length) {
            const { entrada, salida } = obtenerEntradaSalida(marcas);
            reporteConFaltas.push({
                "ID de usuario": trabajador.id,
                "Nombre": trabajador.nombre,
                "Departamento": trabajador.departamento,
                "Fecha": dia,
                "Entrada": entrada,
                "Salida": salida,
                "Estado": "ASISTIÓ"
            });
        } else {
            reporteConFaltas.push({
                "ID de usuario": trabajador.id,
                "Nombre": trabajador.nombre,
                "Departamento": trabajador.departamento,
                "Fecha": dia,
                "Entrada": '',
                "Salida": '',
                "Estado": "FALTÓ"
            });
        }
    }
}

// Ahora puedes guardar 'reporteConFaltas' en un nuevo Excel
const nuevaHoja = xlsx.utils.json_to_sheet(reporteConFaltas);
const nuevoLibro = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(nuevoLibro, nuevaHoja, "Asistencia Completa");
xlsx.writeFile(nuevoLibro, "reporte-final-asistencia.xlsx");