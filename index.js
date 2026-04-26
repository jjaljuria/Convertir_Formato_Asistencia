const xlsx = require('xlsx')
const path = require('path')
const { worker } = require('cluster')


const filePath = path.join(__dirname, 'asistencia-01-04-2026_24-04-2026.xls')

const book = xlsx.readFile(filePath, {cellDates: true})
const sheetName = book.SheetNames[0]

const data = xlsx.utils.sheet_to_json(book.Sheets[sheetName])

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
    const fecha = formatearFecha(reg["Fecha/Hora"])
    const llave = `${reg["ID de usuario"]}-${fecha}`;
    reg["Fecha/Hora"] = fecha
    mapaAsistencia[llave] = reg; 
});

const reporteConFaltas = [];

for (const trabajador of trabajadoresUnicos) {
    for (const dia of calendario) {
        const llaveBusqueda = `${trabajador.id}-${dia}`;
        const registroExistente = mapaAsistencia[llaveBusqueda];

        if (registroExistente) {
            // Si asistió, agregamos su registro original
            reporteConFaltas.push(registroExistente);
        } else {
            // Si NO asistió, creamos la fila de falta
            reporteConFaltas.push({
                "ID de usuario": trabajador.id,
                "Nombre": trabajador.nombre,
                "Fecha/Hora": dia,
                "Departamento": trabajador.departamento,
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