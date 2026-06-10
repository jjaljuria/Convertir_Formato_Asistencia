import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { generarReporte } from './index.js';
import ExcelJS from 'exceljs'; // Importar después de mockear para usar el mock
import xlsx from 'xlsx';
import fs from 'node:fs';

// Mock de las librerías externas (deben declararse antes de importar el módulo bajo prueba)
vi.mock('node:fs', async (importOrigin) => {
    const actual = await importOrigin();
    const existsSyncMock = vi.fn();
    return {
        ...actual,
        default: {
            ...actual.default,
            existsSync: existsSyncMock,
        },
        existsSync: existsSyncMock,
    };
});

vi.mock('xlsx', async (importOriginal) => {
    const actual = await importOriginal();
    const readFileMock = vi.fn();
    const sheetToJsonMock = vi.fn();
    return {
        ...actual,
        default: {
            ...(actual && actual.default ? actual.default : {}),
            readFile: readFileMock,
            utils: {
                ...((actual && actual.default && actual.default.utils) || {}),
                sheet_to_json: sheetToJsonMock,
            },
        },
        readFile: readFileMock,
        utils: {
            sheet_to_json: sheetToJsonMock,
        },
    };
});

vi.mock('exceljs', () => {
    const mockWorksheet = {
        columns: [],
        getRow: vi.fn(() => ({
            font: {},
            fill: {},
        })),
        addRow: vi.fn((rowData) => ({
            getCell: vi.fn((cell) => ({
                font: {},
            })),
            ...rowData,
        })),
    };

    const mockWorkbookInstance = {
        xlsx: {
            writeFile: vi.fn(),
        },
        addWorksheet: vi.fn(() => mockWorksheet),
    };

    // Use a regular (non-arrow) function wrapped with vi.fn so it can be used as a constructor
    const MockedWorkbook = vi.fn(function MockedWorkbook() {
        return mockWorkbookInstance;
    });

    return {
        __esModule: true,
        default: { Workbook: MockedWorkbook },
        Workbook: MockedWorkbook, // For destructured imports
    };
});



describe('generarReporte', () => {
    // Limpiar mocks antes de cada test
    beforeEach(() => {
        vi.clearAllMocks();
    });

    it('debería ser una función', () => {
        expect(typeof generarReporte).toBe('function');
    });

    it('debería registrar un error si falta rutaEntrada', async () => {
        const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => { });
        const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => { });

        await generarReporte(undefined, 'ruta/salida.xlsx');

        expect(consoleErrorSpy).toHaveBeenCalledWith("❌ ERROR: Faltan argumentos.");
        expect(consoleLogSpy).toHaveBeenCalledWith("Uso correcto: node index.js <ruta_entrada> <ruta_salida>");
        consoleErrorSpy.mockRestore();
        consoleLogSpy.mockRestore();
    });

    it('debería registrar un error si falta rutaSalida', async () => {
        const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => { });
        const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => { });

        await generarReporte('ruta/entrada.xls', undefined);

        expect(consoleErrorSpy).toHaveBeenCalledWith("❌ ERROR: Faltan argumentos.");
        expect(consoleLogSpy).toHaveBeenCalledWith("Uso correcto: node index.js <ruta_entrada> <ruta_salida>");
        consoleErrorSpy.mockRestore();
        consoleLogSpy.mockRestore();
    });

    it('debería registrar un error si el archivo de entrada no existe', async () => {
        const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => { });
        fs.existsSync.mockReturnValue(false);

        await generarReporte('ruta/entrada.xls', 'ruta/salida.xlsx');

        expect(fs.existsSync).toHaveBeenCalledWith('ruta/entrada.xls');
        expect(consoleErrorSpy).toHaveBeenCalledWith("❌ ERROR: El archivo de entrada no existe en: ruta/entrada.xls");
        consoleErrorSpy.mockRestore();
    });

    it('debería registrar un error si el archivo de entrada está vacío', async () => {
        const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => { });
        fs.existsSync.mockReturnValue(true);
        xlsx.readFile.mockReturnValue({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
        xlsx.utils.sheet_to_json.mockReturnValue([]);

        await generarReporte('ruta/entrada.xls', 'ruta/salida.xlsx');

        expect(xlsx.readFile).toHaveBeenCalledWith('ruta/entrada.xls', { cellDates: true });
        expect(xlsx.utils.sheet_to_json).toHaveBeenCalledWith(expect.any(Object));
        expect(consoleErrorSpy).toHaveBeenCalledWith("❌ ERROR CRÍTICO:", "El archivo de entrada está vacío o no se leyó correctamente.");
        consoleErrorSpy.mockRestore();
    });

    it('debería generar un reporte correctamente con datos válidos', async () => {
        const mockData = [
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-01T08:00:00.000Z'), 'Departamento': 'IT' },
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-01T17:00:00.000Z'), 'Departamento': 'IT' },
            { 'ID de usuario': 2, 'Nombre': 'Maria', 'Fecha/Hora': new Date('2026-04-01T09:00:00.000Z'), 'Departamento': 'HR' },
        ];
        const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => { });

        fs.existsSync.mockReturnValue(true);
        xlsx.readFile.mockReturnValue({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
        xlsx.utils.sheet_to_json.mockReturnValue(mockData);

        await generarReporte('ruta/entrada.xls', 'ruta/salida.xlsx');

        expect(fs.existsSync).toHaveBeenCalledWith('ruta/entrada.xls');
        expect(xlsx.readFile).toHaveBeenCalledWith('ruta/entrada.xls', { cellDates: true });
        expect(xlsx.utils.sheet_to_json).toHaveBeenCalledWith(expect.any(Object));
        expect(ExcelJS.Workbook).toHaveBeenCalledTimes(1);
        const mockWorkbookInstance = ExcelJS.Workbook.mock.results[0].value; // Obtener la instancia mockeada
        expect(mockWorkbookInstance.addWorksheet).toHaveBeenCalledWith('IT');
        expect(mockWorkbookInstance.addWorksheet).toHaveBeenCalledWith('HR');
        expect(mockWorkbookInstance.xlsx.writeFile).toHaveBeenCalledWith('ruta/salida.xlsx');
        expect(consoleLogSpy).toHaveBeenCalledWith("✨ ¡PROCESO COMPLETADO! El archivo se generó correctamente.");

        consoleLogSpy.mockRestore();
    });

    it('debería usar la configuración de usuario para los nombres de las columnas', async () => {
        const mockData = [
            { 'Custom ID': 1, 'Custom Name': 'Jose', 'Custom DateTime': new Date('2026-04-01T08:00:00.000Z'), 'Custom Dept': 'IT' },
        ];
        const usuarioConfig = {
            entrada: {
                id: 'Custom ID',
                nombre: 'Custom Name',
                fechaHora: 'Custom DateTime',
                departamento: 'Custom Dept'
            },
            salida: {
                id: 'ID Personal',
                nombre: 'Nombre Completo',
                fecha: 'Fecha de Asistencia',
                entrada: 'Hora de Entrada',
                salida: 'Hora de Salida',
                departamento: 'Departamento de la Empresa',
                estado: 'Estado de Asistencia'
            }
        };
        const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => { });

        fs.existsSync.mockReturnValue(true);
        xlsx.readFile.mockReturnValue({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
        xlsx.utils.sheet_to_json.mockReturnValue(mockData);

        await generarReporte('ruta/entrada.xls', 'ruta/salida.xlsx', usuarioConfig);

        expect(fs.existsSync).toHaveBeenCalledWith('ruta/entrada.xls');
        expect(xlsx.readFile).toHaveBeenCalledWith('ruta/entrada.xls', { cellDates: true });
        expect(xlsx.utils.sheet_to_json).toHaveBeenCalledWith(expect.any(Object));

        expect(ExcelJS.Workbook).toHaveBeenCalledTimes(1);
        const mockWorkbookInstance = ExcelJS.Workbook.mock.results[0].value; // Obtener la instancia mockeada
        expect(mockWorkbookInstance.addWorksheet).toHaveBeenCalledWith('IT');

        expect(mockWorkbookInstance.xlsx.writeFile).toHaveBeenCalledWith('ruta/salida.xlsx');
        expect(consoleLogSpy).toHaveBeenCalledWith("✨ ¡PROCESO COMPLETADO! El archivo se generó correctamente.");

        consoleLogSpy.mockRestore();
    });

    it('debería usar mensajes personalizados para falta y fin de semana', async () => {
        const mockData = [
            // 2026-04-01 (miércoles) con asistencia
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-01T08:00:00.000Z'), 'Departamento': 'IT' },
            // 2026-04-04 (sábado) con asistencia
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-04T08:00:00.000Z'), 'Departamento': 'IT' },
            // 2026-04-08 (miércoles) con asistencia
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-08T08:00:00.000Z'), 'Departamento': 'IT' },
            // 2026-04-10 (viernes) con asistencia
            { 'ID de usuario': 1, 'Nombre': 'Jose', 'Fecha/Hora': new Date('2026-04-10T08:00:00.000Z'), 'Departamento': 'IT' },
        ];
        const usuarioConfig = {
            mensajes: {
                falto: 'NO ASISTIÓ',
                finDeSemana: 'WEEKEND'
            }
        };
        const consoleLogSpy = vi.spyOn(console, 'log').mockImplementation(() => { });

        fs.existsSync.mockReturnValue(true);
        xlsx.readFile.mockReturnValue({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } });
        xlsx.utils.sheet_to_json.mockReturnValue(mockData);

        await generarReporte('ruta/entrada.xls', 'ruta/salida.xlsx', usuarioConfig);

        expect(fs.existsSync).toHaveBeenCalledWith('ruta/entrada.xls');
        const mockWorkbookInstance = ExcelJS.Workbook.mock.results[0].value;
        
        // Verificar que addRow fue llamado con datos que contienen los mensajes personalizados
        const addRowCalls = mockWorkbookInstance.addWorksheet.mock.results[0].value.addRow.mock.calls;
        
        // Buscar filas con ausencia
        // Días laborales sin asistencia (02, 03, 06, 07, 09) = 'NO ASISTIÓ'
        // Fines de semana sin asistencia (05, 11, 12) = 'WEEKEND'
        expect(addRowCalls.some(call => call[0].estado === 'NO ASISTIÓ')).toBe(true);
        expect(addRowCalls.some(call => call[0].estado === 'WEEKEND')).toBe(true);
        
        expect(mockWorkbookInstance.xlsx.writeFile).toHaveBeenCalledWith('ruta/salida.xlsx');
        expect(consoleLogSpy).toHaveBeenCalledWith("✨ ¡PROCESO COMPLETADO! El archivo se generó correctamente.");

        consoleLogSpy.mockRestore();
    });
});
