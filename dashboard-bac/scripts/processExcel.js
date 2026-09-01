// Script para procesar Excel y generar JSON
// Ejecutar: node scripts/processExcel.js
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const INPUT_PATH = path.join(__dirname, "..", "data", "raw", "Bitacora_Final_Dashboard_BAC.xlsx");
const OUTPUT_PATH = path.join(__dirname, "..", "data", "processed", "bitacora.json");

function parseFecha(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return valor.toISOString().split("T")[0];
  }
  if (typeof valor === "string") {
    const parts = valor.split("/");
    if (parts.length === 3) {
      const [d, m, y] = parts;
      const date = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
      if (!isNaN(date.getTime())) {
        return date.toISOString().split("T")[0];
      }
    }
    return valor;
  }
  if (typeof valor === "number") {
    const date = XLSX.SSF.parse_date_code(valor);
    if (date) {
      return new Date(date.y, date.m - 1, date.d).toISOString().split("T")[0];
    }
  }
  return null;
}

function parseNumero(valor) {
  if (typeof valor === "number") return valor;
  if (typeof valor === "string") {
    const num = parseFloat(valor);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

function procesarExcel() {
  try {
    console.log("Procesando Excel...");

    if (!fs.existsSync(INPUT_PATH)) {
      console.error(`No se encontró el archivo: ${INPUT_PATH}`);
      process.exit(1);
    }

    const workbook = XLSX.readFile(INPUT_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const sban = parseNumero(row[0]);
      if (!sban) continue;

      data.push({
        SBAN: sban,
        Nombre_Oficina: String(row[1] || "").trim(),
        Municipio: String(row[2] || "").trim(),
        Departamento: String(row[3] || "").trim(),
        Region: String(row[4] || "").trim(),
        Ingeniero_BAC: String(row[5] || "").trim(),
        Tecnico_Calidad: String(row[6] || "PMU COLSOF").trim(),
        Fecha_Programada_Inicio: parseFecha(row[7]),
        Fecha_Programada_Fin: parseFecha(row[8]),
        Fecha_Inicio_Real: parseFecha(row[9]),
        Fecha_Salida_Real: parseFecha(row[10]),
        Dias_Desviacion: parseNumero(row[11]) || null,
        Fecha_Cierre_Operativo: parseFecha(row[12]),
        Estado_Mantenimiento: String(row[13] || "Programada").trim(),
        Causal_Desviacion: row[14] ? String(row[14]).trim() : null,
        Estatus_Acta: row[15] ? String(row[15]).trim() : null,
        Fecha_Cierre_Administrativo: parseFecha(row[16]),
        Cantidad_Equipos: parseNumero(row[17]),
        Observaciones_Calidad: String(row[18] || "").trim(),
      });
    }

    // Crear directorio si no existe
    const outputDir = path.dirname(OUTPUT_PATH);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    // Escribir JSON
    const output = {
      ultimaActualizacion: new Date().toISOString(),
      totalRegistros: data.length,
      registros: data,
    };

    fs.writeFileSync(OUTPUT_PATH, JSON.stringify(output, null, 2));
    console.log(`Procesados ${data.length} registros`);
    console.log(`JSON generado en: ${OUTPUT_PATH}`);
  } catch (error) {
    console.error("Error al procesar Excel:", error);
    process.exit(1);
  }
}

procesarExcel();
