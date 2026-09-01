// Script para procesar CSV y generar JSON para el dashboard
// Ejecutar: node scripts/processCsv.js
const fs = require("fs");
const path = require("path");

const INPUT_PATH = path.join(__dirname, "..", "data", "raw", "CONTROL DE EQUIPOS MANTENIMIENTO(Hoja3).csv");
const OUTPUT_PATH = path.join(__dirname, "..", "data", "processed", "bitacora.json");
const PUBLIC_OUTPUT_PATH = path.join(__dirname, "..", "public", "data", "bitacora.json");

// Mapeo de meses
const MESES = {
  "January": "01", "February": "02", "March": "03", "April": "04",
  "May": "05", "June": "06", "July": "07", "August": "08",
  "September": "09", "October": "10", "November": "11", "December": "12",
};

function parseFechaIngles(fechaStr) {
  if (!fechaStr || fechaStr.trim() === "") return null;
  // Formato: "Wednesday, October 7, 2026"
  const partes = fechaStr.trim().split(", ");
  if (partes.length !== 3) return null;
  
  const mesDia = partes[1].split(" ");
  if (mesDia.length !== 2) return null;
  
  const mes = MESES[mesDia[0]];
  if (!mes) return null;
  
  const dia = mesDia[1].padStart(2, "0");
  const anio = partes[2];
  
  return `${anio}-${mes}-${dia}`;
}

function parseCSV() {
  console.log("Procesando CSV...");
  
  if (!fs.existsSync(INPUT_PATH)) {
    console.error(`No se encontró el archivo: ${INPUT_PATH}`);
    process.exit(1);
  }

  const content = fs.readFileSync(INPUT_PATH, "utf-8");
  const lines = content.split("\n");
  
  const data = [];
  let procesados = 0;
  let saltados = 0;

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    
    // Saltar líneas que son solo comas vacías o leyendas de estados
    if (line.startsWith(",") && line.split(",").every(c => c.trim() === "")) {
      saltados++;
      continue;
    }
    
    if (line.includes("Programada") || line.includes("En proceso") || 
        line.includes("Finalizada") || line.includes("Reprogramada") || 
        line.includes("Cerrada") || line.includes("Cancelada")) {
      saltados++;
      continue;
    }

    // Parsear CSV considerando comas dentro de comillas
    const cols = [];
    let current = "";
    let inQuotes = false;
    
    for (let j = 0; j < line.length; j++) {
      const char = line[j];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === "," && !inQuotes) {
        cols.push(current);
        current = "";
      } else {
        current += char;
      }
    }
    cols.push(current); // última columna

    // Mapear columnas del CSV
    // 0: INGENIERO BAC, 1: INGENIERO COLSOF, 2: SBAN, 3: TIPO, 4: Nombre Oficina
    // 5: MJNICIPIO, 6: DEPARTAMENTO, 7: Jefaturas Operaciones Regional, 8: ALIADO
    // 9: OBSERVACIONES, 10: FECHA INICIO CRONOGRAMANA, 11: FECHA FIN CRONOGRMA
    // 12-17: Campos vacíos (estados, fechas reales, etc.)

    if (cols.length < 12) {
      saltados++;
      continue;
    }

    const sban = parseInt(cols[2].trim(), 10);
    if (isNaN(sban) || sban === 0) {
      saltados++;
      continue;
    }

    const fechaInicioProg = parseFechaIngles(cols[10]);
    const fechaFinProg = parseFechaIngles(cols[11]);

    // Solo procesar si tiene fechas válidas
    if (!fechaInicioProg && !fechaFinProg) {
      saltados++;
      continue;
    }

    data.push({
      SBAN: sban,
      Nombre_Oficina: cols[4].trim(),
      Municipio: cols[5].trim(),
      Departamento: cols[6].trim(),
      Region: cols[7].trim(),
      Ingeniero_BAC: cols[1].trim() || "SIN ASIGNAR",
      Tecnico_Calidad: "PMU COLSOF",
      Fecha_Programada_Inicio: fechaInicioProg,
      Fecha_Programada_Fin: fechaFinProg,
      Fecha_Inicio_Real: null,
      Fecha_Salida_Real: null,
      Dias_Desviacion: null,
      Fecha_Cierre_Operativo: null,
      Estado_Mantenimiento: "Programada",
      Causal_Desviacion: null,
      Estatus_Acta: null,
      Fecha_Cierre_Administrativo: null,
      Cantidad_Equipos: 0,
      Observaciones_Calidad: cols[9] ? cols[9].trim() : "",
    });

    procesados++;
  }

  console.log(`Registros procesados: ${procesados}`);
  console.log(`Registros saltados: ${saltados}`);

  // Crear directorios si no existen
  const outputDir = path.dirname(OUTPUT_PATH);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const publicDir = path.dirname(PUBLIC_OUTPUT_PATH);
  if (!fs.existsSync(publicDir)) {
    fs.mkdirSync(publicDir, { recursive: true });
  }

  // Escribir JSON
  const output = {
    ultimaActualizacion: new Date().toISOString(),
    totalRegistros: data.length,
    registros: data,
  };

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(output, null, 2));
  fs.writeFileSync(PUBLIC_OUTPUT_PATH, JSON.stringify(output, null, 2));
  
  console.log(`JSON generado en: ${OUTPUT_PATH}`);
  console.log(`JSON generado en: ${PUBLIC_OUTPUT_PATH}`);
  console.log(`Total registros: ${data.length}`);
}

parseCSV();