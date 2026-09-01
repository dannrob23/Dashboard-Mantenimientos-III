// Script para generar datos de prueba mientras se procesa el Excel real
const fs = require("fs");
const path = require("path");

const ESTADOS = ["Programada", "En Proceso", "Finalizada", "Sede_Cerrada", "Reprogramada", "Cancelada"];
const CAUSALES = [
  "N/A - A Tiempo",
  "Repuesto Faltante",
  "Director Oficina No Disponible",
  "Clima",
  "Falla Técnica",
  "Reprogramación Cliente",
  "Logística/Acceso",
];
const ESTATUS_ACTA = ["Pendiente Firma", "Firmada"];
const REGIONES = ["Oriente", "Occidente", "Costa", "Sur", "Antioquia", "Cafetera", "Bogota", "Santanderes"];
const INGENIEROS = ["Yeferson", "Sergio", "Laura", "Santiago", "Yuliana", "Luis", "Andres"];

const CIUDADES = [
  { nombre: "Bogota Centro", municipio: "Bogotá", departamento: "Cundinamarca" },
  { nombre: "Soacha", municipio: "Soacha", departamento: "Cundinamarca" },
  { nombre: "Medellin", municipio: "Medellín", departamento: "Antioquia" },
  { nombre: "Cali", municipio: "Cali", departamento: "Valle del Cauca" },
  { nombre: "Barranquilla", municipio: "Barranquilla", departamento: "Atlántico" },
  { nombre: "Cartagena", municipio: "Cartagena", departamento: "Bolívar" },
  { nombre: "Bucaramanga", municipio: "Bucaramanga", departamento: "Santander" },
  { nombre: "Pereira", municipio: "Pereira", departamento: "Risaralda" },
  { nombre: "Manizales", municipio: "Manizales", departamento: "Caldas" },
  { nombre: "Ibague", municipio: "Ibagué", departamento: "Tolima" },
  { nombre: "Tunja", municipio: "Tunja", departamento: "Boyacá" },
  { nombre: "Pasto", municipio: "Pasto", departamento: "Nariño" },
  { nombre: "Neiva", municipio: "Neiva", departamento: "Huila" },
  { nombre: "Popayan", municipio: "Popayán", departamento: "Cauca" },
  { nombre: "Valledupar", municipio: "Valledupar", departamento: "Cesar" },
  { nombre: "Monteria", municipio: "Montería", departamento: "Córdoba" },
  { nombre: "Sincelejo", municipio: "Sincelejo", departamento: "Sucre" },
  { nombre: "Villavicencio", municipio: "Villavicencio", departamento: "Meta" },
  { nombre: "Yopal", municipio: "Yopal", departamento: "Casanare" },
  { nombre: "Arauca", municipio: "Arauca", departamento: "Arauca" },
  { nombre: "Mocoa", municipio: "Mocoa", departamento: "Putumayo" },
  { nombre: "Leticia", municipio: "Leticia", departamento: "Amazonas" },
  { nombre: "San Andres", municipio: "San Andrés", departamento: "San Andrés" },
  { nombre: "Puerto Inirida", municipio: "Inírida", departamento: "Guainía" },
];

function fechaAleatoria(inicio, fin) {
  const fecha = new Date(inicio.getTime() + Math.random() * (fin.getTime() - inicio.getTime()));
  return fecha.toISOString().split("T")[0];
}

function generarDatos(numRegistros = 200) {
  const registros = [];
  const fechaInicio = new Date("2026-08-01");
  const fechaFin = new Date("2026-12-31");

  for (let i = 0; i < numRegistros; i++) {
    const ciudad = CIUDADES[Math.floor(Math.random() * CIUDADES.length)];
    const estado = ESTADOS[Math.floor(Math.random() * ESTADOS.length)];
    const region = REGIONES[Math.floor(Math.random() * REGIONES.length)];
    const ingeniero = INGENIEROS[Math.floor(Math.random() * INGENIEROS.length)];

    // Sesgo: más probabilidad de estados pasados
    let estadoFinal = estado;
    if (Math.random() < 0.6) {
      estadoFinal = Math.random() < 0.5 ? "Finalizada" : "Sede_Cerrada";
    }

    const fechaProgInicio = fechaAleatoria(fechaInicio, fechaFin);
    const fechaFinProgramada = fechaAleatoria(new Date(fechaProgInicio), new Date(calcularFechaFin(fechaProgInicio)));

    let fechaInicioReal = null;
    let fechaSalidaReal = null;
    let fechaCierreOperativo = null;
    let fechaCierreAdmin = null;
    let diasDesviacion = 0;
    let estatusActa = null;
    let causal = null;

    if (estadoFinal === "Finalizada" || estadoFinal === "Sede_Cerrada") {
      fechaInicioReal = fechaAleatoria(new Date(fechaProgInicio), new Date(fechaFinProgramada));
      fechaSalidaReal = fechaAleatoria(new Date(fechaInicioReal), new Date(fechaFinProgramada));
      fechaCierreOperativo = fechaSalidaReal;
      diasDesviacion = Math.max(0, Math.floor((new Date(fechaSalidaReal).getTime() - new Date(fechaFinProgramada).getTime()) / (1000 * 60 * 60 * 24)));

      if (estadoFinal === "Sede_Cerrada") {
        fechaCierreAdmin = fechaAleatoria(new Date(fechaSalidaReal), new Date(fechaFinProgramada));
        estatusActa = "Firmada";
        causal = "N/A - A Tiempo";
      } else {
        estatusActa = Math.random() < 0.5 ? "Firmada" : "Pendiente Firma";
        causal = diasDesviacion === 0 ? "N/A - A Tiempo" : CAUSALES[Math.floor(Math.random() * (CAUSALES.length - 1)) + 1];
      }
    } else if (estadoFinal === "En Proceso") {
      fechaInicioReal = fechaProgInicio;
    } else if (estadoFinal === "Reprogramada") {
      causal = "Reprogramación Cliente";
    }

    registros.push({
      SBAN: 10 + i,
      Nombre_Oficina: ciudad.nombre,
      Municipio: ciudad.municipio,
      Departamento: ciudad.departamento,
      Region: region,
      Ingeniero_BAC: ingeniero,
      Tecnico_Calidad: "PMU COLSOF",
      Fecha_Programada_Inicio: fechaProgInicio,
      Fecha_Programada_Fin: fechaFinProgramada,
      Fecha_Inicio_Real: fechaInicioReal,
      Fecha_Salida_Real: fechaSalidaReal,
      Dias_Desviacion: diasDesviacion,
      Fecha_Cierre_Operativo: fechaCierreOperativo,
      Estado_Mantenimiento: estadoFinal,
      Causal_Desviacion: causal,
      Estatus_Acta: estatusActa,
      Fecha_Cierre_Administrativo: fechaCierreAdmin,
      Cantidad_Equipos: Math.floor(Math.random() * 20) + 5,
      Observaciones_Calidad: "",
    });
  }

  return registros;
}

function calcularFechaFin(fechaInicio) {
  const fecha = new Date(fechaInicio);
  fecha.setDate(fecha.getDate() + Math.floor(Math.random() * 5) + 1);
  return fecha;
}

// Generar y guardar
const registros = generarDatos(200);
const output = {
  ultimaActualizacion: new Date().toISOString(),
  totalRegistros: registros.length,
  registros,
};

const outputPath = path.join(__dirname, "..", "public", "data", "bitacora.json");
const outputDir = path.dirname(outputPath);
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

fs.writeFileSync(outputPath, JSON.stringify(output, null, 2));
console.log(`Generados ${registros.length} registros de prueba en ${outputPath}`);
