import type { Mantenimiento, EstadoMantenimiento, EstatusActa } from "@/types";

function parseFecha(valor: unknown): string | null {
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
  return null;
}

function parseNumero(valor: unknown): number {
  if (typeof valor === "number") return valor;
  if (typeof valor === "string") {
    const num = parseFloat(valor);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

function parseEstado(valor: unknown): EstadoMantenimiento {
  const s = String(valor || "").trim();
  const estadosValidos: EstadoMantenimiento[] = [
    "Programada",
    "En Proceso",
    "Finalizada",
    "Sede_Cerrada",
    "Reprogramada",
    "Cancelada",
  ];
  if (estadosValidos.includes(s as EstadoMantenimiento)) {
    return s as EstadoMantenimiento;
  }
  return "Programada";
}

// Parser para datos en JSON (usado en el frontend)
export function parseBitacoraJSON(data: Record<string, unknown>[]): Mantenimiento[] {
  return data
    .filter((row) => row.SBAN)
    .map((row) => ({
      SBAN: parseNumero(row.SBAN),
      Nombre_Oficina: String(row.Nombre_Oficina || "").trim(),
      Municipio: String(row.Municipio || "").trim(),
      Departamento: String(row.Departamento || "").trim(),
      Region: String(row.Region || "").trim(),
      Ingeniero_BAC: String(row.Ingeniero_BAC || "").trim(),
      Tecnico_Calidad: String(row.Tecnico_Calidad || "PMU COLSOF").trim(),
      Fecha_Programada_Inicio: parseFecha(row.Fecha_Programada_Inicio),
      Fecha_Programada_Fin: parseFecha(row.Fecha_Programada_Fin),
      Fecha_Inicio_Real: parseFecha(row.Fecha_Inicio_Real),
      Fecha_Salida_Real: parseFecha(row.Fecha_Salida_Real),
      Dias_Desviacion: parseNumero(row.Dias_Desviacion) || null,
      Fecha_Cierre_Operativo: parseFecha(row.Fecha_Cierre_Operativo),
      Estado_Mantenimiento: parseEstado(row.Estado_Mantenimiento),
      Causal_Desviacion: row.Causal_Desviacion ? String(row.Causal_Desviacion).trim() : null,
      Estatus_Acta: row.Estatus_Acta ? (String(row.Estatus_Acta).trim() as EstatusActa) : null,
      Fecha_Cierre_Administrativo: parseFecha(row.Fecha_Cierre_Administrativo),
      Cantidad_Equipos: parseNumero(row.Cantidad_Equipos),
      Observaciones_Calidad: String(row.Observaciones_Calidad || "").trim(),
    }));
}
