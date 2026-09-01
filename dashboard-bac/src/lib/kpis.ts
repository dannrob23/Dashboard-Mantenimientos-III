import type {
  Mantenimiento,
  KPIsTier1,
  EstadoCount,
  RegionCount,
  CausalCount,
  TecnicoRanking,
  SedeRiesgo,
  EstadoMantenimiento,
} from "@/types";

// =====================================================
// Tier 1: KPIs Esenciales
// =====================================================

export function calcularKPIsTier1(data: Mantenimiento[]): KPIsTier1 {
  const total = data.length;

  // Cumplimiento: Finalizada + Sede_Cerrada / Total
  const finalizadas = data.filter(
    (m) => m.Estado_Mantenimiento === "Finalizada" || m.Estado_Mantenimiento === "Sede_Cerrada"
  ).length;
  const cumplimiento = total > 0 ? (finalizadas / total) * 100 : 0;

  // Total equipos
  const totalEquipos = data.reduce((sum, m) => sum + (m.Cantidad_Equipos || 0), 0);

  // Equipos intervenidos (Finalizada + Sede_Cerrada)
  const equiposIntervenidos = data
    .filter((m) => m.Estado_Mantenimiento === "Finalizada" || m.Estado_Mantenimiento === "Sede_Cerrada")
    .reduce((sum, m) => sum + (m.Cantidad_Equipos || 0), 0);

  // Promedio desviación
  const conDesviacion = data.filter((m) => m.Dias_Desviacion !== null && m.Dias_Desviacion > 0);
  const promedioDesviacion =
    conDesviacion.length > 0
      ? conDesviacion.reduce((sum, m) => sum + (m.Dias_Desviacion || 0), 0) / conDesviacion.length
      : 0;

  return {
    cumplimientoGeneral: Math.round(cumplimiento * 10) / 10,
    totalMantenimientos: total,
    totalEquipos,
    equiposIntervenidos,
    promedioDesviacion: Math.round(promedioDesviacion * 10) / 10,
  };
}

// =====================================================
// Tier 2: Distribución por Estado
// =====================================================

export function calcularEstados(data: Mantenimiento[]): EstadoCount[] {
  const total = data.length;
  const orden: EstadoMantenimiento[] = [
    "Programada",
    "En Proceso",
    "Finalizada",
    "Sede_Cerrada",
    "Reprogramada",
    "Cancelada",
  ];

  const counts = data.reduce((acc, m) => {
    acc[m.Estado_Mantenimiento] = (acc[m.Estado_Mantenimiento] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  return orden.map((estado) => ({
    estado,
    cantidad: counts[estado] || 0,
    porcentaje: total > 0 ? Math.round(((counts[estado] || 0) / total) * 1000) / 10 : 0,
  }));
}

// =====================================================
// Tier 2: Cumplimiento por Región
// =====================================================

export function calcularRegiones(data: Mantenimiento[]): RegionCount[] {
  const porRegion = data.reduce((acc, m) => {
    const region = m.Region || "Sin región";
    if (!acc[region]) {
      acc[region] = { total: 0, finalizadas: 0 };
    }
    acc[region].total += 1;
    if (m.Estado_Mantenimiento === "Finalizada" || m.Estado_Mantenimiento === "Sede_Cerrada") {
      acc[region].finalizadas += 1;
    }
    return acc;
  }, {} as Record<string, { total: number; finalizadas: number }>);

  return Object.entries(porRegion)
    .map(([region, vals]) => ({
      region,
      total: vals.total,
      finalizadas: vals.finalizadas,
      cumplimiento: vals.total > 0 ? Math.round((vals.finalizadas / vals.total) * 1000) / 10 : 0,
    }))
    .sort((a, b) => b.cumplimiento - a.cumplimiento);
}

// =====================================================
// Tier 2: Top Causales
// =====================================================

export function calcularCausales(data: Mantenimiento[]): CausalCount[] {
  const total = data.length;
  const counts = data.reduce((acc, m) => {
    const causal = m.Causal_Desviacion || "Sin causal";
    acc[causal] = (acc[causal] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  return Object.entries(counts)
    .map(([causal, cantidad]) => ({
      causal,
      cantidad,
      porcentaje: total > 0 ? Math.round((cantidad / total) * 1000) / 10 : 0,
    }))
    .sort((a, b) => b.cantidad - a.cantidad)
    .slice(0, 6);
}

// =====================================================
// Tier 2: Ranking Técnicos
// =====================================================

export function calcularTecnicos(data: Mantenimiento[]): TecnicoRanking[] {
  const porTecnico = data.reduce((acc, m) => {
    const tecnico = m.Ingeniero_BAC || "Sin asignar";
    if (!acc[tecnico]) {
      acc[tecnico] = { total: 0, finalizadas: 0 };
    }
    acc[tecnico].total += 1;
    if (m.Estado_Mantenimiento === "Finalizada" || m.Estado_Mantenimiento === "Sede_Cerrada") {
      acc[tecnico].finalizadas += 1;
    }
    return acc;
  }, {} as Record<string, { total: number; finalizadas: number }>);

  return Object.entries(porTecnico)
    .map(([tecnico, vals]) => ({
      tecnico,
      total: vals.total,
      finalizadas: vals.finalizadas,
      cumplimiento: vals.total > 0 ? Math.round((vals.finalizadas / vals.total) * 1000) / 10 : 0,
    }))
    .sort((a, b) => b.cumplimiento - a.cumplimiento)
    .slice(0, 10);
}

// =====================================================
// Tier 2: Sedes con Riesgo
// =====================================================

export function calcularSedesRiesgo(data: Mantenimiento[]): SedeRiesgo[] {
  return data
    .filter((m) => m.Estado_Mantenimiento !== "Sede_Cerrada" && m.Estado_Mantenimiento !== "Cancelada")
    .map((m) => ({
      sban: m.SBAN,
      oficina: m.Nombre_Oficina,
      municipio: m.Municipio,
      departamento: m.Departamento,
      region: m.Region,
      estado: m.Estado_Mantenimiento,
      diasDesviacion: m.Dias_Desviacion || 0,
      fechaProgramada: m.Fecha_Programada_Fin || m.Fecha_Programada_Inicio || "",
      equipos: m.Cantidad_Equipos || 0,
    }))
    .sort((a, b) => b.diasDesviacion - a.diasDesviacion)
    .slice(0, 10);
}
