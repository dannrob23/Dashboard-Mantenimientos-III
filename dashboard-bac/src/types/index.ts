// Tipos basados en Bitacora_Simplificada_Dashboard_BAC.xlsx

export type EstadoMantenimiento =
  | "Programada"
  | "En Proceso"
  | "Finalizada"
  | "Sede_Cerrada"
  | "Reprogramada"
  | "Cancelada";

export type EstatusActa = "Pendiente Firma" | "Firmada";

export type CausalDesviacion =
  | "N/A - A Tiempo"
  | "Repuesto Faltante"
  | "Director Oficina No Disponible"
  | "Clima"
  | "Falla Técnica"
  | "Reprogramación Cliente"
  | "Logística/Acceso";

export type Region =
  | "Oriente"
  | "Occidente"
  | "Costa"
  | "Sur"
  | "Antioquia"
  | "Cafetera"
  | "Santanderes"
  | "Bogota";

export interface Mantenimiento {
  SBAN: number;
  Nombre_Oficina: string;
  Municipio: string;
  Departamento: string;
  Region: Region | string;
  Ingeniero_BAC: string;
  Tecnico_Calidad: string;
  Fecha_Programada_Inicio: string | null;
  Fecha_Programada_Fin: string | null;
  Fecha_Inicio_Real: string | null;
  Fecha_Salida_Real: string | null;
  Dias_Desviacion: number | null;
  Fecha_Cierre_Operativo: string | null;
  Estado_Mantenimiento: EstadoMantenimiento;
  Causal_Desviacion: CausalDesviacion | string | null;
  Estatus_Acta: EstatusActa | null;
  Fecha_Cierre_Administrativo: string | null;
  Cantidad_Equipos: number;
  Observaciones_Calidad: string;
}

// =====================================================
// KPIs - Tier 1 (Esenciales)
// =====================================================

export interface KPIsTier1 {
  cumplimientoGeneral: number;        // % de cumplimiento (0-100)
  totalMantenimientos: number;        // Total de registros
  totalEquipos: number;              // Total de equipos sumados
  equiposIntervenidos: number;       // Equipos en estados finalizados
  promedioDesviacion: number;         // Promedio días desviación
}

// =====================================================
// KPIs - Tier 2 (Operativos)
// =====================================================

export interface EstadoCount {
  estado: EstadoMantenimiento;
  cantidad: number;
  porcentaje: number;
}

export interface RegionCount {
  region: string;
  total: number;
  finalizadas: number;
  cumplimiento: number;
}

export interface CausalCount {
  causal: string;
  cantidad: number;
  porcentaje: number;
}

export interface TecnicoRanking {
  tecnico: string;
  total: number;
  finalizadas: number;
  cumplimiento: number;
}

export interface SedeRiesgo {
  sban: number;
  oficina: string;
  municipio: string;
  departamento: string;
  region: string;
  estado: string;
  diasDesviacion: number;
  fechaProgramada: string;
  equipos: number;
}

// =====================================================
// Datos procesados completos
// =====================================================

export interface DashboardData {
  kpis: KPIsTier1;
  estados: EstadoCount[];
  regiones: RegionCount[];
  causales: CausalCount[];
  tecnicos: TecnicoRanking[];
  sedesRiesgo: SedeRiesgo[];
  ultimaActualizacion: string;
  totalRegistros: number;
}

// =====================================================
// Filtros
// =====================================================

export interface DashboardFilters {
  region: string | null;
  estado: EstadoMantenimiento | null;
  fechaDesde: string | null;
  fechaHasta: string | null;
  ingeniero: string | null;
}
