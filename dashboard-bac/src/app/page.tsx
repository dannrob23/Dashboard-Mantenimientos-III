"use client";

import { useEffect, useState } from "react";
import {
  Target,
  ListTodo,
  Server,
  Clock,
  Activity,
  BarChart3,
  Map,
  AlertTriangle,
} from "lucide-react";
import { Header } from "@/components/layout/Header";
import { KPICard } from "@/components/cards/KPICard";
import { Card, CardHeader } from "@/components/ui/Card";
import { RadialRing } from "@/components/charts/RadialRing";
import { EstadosChart } from "@/components/charts/EstadosChart";
import { RegionesChart } from "@/components/charts/RegionesChart";
import { CausalesChart } from "@/components/charts/CausalesChart";
import { SedesRiesgoTable } from "@/components/tables/SedesRiesgoTable";
import { parseBitacoraJSON } from "@/lib/parser";
import {
  calcularKPIsTier1,
  calcularEstados,
  calcularRegiones,
  calcularCausales,
  calcularSedesRiesgo,
} from "@/lib/kpis";
import type { Mantenimiento } from "@/types";
import { formatNumber, formatPercent } from "@/lib/utils";

export default function DashboardPage() {
  const [data, setData] = useState<Mantenimiento[]>([]);
  const [loading, setLoading] = useState(true);
  const [ultimaActualizacion, setUltimaActualizacion] = useState("");

  useEffect(() => {
    cargarDatos();
  }, []);

  async function cargarDatos() {
    try {
      setLoading(true);
      const response = await fetch("/data/bitacora.json");
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: No se pudieron cargar los datos`);
      }
      const json = await response.json();
      const registros = parseBitacoraJSON(json.registros || []);
      setData(registros);
      if (json.ultimaActualizacion) {
        setUltimaActualizacion(
          new Date(json.ultimaActualizacion).toLocaleDateString("es-CO", {
            day: "2-digit",
            month: "short",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
          })
        );
      } else {
        setUltimaActualizacion(new Date().toLocaleDateString("es-CO"));
      }
    } catch (error) {
      console.error("Error cargando datos:", error);
      setData([]);
      setUltimaActualizacion("Sin datos");
    } finally {
      setLoading(false);
    }
  }

  // Calcular KPIs
  const kpis = calcularKPIsTier1(data);
  const estados = calcularEstados(data);
  const regiones = calcularRegiones(data).slice(0, 8);
  const causales = calcularCausales(data);
  const sedesRiesgo = calcularSedesRiesgo(data);

  const avanceEquipos =
    kpis.totalEquipos > 0 ? (kpis.equiposIntervenidos / kpis.totalEquipos) * 100 : 0;

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <div className="text-center">
          <div className="w-12 h-12 border-4 border-primary-200 border-t-primary-600 rounded-full animate-spin mx-auto" />
          <p className="mt-4 text-gray-600">Cargando datos...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <Header ultimaActualizacion={ultimaActualizacion} onRefresh={cargarDatos} />

      <main className="max-w-[1440px] mx-auto px-6 py-8 space-y-6">
        {/* Título y descripción */}
        <div className="animate-slide-up">
          <h2 className="text-2xl font-bold text-gray-900">
            Resumen de Mantenimientos Preventivos
          </h2>
          <p className="text-sm text-gray-500 mt-1">
            Visualización en tiempo real del estado y cumplimiento de mantenimientos
          </p>
        </div>

        {/* KPI Cards - Tier 1 */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 animate-slide-up">
          <KPICard
            title="% Cumplimiento General"
            value={formatPercent(kpis.cumplimientoGeneral)}
            subtitle="Finalizadas + Cerradas"
            icon={<Target className="w-5 h-5" />}
            trend={kpis.cumplimientoGeneral >= 80 ? "up" : "down"}
            trendValue={kpis.cumplimientoGeneral >= 80 ? "Óptimo" : "Mejorar"}
            gradient="blue"
          />
          <KPICard
            title="Total Mantenimientos"
            value={formatNumber(kpis.totalMantenimientos)}
            subtitle="Registros en sistema"
            icon={<ListTodo className="w-5 h-5" />}
            gradient="success"
          />
          <KPICard
            title="% Avance Equipos"
            value={formatPercent(avanceEquipos)}
            subtitle={`${formatNumber(kpis.equiposIntervenidos)} / ${formatNumber(kpis.totalEquipos)} equipos`}
            icon={<Server className="w-5 h-5" />}
            gradient="warning"
            progress={avanceEquipos}
          />
          <KPICard
            title="Días Desviación"
            value={`${kpis.promedioDesviacion}d`}
            subtitle="Promedio de retraso"
            icon={<Clock className="w-5 h-5" />}
            trend={kpis.promedioDesviacion > 5 ? "down" : "up"}
            trendValue={kpis.promedioDesviacion > 5 ? "Atención" : "OK"}
            gradient="danger"
          />
        </div>

        {/* Sección 1: Cumplimiento y Estados */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4 animate-slide-up">
          <Card className="lg:col-span-1">
            <CardHeader
              title="Cumplimiento General"
              subtitle="Indicador principal"
              icon={<Activity className="w-5 h-5" />}
            />
            <div className="flex items-center justify-center py-4">
              <RadialRing percentage={kpis.cumplimientoGeneral} label="Cumplimiento" />
            </div>
            <div className="grid grid-cols-3 gap-2 mt-4 pt-4 border-t border-gray-100">
              <div className="text-center">
                <p className="text-xs text-gray-500">Finalizadas</p>
                <p className="text-lg font-bold text-success-600">
                  {estados.find((e) => e.estado === "Finalizada")?.cantidad || 0}
                </p>
              </div>
              <div className="text-center">
                <p className="text-xs text-gray-500">En Proceso</p>
                <p className="text-lg font-bold text-primary-600">
                  {estados.find((e) => e.estado === "En Proceso")?.cantidad || 0}
                </p>
              </div>
              <div className="text-center">
                <p className="text-xs text-gray-500">Programadas</p>
                <p className="text-lg font-bold text-info-600">
                  {estados.find((e) => e.estado === "Programada")?.cantidad || 0}
                </p>
              </div>
            </div>
          </Card>

          <Card className="lg:col-span-2">
            <CardHeader
              title="Distribución por Estado"
              subtitle="Cantidad de mantenimientos"
              icon={<BarChart3 className="w-5 h-5" />}
            />
            <EstadosChart data={estados} />
          </Card>
        </div>

        {/* Sección 2: Regiones y Causales */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 animate-slide-up">
          <Card>
            <CardHeader
              title="Cumplimiento por Región"
              subtitle="% de cumplimiento"
              icon={<Map className="w-5 h-5" />}
            />
            <RegionesChart data={regiones} />
          </Card>

          <Card>
            <CardHeader
              title="Causales de Desviación"
              subtitle="Top motivos"
              icon={<AlertTriangle className="w-5 h-5" />}
            />
            <CausalesChart data={causales} />
          </Card>
        </div>

        {/* Sección 3: Sedes con Riesgo */}
        <Card className="animate-slide-up">
          <CardHeader
            title="Sedes con Mayor Riesgo"
            subtitle="Top 10 con más días de desviación"
            icon={<AlertTriangle className="w-5 h-5" />}
          />
          <SedesRiesgoTable data={sedesRiesgo} />
        </Card>

        {/* Footer */}
        <footer className="text-center py-6 text-xs text-gray-400">
          Dashboard Mantenimientos Preventivos BAC • v1.0.0 • Datos en tiempo real
        </footer>
      </main>
    </div>
  );
}
