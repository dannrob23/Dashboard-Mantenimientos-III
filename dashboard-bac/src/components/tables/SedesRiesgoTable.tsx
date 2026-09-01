"use client";

import { AlertTriangle, MapPin } from "lucide-react";
import type { SedeRiesgo } from "@/types";

interface SedesRiesgoTableProps {
  data: SedeRiesgo[];
}

export function SedesRiesgoTable({ data }: SedesRiesgoTableProps) {
  const getEstadoColor = (estado: string) => {
    switch (estado) {
      case "Programada":
        return "bg-info-50 text-info-600";
      case "En Proceso":
        return "bg-primary-50 text-primary-600";
      case "Finalizada":
        return "bg-success-50 text-success-600";
      case "Sede_Cerrada":
        return "bg-purple-50 text-purple-600";
      case "Reprogramada":
        return "bg-warning-50 text-warning-600";
      default:
        return "bg-gray-50 text-gray-600";
    }
  };

  const getRiesgoColor = (dias: number) => {
    if (dias > 30) return "text-danger-600 font-semibold";
    if (dias > 15) return "text-warning-600 font-medium";
    return "text-gray-700";
  };

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-gray-200">
            <th className="text-left py-3 px-3 font-semibold text-gray-600 text-xs">SBAN</th>
            <th className="text-left py-3 px-3 font-semibold text-gray-600 text-xs">Oficina</th>
            <th className="text-left py-3 px-3 font-semibold text-gray-600 text-xs hidden md:table-cell">
              Ubicación
            </th>
            <th className="text-left py-3 px-3 font-semibold text-gray-600 text-xs">Estado</th>
            <th className="text-right py-3 px-3 font-semibold text-gray-600 text-xs">Días</th>
            <th className="text-right py-3 px-3 font-semibold text-gray-600 text-xs hidden md:table-cell">
              Equipos
            </th>
          </tr>
        </thead>
        <tbody>
          {data.length === 0 ? (
            <tr>
              <td colSpan={6} className="text-center py-8 text-gray-500 text-sm">
                No hay sedes con riesgo
              </td>
            </tr>
          ) : (
            data.map((sede, idx) => (
              <tr
                key={sede.sban}
                className="border-b border-gray-100 hover:bg-gray-50 transition-colors"
              >
                <td className="py-3 px-3">
                  <span className="font-mono text-xs text-gray-900 font-medium">
                    {sede.sban}
                  </span>
                </td>
                <td className="py-3 px-3">
                  <div>
                    <p className="font-medium text-gray-900 text-sm">{sede.oficina}</p>
                    <p className="text-xs text-gray-500 md:hidden flex items-center gap-1 mt-0.5">
                      <MapPin className="w-3 h-3" />
                      {sede.municipio}
                    </p>
                  </div>
                </td>
                <td className="py-3 px-3 text-gray-600 text-xs hidden md:table-cell">
                  <div className="flex items-center gap-1">
                    <MapPin className="w-3 h-3 text-gray-400" />
                    <span>{sede.municipio}, {sede.departamento}</span>
                  </div>
                </td>
                <td className="py-3 px-3">
                  <span
                    className={`inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium ${getEstadoColor(
                      sede.estado
                    )}`}
                  >
                    {sede.estado.replace("_", " ")}
                  </span>
                </td>
                <td className="py-3 px-3 text-right">
                  <div className="flex items-center justify-end gap-1">
                    {sede.diasDesviacion > 15 && (
                      <AlertTriangle className="w-3 h-3 text-danger-500" />
                    )}
                    <span className={getRiesgoColor(sede.diasDesviacion)}>
                      {sede.diasDesviacion}
                    </span>
                  </div>
                </td>
                <td className="py-3 px-3 text-right text-gray-700 text-sm hidden md:table-cell">
                  {sede.equipos}
                </td>
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  );
}
