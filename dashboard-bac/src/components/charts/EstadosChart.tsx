"use client";

import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
} from "recharts";
import type { EstadoCount } from "@/types";

interface EstadosChartProps {
  data: EstadoCount[];
}

const COLORES_ESTADO: Record<string, string> = {
  Programada: "#06B6D4",
  "En Proceso": "#3B82F6",
  Finalizada: "#10B981",
  Sede_Cerrada: "#8B5CF6",
  Reprogramada: "#F59E0B",
  Cancelada: "#EF4444",
};

export function EstadosChart({ data }: EstadosChartProps) {
  return (
    <ResponsiveContainer width="100%" height={280}>
      <BarChart data={data} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
        <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6" vertical={false} />
        <XAxis
          dataKey="estado"
          tick={{ fontSize: 11, fill: "#6B7280" }}
          axisLine={{ stroke: "#E5E7EB" }}
          tickLine={false}
        />
        <YAxis
          tick={{ fontSize: 11, fill: "#6B7280" }}
          axisLine={false}
          tickLine={false}
        />
        <Tooltip
          contentStyle={{
            backgroundColor: "white",
            border: "1px solid #E5E7EB",
            borderRadius: "8px",
            fontSize: "12px",
            boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)",
          }}
          cursor={{ fill: "#F9FAFB" }}
        />
        <Bar dataKey="cantidad" radius={[6, 6, 0, 0]}>
          {data.map((entry, index) => (
            <Cell key={`cell-${index}`} fill={COLORES_ESTADO[entry.estado] || "#6B7280"} />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}
