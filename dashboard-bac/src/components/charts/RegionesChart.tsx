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
import type { RegionCount } from "@/types";

interface RegionesChartProps {
  data: RegionCount[];
}

export function RegionesChart({ data }: RegionesChartProps) {
  const getColor = (pct: number) => {
    if (pct >= 80) return "#10B981";
    if (pct >= 60) return "#F59E0B";
    return "#EF4444";
  };

  return (
    <ResponsiveContainer width="100%" height={300}>
      <BarChart
        data={data}
        layout="vertical"
        margin={{ top: 5, right: 30, left: 80, bottom: 5 }}
      >
        <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6" horizontal={false} />
        <XAxis
          type="number"
          domain={[0, 100]}
          tick={{ fontSize: 11, fill: "#6B7280" }}
          axisLine={false}
          tickLine={false}
        />
        <YAxis
          type="category"
          dataKey="region"
          tick={{ fontSize: 11, fill: "#6B7280" }}
          axisLine={false}
          tickLine={false}
          width={75}
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
          formatter={(value: number) => [`${value}%`, "Cumplimiento"]}
        />
        <Bar dataKey="cumplimiento" radius={[0, 6, 6, 0]}>
          {data.map((entry, index) => (
            <Cell key={`cell-${index}`} fill={getColor(entry.cumplimiento)} />
          ))}
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}
