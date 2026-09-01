"use client";

import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, Legend } from "recharts";
import type { CausalCount } from "@/types";

interface CausalesChartProps {
  data: CausalCount[];
}

const COLORS = ["#3B82F6", "#10B981", "#F59E0B", "#EF4444", "#8B5CF6", "#06B6D4"];

export function CausalesChart({ data }: CausalesChartProps) {
  return (
    <ResponsiveContainer width="100%" height={280}>
      <PieChart>
        <Pie
          data={data}
          cx="50%"
          cy="50%"
          innerRadius={50}
          outerRadius={90}
          paddingAngle={2}
          dataKey="cantidad"
        >
          {data.map((entry, index) => (
            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
          ))}
        </Pie>
        <Tooltip
          contentStyle={{
            backgroundColor: "white",
            border: "1px solid #E5E7EB",
            borderRadius: "8px",
            fontSize: "12px",
            boxShadow: "0 4px 6px -1px rgb(0 0 0 / 0.1)",
          }}
        />
        <Legend
          verticalAlign="bottom"
          height={50}
          iconType="circle"
          wrapperStyle={{ fontSize: "11px" }}
          formatter={(value) => {
            const item = data.find((d) => d.causal === value);
            return `${value} (${item?.cantidad || 0})`;
          }}
        />
      </PieChart>
    </ResponsiveContainer>
  );
}
