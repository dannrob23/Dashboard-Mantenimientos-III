"use client";

import { TrendingUp, TrendingDown, Minus } from "lucide-react";
import { cn } from "@/lib/utils";

interface KPICardProps {
  title: string;
  value: string | number;
  subtitle?: string;
  icon: React.ReactNode;
  trend?: "up" | "down" | "neutral";
  trendValue?: string;
  gradient: "blue" | "success" | "warning" | "danger";
  progress?: number;
}

export function KPICard({
  title,
  value,
  subtitle,
  icon,
  trend,
  trendValue,
  gradient,
  progress,
}: KPICardProps) {
  const gradientClass = {
    blue: "gradient-blue",
    success: "gradient-success",
    warning: "gradient-warning",
    danger: "gradient-danger",
  }[gradient];

  const TrendIcon = trend === "up" ? TrendingUp : trend === "down" ? TrendingDown : Minus;

  const trendColor =
    trend === "up" ? "text-success-600" : trend === "down" ? "text-danger-600" : "text-gray-500";

  return (
    <div className="card card-hover p-6 relative overflow-hidden animate-fade-in">
      {/* Icono de fondo decorativo */}
      <div
        className={cn(
          "absolute -top-4 -right-4 w-24 h-24 rounded-full opacity-10",
          gradientClass
        )}
      />

      <div className="relative z-10">
        <div className="flex items-center justify-between mb-4">
          <div
            className={cn(
              "w-12 h-12 rounded-xl flex items-center justify-center text-white shadow-md",
              gradientClass
            )}
          >
            {icon}
          </div>
          {trend && trendValue && (
            <div
              className={cn(
                "flex items-center gap-1 px-2 py-1 rounded-md text-xs font-medium bg-gray-50",
                trendColor
              )}
            >
              <TrendIcon className="w-3 h-3" />
              {trendValue}
            </div>
          )}
        </div>

        <div>
          <p className="text-sm font-medium text-gray-600 mb-1">{title}</p>
          <p className="text-3xl font-bold text-gray-900 tracking-tight">{value}</p>
          {subtitle && <p className="text-xs text-gray-500 mt-1.5">{subtitle}</p>}
        </div>

        {progress !== undefined && (
          <div className="mt-4">
            <div className="h-2 bg-gray-100 rounded-full overflow-hidden">
              <div
                className={cn("h-full rounded-full transition-all duration-700", gradientClass)}
                style={{ width: `${Math.min(100, Math.max(0, progress))}%` }}
              />
            </div>
            <p className="text-xs text-gray-500 mt-1.5">
              {progress.toFixed(1)}% completado
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
