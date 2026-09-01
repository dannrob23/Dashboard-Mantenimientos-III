import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Formato de números colombianos
export function formatNumber(num: number): string {
  return new Intl.NumberFormat("es-CO").format(num);
}

// Formato de porcentaje
export function formatPercent(num: number, decimals: number = 1): string {
  return `${num.toFixed(decimals)}%`;
}

// Formato de fecha
export function formatDate(date: string | Date | null): string {
  if (!date) return "—";
  const d = typeof date === "string" ? new Date(date) : date;
  if (isNaN(d.getTime())) return "—";
  return new Intl.DateTimeFormat("es-CO", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(d);
}

// Formato de fecha relativa
export function formatRelativeDate(date: string | Date | null): string {
  if (!date) return "—";
  const d = typeof date === "string" ? new Date(date) : date;
  if (isNaN(d.getTime())) return "—";
  return new Intl.DateTimeFormat("es-CO", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(d);
}
