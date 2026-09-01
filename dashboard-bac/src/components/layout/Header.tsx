"use client";

import { Building2, RefreshCw, Download, Calendar } from "lucide-react";
import { Button } from "@/components/ui/Button";

interface HeaderProps {
  ultimaActualizacion: string;
  onRefresh?: () => void;
}

export function Header({ ultimaActualizacion, onRefresh }: HeaderProps) {
  return (
    <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-md border-b border-gray-200">
      <div className="max-w-[1440px] mx-auto px-6 py-4">
        <div className="flex items-center justify-between">
          {/* Logo y Título */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl gradient-blue flex items-center justify-center shadow-md">
              <Building2 className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-gray-900">
                Mantenimientos Preventivos
              </h1>
              <p className="text-xs text-gray-500">Banco Agrario de Colombia</p>
            </div>
          </div>

          {/* Info y Acciones */}
          <div className="flex items-center gap-3">
            <div className="hidden md:flex items-center gap-2 px-3 py-2 bg-gray-50 rounded-lg">
              <Calendar className="w-4 h-4 text-gray-500" />
              <span className="text-xs text-gray-600">
                Actualizado: {ultimaActualizacion}
              </span>
            </div>

            <Button variant="outline" size="sm" onClick={onRefresh}>
              <RefreshCw className="w-4 h-4 mr-1.5" />
              Actualizar
            </Button>

            <Button variant="primary" size="sm">
              <Download className="w-4 h-4 mr-1.5" />
              Exportar
            </Button>
          </div>
        </div>
      </div>
    </header>
  );
}
