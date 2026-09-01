"use client";

import { useEffect, useState } from "react";

interface RadialRingProps {
  percentage: number;
  label: string;
  size?: number;
  strokeWidth?: number;
}

export function RadialRing({
  percentage,
  label,
  size = 180,
  strokeWidth = 14,
}: RadialRingProps) {
  const [animated, setAnimated] = useState(0);
  const radius = (size - strokeWidth) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (animated / 100) * circumference;

  useEffect(() => {
    const timer = setTimeout(() => setAnimated(percentage), 100);
    return () => clearTimeout(timer);
  }, [percentage]);

  const getColor = () => {
    if (percentage >= 80) return "#10B981";
    if (percentage >= 60) return "#F59E0B";
    return "#EF4444";
  };

  return (
    <div className="flex flex-col items-center justify-center">
      <div className="relative" style={{ width: size, height: size }}>
        <svg width={size} height={size} className="transform -rotate-90">
          <circle
            cx={size / 2}
            cy={size / 2}
            r={radius}
            stroke="#F3F4F6"
            strokeWidth={strokeWidth}
            fill="none"
          />
          <circle
            cx={size / 2}
            cy={size / 2}
            r={radius}
            stroke={getColor()}
            strokeWidth={strokeWidth}
            fill="none"
            strokeDasharray={circumference}
            strokeDashoffset={offset}
            strokeLinecap="round"
            style={{ transition: "stroke-dashoffset 1s ease-out" }}
          />
        </svg>
        <div className="absolute inset-0 flex flex-col items-center justify-center">
          <span className="text-3xl font-bold text-gray-900">
            {Math.round(percentage)}%
          </span>
          <span className="text-xs text-gray-500 mt-1">{label}</span>
        </div>
      </div>
    </div>
  );
}
