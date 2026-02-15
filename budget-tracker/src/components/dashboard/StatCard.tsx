"use client";

import { clsx } from "clsx";
import { TrendingUp, TrendingDown, Minus } from "lucide-react";

interface StatCardProps {
  icon: React.ReactNode;
  title: string;
  value: string;
  change: string;
  trend: "up" | "down" | "neutral";
}

const trendConfig = {
  up: {
    color: "text-green-600 dark:text-green-400",
    bg: "bg-green-50 dark:bg-green-900/20",
    Icon: TrendingUp,
  },
  down: {
    color: "text-red-600 dark:text-red-400",
    bg: "bg-red-50 dark:bg-red-900/20",
    Icon: TrendingDown,
  },
  neutral: {
    color: "text-gray-500 dark:text-gray-400",
    bg: "bg-gray-50 dark:bg-gray-800",
    Icon: Minus,
  },
};

export function StatCard({ icon, title, value, change, trend }: StatCardProps) {
  const { color, bg, Icon } = trendConfig[trend];

  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 transition-shadow hover:shadow-md">
      <div className="flex items-center justify-between">
        <div className="flex-1 min-w-0">
          <p className="text-sm font-medium text-gray-500 dark:text-gray-400 truncate">
            {title}
          </p>
          <p className="mt-2 text-2xl font-bold text-gray-900 dark:text-white truncate">
            {value}
          </p>
          <div className={clsx("flex items-center gap-1 mt-2", color)}>
            <Icon className="w-4 h-4 shrink-0" />
            <span className="text-sm font-medium">{change}</span>
          </div>
        </div>
        <div
          className={clsx(
            "flex items-center justify-center w-12 h-12 rounded-xl shrink-0 ml-4",
            bg
          )}
        >
          {icon}
        </div>
      </div>
    </div>
  );
}
