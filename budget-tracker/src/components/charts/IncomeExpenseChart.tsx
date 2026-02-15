"use client";

import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
} from "recharts";

interface ChartDataPoint {
  month: string;
  income: number;
  expense: number;
}

interface IncomeExpenseChartProps {
  data: ChartDataPoint[];
}

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});

function CustomTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ value: number; dataKey: string; color: string }>;
  label?: string;
}) {
  if (!active || !payload) return null;

  return (
    <div className="bg-white dark:bg-slate-800 rounded-lg shadow-lg border border-gray-200 dark:border-slate-700 p-3">
      <p className="text-sm font-medium text-gray-900 dark:text-white mb-2">
        {label}
      </p>
      {payload.map((entry) => (
        <div key={entry.dataKey} className="flex items-center gap-2 text-sm">
          <div
            className="w-3 h-3 rounded-sm"
            style={{ backgroundColor: entry.color }}
          />
          <span className="text-gray-600 dark:text-gray-400">
            {entry.dataKey === "income" ? "Einnahmen" : "Ausgaben"}:
          </span>
          <span className="font-medium text-gray-900 dark:text-white">
            {formatCurrency.format(entry.value)}
          </span>
        </div>
      ))}
    </div>
  );
}

export function IncomeExpenseChart({ data }: IncomeExpenseChartProps) {
  return (
    <ResponsiveContainer width="100%" height={350}>
      <BarChart
        data={data}
        margin={{ top: 10, right: 10, left: 10, bottom: 0 }}
      >
        <CartesianGrid
          strokeDasharray="3 3"
          stroke="#e5e7eb"
          className="dark:opacity-20"
          vertical={false}
        />
        <XAxis
          dataKey="month"
          axisLine={false}
          tickLine={false}
          tick={{ fontSize: 12, fill: "#9ca3af" }}
          dy={8}
        />
        <YAxis
          axisLine={false}
          tickLine={false}
          tick={{ fontSize: 12, fill: "#9ca3af" }}
          tickFormatter={(value: number) =>
            `${(value / 1000).toFixed(value >= 1000 ? 1 : 0)}k`
          }
          width={45}
        />
        <Tooltip
          content={<CustomTooltip />}
          cursor={{ fill: "rgba(0,0,0,0.04)" }}
        />
        <Bar
          dataKey="income"
          fill="#22c55e"
          radius={[4, 4, 0, 0]}
          barSize={20}
          name="Einnahmen"
        />
        <Bar
          dataKey="expense"
          fill="#ef4444"
          radius={[4, 4, 0, 0]}
          barSize={20}
          name="Ausgaben"
        />
      </BarChart>
    </ResponsiveContainer>
  );
}
