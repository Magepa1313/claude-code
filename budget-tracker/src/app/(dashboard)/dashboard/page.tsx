"use client";

import { useEffect, useState } from "react";
import {
  Wallet,
  TrendingUp,
  TrendingDown,
  ArrowLeftRight,
} from "lucide-react";
import { Card } from "@/components/ui/Card";
import { StatCard } from "@/components/dashboard/StatCard";
import { RecentTransactions } from "@/components/dashboard/RecentTransactions";
import { IncomeExpenseChart } from "@/components/charts/IncomeExpenseChart";
import { CategoryPieChart } from "@/components/charts/CategoryPieChart";

interface DashboardData {
  totalBalance: number;
  income: number;
  expenses: number;
  transactionCount: number;
  incomeChange: string;
  expenseChange: string;
  balanceChange: string;
  transactionChange: string;
  incomeTrend: "up" | "down" | "neutral";
  expenseTrend: "up" | "down" | "neutral";
  balanceTrend: "up" | "down" | "neutral";
  transactionTrend: "up" | "down" | "neutral";
  monthlyData: Array<{ month: string; income: number; expense: number }>;
  categoryData: Array<{ name: string; value: number; color: string }>;
  recentTransactions: Array<{
    id: string;
    description: string;
    amount: number;
    type: "INCOME" | "EXPENSE";
    date: string;
    category: { name: string; color: string; icon: string };
    account: { name: string };
  }>;
}

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
});

function SkeletonCard() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 animate-pulse">
      <div className="flex items-center justify-between">
        <div className="flex-1">
          <div className="h-4 w-24 bg-gray-200 dark:bg-slate-700 rounded" />
          <div className="h-8 w-36 bg-gray-200 dark:bg-slate-700 rounded mt-2" />
          <div className="h-4 w-16 bg-gray-200 dark:bg-slate-700 rounded mt-2" />
        </div>
        <div className="w-12 h-12 bg-gray-200 dark:bg-slate-700 rounded-xl" />
      </div>
    </div>
  );
}

function SkeletonChart() {
  return (
    <Card>
      <div className="animate-pulse">
        <div className="h-5 w-40 bg-gray-200 dark:bg-slate-700 rounded mb-6" />
        <div className="h-[350px] bg-gray-100 dark:bg-slate-700/50 rounded-lg" />
      </div>
    </Card>
  );
}

function SkeletonTransactions() {
  return (
    <Card>
      <div className="animate-pulse">
        <div className="h-5 w-48 bg-gray-200 dark:bg-slate-700 rounded mb-6" />
        <div className="space-y-4">
          {Array.from({ length: 5 }).map((_, i) => (
            <div key={i} className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-gray-200 dark:bg-slate-700 rounded-full" />
                <div>
                  <div className="h-4 w-32 bg-gray-200 dark:bg-slate-700 rounded" />
                  <div className="h-3 w-24 bg-gray-200 dark:bg-slate-700 rounded mt-1.5" />
                </div>
              </div>
              <div className="h-4 w-20 bg-gray-200 dark:bg-slate-700 rounded" />
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
}

export default function DashboardPage() {
  const [data, setData] = useState<DashboardData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    async function fetchDashboard() {
      try {
        setLoading(true);
        const response = await fetch("/api/dashboard");
        if (!response.ok) {
          throw new Error("Daten konnten nicht geladen werden");
        }
        const dashboardData = await response.json();
        setData(dashboardData);
      } catch (err) {
        setError(
          err instanceof Error
            ? err.message
            : "Ein unbekannter Fehler ist aufgetreten"
        );
      } finally {
        setLoading(false);
      }
    }

    fetchDashboard();
  }, []);

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-[400px]">
        <div className="text-center">
          <p className="text-red-500 dark:text-red-400 text-sm font-medium">
            {error}
          </p>
          <button
            onClick={() => window.location.reload()}
            className="mt-3 text-sm text-blue-600 dark:text-blue-400 hover:underline"
          >
            Erneut versuchen
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Page header */}
      <div>
        <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
          Dashboard
        </h1>
        <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
          Willkommen! Hier ist Ihre finanzielle Übersicht.
        </p>
      </div>

      {/* Stat cards */}
      {loading ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {Array.from({ length: 4 }).map((_, i) => (
            <SkeletonCard key={i} />
          ))}
        </div>
      ) : data ? (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <StatCard
            icon={<Wallet className="w-6 h-6 text-blue-600 dark:text-blue-400" />}
            title="Gesamtguthaben"
            value={formatCurrency.format(data.totalBalance)}
            change={data.balanceChange}
            trend={data.balanceTrend}
          />
          <StatCard
            icon={
              <TrendingUp className="w-6 h-6 text-green-600 dark:text-green-400" />
            }
            title="Einnahmen"
            value={formatCurrency.format(data.income)}
            change={data.incomeChange}
            trend={data.incomeTrend}
          />
          <StatCard
            icon={
              <TrendingDown className="w-6 h-6 text-red-600 dark:text-red-400" />
            }
            title="Ausgaben"
            value={formatCurrency.format(data.expenses)}
            change={data.expenseChange}
            trend={data.expenseTrend}
          />
          <StatCard
            icon={
              <ArrowLeftRight className="w-6 h-6 text-purple-600 dark:text-purple-400" />
            }
            title="Transaktionen"
            value={String(data.transactionCount)}
            change={data.transactionChange}
            trend={data.transactionTrend}
          />
        </div>
      ) : null}

      {/* Charts */}
      {loading ? (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <SkeletonChart />
          <SkeletonChart />
        </div>
      ) : data ? (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <Card>
            <h2 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
              Einnahmen vs. Ausgaben
            </h2>
            <IncomeExpenseChart data={data.monthlyData} />
          </Card>
          <Card>
            <h2 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
              Ausgaben nach Kategorie
            </h2>
            <CategoryPieChart data={data.categoryData} />
          </Card>
        </div>
      ) : null}

      {/* Recent transactions */}
      {loading ? (
        <SkeletonTransactions />
      ) : data ? (
        <Card>
          <h2 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
            Letzte Transaktionen
          </h2>
          <RecentTransactions transactions={data.recentTransactions} />
        </Card>
      ) : null}
    </div>
  );
}
