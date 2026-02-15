"use client";

import { useEffect, useState, useCallback } from "react";
import {
  Plus,
  ChevronLeft,
  ChevronRight,
  Pencil,
  Trash2,
  PieChart,
  Target,
} from "lucide-react";
import { Card } from "@/components/ui/Card";
import { Button } from "@/components/ui/Button";
import { Modal } from "@/components/ui/Modal";
import { BudgetForm } from "@/components/budgets/BudgetForm";

const MONTH_NAMES = [
  "Januar",
  "Februar",
  "März",
  "April",
  "Mai",
  "Juni",
  "Juli",
  "August",
  "September",
  "Oktober",
  "November",
  "Dezember",
];

interface Category {
  id: string;
  name: string;
  color: string;
  icon: string;
  type: string;
}

interface Budget {
  id: string;
  amount: number;
  month: number;
  year: number;
  categoryId: string;
  category: Category;
  spent: number;
}

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
});

function getProgressColor(percentage: number): {
  bar: string;
  bg: string;
  text: string;
} {
  if (percentage > 80) {
    return {
      bar: "bg-red-500 dark:bg-red-400",
      bg: "bg-red-100 dark:bg-red-900/30",
      text: "text-red-600 dark:text-red-400",
    };
  }
  if (percentage > 50) {
    return {
      bar: "bg-yellow-500 dark:bg-yellow-400",
      bg: "bg-yellow-100 dark:bg-yellow-900/30",
      text: "text-yellow-600 dark:text-yellow-400",
    };
  }
  return {
    bar: "bg-green-500 dark:bg-green-400",
    bg: "bg-green-100 dark:bg-green-900/30",
    text: "text-green-600 dark:text-green-400",
  };
}

function SkeletonCard() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 animate-pulse">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-3">
          <div className="w-3 h-3 rounded-full bg-gray-200 dark:bg-slate-700" />
          <div className="h-5 w-32 bg-gray-200 dark:bg-slate-700 rounded" />
        </div>
        <div className="flex gap-2">
          <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
          <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
        </div>
      </div>
      <div className="h-3 bg-gray-200 dark:bg-slate-700 rounded-full mb-3" />
      <div className="flex justify-between">
        <div className="h-4 w-40 bg-gray-200 dark:bg-slate-700 rounded" />
        <div className="h-4 w-12 bg-gray-200 dark:bg-slate-700 rounded" />
      </div>
    </div>
  );
}

function SkeletonSummary() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 animate-pulse">
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        {Array.from({ length: 3 }).map((_, i) => (
          <div key={i} className="text-center">
            <div className="h-4 w-24 bg-gray-200 dark:bg-slate-700 rounded mx-auto mb-2" />
            <div className="h-7 w-32 bg-gray-200 dark:bg-slate-700 rounded mx-auto" />
          </div>
        ))}
      </div>
    </div>
  );
}

export default function BudgetsPage() {
  const now = new Date();
  const [selectedMonth, setSelectedMonth] = useState(now.getMonth() + 1);
  const [selectedYear, setSelectedYear] = useState(now.getFullYear());
  const [budgets, setBudgets] = useState<Budget[]>([]);
  const [categories, setCategories] = useState<Category[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [editingBudget, setEditingBudget] = useState<Budget | null>(null);
  const [deleting, setDeleting] = useState<string | null>(null);

  const fetchBudgets = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const response = await fetch(
        `/api/budgets?month=${selectedMonth}&year=${selectedYear}`
      );
      if (!response.ok) throw new Error("Budgets konnten nicht geladen werden");
      const data = await response.json();
      setBudgets(data);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setLoading(false);
    }
  }, [selectedMonth, selectedYear]);

  const fetchCategories = useCallback(async () => {
    try {
      const response = await fetch("/api/categories?type=expense");
      if (!response.ok) throw new Error("Kategorien konnten nicht geladen werden");
      const data = await response.json();
      setCategories(data);
    } catch {
      // Categories are needed for the form, silently fail
    }
  }, []);

  useEffect(() => {
    fetchBudgets();
  }, [fetchBudgets]);

  useEffect(() => {
    fetchCategories();
  }, [fetchCategories]);

  function goToPreviousMonth() {
    if (selectedMonth === 1) {
      setSelectedMonth(12);
      setSelectedYear((y) => y - 1);
    } else {
      setSelectedMonth((m) => m - 1);
    }
  }

  function goToNextMonth() {
    if (selectedMonth === 12) {
      setSelectedMonth(1);
      setSelectedYear((y) => y + 1);
    } else {
      setSelectedMonth((m) => m + 1);
    }
  }

  function openCreateModal() {
    setEditingBudget(null);
    setModalOpen(true);
  }

  function openEditModal(budget: Budget) {
    setEditingBudget(budget);
    setModalOpen(true);
  }

  function closeModal() {
    setModalOpen(false);
    setEditingBudget(null);
  }

  async function handleSubmit(data: { categoryId: string; amount: number }) {
    if (editingBudget) {
      const response = await fetch(`/api/budgets/${editingBudget.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount: data.amount }),
      });
      if (!response.ok) throw new Error("Budget konnte nicht aktualisiert werden");
    } else {
      const response = await fetch("/api/budgets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          categoryId: data.categoryId,
          amount: data.amount,
          month: selectedMonth,
          year: selectedYear,
        }),
      });
      if (!response.ok) throw new Error("Budget konnte nicht erstellt werden");
    }
    closeModal();
    fetchBudgets();
  }

  async function handleDelete(id: string) {
    if (!confirm("Budget wirklich löschen?")) return;

    setDeleting(id);
    try {
      const response = await fetch(`/api/budgets/${id}`, {
        method: "DELETE",
      });
      if (!response.ok) throw new Error("Budget konnte nicht gelöscht werden");
      fetchBudgets();
    } catch {
      // silently fail
    } finally {
      setDeleting(null);
    }
  }

  const totalBudgeted = budgets.reduce((sum, b) => sum + b.amount, 0);
  const totalSpent = budgets.reduce((sum, b) => sum + b.spent, 0);
  const totalRemaining = totalBudgeted - totalSpent;

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-[400px]">
        <div className="text-center">
          <p className="text-red-500 dark:text-red-400 text-sm font-medium">
            {error}
          </p>
          <button
            onClick={() => fetchBudgets()}
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
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
            Budgets
          </h1>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
            Verwalten Sie Ihre monatlichen Budgets.
          </p>
        </div>
        <Button
          variant="primary"
          icon={<Plus className="w-4 h-4" />}
          onClick={openCreateModal}
        >
          Budget hinzufügen
        </Button>
      </div>

      {/* Month selector */}
      <div className="flex items-center justify-center gap-4">
        <button
          onClick={goToPreviousMonth}
          className="p-2 rounded-lg text-gray-500 hover:text-gray-700 hover:bg-gray-100 dark:text-gray-400 dark:hover:text-gray-200 dark:hover:bg-slate-700 transition-colors"
        >
          <ChevronLeft className="w-5 h-5" />
        </button>
        <h2 className="text-lg font-semibold text-gray-900 dark:text-white min-w-[200px] text-center">
          {MONTH_NAMES[selectedMonth - 1]} {selectedYear}
        </h2>
        <button
          onClick={goToNextMonth}
          className="p-2 rounded-lg text-gray-500 hover:text-gray-700 hover:bg-gray-100 dark:text-gray-400 dark:hover:text-gray-200 dark:hover:bg-slate-700 transition-colors"
        >
          <ChevronRight className="w-5 h-5" />
        </button>
      </div>

      {/* Summary */}
      {loading ? (
        <SkeletonSummary />
      ) : budgets.length > 0 ? (
        <Card>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 sm:gap-6">
            <div className="text-center">
              <p className="text-sm font-medium text-gray-500 dark:text-gray-400">
                Gesamtbudget
              </p>
              <p className="text-xl font-bold text-gray-900 dark:text-white mt-1">
                {formatCurrency.format(totalBudgeted)}
              </p>
            </div>
            <div className="text-center">
              <p className="text-sm font-medium text-gray-500 dark:text-gray-400">
                Ausgegeben
              </p>
              <p className="text-xl font-bold text-red-600 dark:text-red-400 mt-1">
                {formatCurrency.format(totalSpent)}
              </p>
            </div>
            <div className="text-center">
              <p className="text-sm font-medium text-gray-500 dark:text-gray-400">
                Verbleibend
              </p>
              <p
                className={`text-xl font-bold mt-1 ${
                  totalRemaining >= 0
                    ? "text-green-600 dark:text-green-400"
                    : "text-red-600 dark:text-red-400"
                }`}
              >
                {formatCurrency.format(totalRemaining)}
              </p>
            </div>
          </div>
        </Card>
      ) : null}

      {/* Budget cards */}
      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {Array.from({ length: 6 }).map((_, i) => (
            <SkeletonCard key={i} />
          ))}
        </div>
      ) : budgets.length === 0 ? (
        <Card className="flex flex-col items-center justify-center py-16">
          <div className="w-16 h-16 rounded-full bg-gray-100 dark:bg-slate-700 flex items-center justify-center mb-4">
            <Target className="w-8 h-8 text-gray-400 dark:text-gray-500" />
          </div>
          <h3 className="text-lg font-semibold text-gray-900 dark:text-white mb-1">
            Keine Budgets vorhanden
          </h3>
          <p className="text-sm text-gray-500 dark:text-gray-400 mb-6 text-center max-w-sm">
            Erstellen Sie Ihr erstes Budget, um Ihre Ausgaben für{" "}
            {MONTH_NAMES[selectedMonth - 1]} {selectedYear} zu planen.
          </p>
          <Button
            variant="primary"
            icon={<Plus className="w-4 h-4" />}
            onClick={openCreateModal}
          >
            Budget hinzufügen
          </Button>
        </Card>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {budgets.map((budget) => {
            const percentage =
              budget.amount > 0
                ? Math.round((budget.spent / budget.amount) * 100)
                : 0;
            const clampedPercentage = Math.min(percentage, 100);
            const colors = getProgressColor(percentage);

            return (
              <Card key={budget.id} className="relative">
                {/* Header row */}
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2.5 min-w-0">
                    <div
                      className="w-3 h-3 rounded-full shrink-0"
                      style={{ backgroundColor: budget.category.color }}
                    />
                    <span className="font-semibold text-gray-900 dark:text-white truncate">
                      {budget.category.name}
                    </span>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      onClick={() => openEditModal(budget)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 transition-colors"
                      title="Bearbeiten"
                    >
                      <Pencil className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => handleDelete(budget.id)}
                      disabled={deleting === budget.id}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:text-red-400 dark:hover:bg-red-900/20 transition-colors disabled:opacity-50"
                      title="Löschen"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Progress bar */}
                <div className={`w-full h-2.5 rounded-full ${colors.bg}`}>
                  <div
                    className={`h-full rounded-full transition-all duration-500 ${colors.bar}`}
                    style={{ width: `${clampedPercentage}%` }}
                  />
                </div>

                {/* Info row */}
                <div className="flex items-center justify-between mt-3">
                  <span className="text-sm text-gray-600 dark:text-gray-400">
                    {formatCurrency.format(budget.spent)} von{" "}
                    {formatCurrency.format(budget.amount)} ausgegeben
                  </span>
                  <span className={`text-sm font-semibold ${colors.text}`}>
                    {percentage}%
                  </span>
                </div>
              </Card>
            );
          })}
        </div>
      )}

      {/* Modal */}
      <Modal
        isOpen={modalOpen}
        onClose={closeModal}
        title={editingBudget ? "Budget bearbeiten" : "Neues Budget"}
      >
        <BudgetForm
          categories={categories}
          onSubmit={handleSubmit}
          onCancel={closeModal}
          initialData={
            editingBudget
              ? {
                  id: editingBudget.id,
                  categoryId: editingBudget.categoryId,
                  amount: editingBudget.amount,
                }
              : undefined
          }
        />
      </Modal>
    </div>
  );
}
