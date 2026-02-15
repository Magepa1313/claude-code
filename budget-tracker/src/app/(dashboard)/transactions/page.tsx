"use client";

import { useState, useEffect, useCallback } from "react";
import { format } from "date-fns";
import { de } from "date-fns/locale";
import {
  Plus,
  Pencil,
  Trash2,
  Search,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Receipt,
} from "lucide-react";
import { clsx } from "clsx";
import { Card } from "@/components/ui/Card";
import { Button } from "@/components/ui/Button";
import { Input } from "@/components/ui/Input";
import { Select } from "@/components/ui/Select";
import { Modal } from "@/components/ui/Modal";
import {
  TransactionForm,
  TransactionFormData,
} from "@/components/transactions/TransactionForm";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface Category {
  id: string;
  name: string;
  type: string;
  color: string;
  icon: string;
}

interface Account {
  id: string;
  name: string;
  type: string;
  color: string;
  icon: string;
}

interface Transaction {
  id: string;
  amount: number;
  type: "income" | "expense";
  description: string;
  date: string;
  categoryId: string;
  accountId: string;
  category: Category;
  account: Account;
}

type SortField = "date" | "amount";
type SortOrder = "asc" | "desc";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
});

// ---------------------------------------------------------------------------
// Skeleton components
// ---------------------------------------------------------------------------

function SkeletonRow() {
  return (
    <div className="flex items-center justify-between py-4 animate-pulse">
      <div className="flex items-center gap-3 flex-1 min-w-0">
        <div className="w-10 h-10 bg-gray-200 dark:bg-slate-700 rounded-full shrink-0" />
        <div className="flex-1 min-w-0">
          <div className="h-4 w-40 bg-gray-200 dark:bg-slate-700 rounded" />
          <div className="h-3 w-24 bg-gray-200 dark:bg-slate-700 rounded mt-1.5" />
        </div>
      </div>
      <div className="h-4 w-20 bg-gray-200 dark:bg-slate-700 rounded ml-4" />
    </div>
  );
}

function SkeletonTable() {
  return (
    <div className="divide-y divide-gray-100 dark:divide-slate-700">
      {Array.from({ length: 6 }).map((_, i) => (
        <SkeletonRow key={i} />
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Main page component
// ---------------------------------------------------------------------------

export default function TransactionsPage() {
  // Data
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [categories, setCategories] = useState<Category[]>([]);
  const [accounts, setAccounts] = useState<Account[]>([]);

  // Loading / error
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Filters
  const [filterType, setFilterType] = useState("");
  const [filterCategory, setFilterCategory] = useState("");
  const [filterAccount, setFilterAccount] = useState("");
  const [filterDateFrom, setFilterDateFrom] = useState("");
  const [filterDateTo, setFilterDateTo] = useState("");
  const [searchQuery, setSearchQuery] = useState("");

  // Sorting
  const [sortField, setSortField] = useState<SortField>("date");
  const [sortOrder, setSortOrder] = useState<SortOrder>("desc");

  // Modals
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [editingTransaction, setEditingTransaction] =
    useState<Transaction | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);

  // -------------------------------------------------------------------------
  // Data fetching
  // -------------------------------------------------------------------------

  const fetchTransactions = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);

      const params = new URLSearchParams();
      if (filterType) params.set("type", filterType);
      if (filterCategory) params.set("categoryId", filterCategory);
      if (filterAccount) params.set("accountId", filterAccount);
      if (filterDateFrom) params.set("startDate", filterDateFrom);
      if (filterDateTo) params.set("endDate", filterDateTo);
      if (searchQuery) params.set("search", searchQuery);
      params.set("sort", sortField);
      params.set("order", sortOrder);

      const response = await fetch(`/api/transactions?${params.toString()}`);
      if (!response.ok) throw new Error("Transaktionen konnten nicht geladen werden");
      const data = await response.json();
      setTransactions(data);
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setLoading(false);
    }
  }, [
    filterType,
    filterCategory,
    filterAccount,
    filterDateFrom,
    filterDateTo,
    searchQuery,
    sortField,
    sortOrder,
  ]);

  const fetchCategories = useCallback(async () => {
    try {
      const response = await fetch("/api/categories");
      if (response.ok) {
        const data = await response.json();
        setCategories(data);
      }
    } catch {
      // Silently fail — categories are supplementary
    }
  }, []);

  const fetchAccounts = useCallback(async () => {
    try {
      const response = await fetch("/api/accounts");
      if (response.ok) {
        const data = await response.json();
        setAccounts(data);
      }
    } catch {
      // Silently fail — accounts are supplementary
    }
  }, []);

  useEffect(() => {
    fetchCategories();
    fetchAccounts();
  }, [fetchCategories, fetchAccounts]);

  useEffect(() => {
    fetchTransactions();
  }, [fetchTransactions]);

  // -------------------------------------------------------------------------
  // Handlers
  // -------------------------------------------------------------------------

  async function handleCreate(formData: TransactionFormData) {
    const response = await fetch("/api/transactions", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(formData),
    });

    if (!response.ok) {
      const body = await response.json().catch(() => null);
      throw new Error(body?.error ?? "Transaktion konnte nicht erstellt werden");
    }

    setShowCreateModal(false);
    fetchTransactions();
    fetchAccounts(); // balances may have changed
  }

  async function handleEdit(formData: TransactionFormData) {
    if (!editingTransaction) return;

    const response = await fetch(`/api/transactions/${editingTransaction.id}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(formData),
    });

    if (!response.ok) {
      const body = await response.json().catch(() => null);
      throw new Error(body?.error ?? "Transaktion konnte nicht aktualisiert werden");
    }

    setEditingTransaction(null);
    fetchTransactions();
    fetchAccounts();
  }

  async function handleDelete(id: string) {
    const response = await fetch(`/api/transactions/${id}`, {
      method: "DELETE",
    });

    if (!response.ok) {
      const body = await response.json().catch(() => null);
      throw new Error(body?.error ?? "Transaktion konnte nicht gelöscht werden");
    }

    setDeletingId(null);
    fetchTransactions();
    fetchAccounts();
  }

  function handleSort(field: SortField) {
    if (sortField === field) {
      setSortOrder((prev) => (prev === "asc" ? "desc" : "asc"));
    } else {
      setSortField(field);
      setSortOrder("desc");
    }
  }

  function SortIcon({ field }: { field: SortField }) {
    if (sortField !== field) {
      return <ArrowUpDown className="w-3.5 h-3.5 text-gray-400" />;
    }
    return sortOrder === "asc" ? (
      <ArrowUp className="w-3.5 h-3.5" />
    ) : (
      <ArrowDown className="w-3.5 h-3.5" />
    );
  }

  // Filtered categories for the filter dropdown (all categories, but optionally filtered by selected type)
  const filterCategoryOptions = filterType
    ? categories.filter((c) => c.type === filterType)
    : categories;

  // -------------------------------------------------------------------------
  // Render
  // -------------------------------------------------------------------------

  return (
    <div className="space-y-6">
      {/* Page header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
            Transaktionen
          </h1>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
            Verwalten Sie Ihre Einnahmen und Ausgaben.
          </p>
        </div>
        <Button
          variant="primary"
          icon={<Plus className="w-4 h-4" />}
          onClick={() => setShowCreateModal(true)}
        >
          Neue Transaktion
        </Button>
      </div>

      {/* Filter bar */}
      <Card className="!p-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-3">
          {/* Search */}
          <Input
            placeholder="Suchen..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            icon={<Search className="w-4 h-4" />}
          />

          {/* Type filter */}
          <Select
            value={filterType}
            onChange={(e) => {
              setFilterType(e.target.value);
              // Reset category filter when type changes
              setFilterCategory("");
            }}
            options={[
              { value: "", label: "Alle Typen" },
              { value: "income", label: "Einnahmen" },
              { value: "expense", label: "Ausgaben" },
            ]}
          />

          {/* Category filter */}
          <Select
            value={filterCategory}
            onChange={(e) => setFilterCategory(e.target.value)}
            options={[
              { value: "", label: "Alle Kategorien" },
              ...filterCategoryOptions.map((cat) => ({
                value: cat.id,
                label: cat.name,
              })),
            ]}
          />

          {/* Account filter */}
          <Select
            value={filterAccount}
            onChange={(e) => setFilterAccount(e.target.value)}
            options={[
              { value: "", label: "Alle Konten" },
              ...accounts.map((acc) => ({
                value: acc.id,
                label: acc.name,
              })),
            ]}
          />

          {/* Date from */}
          <Input
            type="date"
            placeholder="Von"
            value={filterDateFrom}
            onChange={(e) => setFilterDateFrom(e.target.value)}
          />

          {/* Date to */}
          <Input
            type="date"
            placeholder="Bis"
            value={filterDateTo}
            onChange={(e) => setFilterDateTo(e.target.value)}
          />
        </div>
      </Card>

      {/* Error state */}
      {error && (
        <div className="flex items-center justify-center min-h-[200px]">
          <div className="text-center">
            <p className="text-red-500 dark:text-red-400 text-sm font-medium">
              {error}
            </p>
            <button
              onClick={() => fetchTransactions()}
              className="mt-3 text-sm text-blue-600 dark:text-blue-400 hover:underline"
            >
              Erneut versuchen
            </button>
          </div>
        </div>
      )}

      {/* Transactions list */}
      {!error && (
        <Card>
          {loading ? (
            <SkeletonTable />
          ) : transactions.length === 0 ? (
            /* Empty state */
            <div className="flex flex-col items-center justify-center py-16 text-center">
              <div className="w-16 h-16 rounded-full bg-gray-100 dark:bg-slate-700 flex items-center justify-center mb-4">
                <Receipt className="w-8 h-8 text-gray-400 dark:text-gray-500" />
              </div>
              <p className="text-gray-500 dark:text-gray-400 font-medium">
                Keine Transaktionen gefunden
              </p>
              <p className="text-sm text-gray-400 dark:text-gray-500 mt-1">
                Erstellen Sie Ihre erste Transaktion oder passen Sie die Filter
                an.
              </p>
            </div>
          ) : (
            <>
              {/* ---- Desktop table ---- */}
              <div className="hidden md:block overflow-x-auto">
                <table className="w-full">
                  <thead>
                    <tr className="border-b border-gray-200 dark:border-slate-700">
                      <th className="text-left pb-3">
                        <button
                          className="inline-flex items-center gap-1 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200 transition-colors"
                          onClick={() => handleSort("date")}
                        >
                          Datum
                          <SortIcon field="date" />
                        </button>
                      </th>
                      <th className="text-left pb-3 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400">
                        Beschreibung
                      </th>
                      <th className="text-left pb-3 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400">
                        Kategorie
                      </th>
                      <th className="text-left pb-3 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400">
                        Konto
                      </th>
                      <th className="text-right pb-3">
                        <button
                          className="inline-flex items-center gap-1 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200 transition-colors ml-auto"
                          onClick={() => handleSort("amount")}
                        >
                          Betrag
                          <SortIcon field="amount" />
                        </button>
                      </th>
                      <th className="text-right pb-3 text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400">
                        Aktionen
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100 dark:divide-slate-700/60">
                    {transactions.map((tx) => (
                      <tr
                        key={tx.id}
                        className="group hover:bg-gray-50 dark:hover:bg-slate-700/30 transition-colors"
                      >
                        {/* Date */}
                        <td className="py-3.5 pr-4 text-sm text-gray-600 dark:text-gray-300 whitespace-nowrap">
                          {format(new Date(tx.date), "dd. MMM yyyy", {
                            locale: de,
                          })}
                        </td>

                        {/* Description */}
                        <td className="py-3.5 pr-4 text-sm font-medium text-gray-900 dark:text-white max-w-[250px] truncate">
                          {tx.description}
                        </td>

                        {/* Category */}
                        <td className="py-3.5 pr-4">
                          <span className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
                            <span
                              className="w-2.5 h-2.5 rounded-full shrink-0"
                              style={{ backgroundColor: tx.category.color }}
                            />
                            {tx.category.name}
                          </span>
                        </td>

                        {/* Account */}
                        <td className="py-3.5 pr-4 text-sm text-gray-600 dark:text-gray-300 whitespace-nowrap">
                          {tx.account.name}
                        </td>

                        {/* Amount */}
                        <td
                          className={clsx(
                            "py-3.5 pr-4 text-sm font-semibold text-right whitespace-nowrap",
                            tx.type === "income"
                              ? "text-green-600 dark:text-green-400"
                              : "text-red-600 dark:text-red-400"
                          )}
                        >
                          {tx.type === "income" ? "+" : "-"}
                          {formatCurrency.format(Math.abs(tx.amount))}
                        </td>

                        {/* Actions */}
                        <td className="py-3.5 text-right whitespace-nowrap">
                          <div className="inline-flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                            <button
                              onClick={() => setEditingTransaction(tx)}
                              className="p-1.5 rounded-lg text-gray-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 transition-colors"
                              title="Bearbeiten"
                            >
                              <Pencil className="w-4 h-4" />
                            </button>
                            <button
                              onClick={() => setDeletingId(tx.id)}
                              className="p-1.5 rounded-lg text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:text-red-400 dark:hover:bg-red-900/20 transition-colors"
                              title="Löschen"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* ---- Mobile card list ---- */}
              <div className="md:hidden divide-y divide-gray-100 dark:divide-slate-700">
                {/* Mobile sort controls */}
                <div className="flex items-center gap-2 pb-3 mb-1">
                  <span className="text-xs text-gray-500 dark:text-gray-400">
                    Sortieren:
                  </span>
                  <button
                    className={clsx(
                      "inline-flex items-center gap-1 text-xs font-medium px-2 py-1 rounded-md transition-colors",
                      sortField === "date"
                        ? "bg-blue-50 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"
                        : "text-gray-500 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-slate-700"
                    )}
                    onClick={() => handleSort("date")}
                  >
                    Datum
                    <SortIcon field="date" />
                  </button>
                  <button
                    className={clsx(
                      "inline-flex items-center gap-1 text-xs font-medium px-2 py-1 rounded-md transition-colors",
                      sortField === "amount"
                        ? "bg-blue-50 text-blue-700 dark:bg-blue-900/30 dark:text-blue-400"
                        : "text-gray-500 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-slate-700"
                    )}
                    onClick={() => handleSort("amount")}
                  >
                    Betrag
                    <SortIcon field="amount" />
                  </button>
                </div>

                {transactions.map((tx) => (
                  <div
                    key={tx.id}
                    className="py-4 first:pt-0 last:pb-0"
                  >
                    <div className="flex items-start justify-between gap-3">
                      {/* Left: icon + info */}
                      <div className="flex items-start gap-3 min-w-0 flex-1">
                        <div
                          className="flex items-center justify-center w-10 h-10 rounded-full shrink-0 mt-0.5"
                          style={{
                            backgroundColor: `${tx.category.color}20`,
                          }}
                        >
                          <span
                            className="w-2.5 h-2.5 rounded-full"
                            style={{ backgroundColor: tx.category.color }}
                          />
                        </div>
                        <div className="min-w-0 flex-1">
                          <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                            {tx.description}
                          </p>
                          <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 mt-0.5">
                            <span className="text-xs text-gray-500 dark:text-gray-400">
                              {format(new Date(tx.date), "dd. MMM yyyy", {
                                locale: de,
                              })}
                            </span>
                            <span className="text-gray-300 dark:text-slate-600">
                              &middot;
                            </span>
                            <span className="text-xs text-gray-500 dark:text-gray-400 truncate">
                              {tx.category.name}
                            </span>
                            <span className="text-gray-300 dark:text-slate-600">
                              &middot;
                            </span>
                            <span className="text-xs text-gray-500 dark:text-gray-400 truncate">
                              {tx.account.name}
                            </span>
                          </div>
                        </div>
                      </div>

                      {/* Right: amount + actions */}
                      <div className="flex flex-col items-end gap-1 shrink-0">
                        <span
                          className={clsx(
                            "text-sm font-semibold",
                            tx.type === "income"
                              ? "text-green-600 dark:text-green-400"
                              : "text-red-600 dark:text-red-400"
                          )}
                        >
                          {tx.type === "income" ? "+" : "-"}
                          {formatCurrency.format(Math.abs(tx.amount))}
                        </span>
                        <div className="flex items-center gap-0.5">
                          <button
                            onClick={() => setEditingTransaction(tx)}
                            className="p-1 rounded text-gray-400 hover:text-blue-600 dark:hover:text-blue-400 transition-colors"
                            title="Bearbeiten"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => setDeletingId(tx.id)}
                            className="p-1 rounded text-gray-400 hover:text-red-600 dark:hover:text-red-400 transition-colors"
                            title="Löschen"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </Card>
      )}

      {/* ----------------------------------------------------------------- */}
      {/* Create Modal                                                      */}
      {/* ----------------------------------------------------------------- */}
      <Modal
        isOpen={showCreateModal}
        onClose={() => setShowCreateModal(false)}
        title="Neue Transaktion"
      >
        <TransactionForm
          onSubmit={handleCreate}
          onCancel={() => setShowCreateModal(false)}
          categories={categories}
          accounts={accounts}
        />
      </Modal>

      {/* ----------------------------------------------------------------- */}
      {/* Edit Modal                                                        */}
      {/* ----------------------------------------------------------------- */}
      <Modal
        isOpen={!!editingTransaction}
        onClose={() => setEditingTransaction(null)}
        title="Transaktion bearbeiten"
      >
        {editingTransaction && (
          <TransactionForm
            onSubmit={handleEdit}
            onCancel={() => setEditingTransaction(null)}
            categories={categories}
            accounts={accounts}
            initialData={{
              amount: editingTransaction.amount,
              type: editingTransaction.type,
              description: editingTransaction.description,
              date: new Date(editingTransaction.date)
                .toISOString()
                .split("T")[0],
              categoryId: editingTransaction.categoryId,
              accountId: editingTransaction.accountId,
            }}
          />
        )}
      </Modal>

      {/* ----------------------------------------------------------------- */}
      {/* Delete confirmation Modal                                         */}
      {/* ----------------------------------------------------------------- */}
      <Modal
        isOpen={!!deletingId}
        onClose={() => setDeletingId(null)}
        title="Transaktion löschen"
      >
        <div className="space-y-4">
          <p className="text-sm text-gray-600 dark:text-gray-300">
            Sind Sie sicher, dass Sie diese Transaktion löschen möchten? Diese
            Aktion kann nicht rückgängig gemacht werden.
          </p>
          <div className="flex items-center justify-end gap-3">
            <Button
              variant="secondary"
              onClick={() => setDeletingId(null)}
            >
              Abbrechen
            </Button>
            <Button
              variant="danger"
              onClick={() => {
                if (deletingId) handleDelete(deletingId);
              }}
            >
              Löschen
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
