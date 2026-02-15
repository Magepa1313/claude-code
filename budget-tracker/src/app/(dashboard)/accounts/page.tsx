"use client";

import { useEffect, useState, useCallback } from "react";
import {
  Plus,
  Pencil,
  Trash2,
  Wallet,
  Landmark,
  CreditCard,
  PiggyBank,
  Banknote,
} from "lucide-react";
import { Card } from "@/components/ui/Card";
import { Button } from "@/components/ui/Button";
import { Badge } from "@/components/ui/Badge";
import { Modal } from "@/components/ui/Modal";
import { AccountForm } from "@/components/accounts/AccountForm";

interface Account {
  id: string;
  name: string;
  type: string;
  balance: number;
  color: string;
  icon: string;
  _count: {
    transactions: number;
  };
}

const ACCOUNT_TYPE_LABELS: Record<string, string> = {
  cash: "Bargeld",
  bank: "Bank",
  credit_card: "Kreditkarte",
  savings: "Sparkonto",
};

const ACCOUNT_TYPE_BADGE_COLORS: Record<
  string,
  "gray" | "blue" | "purple" | "green" | "orange" | "teal" | "red" | "yellow" | "pink"
> = {
  cash: "green",
  bank: "blue",
  credit_card: "purple",
  savings: "teal",
};

function getIconComponent(iconName: string, className: string) {
  switch (iconName) {
    case "wallet":
      return <Wallet className={className} />;
    case "landmark":
      return <Landmark className={className} />;
    case "credit-card":
      return <CreditCard className={className} />;
    case "piggy-bank":
      return <PiggyBank className={className} />;
    case "banknote":
      return <Banknote className={className} />;
    default:
      return <Wallet className={className} />;
  }
}

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
});

function SkeletonCard() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 overflow-hidden animate-pulse">
      <div className="h-1.5 bg-gray-200 dark:bg-slate-700" />
      <div className="p-6">
        <div className="flex items-start justify-between mb-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gray-200 dark:bg-slate-700 rounded-xl" />
            <div>
              <div className="h-5 w-28 bg-gray-200 dark:bg-slate-700 rounded" />
              <div className="h-4 w-16 bg-gray-200 dark:bg-slate-700 rounded mt-1.5" />
            </div>
          </div>
          <div className="flex gap-1">
            <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
            <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
          </div>
        </div>
        <div className="h-8 w-36 bg-gray-200 dark:bg-slate-700 rounded mt-4" />
        <div className="h-4 w-28 bg-gray-200 dark:bg-slate-700 rounded mt-3" />
      </div>
    </div>
  );
}

function SkeletonSummary() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-6 animate-pulse">
      <div className="flex items-center justify-between">
        <div>
          <div className="h-4 w-28 bg-gray-200 dark:bg-slate-700 rounded" />
          <div className="h-9 w-44 bg-gray-200 dark:bg-slate-700 rounded mt-2" />
        </div>
        <div className="w-12 h-12 bg-gray-200 dark:bg-slate-700 rounded-xl" />
      </div>
    </div>
  );
}

export default function AccountsPage() {
  const [accounts, setAccounts] = useState<Account[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [editingAccount, setEditingAccount] = useState<Account | null>(null);
  const [deleting, setDeleting] = useState<string | null>(null);

  const fetchAccounts = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const response = await fetch("/api/accounts");
      if (!response.ok) throw new Error("Konten konnten nicht geladen werden");
      const data = await response.json();
      setAccounts(data);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchAccounts();
  }, [fetchAccounts]);

  function openCreateModal() {
    setEditingAccount(null);
    setModalOpen(true);
  }

  function openEditModal(account: Account) {
    setEditingAccount(account);
    setModalOpen(true);
  }

  function closeModal() {
    setModalOpen(false);
    setEditingAccount(null);
  }

  async function handleSubmit(data: {
    name: string;
    type: string;
    balance?: number;
    color: string;
    icon: string;
  }) {
    if (editingAccount) {
      const response = await fetch(`/api/accounts/${editingAccount.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: data.name,
          type: data.type,
          color: data.color,
          icon: data.icon,
        }),
      });
      if (!response.ok) throw new Error("Konto konnte nicht aktualisiert werden");
    } else {
      const response = await fetch("/api/accounts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });
      if (!response.ok) throw new Error("Konto konnte nicht erstellt werden");
    }
    closeModal();
    fetchAccounts();
  }

  async function handleDelete(id: string) {
    if (!confirm("Konto wirklich löschen? Alle zugehörigen Transaktionen werden ebenfalls gelöscht.")) return;

    setDeleting(id);
    try {
      const response = await fetch(`/api/accounts/${id}`, {
        method: "DELETE",
      });
      if (!response.ok) throw new Error("Konto konnte nicht gelöscht werden");
      fetchAccounts();
    } catch {
      // silently fail
    } finally {
      setDeleting(null);
    }
  }

  const totalBalance = accounts.reduce((sum, a) => sum + a.balance, 0);

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-[400px]">
        <div className="text-center">
          <p className="text-red-500 dark:text-red-400 text-sm font-medium">
            {error}
          </p>
          <button
            onClick={() => fetchAccounts()}
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
            Konten
          </h1>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
            Verwalten Sie Ihre Bankkonten und Geldquellen.
          </p>
        </div>
        <Button
          variant="primary"
          icon={<Plus className="w-4 h-4" />}
          onClick={openCreateModal}
        >
          Neues Konto
        </Button>
      </div>

      {/* Total balance summary */}
      {loading ? (
        <SkeletonSummary />
      ) : accounts.length > 0 ? (
        <Card>
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm font-medium text-gray-500 dark:text-gray-400">
                Gesamtguthaben
              </p>
              <p
                className={`text-2xl sm:text-3xl font-bold mt-1 ${
                  totalBalance >= 0
                    ? "text-green-600 dark:text-green-400"
                    : "text-red-600 dark:text-red-400"
                }`}
              >
                {formatCurrency.format(totalBalance)}
              </p>
            </div>
            <div className="w-12 h-12 rounded-xl bg-blue-100 dark:bg-blue-900/30 flex items-center justify-center">
              <Wallet className="w-6 h-6 text-blue-600 dark:text-blue-400" />
            </div>
          </div>
        </Card>
      ) : null}

      {/* Account cards */}
      {loading ? (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {Array.from({ length: 3 }).map((_, i) => (
            <SkeletonCard key={i} />
          ))}
        </div>
      ) : accounts.length === 0 ? (
        <Card className="flex flex-col items-center justify-center py-16">
          <div className="w-16 h-16 rounded-full bg-gray-100 dark:bg-slate-700 flex items-center justify-center mb-4">
            <Wallet className="w-8 h-8 text-gray-400 dark:text-gray-500" />
          </div>
          <h3 className="text-lg font-semibold text-gray-900 dark:text-white mb-1">
            Keine Konten vorhanden
          </h3>
          <p className="text-sm text-gray-500 dark:text-gray-400 mb-6 text-center max-w-sm">
            Erstellen Sie Ihr erstes Konto, um mit der Verwaltung Ihrer Finanzen
            zu beginnen.
          </p>
          <Button
            variant="primary"
            icon={<Plus className="w-4 h-4" />}
            onClick={openCreateModal}
          >
            Neues Konto
          </Button>
        </Card>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {accounts.map((account) => (
            <div
              key={account.id}
              className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 overflow-hidden"
            >
              {/* Colored top border */}
              <div
                className="h-1.5"
                style={{ backgroundColor: account.color }}
              />

              <div className="p-6">
                {/* Header with icon, name, and actions */}
                <div className="flex items-start justify-between mb-4">
                  <div className="flex items-center gap-3 min-w-0">
                    <div
                      className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0"
                      style={{
                        backgroundColor: account.color + "20",
                      }}
                    >
                      {getIconComponent(
                        account.icon,
                        "w-5 h-5"
                      )}
                    </div>
                    <div className="min-w-0">
                      <h3 className="font-semibold text-gray-900 dark:text-white truncate">
                        {account.name}
                      </h3>
                      <Badge
                        color={
                          ACCOUNT_TYPE_BADGE_COLORS[account.type] || "gray"
                        }
                      >
                        {ACCOUNT_TYPE_LABELS[account.type] || account.type}
                      </Badge>
                    </div>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button
                      onClick={() => openEditModal(account)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 transition-colors"
                      title="Bearbeiten"
                    >
                      <Pencil className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => handleDelete(account.id)}
                      disabled={deleting === account.id}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:text-red-400 dark:hover:bg-red-900/20 transition-colors disabled:opacity-50"
                      title="Löschen"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Balance */}
                <div className="mt-4">
                  <p
                    className={`text-2xl font-bold ${
                      account.balance >= 0
                        ? "text-green-600 dark:text-green-400"
                        : "text-red-600 dark:text-red-400"
                    }`}
                  >
                    {formatCurrency.format(account.balance)}
                  </p>
                </div>

                {/* Transaction count */}
                <p className="text-sm text-gray-500 dark:text-gray-400 mt-2">
                  {account._count.transactions}{" "}
                  {account._count.transactions === 1
                    ? "Transaktion"
                    : "Transaktionen"}
                </p>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Modal */}
      <Modal
        isOpen={modalOpen}
        onClose={closeModal}
        title={editingAccount ? "Konto bearbeiten" : "Neues Konto"}
      >
        <AccountForm
          onSubmit={handleSubmit}
          onCancel={closeModal}
          initialData={
            editingAccount
              ? {
                  id: editingAccount.id,
                  name: editingAccount.name,
                  type: editingAccount.type,
                  color: editingAccount.color,
                  icon: editingAccount.icon,
                }
              : undefined
          }
        />
      </Modal>
    </div>
  );
}
