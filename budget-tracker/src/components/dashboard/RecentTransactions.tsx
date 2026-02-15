"use client";

import { format } from "date-fns";
import { clsx } from "clsx";

interface Transaction {
  id: string;
  description: string;
  amount: number;
  type: "INCOME" | "EXPENSE";
  date: string;
  category: {
    name: string;
    color: string;
    icon: string;
  };
  account: {
    name: string;
  };
}

interface RecentTransactionsProps {
  transactions: Transaction[];
}

const formatCurrency = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
});

export function RecentTransactions({ transactions }: RecentTransactionsProps) {
  if (transactions.length === 0) {
    return (
      <div className="text-center py-12">
        <p className="text-gray-500 dark:text-gray-400 text-sm">
          Keine Transaktionen vorhanden
        </p>
      </div>
    );
  }

  return (
    <div className="divide-y divide-gray-100 dark:divide-slate-700">
      {transactions.map((transaction) => (
        <div
          key={transaction.id}
          className="flex items-center justify-between py-4 first:pt-0 last:pb-0"
        >
          <div className="flex items-center gap-3 min-w-0 flex-1">
            {/* Category color dot and icon */}
            <div
              className="flex items-center justify-center w-10 h-10 rounded-full shrink-0"
              style={{ backgroundColor: `${transaction.category.color}20` }}
            >
              <div
                className="w-2.5 h-2.5 rounded-full"
                style={{ backgroundColor: transaction.category.color }}
              />
            </div>

            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                {transaction.description}
              </p>
              <div className="flex items-center gap-2 mt-0.5">
                <span className="text-xs text-gray-500 dark:text-gray-400">
                  {format(new Date(transaction.date), "dd MMM yyyy")}
                </span>
                <span className="text-gray-300 dark:text-slate-600">&middot;</span>
                <span className="text-xs text-gray-500 dark:text-gray-400 truncate">
                  {transaction.account.name}
                </span>
              </div>
            </div>
          </div>

          <span
            className={clsx(
              "text-sm font-semibold shrink-0 ml-4",
              transaction.type === "INCOME"
                ? "text-green-600 dark:text-green-400"
                : "text-red-600 dark:text-red-400"
            )}
          >
            {transaction.type === "INCOME" ? "+" : "-"}
            {formatCurrency.format(Math.abs(transaction.amount))}
          </span>
        </div>
      ))}
    </div>
  );
}
