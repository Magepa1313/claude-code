"use client";

import { useState, useEffect, FormEvent } from "react";
import { Input } from "@/components/ui/Input";
import { Select } from "@/components/ui/Select";
import { Button } from "@/components/ui/Button";

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
}

export interface TransactionFormData {
  amount: number;
  type: "income" | "expense";
  description: string;
  date: string;
  categoryId: string;
  accountId: string;
}

interface TransactionFormProps {
  onSubmit: (data: TransactionFormData) => void | Promise<void>;
  onCancel: () => void;
  initialData?: Partial<TransactionFormData>;
  categories: Category[];
  accounts: Account[];
}

export function TransactionForm({
  onSubmit,
  onCancel,
  initialData,
  categories,
  accounts,
}: TransactionFormProps) {
  const [type, setType] = useState<"income" | "expense">(
    initialData?.type ?? "expense"
  );
  const [amount, setAmount] = useState<string>(
    initialData?.amount ? String(initialData.amount) : ""
  );
  const [description, setDescription] = useState(
    initialData?.description ?? ""
  );
  const [date, setDate] = useState(
    initialData?.date ?? new Date().toISOString().split("T")[0]
  );
  const [categoryId, setCategoryId] = useState(initialData?.categoryId ?? "");
  const [accountId, setAccountId] = useState(initialData?.accountId ?? "");
  const [submitting, setSubmitting] = useState(false);

  const [errors, setErrors] = useState<Record<string, string>>({});

  // Filter categories by selected type
  const filteredCategories = categories.filter((cat) => cat.type === type);

  // When type changes, reset category if current selection doesn't match
  useEffect(() => {
    const currentCategoryValid = filteredCategories.some(
      (cat) => cat.id === categoryId
    );
    if (!currentCategoryValid) {
      setCategoryId("");
    }
  }, [type]); // eslint-disable-line react-hooks/exhaustive-deps

  function validate(): boolean {
    const newErrors: Record<string, string> = {};

    const parsedAmount = parseFloat(amount);
    if (!amount || isNaN(parsedAmount) || parsedAmount <= 0) {
      newErrors.amount = "Betrag muss größer als 0 sein";
    }

    if (!description.trim()) {
      newErrors.description = "Beschreibung ist erforderlich";
    }

    if (!date) {
      newErrors.date = "Datum ist erforderlich";
    }

    if (!categoryId) {
      newErrors.categoryId = "Kategorie ist erforderlich";
    }

    if (!accountId) {
      newErrors.accountId = "Konto ist erforderlich";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();

    if (!validate()) return;

    setSubmitting(true);
    try {
      await onSubmit({
        amount: parseFloat(amount),
        type,
        description: description.trim(),
        date,
        categoryId,
        accountId,
      });
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-4">
      {/* Type */}
      <Select
        label="Typ"
        value={type}
        onChange={(e) => setType(e.target.value as "income" | "expense")}
        options={[
          { value: "income", label: "Einnahme" },
          { value: "expense", label: "Ausgabe" },
        ]}
      />

      {/* Amount */}
      <Input
        label="Betrag"
        type="number"
        step="0.01"
        min="0.01"
        placeholder="0,00"
        value={amount}
        onChange={(e) => setAmount(e.target.value)}
        error={errors.amount}
      />

      {/* Description */}
      <Input
        label="Beschreibung"
        type="text"
        placeholder="z.B. Einkauf im Supermarkt"
        value={description}
        onChange={(e) => setDescription(e.target.value)}
        error={errors.description}
      />

      {/* Date */}
      <Input
        label="Datum"
        type="date"
        value={date}
        onChange={(e) => setDate(e.target.value)}
        error={errors.date}
      />

      {/* Category */}
      <Select
        label="Kategorie"
        value={categoryId}
        onChange={(e) => setCategoryId(e.target.value)}
        placeholder="Kategorie wählen"
        options={filteredCategories.map((cat) => ({
          value: cat.id,
          label: cat.name,
        }))}
        error={errors.categoryId}
      />

      {/* Account */}
      <Select
        label="Konto"
        value={accountId}
        onChange={(e) => setAccountId(e.target.value)}
        placeholder="Konto wählen"
        options={accounts.map((acc) => ({
          value: acc.id,
          label: acc.name,
        }))}
        error={errors.accountId}
      />

      {/* Actions */}
      <div className="flex items-center justify-end gap-3 pt-2">
        <Button
          type="button"
          variant="secondary"
          onClick={onCancel}
          disabled={submitting}
        >
          Abbrechen
        </Button>
        <Button type="submit" variant="primary" loading={submitting}>
          {initialData ? "Speichern" : "Erstellen"}
        </Button>
      </div>
    </form>
  );
}
