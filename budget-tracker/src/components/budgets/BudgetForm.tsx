"use client";

import { useState, useEffect } from "react";
import { Input } from "@/components/ui/Input";
import { Select } from "@/components/ui/Select";
import { Button } from "@/components/ui/Button";

interface Category {
  id: string;
  name: string;
  color: string;
  icon: string;
  type: string;
}

interface BudgetFormData {
  categoryId: string;
  amount: number;
}

interface BudgetFormProps {
  onSubmit: (data: BudgetFormData) => Promise<void>;
  onCancel: () => void;
  categories: Category[];
  initialData?: {
    id: string;
    categoryId: string;
    amount: number;
  };
}

export function BudgetForm({
  onSubmit,
  onCancel,
  categories,
  initialData,
}: BudgetFormProps) {
  const [categoryId, setCategoryId] = useState(initialData?.categoryId ?? "");
  const [amount, setAmount] = useState(
    initialData?.amount?.toString() ?? ""
  );
  const [errors, setErrors] = useState<{ categoryId?: string; amount?: string }>({});
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (initialData) {
      setCategoryId(initialData.categoryId);
      setAmount(initialData.amount.toString());
    }
  }, [initialData]);

  function validate(): boolean {
    const newErrors: { categoryId?: string; amount?: string } = {};

    if (!categoryId) {
      newErrors.categoryId = "Bitte eine Kategorie auswählen";
    }

    const parsedAmount = parseFloat(amount);
    if (!amount || isNaN(parsedAmount) || parsedAmount <= 0) {
      newErrors.amount = "Bitte einen gültigen Betrag eingeben";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();

    if (!validate()) return;

    setSubmitting(true);
    try {
      await onSubmit({
        categoryId,
        amount: parseFloat(amount),
      });
    } finally {
      setSubmitting(false);
    }
  }

  const categoryOptions = categories.map((cat) => ({
    value: cat.id,
    label: cat.name,
  }));

  return (
    <form onSubmit={handleSubmit} className="space-y-5">
      <Select
        label="Kategorie"
        options={categoryOptions}
        placeholder="Kategorie auswählen..."
        value={categoryId}
        onChange={(e) => setCategoryId(e.target.value)}
        error={errors.categoryId}
        disabled={!!initialData}
      />

      <Input
        label="Betrag"
        type="number"
        placeholder="0,00"
        step="0.01"
        min="0.01"
        value={amount}
        onChange={(e) => setAmount(e.target.value)}
        error={errors.amount}
        icon={<span className="text-sm font-medium">&euro;</span>}
      />

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
          {initialData ? "Speichern" : "Hinzufügen"}
        </Button>
      </div>
    </form>
  );
}
