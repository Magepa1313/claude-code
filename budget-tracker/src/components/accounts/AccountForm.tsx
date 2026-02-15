"use client";

import { useState, useEffect } from "react";
import { Input } from "@/components/ui/Input";
import { Select } from "@/components/ui/Select";
import { Button } from "@/components/ui/Button";

const ACCOUNT_TYPES = [
  { value: "cash", label: "Bargeld" },
  { value: "bank", label: "Bank" },
  { value: "credit_card", label: "Kreditkarte" },
  { value: "savings", label: "Sparkonto" },
];

const ICON_OPTIONS = [
  { value: "wallet", label: "Wallet" },
  { value: "landmark", label: "Landmark" },
  { value: "credit-card", label: "Credit Card" },
  { value: "piggy-bank", label: "Piggy Bank" },
  { value: "banknote", label: "Banknote" },
];

const COLOR_PRESETS = [
  "#3b82f6",
  "#10b981",
  "#f59e0b",
  "#ef4444",
  "#8b5cf6",
  "#ec4899",
  "#06b6d4",
  "#f97316",
  "#14b8a6",
  "#6366f1",
  "#84cc16",
  "#64748b",
];

interface AccountFormData {
  name: string;
  type: string;
  balance?: number;
  color: string;
  icon: string;
}

interface AccountFormProps {
  onSubmit: (data: AccountFormData) => Promise<void>;
  onCancel: () => void;
  initialData?: {
    id: string;
    name: string;
    type: string;
    color: string;
    icon: string;
  };
}

export function AccountForm({ onSubmit, onCancel, initialData }: AccountFormProps) {
  const [name, setName] = useState(initialData?.name ?? "");
  const [type, setType] = useState(initialData?.type ?? "bank");
  const [balance, setBalance] = useState("0");
  const [color, setColor] = useState(initialData?.color ?? "#3b82f6");
  const [icon, setIcon] = useState(initialData?.icon ?? "wallet");
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [submitting, setSubmitting] = useState(false);

  useEffect(() => {
    if (initialData) {
      setName(initialData.name);
      setType(initialData.type);
      setColor(initialData.color);
      setIcon(initialData.icon);
    }
  }, [initialData]);

  function validate(): boolean {
    const newErrors: Record<string, string> = {};

    if (!name.trim()) {
      newErrors.name = "Bitte einen Namen eingeben";
    }

    if (!type) {
      newErrors.type = "Bitte einen Typ auswählen";
    }

    if (!initialData) {
      const parsedBalance = parseFloat(balance);
      if (isNaN(parsedBalance)) {
        newErrors.balance = "Bitte einen gültigen Betrag eingeben";
      }
    }

    if (!color) {
      newErrors.color = "Bitte eine Farbe auswählen";
    }

    if (!icon) {
      newErrors.icon = "Bitte ein Icon auswählen";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();

    if (!validate()) return;

    setSubmitting(true);
    try {
      const data: AccountFormData = {
        name: name.trim(),
        type,
        color,
        icon,
      };

      if (!initialData) {
        data.balance = parseFloat(balance);
      }

      await onSubmit(data);
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-5">
      <Input
        label="Name"
        type="text"
        placeholder="z.B. Girokonto"
        value={name}
        onChange={(e) => setName(e.target.value)}
        error={errors.name}
      />

      <Select
        label="Typ"
        options={ACCOUNT_TYPES}
        value={type}
        onChange={(e) => setType(e.target.value)}
        error={errors.type}
      />

      {!initialData && (
        <Input
          label="Startguthaben"
          type="number"
          placeholder="0,00"
          step="0.01"
          value={balance}
          onChange={(e) => setBalance(e.target.value)}
          error={errors.balance}
          icon={<span className="text-sm font-medium">&euro;</span>}
        />
      )}

      {/* Color picker */}
      <div className="w-full">
        <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1.5">
          Farbe
        </label>
        <div className="flex flex-wrap gap-2">
          {COLOR_PRESETS.map((preset) => (
            <button
              key={preset}
              type="button"
              onClick={() => setColor(preset)}
              className={`w-8 h-8 rounded-lg border-2 transition-all duration-150 ${
                color === preset
                  ? "border-gray-900 dark:border-white scale-110 shadow-md"
                  : "border-transparent hover:scale-105"
              }`}
              style={{ backgroundColor: preset }}
              title={preset}
            />
          ))}
          <div className="relative">
            <input
              type="color"
              value={color}
              onChange={(e) => setColor(e.target.value)}
              className="w-8 h-8 rounded-lg cursor-pointer border-2 border-dashed border-gray-300 dark:border-slate-600"
              title="Benutzerdefinierte Farbe"
            />
          </div>
        </div>
        {errors.color && (
          <p className="mt-1.5 text-sm text-red-500 dark:text-red-400">
            {errors.color}
          </p>
        )}
      </div>

      <Select
        label="Icon"
        options={ICON_OPTIONS}
        value={icon}
        onChange={(e) => setIcon(e.target.value)}
        error={errors.icon}
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
          {initialData ? "Speichern" : "Erstellen"}
        </Button>
      </div>
    </form>
  );
}
