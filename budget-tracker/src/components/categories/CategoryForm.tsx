"use client";

import { useState } from "react";
import { Input } from "@/components/ui/Input";
import { Select } from "@/components/ui/Select";
import { Button } from "@/components/ui/Button";

const PREDEFINED_COLORS = [
  "#ef4444",
  "#f97316",
  "#eab308",
  "#22c55e",
  "#14b8a6",
  "#3b82f6",
  "#6366f1",
  "#8b5cf6",
  "#ec4899",
  "#f43f5e",
  "#d946ef",
  "#6b7280",
];

const PREDEFINED_ICONS = [
  "shopping-cart",
  "car",
  "home",
  "film",
  "heart",
  "book-open",
  "shirt",
  "utensils",
  "banknote",
  "laptop",
  "trending-up",
  "plus-circle",
  "tag",
  "gift",
  "coffee",
  "music",
  "plane",
  "briefcase",
  "phone",
  "gamepad-2",
];

const TYPE_OPTIONS = [
  { value: "income", label: "Einnahme" },
  { value: "expense", label: "Ausgabe" },
];

interface CategoryFormData {
  name: string;
  type: string;
  color: string;
  icon: string;
}

interface CategoryFormProps {
  onSubmit: (data: CategoryFormData) => void;
  onCancel: () => void;
  initialData?: Partial<CategoryFormData>;
}

export function CategoryForm({
  onSubmit,
  onCancel,
  initialData,
}: CategoryFormProps) {
  const [name, setName] = useState(initialData?.name ?? "");
  const [type, setType] = useState(initialData?.type ?? "");
  const [color, setColor] = useState(initialData?.color ?? PREDEFINED_COLORS[0]);
  const [icon, setIcon] = useState(initialData?.icon ?? PREDEFINED_ICONS[0]);
  const [errors, setErrors] = useState<{ name?: string; type?: string }>({});

  function validate(): boolean {
    const newErrors: { name?: string; type?: string } = {};
    if (!name.trim()) {
      newErrors.name = "Name ist erforderlich";
    }
    if (!type) {
      newErrors.type = "Typ ist erforderlich";
    }
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  }

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    if (!validate()) return;
    onSubmit({ name: name.trim(), type, color, icon });
  }

  return (
    <form onSubmit={handleSubmit} className="space-y-5">
      {/* Name field */}
      <Input
        label="Name"
        value={name}
        onChange={(e) => setName(e.target.value)}
        placeholder="z.B. Lebensmittel, Gehalt..."
        error={errors.name}
      />

      {/* Type select */}
      <Select
        label="Typ"
        value={type}
        onChange={(e) => setType(e.target.value)}
        options={TYPE_OPTIONS}
        placeholder="Typ auswählen"
        error={errors.type}
      />

      {/* Color picker */}
      <div>
        <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1.5">
          Farbe
        </label>
        <div className="flex flex-wrap gap-2">
          {PREDEFINED_COLORS.map((c) => (
            <button
              key={c}
              type="button"
              onClick={() => setColor(c)}
              className={`w-8 h-8 rounded-full transition-all duration-150 ${
                color === c
                  ? "ring-2 ring-offset-2 ring-offset-white dark:ring-offset-slate-800 ring-blue-500 scale-110"
                  : "hover:scale-110"
              }`}
              style={{ backgroundColor: c }}
              aria-label={`Farbe ${c}`}
            />
          ))}
        </div>
      </div>

      {/* Icon picker */}
      <div>
        <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-1.5">
          Icon
        </label>
        <div className="grid grid-cols-4 gap-2 max-h-48 overflow-y-auto">
          {PREDEFINED_ICONS.map((ic) => (
            <button
              key={ic}
              type="button"
              onClick={() => setIcon(ic)}
              className={`px-2 py-1.5 rounded-lg text-xs font-medium transition-colors duration-150 border ${
                icon === ic
                  ? "bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300 border-blue-400 dark:border-blue-500"
                  : "bg-gray-50 dark:bg-slate-700 text-gray-600 dark:text-gray-300 border-gray-200 dark:border-slate-600 hover:bg-gray-100 dark:hover:bg-slate-600"
              }`}
            >
              {ic}
            </button>
          ))}
        </div>
      </div>

      {/* Action buttons */}
      <div className="flex items-center justify-end gap-3 pt-2">
        <Button type="button" variant="secondary" onClick={onCancel}>
          Abbrechen
        </Button>
        <Button type="submit">Speichern</Button>
      </div>
    </form>
  );
}
