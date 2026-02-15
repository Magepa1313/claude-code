"use client";

import { useEffect, useState, useCallback } from "react";
import { Plus, Pencil, Trash2 } from "lucide-react";
import { Card } from "@/components/ui/Card";
import { Button } from "@/components/ui/Button";
import { Modal } from "@/components/ui/Modal";
import { CategoryForm } from "@/components/categories/CategoryForm";

interface Category {
  id: string;
  name: string;
  type: string;
  color: string;
  icon: string;
}

interface CategoryFormData {
  name: string;
  type: string;
  color: string;
  icon: string;
}

function SkeletonCard() {
  return (
    <div className="bg-white dark:bg-slate-800 rounded-xl shadow-sm border border-gray-200 dark:border-slate-700 p-4 animate-pulse">
      <div className="flex items-center gap-3">
        <div className="w-1 h-10 rounded bg-gray-200 dark:bg-slate-700" />
        <div className="flex-1">
          <div className="h-4 w-20 bg-gray-200 dark:bg-slate-700 rounded" />
          <div className="h-3 w-16 bg-gray-200 dark:bg-slate-700 rounded mt-1.5" />
        </div>
        <div className="flex gap-1">
          <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
          <div className="w-8 h-8 bg-gray-200 dark:bg-slate-700 rounded-lg" />
        </div>
      </div>
    </div>
  );
}

function EmptyState({ type }: { type: string }) {
  return (
    <div className="col-span-full text-center py-10">
      <div className="text-gray-400 dark:text-gray-500 mb-2">
        <Plus className="w-10 h-10 mx-auto" />
      </div>
      <p className="text-sm text-gray-500 dark:text-gray-400">
        Keine {type === "income" ? "Einnahmen" : "Ausgaben"}-Kategorien
        vorhanden.
      </p>
      <p className="text-xs text-gray-400 dark:text-gray-500 mt-1">
        Erstellen Sie eine neue Kategorie, um loszulegen.
      </p>
    </div>
  );
}

export default function CategoriesPage() {
  const [categories, setCategories] = useState<Category[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingCategory, setEditingCategory] = useState<Category | null>(null);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);

  const fetchCategories = useCallback(async () => {
    try {
      setLoading(true);
      const response = await fetch("/api/categories");
      if (!response.ok) {
        throw new Error("Kategorien konnten nicht geladen werden");
      }
      const data = await response.json();
      setCategories(data);
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
    fetchCategories();
  }, [fetchCategories]);

  async function handleCreate(data: CategoryFormData) {
    try {
      setSubmitting(true);
      const response = await fetch("/api/categories", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });
      if (!response.ok) {
        throw new Error("Kategorie konnte nicht erstellt werden");
      }
      setIsModalOpen(false);
      await fetchCategories();
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setSubmitting(false);
    }
  }

  async function handleUpdate(data: CategoryFormData) {
    if (!editingCategory) return;
    try {
      setSubmitting(true);
      const response = await fetch(`/api/categories/${editingCategory.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });
      if (!response.ok) {
        throw new Error("Kategorie konnte nicht aktualisiert werden");
      }
      setEditingCategory(null);
      setIsModalOpen(false);
      await fetchCategories();
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDelete(id: string) {
    try {
      setSubmitting(true);
      const response = await fetch(`/api/categories/${id}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        throw new Error("Kategorie konnte nicht gelöscht werden");
      }
      setDeleteConfirmId(null);
      await fetchCategories();
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : "Ein unbekannter Fehler ist aufgetreten"
      );
    } finally {
      setSubmitting(false);
    }
  }

  function openCreateModal() {
    setEditingCategory(null);
    setIsModalOpen(true);
  }

  function openEditModal(category: Category) {
    setEditingCategory(category);
    setIsModalOpen(true);
  }

  function closeModal() {
    setIsModalOpen(false);
    setEditingCategory(null);
  }

  const incomeCategories = categories.filter((c) => c.type === "income");
  const expenseCategories = categories.filter((c) => c.type === "expense");

  if (error) {
    return (
      <div className="flex items-center justify-center min-h-[400px]">
        <div className="text-center">
          <p className="text-red-500 dark:text-red-400 text-sm font-medium">
            {error}
          </p>
          <button
            onClick={() => {
              setError(null);
              fetchCategories();
            }}
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
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-gray-900 dark:text-white">
            Kategorien
          </h1>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
            Verwalten Sie Ihre Einnahmen- und Ausgaben-Kategorien.
          </p>
        </div>
        <Button
          onClick={openCreateModal}
          icon={<Plus className="w-4 h-4" />}
        >
          Neue Kategorie
        </Button>
      </div>

      {/* Income categories section */}
      <div>
        <h2 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
          Einnahmen-Kategorien
        </h2>
        {loading ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {Array.from({ length: 4 }).map((_, i) => (
              <SkeletonCard key={i} />
            ))}
          </div>
        ) : incomeCategories.length === 0 ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            <EmptyState type="income" />
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {incomeCategories.map((category) => (
              <Card key={category.id} className="!p-0 overflow-hidden">
                <div className="flex items-center gap-3 p-4">
                  <div
                    className="w-1 self-stretch rounded-full flex-shrink-0"
                    style={{ backgroundColor: category.color }}
                  />
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                      {category.name}
                    </p>
                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-0.5">
                      {category.icon}
                    </p>
                  </div>
                  <div className="flex items-center gap-1 flex-shrink-0">
                    <button
                      onClick={() => openEditModal(category)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 transition-colors"
                      title="Bearbeiten"
                    >
                      <Pencil className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => setDeleteConfirmId(category.id)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:text-red-400 dark:hover:bg-red-900/20 transition-colors"
                      title="Löschen"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </Card>
            ))}
          </div>
        )}
      </div>

      {/* Expense categories section */}
      <div>
        <h2 className="text-lg font-semibold text-gray-900 dark:text-white mb-4">
          Ausgaben-Kategorien
        </h2>
        {loading ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {Array.from({ length: 4 }).map((_, i) => (
              <SkeletonCard key={i} />
            ))}
          </div>
        ) : expenseCategories.length === 0 ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            <EmptyState type="expense" />
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {expenseCategories.map((category) => (
              <Card key={category.id} className="!p-0 overflow-hidden">
                <div className="flex items-center gap-3 p-4">
                  <div
                    className="w-1 self-stretch rounded-full flex-shrink-0"
                    style={{ backgroundColor: category.color }}
                  />
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-gray-900 dark:text-white truncate">
                      {category.name}
                    </p>
                    <p className="text-xs text-gray-500 dark:text-gray-400 mt-0.5">
                      {category.icon}
                    </p>
                  </div>
                  <div className="flex items-center gap-1 flex-shrink-0">
                    <button
                      onClick={() => openEditModal(category)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:text-blue-400 dark:hover:bg-blue-900/20 transition-colors"
                      title="Bearbeiten"
                    >
                      <Pencil className="w-4 h-4" />
                    </button>
                    <button
                      onClick={() => setDeleteConfirmId(category.id)}
                      className="p-1.5 rounded-lg text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:text-red-400 dark:hover:bg-red-900/20 transition-colors"
                      title="Löschen"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </Card>
            ))}
          </div>
        )}
      </div>

      {/* Create/Edit Modal */}
      <Modal
        isOpen={isModalOpen}
        onClose={closeModal}
        title={editingCategory ? "Kategorie bearbeiten" : "Neue Kategorie"}
      >
        <CategoryForm
          onSubmit={editingCategory ? handleUpdate : handleCreate}
          onCancel={closeModal}
          initialData={
            editingCategory
              ? {
                  name: editingCategory.name,
                  type: editingCategory.type,
                  color: editingCategory.color,
                  icon: editingCategory.icon,
                }
              : undefined
          }
        />
      </Modal>

      {/* Delete confirmation Modal */}
      <Modal
        isOpen={deleteConfirmId !== null}
        onClose={() => setDeleteConfirmId(null)}
        title="Kategorie löschen"
      >
        <div className="space-y-4">
          <p className="text-sm text-gray-600 dark:text-gray-300">
            Sind Sie sicher, dass Sie diese Kategorie löschen möchten? Alle
            zugehörigen Transaktionen werden ebenfalls gelöscht. Diese Aktion
            kann nicht rückgängig gemacht werden.
          </p>
          <div className="flex items-center justify-end gap-3">
            <Button
              variant="secondary"
              onClick={() => setDeleteConfirmId(null)}
            >
              Abbrechen
            </Button>
            <Button
              variant="danger"
              loading={submitting}
              onClick={() => {
                if (deleteConfirmId) handleDelete(deleteConfirmId);
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
