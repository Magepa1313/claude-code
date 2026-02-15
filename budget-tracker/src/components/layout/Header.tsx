"use client";

import { usePathname } from "next/navigation";
import { useSession } from "next-auth/react";
import { useTheme } from "next-themes";
import { Menu, Sun, Moon, Bell, Search } from "lucide-react";
import { useEffect, useState } from "react";

interface HeaderProps {
  onMenuToggle: () => void;
}

const pageTitles: Record<string, string> = {
  "/dashboard": "Dashboard",
  "/transactions": "Transaktionen",
  "/categories": "Kategorien",
  "/budgets": "Budgets",
  "/accounts": "Konten",
};

function getPageTitle(pathname: string | null): string {
  if (!pathname) return "Dashboard";

  // Check for exact match first
  if (pageTitles[pathname]) return pageTitles[pathname];

  // Check for prefix match (e.g. /transactions/new)
  for (const [path, title] of Object.entries(pageTitles)) {
    if (pathname.startsWith(path + "/")) return title;
  }

  return "Dashboard";
}

export default function Header({ onMenuToggle }: HeaderProps) {
  const pathname = usePathname();
  const { data: session } = useSession();
  const { theme, setTheme } = useTheme();
  const [mounted, setMounted] = useState(false);

  // Avoid hydration mismatch with theme
  useEffect(() => {
    setMounted(true);
  }, []);

  const pageTitle = getPageTitle(pathname);

  const userInitial = session?.user?.name
    ? session.user.name.charAt(0).toUpperCase()
    : session?.user?.email?.charAt(0).toUpperCase() ?? "U";

  return (
    <header className="sticky top-0 z-30 flex h-16 items-center gap-4 border-b border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 px-4 sm:px-6">
      {/* Mobile menu button */}
      <button
        onClick={onMenuToggle}
        className="rounded-md p-2 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800 hover:text-slate-700 dark:hover:text-slate-200 lg:hidden"
      >
        <Menu className="h-5 w-5" />
      </button>

      {/* Page title */}
      <h1 className="text-lg font-semibold text-slate-900 dark:text-white sm:text-xl">
        {pageTitle}
      </h1>

      {/* Spacer */}
      <div className="flex-1" />

      {/* Search input */}
      <div className="hidden md:flex items-center">
        <div className="relative">
          <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Suchen..."
            className="h-9 w-64 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 pl-9 pr-4 text-sm text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:border-primary-500 focus:outline-none focus:ring-1 focus:ring-primary-500"
          />
        </div>
      </div>

      {/* Notification bell */}
      <button className="relative rounded-md p-2 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800 hover:text-slate-700 dark:hover:text-slate-200">
        <Bell className="h-5 w-5" />
        <span className="absolute right-1.5 top-1.5 h-2 w-2 rounded-full bg-red-500" />
      </button>

      {/* Theme toggle */}
      <button
        onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
        className="rounded-md p-2 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-800 hover:text-slate-700 dark:hover:text-slate-200"
        title="Farbmodus wechseln"
      >
        {mounted ? (
          theme === "dark" ? (
            <Sun className="h-5 w-5" />
          ) : (
            <Moon className="h-5 w-5" />
          )
        ) : (
          <div className="h-5 w-5" />
        )}
      </button>

      {/* User avatar / name */}
      <div className="flex items-center gap-3">
        <div className="hidden sm:block text-right">
          <p className="text-sm font-medium text-slate-700 dark:text-slate-200">
            {session?.user?.name ?? "Benutzer"}
          </p>
        </div>
        <div className="flex h-8 w-8 items-center justify-center rounded-full bg-primary-600 text-sm font-semibold text-white">
          {userInitial}
        </div>
      </div>
    </header>
  );
}
