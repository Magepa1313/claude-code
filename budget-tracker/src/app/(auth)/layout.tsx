import { Wallet } from "lucide-react";

export default function AuthLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gray-50 dark:bg-gray-950 px-4 py-12">
      <div className="mb-8 flex flex-col items-center gap-3">
        <div className="flex items-center justify-center w-14 h-14 rounded-2xl bg-blue-500 text-white shadow-lg shadow-blue-500/30">
          <Wallet className="w-7 h-7" />
        </div>
        <h1 className="text-2xl font-bold text-gray-900 dark:text-white tracking-tight">
          Budget Tracker
        </h1>
        <p className="text-sm text-gray-500 dark:text-gray-400">
          Deine Finanzen im Griff
        </p>
      </div>
      {children}
    </div>
  );
}
