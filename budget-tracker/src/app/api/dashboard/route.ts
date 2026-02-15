import { NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET() {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const userId = session.user.id;
    const now = new Date();
    const currentMonth = now.getMonth(); // 0-indexed
    const currentYear = now.getFullYear();

    // Current month date range
    const monthStart = new Date(currentYear, currentMonth, 1);
    const monthEnd = new Date(currentYear, currentMonth + 1, 0, 23, 59, 59, 999);

    // Total balance (sum of all account balances)
    const accounts = await prisma.account.findMany({
      where: { userId },
      select: { balance: true },
    });
    const totalBalance = accounts.reduce((sum, account) => sum + account.balance, 0);

    // Total income for current month
    const incomeAggregate = await prisma.transaction.aggregate({
      where: {
        userId,
        type: "income",
        date: {
          gte: monthStart,
          lte: monthEnd,
        },
      },
      _sum: {
        amount: true,
      },
    });
    const totalIncome = incomeAggregate._sum.amount || 0;

    // Total expenses for current month
    const expenseAggregate = await prisma.transaction.aggregate({
      where: {
        userId,
        type: "expense",
        date: {
          gte: monthStart,
          lte: monthEnd,
        },
      },
      _sum: {
        amount: true,
      },
    });
    const totalExpenses = expenseAggregate._sum.amount || 0;

    // Recent transactions (last 5)
    const recentTransactions = await prisma.transaction.findMany({
      where: { userId },
      include: {
        category: true,
        account: true,
      },
      orderBy: {
        date: "desc",
      },
      take: 5,
    });

    // Monthly data (last 6 months of income vs expense totals)
    const monthlyData = [];
    for (let i = 5; i >= 0; i--) {
      const date = new Date(currentYear, currentMonth - i, 1);
      const mStart = new Date(date.getFullYear(), date.getMonth(), 1);
      const mEnd = new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59, 999);

      const [monthIncome, monthExpense] = await Promise.all([
        prisma.transaction.aggregate({
          where: {
            userId,
            type: "income",
            date: {
              gte: mStart,
              lte: mEnd,
            },
          },
          _sum: { amount: true },
        }),
        prisma.transaction.aggregate({
          where: {
            userId,
            type: "expense",
            date: {
              gte: mStart,
              lte: mEnd,
            },
          },
          _sum: { amount: true },
        }),
      ]);

      monthlyData.push({
        month: mStart.toLocaleString("default", { month: "short" }),
        year: mStart.getFullYear(),
        income: monthIncome._sum.amount || 0,
        expenses: monthExpense._sum.amount || 0,
      });
    }

    // Category breakdown (expense totals by category for current month)
    const categoryExpenses = await prisma.transaction.groupBy({
      by: ["categoryId"],
      where: {
        userId,
        type: "expense",
        date: {
          gte: monthStart,
          lte: monthEnd,
        },
      },
      _sum: {
        amount: true,
      },
    });

    // Fetch category details for the breakdown
    const categoryIds = categoryExpenses.map((ce) => ce.categoryId);
    const categories = await prisma.category.findMany({
      where: {
        id: { in: categoryIds },
      },
    });

    const categoryMap = new Map(categories.map((c) => [c.id, c]));
    const categoryBreakdown = categoryExpenses.map((ce) => ({
      category: categoryMap.get(ce.categoryId),
      amount: ce._sum.amount || 0,
    }));

    return NextResponse.json({
      totalBalance,
      totalIncome,
      totalExpenses,
      recentTransactions,
      monthlyData,
      categoryBreakdown,
    });
  } catch (error) {
    console.error("Error fetching dashboard data:", error);
    return NextResponse.json(
      { error: "Failed to fetch dashboard data" },
      { status: 500 }
    );
  }
}
