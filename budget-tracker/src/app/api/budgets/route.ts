import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const { searchParams } = new URL(request.url);
    const month = searchParams.get("month");
    const year = searchParams.get("year");

    if (!month || !year) {
      return NextResponse.json(
        { error: "Missing required query params: month, year" },
        { status: 400 }
      );
    }

    const monthNum = parseInt(month, 10);
    const yearNum = parseInt(year, 10);

    // Get the start and end dates for the given month
    const startDate = new Date(yearNum, monthNum - 1, 1);
    const endDate = new Date(yearNum, monthNum, 0, 23, 59, 59, 999);

    const budgets = await prisma.budget.findMany({
      where: {
        userId: session.user.id,
        month: monthNum,
        year: yearNum,
      },
      include: {
        category: true,
      },
      orderBy: {
        category: {
          name: "asc",
        },
      },
    });

    // Calculate spent amount for each budget from transactions in that month
    const budgetsWithSpent = await Promise.all(
      budgets.map(async (budget) => {
        const spent = await prisma.transaction.aggregate({
          where: {
            userId: session.user.id,
            categoryId: budget.categoryId,
            type: "expense",
            date: {
              gte: startDate,
              lte: endDate,
            },
          },
          _sum: {
            amount: true,
          },
        });

        return {
          ...budget,
          spent: spent._sum.amount || 0,
        };
      })
    );

    return NextResponse.json(budgetsWithSpent);
  } catch (error) {
    console.error("Error fetching budgets:", error);
    return NextResponse.json(
      { error: "Failed to fetch budgets" },
      { status: 500 }
    );
  }
}

export async function POST(request: NextRequest) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const body = await request.json();
    const { amount, month, year, categoryId } = body;

    if (amount === undefined || !month || !year || !categoryId) {
      return NextResponse.json(
        { error: "Missing required fields: amount, month, year, categoryId" },
        { status: 400 }
      );
    }

    // Verify the category belongs to the user
    const category = await prisma.category.findFirst({
      where: { id: categoryId, userId: session.user.id },
    });

    if (!category) {
      return NextResponse.json(
        { error: "Category not found" },
        { status: 404 }
      );
    }

    // Upsert: create or update budget by categoryId + month + year
    const budget = await prisma.budget.upsert({
      where: {
        categoryId_month_year: {
          categoryId,
          month,
          year,
        },
      },
      update: {
        amount,
      },
      create: {
        amount,
        month,
        year,
        categoryId,
        userId: session.user.id,
      },
      include: {
        category: true,
      },
    });

    return NextResponse.json(budget, { status: 201 });
  } catch (error) {
    console.error("Error creating/updating budget:", error);
    return NextResponse.json(
      { error: "Failed to create/update budget" },
      { status: 500 }
    );
  }
}
