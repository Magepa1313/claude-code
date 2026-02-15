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
    const type = searchParams.get("type");
    const categoryId = searchParams.get("categoryId");
    const accountId = searchParams.get("accountId");
    const startDate = searchParams.get("startDate");
    const endDate = searchParams.get("endDate");
    const search = searchParams.get("search");
    const sort = searchParams.get("sort") || "date";
    const order = searchParams.get("order") || "desc";

    const where: Record<string, unknown> = {
      userId: session.user.id,
    };

    if (type) {
      where.type = type;
    }

    if (categoryId) {
      where.categoryId = categoryId;
    }

    if (accountId) {
      where.accountId = accountId;
    }

    if (startDate || endDate) {
      where.date = {};
      if (startDate) {
        (where.date as Record<string, unknown>).gte = new Date(startDate);
      }
      if (endDate) {
        (where.date as Record<string, unknown>).lte = new Date(endDate);
      }
    }

    if (search) {
      where.description = {
        contains: search,
      };
    }

    const orderBy: Record<string, string> = {};
    if (sort === "amount") {
      orderBy.amount = order;
    } else {
      orderBy.date = order;
    }

    const transactions = await prisma.transaction.findMany({
      where,
      include: {
        category: true,
        account: true,
      },
      orderBy,
    });

    return NextResponse.json(transactions);
  } catch (error) {
    console.error("Error fetching transactions:", error);
    return NextResponse.json(
      { error: "Failed to fetch transactions" },
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
    const { amount, type, description, date, categoryId, accountId } = body;

    if (!amount || !type || !description || !date || !categoryId || !accountId) {
      return NextResponse.json(
        { error: "Missing required fields: amount, type, description, date, categoryId, accountId" },
        { status: 400 }
      );
    }

    if (type !== "income" && type !== "expense") {
      return NextResponse.json(
        { error: "Type must be 'income' or 'expense'" },
        { status: 400 }
      );
    }

    // Verify the account belongs to the user
    const account = await prisma.account.findFirst({
      where: { id: accountId, userId: session.user.id },
    });

    if (!account) {
      return NextResponse.json(
        { error: "Account not found" },
        { status: 404 }
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

    // Create transaction and update account balance in a transaction
    const balanceChange = type === "income" ? amount : -amount;

    const [transaction] = await prisma.$transaction([
      prisma.transaction.create({
        data: {
          amount,
          type,
          description,
          date: new Date(date),
          categoryId,
          accountId,
          userId: session.user.id,
        },
        include: {
          category: true,
          account: true,
        },
      }),
      prisma.account.update({
        where: { id: accountId },
        data: {
          balance: {
            increment: balanceChange,
          },
        },
      }),
    ]);

    return NextResponse.json(transaction, { status: 201 });
  } catch (error) {
    console.error("Error creating transaction:", error);
    return NextResponse.json(
      { error: "Failed to create transaction" },
      { status: 500 }
    );
  }
}
