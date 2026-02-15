import { NextRequest, NextResponse } from "next/server";
import { getServerSession } from "next-auth";
import { authOptions } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export async function GET(
  request: NextRequest,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const transaction = await prisma.transaction.findFirst({
      where: {
        id: params.id,
        userId: session.user.id,
      },
      include: {
        category: true,
        account: true,
      },
    });

    if (!transaction) {
      return NextResponse.json(
        { error: "Transaction not found" },
        { status: 404 }
      );
    }

    return NextResponse.json(transaction);
  } catch (error) {
    console.error("Error fetching transaction:", error);
    return NextResponse.json(
      { error: "Failed to fetch transaction" },
      { status: 500 }
    );
  }
}

export async function PUT(
  request: NextRequest,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    // Fetch existing transaction
    const existingTransaction = await prisma.transaction.findFirst({
      where: {
        id: params.id,
        userId: session.user.id,
      },
    });

    if (!existingTransaction) {
      return NextResponse.json(
        { error: "Transaction not found" },
        { status: 404 }
      );
    }

    const body = await request.json();
    const { amount, type, description, date, categoryId, accountId } = body;

    // Determine new values (use existing if not provided)
    const newAmount = amount ?? existingTransaction.amount;
    const newType = type ?? existingTransaction.type;
    const newAccountId = accountId ?? existingTransaction.accountId;

    // Validate type if provided
    if (type && type !== "income" && type !== "expense") {
      return NextResponse.json(
        { error: "Type must be 'income' or 'expense'" },
        { status: 400 }
      );
    }

    // Verify new account belongs to user if account is changing
    if (accountId && accountId !== existingTransaction.accountId) {
      const account = await prisma.account.findFirst({
        where: { id: accountId, userId: session.user.id },
      });
      if (!account) {
        return NextResponse.json(
          { error: "Account not found" },
          { status: 404 }
        );
      }
    }

    // Verify new category belongs to user if category is changing
    if (categoryId && categoryId !== existingTransaction.categoryId) {
      const category = await prisma.category.findFirst({
        where: { id: categoryId, userId: session.user.id },
      });
      if (!category) {
        return NextResponse.json(
          { error: "Category not found" },
          { status: 404 }
        );
      }
    }

    // Calculate balance adjustments
    // Step 1: Reverse the old transaction's effect on the old account
    const oldBalanceChange =
      existingTransaction.type === "income"
        ? -existingTransaction.amount
        : existingTransaction.amount;

    // Step 2: Apply the new transaction's effect on the new account
    const newBalanceChange = newType === "income" ? newAmount : -newAmount;

    const operations = [];

    // Reverse old transaction on old account
    operations.push(
      prisma.account.update({
        where: { id: existingTransaction.accountId },
        data: {
          balance: {
            increment: oldBalanceChange,
          },
        },
      })
    );

    // Apply new transaction on new account
    // If the account is the same, both increments will apply
    operations.push(
      prisma.account.update({
        where: { id: newAccountId },
        data: {
          balance: {
            increment: newBalanceChange,
          },
        },
      })
    );

    // Update the transaction
    const updateData: Record<string, unknown> = {};
    if (amount !== undefined) updateData.amount = amount;
    if (type !== undefined) updateData.type = type;
    if (description !== undefined) updateData.description = description;
    if (date !== undefined) updateData.date = new Date(date);
    if (categoryId !== undefined) updateData.categoryId = categoryId;
    if (accountId !== undefined) updateData.accountId = accountId;

    operations.push(
      prisma.transaction.update({
        where: { id: params.id },
        data: updateData,
        include: {
          category: true,
          account: true,
        },
      })
    );

    const results = await prisma.$transaction(operations);
    const updatedTransaction = results[results.length - 1];

    return NextResponse.json(updatedTransaction);
  } catch (error) {
    console.error("Error updating transaction:", error);
    return NextResponse.json(
      { error: "Failed to update transaction" },
      { status: 500 }
    );
  }
}

export async function DELETE(
  request: NextRequest,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerSession(authOptions);
    if (!session?.user?.id) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
    }

    const transaction = await prisma.transaction.findFirst({
      where: {
        id: params.id,
        userId: session.user.id,
      },
    });

    if (!transaction) {
      return NextResponse.json(
        { error: "Transaction not found" },
        { status: 404 }
      );
    }

    // Reverse the balance change
    const balanceChange =
      transaction.type === "income" ? -transaction.amount : transaction.amount;

    await prisma.$transaction([
      prisma.account.update({
        where: { id: transaction.accountId },
        data: {
          balance: {
            increment: balanceChange,
          },
        },
      }),
      prisma.transaction.delete({
        where: { id: params.id },
      }),
    ]);

    return NextResponse.json({ message: "Transaction deleted successfully" });
  } catch (error) {
    console.error("Error deleting transaction:", error);
    return NextResponse.json(
      { error: "Failed to delete transaction" },
      { status: 500 }
    );
  }
}
