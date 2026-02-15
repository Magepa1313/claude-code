import { NextResponse } from "next/server";
import { hash } from "bcryptjs";
import { prisma } from "@/lib/prisma";

export async function POST(req: Request) {
  try {
    const { name, email, password } = await req.json();

    if (!name || !email || !password) {
      return NextResponse.json(
        { error: "Name, email and password are required" },
        { status: 400 }
      );
    }

    const existingUser = await prisma.user.findUnique({
      where: { email },
    });

    if (existingUser) {
      return NextResponse.json(
        { error: "User with this email already exists" },
        { status: 400 }
      );
    }

    const hashedPassword = await hash(password, 12);

    const user = await prisma.user.create({
      data: {
        name,
        email,
        password: hashedPassword,
      },
    });

    // Create default categories for the user
    const defaultCategories = [
      { name: "Gehalt", type: "income", color: "#22c55e", icon: "banknote" },
      { name: "Freelance", type: "income", color: "#3b82f6", icon: "laptop" },
      { name: "Investitionen", type: "income", color: "#8b5cf6", icon: "trending-up" },
      { name: "Sonstiges", type: "income", color: "#6b7280", icon: "plus-circle" },
      { name: "Lebensmittel", type: "expense", color: "#ef4444", icon: "shopping-cart" },
      { name: "Transport", type: "expense", color: "#f97316", icon: "car" },
      { name: "Wohnung", type: "expense", color: "#eab308", icon: "home" },
      { name: "Unterhaltung", type: "expense", color: "#ec4899", icon: "film" },
      { name: "Gesundheit", type: "expense", color: "#14b8a6", icon: "heart" },
      { name: "Bildung", type: "expense", color: "#6366f1", icon: "book-open" },
      { name: "Kleidung", type: "expense", color: "#f43f5e", icon: "shirt" },
      { name: "Restaurant", type: "expense", color: "#d946ef", icon: "utensils" },
    ];

    await prisma.category.createMany({
      data: defaultCategories.map((cat) => ({
        ...cat,
        userId: user.id,
      })),
    });

    // Create default account
    await prisma.account.create({
      data: {
        name: "Hauptkonto",
        type: "bank",
        balance: 0,
        color: "#3b82f6",
        icon: "landmark",
        userId: user.id,
      },
    });

    return NextResponse.json(
      { message: "User created successfully" },
      { status: 201 }
    );
  } catch (error) {
    console.error("Registration error:", error);
    return NextResponse.json(
      { error: "Something went wrong" },
      { status: 500 }
    );
  }
}
