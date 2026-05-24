import { NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET() {

  const participants =
    await prisma.participant.findMany();

  return NextResponse.json(participants);
}