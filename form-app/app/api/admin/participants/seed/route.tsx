import { NextResponse } from "next/server";

import prisma from "@/lib/prisma";

const PARTICIPANTS = [
  {
    fullName: "Marcus Johnson",
    email: "marcus@gmail.com",
    department: "Engineering",
    designation: "Senior Manager",
  },
  {
    fullName: "Sarah Williams",
    email: "sarah@gmail.com",
    department: "HR",
    designation: "Director",
  },
  {
    fullName: "David Miller",
    email: "david@gmail.com",
    department: "Finance",
    designation: "Lead",
  },
];

export async function POST() {

  await prisma.participant.createMany({
    data: PARTICIPANTS,
    skipDuplicates: true,
  });

  return NextResponse.json({
    success: true,
  });
}