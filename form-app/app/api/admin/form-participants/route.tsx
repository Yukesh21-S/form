import { NextRequest, NextResponse }
from "next/server";

import prisma from "@/lib/prisma";

//  
export async function GET(
  req: NextRequest
) {

  const { searchParams } =
    new URL(req.url);

  const formId =
    searchParams.get("formId");

  if (!formId) {

    return NextResponse.json(
      { error: "formId required" },
      { status: 400 }
    );
  }

  const invites =
    await prisma.inviteToken.findMany({

      where: {
        formId,
        participantId: {
          not: null,
        },
      },

      distinct: ["participantId"],
    });

  const participantIds =
    invites
      .map((i) => i.participantId)
      .filter(Boolean);

  const participants =
    await prisma.participant.findMany({

      where: {
        id: {
          in: participantIds as string[],
        },
      },
    });

  return NextResponse.json(
    participants
  );
}