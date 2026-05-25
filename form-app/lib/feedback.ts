import prisma from "./prisma";

export const QUESTION_OPTIONS = [
  { label: "Rarely", value: 25 },
  { label: "Sometimes", value: 50 },
  { label: "Often", value: 75 },
  { label: "Always", value: 100 },
  { label: "Insufficient Exposure", value: null },
] as const;

export const INVITE_TOKEN_LIFETIME_HOURS = 48;

export async function getValidInviteToken(token: string) {
  const invite = await prisma.inviteToken.findUnique({
    where: { token },
    include: {
      participant: true,
      form: {
        include: {
          questions: {
            include: {
              options: true,
            },
          },
        },
      },
    },
  });

  if (!invite) {
    return null;
  }

  if (invite.isUsed) {
    return null;
  }

  if (invite.expiresAt < new Date()) {
    return null;
  }

  return invite;
}
