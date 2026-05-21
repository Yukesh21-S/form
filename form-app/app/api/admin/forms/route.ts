import { NextResponse } from "next/server";
import { z } from "zod";
import prisma from "@/lib/prisma";
import { QUESTION_OPTIONS } from "@/lib/feedback";

const createFormSchema = z.object({
  title: z.string().min(1),
  questions: z.array(z.object({ text: z.string().min(1) })).min(1),
});

const jsonError = (message: string, status = 400) =>
  NextResponse.json({ error: message }, { status });

export async function GET() {
  const forms = await prisma.form.findMany({
    include: {
      questions: {
        include: {
          options: true,
        },
      },
    },
  });

  return NextResponse.json(forms);
}

export async function POST(request: Request) {
  const body = await request.json().catch(() => null);

  if (!body) {
    return jsonError("Invalid JSON payload", 400);
  }

  const parseResult = createFormSchema.safeParse(body);
  if (!parseResult.success) {
    return jsonError(parseResult.error.message, 422);
  }

  const { title, questions } = parseResult.data;

  const form = await prisma.form.create({
    data: {
      title,
      questions: {
        create: questions.map((question) => ({
          text: question.text,
          options: {
            create: QUESTION_OPTIONS.map((option) => ({
              label: option.label,
              value: option.value,
            })),
          },
        })),
      },
    },
    include: {
      questions: {
        include: {
          options: true,
        },
      },
    },
  });

  return NextResponse.json(form, { status: 201 });
}
