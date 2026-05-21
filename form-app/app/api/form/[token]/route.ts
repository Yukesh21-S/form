import { NextResponse } from "next/server";
import { getValidInviteToken } from "@/lib/feedback";

const jsonError = (message: string, status = 400) =>
  NextResponse.json({ error: message }, { status });

export async function GET(
  _request: Request,
  { params }: { params: Promise<{ token: string }> }
) {
  const { token } = await params;
  if (!token) {
    return jsonError("Token is required", 400);
  }

  const invite = await getValidInviteToken(token);
  if (!invite) {
    return jsonError("Invalid or expired token", 404);
  }

  const form = invite.form;

  return NextResponse.json({
    formId: form.id,
    title: form.title,
    questions: form.questions.map((question: any) => ({
      questionId: question.id,
      text: question.text,
      options: question.options.map((option: any) => ({
        optionId: option.id,
        label: option.label,
        value: option.value,
      })),
    })),
  });
}
