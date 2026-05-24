"use client";

import { Survey } from "survey-react-ui";

export default function SurveyRenderer({ form, onSubmit }: any) {
  const surveyJson = {
    elements: form.questions.map((q: any) => ({
      type: "radiogroup",
      name: q.questionId,
      title: q.text,
      isRequired: true,
      choices: q.options.map((opt: any) => ({
        value: opt.optionId,
        text: opt.label,
      })),
    })),
  };

  return <Survey json={surveyJson} onComplete={onSubmit} />;
}