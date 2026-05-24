"use client";

import { useParams } from "next/navigation";
import { useEffect, useState } from "react";

export default function UserFormPage() {
  const params = useParams();
  const token = params.token as string;

  const [form, setForm] = useState<any>(null);
  const [answers, setAnswers] = useState<{ [key: string]: string }>({});
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState("");
  const [alreadySubmitted, setAlreadySubmitted] = useState(false);

  //  Fetch form
  useEffect(() => {
    const fetchForm = async () => {
      try {
        const res = await fetch(`/api/form/${token}`);
        const data = await res.json();

        if (data.alreadySubmitted) {
          setAlreadySubmitted(true);
        }

        setForm(data);
      } catch {
        setForm(null);
      } finally {
        setLoading(false);
      }
    };

    if (token) fetchForm();
  }, [token]);

  //  Handle selection
  const handleSelect = (questionId: string, optionId: string) => {
    setAnswers((prev) => ({
      ...prev,
      [questionId]: optionId,
    }));
  };

  // Submit
  const handleSubmit = async () => {
    if (!form?.questions) return;

    if (Object.keys(answers).length !== form.questions.length) {
      setMessage(" Please answer all questions");
      return;
    }

    const payload = {
      answers: Object.entries(answers).map(
        ([questionId, optionId]) => ({
          questionId,
          optionId,
        })
      ),
    };

    try {
      const res = await fetch(`/api/form/${token}/submit`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      if (!res.ok) throw new Error();

      setAlreadySubmitted(true);
    } catch {
      setMessage(" Submission failed");
    }
  };

  //  Loading
  if (loading) return <div className="p-6">Loading...</div>;

  //  Invalid token
  if (!form || !form.questions)
    return <div className="p-6">Invalid or expired link</div>;

  //  Already submitted page
  if (alreadySubmitted) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-100">
        <div className="bg-white p-8 rounded shadow text-center">
          <h1 className="text-2xl font-bold mb-4">
            Submitted
          </h1>
          <p className="text-gray-600">
            You have already submitted this form.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-200 p-8">

      <div className="max-w-5xl mx-auto">

        <h1 className="text-2xl font-bold mb-6">
          {form.title}
        </h1>

        {/* Questions */}
        {form.questions.map((q: any, index: number) => (
          <div
            key={q.questionId}
            className={`mb-6 p-5 rounded ${
              index % 2 === 0 ? "bg-gray-100" : "bg-gray-200"
            }`}
          >
            {/* Question */}
            <p className="font-semibold mb-1">
              {index + 1}. {q.text}{" "}
              <span className="text-red-500">*</span>
            </p>

            {/* Subtext */}
            <p className="text-sm text-gray-600 mb-3">
              How consistently do you observe this behavior?
            </p>

            {/* Label row */}
            <div className="grid grid-cols-5 text-center text-xs text-gray-500 mb-2">
              {q.options.map((opt: any) => (
                <span key={opt.optionId}>
                  {opt.label}
                  {opt.value !== null && ` (${opt.value}%)`}
                </span>
              ))}
            </div>

            {/*  Radio row */}
            <div className="grid grid-cols-5 text-center">
              {q.options.map((opt: any) => (
                <input
                  key={opt.optionId}
                  type="radio"
                  name={q.questionId}
                  className="mx-auto scale-110"
                  checked={answers[q.questionId] === opt.optionId}
                  onChange={() =>
                    handleSelect(q.questionId, opt.optionId)
                  }
                />
              ))}
            </div>
          </div>
        ))}

        {/* Submit */}
        <button
          onClick={handleSubmit}
          className="mt-6 w-full bg-blue-600 text-white py-3 rounded hover:bg-blue-700"
        >
          Submit
        </button>

        {/* Message */}
        {message && (
          <p className="text-center mt-4 text-red-500">
            {message}
          </p>
        )}
      </div>
    </div>
  );
}