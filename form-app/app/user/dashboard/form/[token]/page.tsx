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
  const [relationshipType, setRelationshipType] = useState("PEER");

  // Determine relationship types based on whether it's a self-assessment
  const RELATIONSHIP_TYPES = form?.isSelf 
    ? ["SELF", "MANAGER", "PEER", "DIRECT_REPORT", "OTHER"]
    : ["MANAGER", "PEER", "DIRECT_REPORT", "OTHER"];

  //  Fetch form
  useEffect(() => {
    const fetchForm = async () => {
      try {
        const res = await fetch(`/api/form/${token}`);
        const data = await res.json();

        if (data.alreadySubmitted) {
          setAlreadySubmitted(true);
        }

        if (data.isSelf) {
          setRelationshipType("SELF");
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
      relationshipType,
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
          <p className="text-gray-800">
            You have already submitted this form.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-200 p-8">

      <div className="max-w-5xl mx-auto">

        <h1 className="text-2xl font-bold mb-2 text-black">
          {form.title}
        </h1>

        <p className="text-gray-800 mb-6 font-medium">
          Providing feedback for: <span className="font-bold text-blue-700">{form.participantName}</span>
        </p>

        {/* Relationship Selector */}
        <div className="bg-white p-6 rounded-xl shadow-sm mb-8 text-black border-l-4 border-blue-500">
          <h2 className="font-semibold mb-3">Your Relationship</h2>
          
          {form.isSelf ? (
            <div className="flex items-center gap-2 text-blue-800 font-bold bg-blue-50 p-3 rounded border border-blue-100">
              <span>👤</span>
              <span>Relationship: SELF (Self-Assessment)</span>
            </div>
          ) : (
            <div>
              <p className="text-sm text-gray-800 mb-3 font-medium">
                Please select your relationship to {form.participantName}
              </p>
              <select
                value={relationshipType}
                onChange={(e) => setRelationshipType(e.target.value)}
                className="w-full md:w-1/2 border-2 border-gray-200 p-2.5 rounded-lg focus:border-blue-500 outline-none bg-gray-50 transition-all font-bold text-gray-900 cursor-pointer"
              >
                {RELATIONSHIP_TYPES.map((type) => (
                  <option key={type} value={type}>
                    {type}
                  </option>
                ))}
              </select>
            </div>
          )}
        </div>

        {/* Questions */}
        {form.questions.map((q: any, index: number) => (
          <div
            key={q.questionId}
            className={`mb-6 p-5 rounded ${
              index % 2 === 0 ? "bg-gray-100" : "bg-gray-200"
            }`}
          >
            {/* Question */}
            <p className="font-bold text-black mb-1">
              {index + 1}. {q.text}{" "}
              <span className="text-red-500">*</span>
            </p>

            {/* Subtext */}
            <p className="text-sm text-black mb-3 italic font-medium">
              How consistently do you observe this behavior?
            </p>

            {/* Label row */}
            <div className="grid grid-cols-5 text-center text-xs text-gray-900 mb-2 font-bold">
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