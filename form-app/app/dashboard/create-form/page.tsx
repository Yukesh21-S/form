"use client";

import { useState } from "react";
import { useRouter } from "next/navigation";

//////////////////////////////////////////////////////
// TYPES
//////////////////////////////////////////////////////

interface QuestionItem {
  text: string;
}

//////////////////////////////////////////////////////
// PAGE
//////////////////////////////////////////////////////

export default function CreateFormPage() {
  const router = useRouter();
  const [title, setTitle] = useState("");
  const [questions, setQuestions] = useState<QuestionItem[]>([
    { text: "" },
  ]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  // ✅ Add new question
  const addQuestion = () => {
    setQuestions([...questions, { text: "" }]);
  };

  // ✅ Remove question
  const removeQuestion = (index: number) => {
    const updated = questions.filter((_, i) => i !== index);
    setQuestions(updated.length ? updated : [{ text: "" }]);
  };

  // ✅ Update question text
  const updateQuestionText = (index: number, value: string) => {
    const updated = [...questions];
    updated[index] = { ...updated[index], text: value };
    setQuestions(updated);
  };


  // ✅ Submit form
  const handleSubmit = async () => {
    if (!title.trim()) {
      setMessage("Form title is required");
      return;
    }

    const filteredQuestions = questions.filter((q) => q.text.trim() !== "");
    if (filteredQuestions.length === 0) {
      setMessage("At least one question is required");
      return;
    }

    setLoading(true);
    setMessage("");

    try {
      const res = await fetch("/api/admin/forms", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          title,
          questions: filteredQuestions.map((q) => ({
            text: q.text,
          })),
        }),
      });

      if (!res.ok) {
        throw new Error("Failed to create form");
      }

      const data = await res.json();

      setMessage("✅ Form created successfully!");
      console.log("Created form:", data);

      // Reset form
      setTitle("");
      setQuestions([{ text: "" }]);
    } catch (error) {
      console.error(error);
      setMessage("❌ Error creating form");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <h1 className="text-3xl font-bold mb-6 text-black">Create Form</h1>

      <div className="bg-white p-8 rounded-2xl shadow-md max-w-2xl mx-auto border border-gray-200">

        {/* ✅ Form Title */}
        <label className="block mb-2 font-bold text-slate-900">Form Title</label>
        <input
          className="w-full border-2 border-gray-100 p-3 rounded-xl mb-8 focus:border-blue-500 outline-none transition-all text-black font-semibold"
          value={title}
          onChange={(e) => setTitle(e.target.value)}
          placeholder="Enter form title"
        />

        {/* ✅ Questions */}
        <h2 className="text-xl font-bold mb-4 text-slate-900 border-b pb-2">Questions</h2>

        {questions.map((q, index) => (
          <div key={index} className="mb-6 border-2 border-gray-50 rounded-2xl p-5 bg-gray-50/50 shadow-sm">

            {/* Question text row */}
            <div className="flex gap-3 items-center mb-4">
              <span className="font-bold text-blue-600 text-lg">#{index + 1}</span>
              <input
                className="flex-1 border-2 border-white p-3 rounded-xl bg-white shadow-sm focus:border-blue-500 outline-none transition-all text-black font-medium"
                value={q.text}
                onChange={(e) => updateQuestionText(index, e.target.value)}
                placeholder="Enter behavioral question text..."
              />

              {questions.length > 1 && (
                <button
                  onClick={() => removeQuestion(index)}
                  className="text-red-500 hover:text-red-700 font-bold text-sm bg-red-50 px-3 py-2 rounded-lg transition-colors"
                >
                  Remove
                </button>
              )}
            </div>


          </div>
        ))}

        {/* ✅ Add Question Button */}
        <button
          onClick={addQuestion}
          className="mb-6 px-4 py-2 bg-gray-200 hover:bg-gray-300 rounded"
        >
          + Add Question
        </button>

        {/* ✅ Submit */}
        <button
          onClick={handleSubmit}
          disabled={loading}
          className="w-full bg-blue-500 text-white py-2 rounded hover:bg-blue-600 disabled:bg-gray-400"
        >
          {loading ? "Creating..." : "Create Form"}
        </button>

        {/* ✅ Message */}
        {message && (
          <p className="mt-4 text-center text-sm">{message}</p>
        )}

      </div>

      <button
        onClick={() => router.push("/dashboard")}
        className="mt-6 w-full max-w-2xl mx-auto block bg-gray-300 py-2 rounded hover:bg-gray-400"
      >
        Back to Dashboard
      </button>

    </div>
  );
}