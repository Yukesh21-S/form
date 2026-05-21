"use client";

import { useState } from "react";
import { useRouter } from "next/navigation"
export default function CreateFormPage() {

  const router = useRouter();
  const [title, setTitle] = useState("");
  const [questions, setQuestions] = useState<string[]>([""]);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  // ✅ Add new question
  const addQuestion = () => {
    setQuestions([...questions, ""]);
  };

  // ✅ Remove question
  const removeQuestion = (index: number) => {
    const updated = questions.filter((_, i) => i !== index);
    setQuestions(updated.length ? updated : [""]);
  };

  // ✅ Update question text
  const updateQuestion = (index: number, value: string) => {
    const updated = [...questions];
    updated[index] = value;
    setQuestions(updated);
  };

  // ✅ Submit form
  const handleSubmit = async () => {
    if (!title.trim()) {
      setMessage("Form title is required");
      return;
    }

    const filteredQuestions = questions.filter((q) => q.trim() !== "");
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
            text: q,
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
      setQuestions([""]);
    } catch (error) {
      console.error(error);
      setMessage("❌ Error creating form");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <h1 className="text-2xl font-bold mb-6">Create Form</h1>

      <div className="bg-white p-6 rounded-xl shadow max-w-2xl mx-auto">

        {/* ✅ Form Title */}
        <label className="block mb-2 font-medium">Form Title</label>
        <input
          className="w-full border p-2 rounded mb-6"
          value={title}
          onChange={(e) => setTitle(e.target.value)}
          placeholder="Enter form title"
        />

        {/* ✅ Questions */}
        <h2 className="text-lg font-semibold mb-3">Questions</h2>

        {questions.map((q, index) => (
          <div key={index} className="mb-4 flex gap-2 items-center">
            <input
              className="flex-1 border p-2 rounded"
              value={q}
              onChange={(e) => updateQuestion(index, e.target.value)}
              placeholder={`Question ${index + 1}`}
            />

            {questions.length > 1 && (
              <button
                onClick={() => removeQuestion(index)}
                className="text-red-500 text-sm"
              >
                Remove
              </button>
            )}
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
  className="mt-6 w-full bg-gray-300 py-2 rounded hover:bg-gray-400"
>
  Back to Dashboard
</button>

    </div>
    
  );
}