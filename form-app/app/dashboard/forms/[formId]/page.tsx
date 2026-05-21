"use client";

import { useParams } from "next/navigation";
import { useEffect, useState } from "react";

export default function FormDetailsPage() {
  const params = useParams();
  const formId = params.formId as string;

  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // ✅ Fetch form result data
  useEffect(() => {
    const fetchData = async () => {
      try {
        const res = await fetch(
          `/api/admin/results?formId=${formId}`
        );

        if (!res.ok) {
          throw new Error("Failed to fetch form results");
        }

        const json = await res.json();
        setData(json);
      } catch (err) {
        console.error(err);
        setError("Error loading data");
      } finally {
        setLoading(false);
      }
    };

    if (formId) fetchData();
  }, [formId]);

  // ✅ Loading
  if (loading) {
    return <div className="p-6">Loading...</div>;
  }

  // ✅ Error
  if (error) {
    return <div className="p-6 text-red-500">{error}</div>;
  }

  // ✅ No Data
  if (!data) {
    return <div className="p-6">No data found</div>;
  }

  return (
    <div className="min-h-screen bg-gray-100 p-6">

      {/* Header */}
      <h1 className="text-2xl font-bold mb-6">{data.title}</h1>

      {/* Stats */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <div className="bg-white p-4 rounded shadow text-center">
          <p>Total Invited</p>
          <h2>{data.totalInvited}</h2>
        </div>

        <div className="bg-white p-4 rounded shadow text-center">
          <p>Responses</p>
          <h2>{data.totalResponses}</h2>
        </div>

        <div className="bg-white p-4 rounded shadow text-center">
          <p>Completion Rate</p>
          <h2>{data.completionRate}%</h2>
        </div>
      </div>

      {/* Questions */}
      {data.questions.map((q: any, index: number) => (
        <div key={q.questionId} className="bg-white p-4 rounded mb-4 shadow">
          <h2>
            {index + 1}. {q.questionText}
          </h2>

          <p>Avg Score: {q.averageScore ?? "N/A"}</p>

          {q.distribution.map((opt: any) => (
            <div key={opt.optionId} className="flex justify-between">
              <span>{opt.label}</span>
              <span>{opt.count}</span>
            </div>
          ))}
        </div>
      ))}

    </div>
  );
}