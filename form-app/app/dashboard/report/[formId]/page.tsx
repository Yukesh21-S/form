"use client";

import { useParams } from "next/navigation";
import { useEffect, useState } from "react";

export default function ReportPage() {
  const { formId } = useParams();

  const [data, setData] = useState<any>(null);

  useEffect(() => {
    const fetchData = async () => {
      const res = await fetch(
        `/api/admin/report?formId=${formId}`
      );
      const json = await res.json();
      setData(json);
    };

    if (formId) fetchData();
  }, [formId]);

  if (!data) return <div>Loading...</div>;

  const getColor = (score: number) => {
    if (score >= 85) return "bg-green-500";
    if (score >= 70) return "bg-orange-400";
    return "bg-red-500";
  };

  return (
    <div className="p-6 bg-gray-100 min-h-screen">

      {/* ✅ Title */}
      <h1 className="text-2xl font-bold mb-2">
        Overall Results
      </h1>
      <p className="text-gray-600 mb-6">
        Survey: {data.title}
      </p>

      {/* ✅ Summary */}
      <div className="flex gap-6 mb-8">
        <p>Invited: {data.totalInvited}</p>
        <p>Responses: {data.totalResponses}</p>
        <p>Completion: {data.completionRate}%</p>
      </div>

      {/* ✅ Overall Bars */}
      <div className="space-y-4 mb-10">
        {data.questions.map((q: any) => (
          <div key={q.questionId}>
            <p className="text-sm mb-1">
              {q.questionText}
            </p>

            <div className="flex items-center gap-3">
              <div className="w-full bg-gray-200 h-4 rounded">
                <div
                  className={`${getColor(
                    q.averageScore
                  )} h-4 rounded`}
                  style={{
                    width: `${q.averageScore}%`,
                  }}
                />
              </div>

              <span>{q.averageScore}%</span>
            </div>
          </div>
        ))}
      </div>

      {/* ✅ Distribution Table */}
      <div className="bg-white p-6 rounded shadow">
        <h2 className="font-bold mb-4">
          Results Distribution
        </h2>

        <table className="w-full text-sm border">
          <thead>
            <tr className="bg-gray-200">
              <th className="p-2 text-left">Question</th>
              <th>Rarely</th>
              <th>Sometimes</th>
              <th>Often</th>
              <th>Always</th>
            </tr>
          </thead>

          <tbody>
            {data.questions.map((q: any) => (
              <tr key={q.questionId}>
                <td className="p-2">
                  {q.questionText}
                </td>
                <td>{q.distribution.Rarely || 0}</td>
                <td>{q.distribution.Sometimes || 0}</td>
                <td>{q.distribution.Often || 0}</td>
                <td>{q.distribution.Always || 0}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

    </div>
  );
}