"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";

type Form = {
  id: string;
  title: string;
  createdAt: string;
};

export default function FormsPage() {
  const [forms, setForms] = useState<Form[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const router = useRouter();

  // ✅ Fetch all forms
  useEffect(() => {
    const fetchForms = async () => {
      try {
        const res = await fetch("/api/admin/forms");

        if (!res.ok) {
          throw new Error("Failed to fetch forms");
        }

        const data = await res.json();
        setForms(data.forms || data);
      } catch (err) {
        console.error(err);
        setError("Failed to load forms");
      } finally {
        setLoading(false);
      }
    };

    fetchForms();
  }, []);

  // ✅ Loading
  if (loading) {
    return <div className="p-6">Loading forms...</div>;
  }

  // ✅ Error
  if (error) {
    return <div className="p-6 text-red-500">{error}</div>;
  }

  return (
    <div className="min-h-screen bg-gray-100 p-6">

      <h1 className="text-2xl font-bold mb-6">All Forms</h1>

      {forms.length === 0 ? (
        <p>No forms found</p>
      ) : (
        <div className="grid gap-4">
          {forms.map((form) => (
            <div
              key={form.id}
              onClick={() =>
                router.push(`/dashboard/forms/${form.id}`)
              }
              className="bg-white p-5 rounded-xl shadow hover:shadow-lg cursor-pointer transition"
            >
              {/* ✅ Title */}
              <h2 className="text-lg font-semibold">
                {form.title}
              </h2>

              {/* ✅ Form ID */}
              <p className="text-sm text-gray-500 mt-1">
                Form ID: {form.id}
              </p>

              {/* ✅ Created Date */}
              <p className="text-sm text-gray-400">
                Created:{" "}
                {new Date(form.createdAt).toLocaleString()}
              </p>

              {/* ✅ Buttons */}
              <div className="mt-3 flex gap-2">

                {/* View Details */}
                <button
                  className="px-4 py-1 bg-blue-500 text-white rounded text-sm"
                  onClick={(e) => {
                    e.stopPropagation();
                    router.push(
                      `/dashboard/forms/${form.id}`
                    );
                  }}
                >
                  View Details
                </button>

                {/* ✅ View Report (NEW) */}
                <button
                  className="px-4 py-1 bg-green-500 text-white rounded text-sm"
                  onClick={(e) => {
                    e.stopPropagation();
                    router.push(
                      `/dashboard/report/${form.id}`
                    );
                  }}
                >
                  View Report
                </button>

              </div>

            </div>
          ))}
        </div>
      )}
    </div>
  );
}
