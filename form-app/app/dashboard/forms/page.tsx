"use client";

import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";

//////////////////////////////////////////////////////
// TYPES
//////////////////////////////////////////////////////

type Form = {
  id: string;
  title: string;
  createdAt: string;
};

export default function FormsPage() {

  //////////////////////////////////////////////////////
  // STATES
  //////////////////////////////////////////////////////

  const [forms, setForms] =
    useState<Form[]>([]);

  const [loading, setLoading] =
    useState(true);

  const [error, setError] =
    useState("");

  const router = useRouter();

  //////////////////////////////////////////////////////
  // FETCH FORMS
  //////////////////////////////////////////////////////

  useEffect(() => {

    const fetchForms = async () => {

      try {

        const res =
          await fetch("/api/admin/forms");

        if (!res.ok) {

          throw new Error(
            "Failed to fetch forms"
          );
        }

        const data =
          await res.json();

        setForms(data.forms || data);

      } catch (err) {

        console.error(err);

        setError(
          "Failed to load forms"
        );

      } finally {

        setLoading(false);
      }
    };

    fetchForms();

  }, []);

  //////////////////////////////////////////////////////
  // LOADING
  //////////////////////////////////////////////////////

  if (loading) {

    return (
      <div className="p-6">
        Loading forms...
      </div>
    );
  }

  //////////////////////////////////////////////////////
  // ERROR
  //////////////////////////////////////////////////////

  if (error) {

    return (
      <div className="p-6 text-red-500">
        {error}
      </div>
    );
  }

  //////////////////////////////////////////////////////
  // UI
  //////////////////////////////////////////////////////

  return (

    <div className="min-h-screen bg-gray-100 p-6">

      <h1 className="text-2xl font-bold mb-6">
        All Forms
      </h1>

      {forms.length === 0 ? (

        <p>No forms found</p>

      ) : (

        <div className="grid gap-4">

          {forms.map((form) => (

            <div
              key={form.id}

              onClick={() =>
                router.push(
                  `/dashboard/forms/${form.id}`
                )
              }

              className="bg-white p-5 rounded-xl shadow hover:shadow-lg cursor-pointer transition"
            >

              {/* //////////////////////////////////////////////////////
              // TITLE
              ////////////////////////////////////////////////////// */}

              <h2 className="text-lg font-semibold">
                {form.title}
              </h2>

              {/* //////////////////////////////////////////////////////
              // FORM ID
              ////////////////////////////////////////////////////// */}

              <p className="text-sm text-gray-500 mt-1">
                Form ID: {form.id}
              </p>

              {/* //////////////////////////////////////////////////////
              // CREATED DATE
              ////////////////////////////////////////////////////// */}

              <p className="text-sm text-gray-400">
                Created{" "}
                {new Date(
                  form.createdAt
                ).toLocaleString()}
              </p>

              {/* //////////////////////////////////////////////////////
              // BUTTONS
              ////////////////////////////////////////////////////// */}

              <div className="mt-4 flex gap-2 flex-wrap">

                {/* VIEW DETAILS */}
                <button
                  className="px-4 py-2 bg-blue-500 text-white rounded text-sm hover:bg-blue-600"

                  onClick={(e) => {

                    e.stopPropagation();

                    router.push(
                      `/dashboard/forms/${form.id}`
                    );
                  }}
                >
                  View Details
                </button>

                {/* OVERALL ANALYTICS */}
                <button
                  className="px-4 py-2 bg-green-500 text-white rounded text-sm hover:bg-green-600"

                  onClick={(e) => {

                    e.stopPropagation();

                    router.push(
                      `/dashboard/report/${form.id}`
                    );
                  }}
                >
                  Overall Analytics
                </button>
{/* 
                //////////////////////////////////////////////////////
                // DROPDOWN
                ////////////////////////////////////////////////////// */}

                <div className="relative">

                  <select
                    className="px-4 py-2 border rounded text-sm bg-white"

                    onClick={(e) =>
                      e.stopPropagation()
                    }

                    onChange={(e) => {

                      const value =
                        e.target.value;

                      if (!value) return;

                      router.push(value);
                    }}
                  >

                    <option value="">
                      More Actions
                    </option>

                    {/* PARTICIPANTS */}
                    <option
                      value={`/dashboard/forms/${form.id}/participants`}
                    >
                      Participants
                    </option>

                    {/* FORM DETAILS */}
                    <option
                      value={`/dashboard/forms/${form.id}`}
                    >
                      Form Details
                    </option>

                    {/* OVERALL ANALYTICS */}
                    <option
                      value={`/dashboard/report/${form.id}`}
                    >
                      Overall Analytics
                    </option>

                  </select>

                </div>

              </div>

            </div>
          ))}
        </div>
      )}
    </div>
  );
}