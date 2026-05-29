"use client";

import { useEffect, useState } from "react";

import { useParams } from "next/navigation";

type Participant = {
  id: string;
  fullName: string;
  email: string;
};

export default function ParticipantsPage() {

  const params = useParams();

  const formId =
    params.formId as string;

  const [participants, setParticipants] =
    useState<Participant[]>([]);

  const [loading, setLoading] =
    useState(true);

  //////////////////////////////////////////////////////
  // FETCH PARTICIPANTS
  //////////////////////////////////////////////////////

  useEffect(() => {

    const fetchParticipants =
      async () => {

        try {

          const res =
            await fetch(
              `/api/admin/form-participants?formId=${formId}`
            );

          const data =
            await res.json();

          setParticipants(data);

        } catch (error) {

          console.error(error);

        } finally {

          setLoading(false);
        }
      };

    if (formId) {

      fetchParticipants();
    }

  }, [formId]);

  //////////////////////////////////////////////////////
  // LOADING
  //////////////////////////////////////////////////////

  if (loading) {

    return (
      <div className="p-6">
        Loading...
      </div>
    );
  }

  //////////////////////////////////////////////////////
  // UI
  //////////////////////////////////////////////////////

  return (

    <div className="min-h-screen bg-gray-100 p-6">

      <h1 className="text-2xl font-bold mb-6">
        Participants
      </h1>

      <div className="grid gap-4">

        {participants.map((participant) => (

          <div
            key={participant.id}

            className="bg-white p-5 rounded-xl shadow"
          >

            <h2 className="text-lg font-semibold">
              {participant.fullName}
            </h2>

            <p className="text-gray-700 text-sm font-medium">
              {participant.email}
            </p>

            <div className="mt-4 flex gap-2">

              {/* GENERATE PPT */}
              <button
                className="px-4 py-2 bg-purple-600 text-white rounded text-sm"

                onClick={() => {

                  window.open(
                    `/api/admin/generate-ppt?formId=${formId}&participantId=${participant.id}`,
                    "_blank"
                  );
                }}
              >
                Generate PPT
              </button>

              {/* GENERATE PDF */}
              <button
                className="px-4 py-2 bg-red-600 text-white rounded text-sm"
                onClick={() => {
                  window.open(
                    `/api/admin/generate-pdf?formId=${formId}&participantId=${participant.id}`,
                    "_blank"
                  );
                }}
              >
                Generate PDF
              </button>

            </div>

          </div>
        ))}
      </div>
    </div>
  );
}