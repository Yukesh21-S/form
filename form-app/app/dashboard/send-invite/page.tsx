"use client";

import { useEffect, useState } from "react";

type Form = {
  id: string;
  title: string;
};

type Participant = {
  id: string;
  fullName: string;
  email: string;
};

export default function SendInvitePage() {

  //////////////////////////////////////////////////////
  // STATES
  //////////////////////////////////////////////////////

  const [forms, setForms] = useState<Form[]>([]);

  const [participants, setParticipants] =
    useState<Participant[]>([]);

  const [selectedForm, setSelectedForm] =
    useState<Form | null>(null);

  const [selectedParticipant, setSelectedParticipant] =
    useState<Participant | null>(null);

  const [emails, setEmails] = useState("");

  const [file, setFile] =
    useState<File | null>(null);

  const [emailLoading, setEmailLoading] =
    useState(false);

  const [uploadLoading, setUploadLoading] =
    useState(false);

  const [message, setMessage] = useState("");

  //////////////////////////////////////////////////////
  // FETCH FORMS + PARTICIPANTS
  //////////////////////////////////////////////////////

  useEffect(() => {

    const fetchData = async () => {

      try {

        //////////////////////////////////////////////////////
        // FETCH FORMS
        //////////////////////////////////////////////////////

        const formsRes =
          await fetch("/api/admin/forms");

        const formsData =
          await formsRes.json();

        setForms(formsData.forms || formsData);

        //////////////////////////////////////////////////////
        // FETCH PARTICIPANTS
        //////////////////////////////////////////////////////

        const participantRes =
          await fetch("/api/admin/participants");

        const participantData =
          await participantRes.json();

        setParticipants(participantData);

      } catch (error) {

        console.error(error);
      }
    };

    fetchData();

  }, []);

  //////////////////////////////////////////////////////
  // SEND EMAIL INVITE
  //////////////////////////////////////////////////////

  const handleSendInvite = async () => {

    if (
      !selectedForm ||
      !selectedParticipant ||
      !emails.trim()
    ) {

      setMessage(
        "Please select form, participant and email"
      );

      return;
    }

    const emailList = emails
      .split(",")
      .map((e) => e.trim())
      .filter(Boolean);

    setEmailLoading(true);

    setMessage("");

    try {

      for (const email of emailList) {

        const res = await fetch(
          "/api/admin/send-invite",
          {
            method: "POST",

            headers: {
              "Content-Type": "application/json",
            },

            body: JSON.stringify({
              email,
              formId: selectedForm.id,
              participantId:
                selectedParticipant.id,
            }),
          }
        );

        if (!res.ok) {
          throw new Error(
            `Failed for ${email}`
          );
        }
      }

      setMessage(
        "Invites sent successfully"
      );

      setEmails("");

    } catch (error) {

      console.error(error);

      setMessage(
        "Error sending invites"
      );

    } finally {

      setEmailLoading(false);
    }
  };

  //////////////////////////////////////////////////////
  // BULK UPLOAD
  //////////////////////////////////////////////////////

  const handleBulkUpload = async () => {

    if (
      !selectedForm ||
      !selectedParticipant ||
      !file
    ) {

      setMessage(
        "Please select form, participant and upload file"
      );

      return;
    }

    if (!file.name.endsWith(".xlsx")) {

      setMessage(
        "Only Excel (.xlsx) files allowed"
      );

      return;
    }

    const formData = new FormData();

    formData.append(
      "formId",
      selectedForm.id
    );

    formData.append(
      "participantId",
      selectedParticipant.id
    );

    formData.append("file", file);

    setUploadLoading(true);

    setMessage("");

    try {

      const res = await fetch(
        "/api/admin/send-invite-bulk",
        {
          method: "POST",
          body: formData,
        }
      );

      if (!res.ok) {
        throw new Error("Bulk failed");
      }

      setMessage(
        "Bulk invites sent successfully"
      );

      setFile(null);

    } catch (error) {

      console.error(error);

      setMessage(
        "Error in bulk upload"
      );

    } finally {

      setUploadLoading(false);
    }
  };

  //////////////////////////////////////////////////////
  // UI
  //////////////////////////////////////////////////////

  return (

    <div className="min-h-screen bg-gray-100 p-6">

      <h1 className="text-2xl font-bold mb-6">
        Send Invite
      </h1>

      <div className="grid gap-6 max-w-2xl mx-auto">

        {/* //////////////////////////////////////////////////////
        // FORM SELECT
        ////////////////////////////////////////////////////// */}

        <div className="bg-white p-5 rounded-xl shadow text-black">

          <label className="block mb-2 font-medium">
            Select Form
          </label>

          <select
            className="w-full border p-2 rounded"

            onChange={(e) => {

              const form = forms.find(
                (f) =>
                  f.id === e.target.value
              );

              setSelectedForm(form || null);
            }}
          >

            <option value="">
              -- Select Form --
            </option>

            {forms.map((form) => (

              <option
                key={form.id}
                value={form.id}
              >
                {form.title}
              </option>

            ))}

          </select>
        </div>

        {/* //////////////////////////////////////////////////////
        // PARTICIPANT SELECT
        ////////////////////////////////////////////////////// */}

        <div className="bg-white p-5 rounded-xl shadow text-black">

          <label className="block mb-2 font-medium">
            Select Participant (The person being reviewed)
          </label>

          <select
            className="w-full border p-2 rounded"

            onChange={(e) => {

              const participant =
                participants.find(
                  (p) =>
                    p.id === e.target.value
                );

              setSelectedParticipant(
                participant || null
              );
            }}
          >

            <option value="">
              -- Select Participant --
            </option>

            {participants.map(
              (participant) => (

                <option
                  key={participant.id}
                  value={participant.id}
                >
                  {participant.fullName}
                </option>

              )
            )}

          </select>
        </div>

{/* 
        //////////////////////////////////////////////////////
        // EMAIL SECTION
        ////////////////////////////////////////////////////// */}

        <div className="bg-white p-5 rounded-xl shadow text-black">

          <h2 className="font-semibold mb-3">
            Send via Email
          </h2>

          <textarea
            className="w-full border p-2 rounded mb-4"

            rows={4}

            value={emails}

            onChange={(e) =>
              setEmails(
                e.target.value
              )
            }

            placeholder="Enter emails separated by comma"
          />

          <button
            onClick={handleSendInvite}

            disabled={emailLoading}

            className="w-full bg-blue-500 text-white py-2 rounded hover:bg-blue-600 disabled:bg-gray-400"
          >

            {emailLoading
              ? "Sending..."
              : "Send Invite"}

          </button>
        </div>

        {/* //////////////////////////////////////////////////////
        // BULK UPLOAD
        ////////////////////////////////////////////////////// */}

        <div className="bg-white p-5 rounded-xl shadow text-black">

          <h2 className="font-semibold mb-3">
            Bulk Upload (Excel)
          </h2>

          <input
            type="file"
            accept=".xlsx"
            className="mb-4"
            onChange={(e) =>
              setFile(
                e.target.files?.[0] ||
                  null
              )
            }
          />

          {file && (

            <div className="flex justify-between items-center bg-gray-100 p-2 rounded mb-4">

              <span className="text-sm truncate">
                {file.name}
              </span>

              <button
                onClick={() =>
                  setFile(null)
                }

                className="text-red-500 text-sm"
              >
                Remove
              </button>

            </div>
          )}

          <button
            onClick={handleBulkUpload}

            disabled={
              uploadLoading || !file
            }

            className="w-full bg-green-500 text-white py-2 rounded hover:bg-green-600 disabled:bg-gray-400"
          >

            {uploadLoading
              ? "Uploading..."
              : "Upload & Send"}

          </button>
        </div>

        {/* //////////////////////////////////////////////////////
        // MESSAGE
        ////////////////////////////////////////////////////// */}

        {message && (

          <div className="text-center text-sm">
            {message}
          </div>

        )}

      </div>
    </div>
  );
}