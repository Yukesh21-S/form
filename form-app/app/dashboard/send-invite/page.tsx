"use client";

import { useState } from "react";

export default function SendInvitePage() {
  const [formId, setFormId] = useState("");
  const [emails, setEmails] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  //  Send invite (single / multiple emails)
  const handleSendInvite = async () => {
    if (!formId || !emails.trim()) {
      setMessage("Please enter Form ID and email(s)");
      return;
    }

    const emailList = emails.split(",").map((e) => e.trim());
    setLoading(true);
    setMessage("");

    try {
      for (const email of emailList) {
        const res = await fetch("/api/admin/send-invite", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            email, //  single email per request
            formId,
          }),
        });

        if (!res.ok) {
          throw new Error(`Failed for ${email}`);
        }
      }

      setMessage(" Invites sent successfully");
      setEmails("");
    } catch (error) {
      console.error(error);
      setMessage(" Error sending invites");
    } finally {
      setLoading(false);
    }
  };

  //  Bulk Excel Upload
  const handleBulkUpload = async () => {
    if (!formId || !file) {
      setMessage("Please provide Form ID and Excel file");
      return;
    }

    const formData = new FormData();
    formData.append("formId", formId);
    formData.append("file", file);

    setLoading(true);
    setMessage("");

    try {
      const res = await fetch("/api/admin/send-invite-bulk", {
        method: "POST",
        body: formData, //  multipart/form-data
      });

      if (!res.ok) {
        throw new Error("Bulk upload failed");
      }

      const data = await res.json();
      console.log(data);

      setMessage(" Bulk invites sent successfully");
      setFile(null);
    } catch (error) {
      console.error(error);
      setMessage(" Error sending bulk invites");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">

      <h1 className="text-2xl font-bold mb-6">Send Invite</h1>

      <div className="grid gap-6 max-w-2xl mx-auto">

        {/* FORM ID INPUT */}
        <div className="bg-white p-5 rounded-xl shadow">
          <label className="block mb-2 font-medium">Form ID</label>
          <input
            className="w-full border p-2 rounded"
            value={formId}
            onChange={(e) => setFormId(e.target.value)}
            placeholder="Enter Form ID"
          />
        </div>

        {/*  EMAIL INPUT SECTION */}
        <div className="bg-white p-5 rounded-xl shadow">
          <h2 className="text-lg font-semibold mb-3">
            Send via Email
          </h2>

          <textarea
            className="w-full border p-2 rounded mb-4"
            rows={4}
            value={emails}
            onChange={(e) => setEmails(e.target.value)}
            placeholder="Enter emails separated by comma"
          />

          <button
            onClick={handleSendInvite}
            disabled={loading}
            className="w-full bg-blue-500 text-white py-2 rounded hover:bg-blue-600 disabled:bg-gray-400"
          >
            {loading ? "Sending..." : "Send Invite"}
          </button>
        </div>

        {/* BULK UPLOAD SECTION */}
        <div className="bg-white p-5 rounded-xl shadow">
          <h2 className="text-lg font-semibold mb-3">
            Bulk Upload (Excel)
          </h2>

          <input
            type="file"
            accept=".xlsx"
            className="mb-4"
            onChange={(e) => setFile(e.target.files?.[0] || null)}
          />

          <button
            onClick={handleBulkUpload}
            disabled={loading}
            className="w-full bg-green-500 text-white py-2 rounded hover:bg-green-600 disabled:bg-gray-400"
          >
            {loading ? "Uploading..." : "Upload & Send"}
          </button>
        </div>

        {/*  MESSAGE */}
        {message && (
          <div className="text-center text-sm">
            {message}
          </div>
        )}

      </div>
    </div>
  );
}
``