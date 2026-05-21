"use client";

import { useEffect, useState } from "react";

type Form = {
  id: string;
  title: string;
};

export default function SendInvitePage() {
  const [forms, setForms] = useState<Form[]>([]);
  const [selectedForm, setSelectedForm] = useState<Form | null>(null);

  const [emails, setEmails] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");

  // ✅ Fetch forms
  useEffect(() => {
    const fetchForms = async () => {
      try {
        const res = await fetch("/api/admin/forms");
        const data = await res.json();
        setForms(data.forms || data);
      } catch (error) {
        console.error(error);
      }
    };

    fetchForms();
  }, []);

  // ✅ Send invite (email)
  const handleSendInvite = async () => {
    if (!selectedForm || !emails.trim()) {
      setMessage("Please select a form and enter email(s)");
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
            email,
            formId: selectedForm.id,
          }),
        });

        if (!res.ok) {
          throw new Error(`Failed for ${email}`);
        }
      }

      setMessage("✅ Invites sent successfully");
      setEmails("");
    } catch (error) {
      console.error(error);
      setMessage("❌ Error sending invites");
    } finally {
      setLoading(false);
    }
  };

  // ✅ Bulk upload
  const handleBulkUpload = async () => {
    if (!selectedForm || !file) {
      setMessage("Please select a form and upload file");
      return;
    }

    // ✅ Optional validation
    if (!file.name.endsWith(".xlsx")) {
      setMessage("⚠️ Only Excel (.xlsx) files allowed");
      return;
    }

    const formData = new FormData();
    formData.append("formId", selectedForm.id);
    formData.append("file", file);

    setLoading(true);
    setMessage("");

    try {
      const res = await fetch("/api/admin/send-invite-bulk", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) throw new Error("Bulk failed");

      setMessage("✅ Bulk invites sent successfully");
      setFile(null);
    } catch (error) {
      console.error(error);
      setMessage("❌ Error in bulk upload");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <h1 className="text-2xl font-bold mb-6">Send Invite</h1>

      <div className="grid gap-6 max-w-2xl mx-auto">

        {/* ✅ FORM SELECT */}
        <div className="bg-white p-5 rounded-xl shadow">
          <label className="block mb-2 font-medium">
            Select Form
          </label>

          <select
            className="w-full border p-2 rounded"
            onChange={(e) => {
              const form = forms.find(f => f.id === e.target.value);
              setSelectedForm(form || null);
            }}
          >
            <option value="">-- Select Form --</option>
            {forms.map((form) => (
              <option key={form.id} value={form.id}>
                {form.title}
              </option>
            ))}
          </select>
        </div>

        {/* ✅ EMAIL SECTION */}
        <div className="bg-white p-5 rounded-xl shadow">
          <h2 className="font-semibold mb-3">Send via Email</h2>

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

        {/* ✅ BULK UPLOAD SECTION */}
        <div className="bg-white p-5 rounded-xl shadow">
          <h2 className="font-semibold mb-3">
            Bulk Upload (Excel)
          </h2>

          {/* File Input */}
          <input
            type="file"
            accept=".xlsx"
            className="mb-4"
            onChange={(e) => setFile(e.target.files?.[0] || null)}
          />

          {/* ✅ Selected File Display */}
          {file && (
            <div className="flex justify-between items-center bg-gray-100 p-2 rounded mb-4">
              <span className="text-sm truncate">{file.name}</span>

              <button
                onClick={() => setFile(null)}
                className="text-red-500 text-sm"
              >
                Remove ❌
              </button>
            </div>
          )}

          {/* Upload button */}
          <button
            onClick={handleBulkUpload}
            disabled={loading || !file}
            className="w-full bg-green-500 text-white py-2 rounded hover:bg-green-600 disabled:bg-gray-400"
          >
            {loading ? "Uploading..." : "Upload & Send"}
          </button>
        </div>

        {/* ✅ MESSAGE */}
        {message && (
          <div className="text-center text-sm">
            {message}
          </div>
        )}

      </div>
    </div>
  );
}