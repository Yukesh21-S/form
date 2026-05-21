"use client";

import { useRouter } from "next/navigation";

export default function DashboardPage() {
  const router = useRouter();

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      
      {/* Header */}
      <h1 className="text-3xl font-bold mb-8 text-gray-800">
        Admin Dashboard
      </h1>

      {/* Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        
        {/* Create Form */}
        <div
          onClick={() => router.push("/dashboard/create-form")}
          className="bg-white p-6 rounded-xl shadow hover:shadow-lg cursor-pointer transition"
        >
          <h2 className="text-xl font-semibold mb-2">Create Form</h2>
          <p className="text-gray-500">
            Create a new feedback form with multiple questions.
          </p>
        </div>

        {/* View Forms */}
        <div
          onClick={() => router.push("/dashboard/forms")}
          className="bg-white p-6 rounded-xl shadow hover:shadow-lg cursor-pointer transition"
        >
          <h2 className="text-xl font-semibold mb-2">View Forms</h2>
          <p className="text-gray-500">
            See all created forms and track responses.
          </p>
        </div>

        {/* Send Invite */}
        <div
          onClick={() => router.push("/dashboard/send-invite")}
          className="bg-white p-6 rounded-xl shadow hover:shadow-lg cursor-pointer transition"
        >
          <h2 className="text-xl font-semibold mb-2">Send Invite</h2>
          <p className="text-gray-500">
            Send form links to users via email or Excel upload.
          </p>
        </div>

      </div>

      
    </div>
  );
}
