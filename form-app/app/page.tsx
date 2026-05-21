import Link from "next/link";

export default function Home() {
  return (
    <div className="flex min-h-screen flex-col items-center justify-center bg-zinc-50 px-4">
      <main className="w-full max-w-lg rounded-xl border border-zinc-200 bg-white p-10 shadow-sm">
        <h1 className="text-2xl font-semibold text-zinc-900">Feedback App</h1>
        <p className="mt-2 text-sm text-zinc-600">
          Leadership feedback forms with invite links and admin analytics.
        </p>

        <div className="mt-8 flex flex-col gap-3">
          <Link
            href="/admin"
            className="rounded-lg bg-zinc-900 px-4 py-3 text-center text-sm font-medium text-white hover:bg-zinc-800"
          >
            Admin dashboard
          </Link>
          <p className="text-center text-xs text-zinc-500">
            Respondents open their unique link from email, e.g.{" "}
            <code className="rounded bg-zinc-100 px-1">/form/[token]</code>
          </p>
        </div>
      </main>
    </div>
  );
}
