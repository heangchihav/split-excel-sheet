import { NextResponse } from "next/server";

export async function GET() {
  const backendUrl = process.env.BACKEND_URL || "http://127.0.0.1:8080";
  try {
    const res = await fetch(`${backendUrl}/api/health`, {
      cache: "no-store",
    });
    const data = await res.json();
    return NextResponse.json(data, { status: res.status });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json(
      { error: "Failed to connect to backend", details: message, backendUrl },
      { status: 502 }
    );
  }
}
