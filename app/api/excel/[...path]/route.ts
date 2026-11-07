import { NextRequest } from "next/server";

export async function POST(
  request: NextRequest,
  context: { params: Promise<{ path: string[] }> }
) {
  const { path } = await context.params;
  const backendUrl = process.env.BACKEND_URL || "http://127.0.0.1:8080";
  const subPath = path.join("/");
  const targetUrl = `${backendUrl}/api/excel/${subPath}`;

  try {
    const headers = new Headers();
    const contentType = request.headers.get("content-type");
    if (contentType) {
      headers.set("content-type", contentType);
    }

    const res = await fetch(targetUrl, {
      method: "POST",
      headers,
      body: request.body,
      // @ts-expect-error duplex is required for fetch with ReadableStream body in Node
      duplex: "half",
    });

    const responseHeaders = new Headers();
    const forwardHeaders = [
      "content-type",
      "content-disposition",
      "content-length",
    ];
    for (const h of forwardHeaders) {
      const val = res.headers.get(h);
      if (val) {
        responseHeaders.set(h, val);
      }
    }

    return new Response(res.body, {
      status: res.status,
      statusText: res.statusText,
      headers: responseHeaders,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return new Response(
      JSON.stringify({ error: "Backend proxy error", details: message, targetUrl }),
      {
        status: 502,
        headers: { "content-type": "application/json" },
      }
    );
  }
}
