// Thin client for the private autoproduct-engine API. This is the only
// place this project talks to the "backend" - no enrichment prompts,
// pricing formula, or CSV mapping logic live here or anywhere in this repo.

export async function runPipeline({ mode, rows }) {
  const baseUrl = process.env.API_BASE_URL;
  const apiKey = process.env.API_KEY;

  if (!baseUrl) {
    throw new Error("API_BASE_URL is not set. Copy .env.example to .env and fill it in.");
  }
  if (!apiKey) {
    throw new Error("API_KEY is not set. Copy .env.example to .env and fill it in.");
  }

  const response = await fetch(`${baseUrl.replace(/\/$/, "")}/api/run`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
    },
    body: JSON.stringify({ mode, rows }),
  });

  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(`API request failed (${response.status}): ${body.error || response.statusText}`);
  }

  return response.json();
}
