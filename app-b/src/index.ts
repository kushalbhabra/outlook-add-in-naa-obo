import "dotenv/config";
import { serve } from "@hono/node-server";
import { Hono } from "hono";
import { cors } from "hono/cors";
import { validateToken } from "./validateToken";
import { exchangeForGraphToken } from "./oboHelper";

const app = new Hono();

app.use("*", cors({ origin: "https://localhost:3000" }));

app.get("/api/mail", async (c) => {
  try {
    const authHeader = c.req.header("Authorization");
    const payload = await validateToken(authHeader);

    const assertion = authHeader!.slice(7);
    const graphToken = await exchangeForGraphToken(assertion);

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/me/messages" +
        "?$select=subject,from,receivedDateTime,isRead" +
        "&$top=10&$orderby=receivedDateTime desc",
      { headers: { Authorization: `Bearer ${graphToken}` } }
    );

    if (!response.ok) {
      const err = await response.text();
      return c.json({ error: `Graph error: ${err}` }, 502);
    }

    const data = (await response.json()) as { value: unknown[] };
    return c.json({
      user: { oid: payload.oid, name: payload.name },
      messages: data.value,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error("[App B] Error:", message);
    return c.json({ error: message }, 401);
  }
});

app.get("/health", (c) => c.json({ status: "ok" }));

const PORT = 4000;
serve({ fetch: app.fetch, port: PORT }, () => {
  console.log(`App B server running on http://localhost:${PORT}`);
});
