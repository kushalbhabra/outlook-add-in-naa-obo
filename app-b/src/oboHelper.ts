import { ConfidentialClientApplication, OnBehalfOfRequest } from "@azure/msal-node";

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.APP_B_CLIENT_ID!,
    clientSecret: process.env.APP_B_CLIENT_SECRET!,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
});

export async function exchangeForGraphToken(assertion: string): Promise<string> {
  const request: OnBehalfOfRequest = {
    oboAssertion: assertion,
    scopes: [
      "https://graph.microsoft.com/Mail.Read",
      "https://graph.microsoft.com/User.Read",
    ],
  };
  const result = await cca.acquireTokenOnBehalfOf(request);
  if (!result?.accessToken) throw new Error("OBO exchange returned no token");
  return result.accessToken;
}
