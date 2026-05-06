import { createRemoteJWKSet, jwtVerify } from "jose";

const TENANT_ID = process.env.TENANT_ID!;
const APP_B_CLIENT_ID = process.env.APP_B_CLIENT_ID!;

const JWKS = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`)
);

export interface TokenPayload {
  oid: string;
  name: string;
  upn?: string;
  scp: string;
}

export async function validateToken(authHeader: string | undefined): Promise<TokenPayload> {
  if (!authHeader?.startsWith("Bearer ")) {
    throw new Error("Missing or invalid Authorization header");
  }
  const token = authHeader.slice(7);
  const { payload } = await jwtVerify(token, JWKS, {
    audience: [APP_B_CLIENT_ID, `api://${APP_B_CLIENT_ID}`],
    issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
  });
  const scp = payload["scp"] as string | undefined;
  if (!scp?.includes("access_as_user")) {
    throw new Error("Token missing required scope: access_as_user");
  }
  return payload as unknown as TokenPayload;
}
