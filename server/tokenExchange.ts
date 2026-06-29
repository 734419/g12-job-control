/**
 * Server-side Microsoft OAuth token exchange endpoint.
 *
 * Mobile apps (Expo Go / native) cannot call the Microsoft token endpoint directly
 * due to CORS restrictions (AADSTS9002326). This endpoint proxies the token exchange
 * server-side and returns the access token to the mobile app.
 *
 * POST /api/auth/token-exchange
 * Body: { code, codeVerifier, redirectUri }
 * Returns: { accessToken, refreshToken, expiresAt }
 */

import type { Express } from "express";

const CLIENT_ID = process.env.EXPO_PUBLIC_AZURE_CLIENT_ID ?? "";
const TENANT_ID = process.env.EXPO_PUBLIC_AZURE_TENANT_ID ?? "common";
const TOKEN_ENDPOINT = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

const SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "Sites.ReadWrite.All",
  "User.Read",
].join(" ");

export function registerTokenExchange(app: Express) {
  app.post("/api/auth/token-exchange", async (req, res) => {
    try {
      const { code, codeVerifier, redirectUri } = req.body as {
        code: string;
        codeVerifier: string;
        redirectUri: string;
      };

      if (!code || !codeVerifier || !redirectUri) {
        res.status(400).json({ error: "Missing required fields: code, codeVerifier, redirectUri" });
        return;
      }

      const params = new URLSearchParams({
        client_id: CLIENT_ID,
        code,
        redirect_uri: redirectUri,
        grant_type: "authorization_code",
        code_verifier: codeVerifier,
        scope: SCOPES,
      });

      const response = await fetch(TOKEN_ENDPOINT, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params.toString(),
      });

      const data = await response.json() as any;

      if (!response.ok) {
        console.error("[token-exchange] Microsoft error:", data);
        res.status(400).json({ error: data.error_description ?? data.error ?? "Token exchange failed" });
        return;
      }

      res.json({
        accessToken: data.access_token,
        refreshToken: data.refresh_token ?? null,
        expiresAt: Date.now() + (data.expires_in ?? 3600) * 1000,
      });
    } catch (err: any) {
      console.error("[token-exchange] Unexpected error:", err);
      res.status(500).json({ error: err.message ?? "Internal server error" });
    }
  });
}
