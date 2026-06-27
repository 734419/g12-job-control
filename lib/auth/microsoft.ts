/**
 * Microsoft Azure AD / MSAL authentication helper for G12 Job Control.
 *
 * Uses expo-auth-session with the Microsoft identity platform (v2.0 endpoint).
 * The Azure App Registration Client ID and Tenant ID are loaded from environment
 * variables so they can be swapped without code changes.
 *
 * Required env vars (set via webdev_request_secrets):
 *   EXPO_PUBLIC_AZURE_CLIENT_ID   — App registration client ID
 *   EXPO_PUBLIC_AZURE_TENANT_ID   — Azure AD tenant ID (or "common")
 */

import * as AuthSession from "expo-auth-session";
import * as WebBrowser from "expo-web-browser";
import * as SecureStore from "expo-secure-store";

WebBrowser.maybeCompleteAuthSession();

const CLIENT_ID = process.env.EXPO_PUBLIC_AZURE_CLIENT_ID ?? "YOUR_CLIENT_ID";
const TENANT_ID = process.env.EXPO_PUBLIC_AZURE_TENANT_ID ?? "common";

const DISCOVERY: AuthSession.DiscoveryDocument = {
  authorizationEndpoint: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`,
  tokenEndpoint: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
  revocationEndpoint: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/logout`,
};

// Scopes required for SharePoint list read/write via Graph API
export const SCOPES = [
  "openid",
  "profile",
  "email",
  "offline_access",
  "Sites.ReadWrite.All",
  "User.Read",
];

const TOKEN_KEY = "g12_ms_token";
const REFRESH_KEY = "g12_ms_refresh";
const USER_KEY = "g12_ms_user";

export interface MSUser {
  id: string;
  displayName: string;
  mail: string;
  jobTitle?: string;
  userPrincipalName: string;
}

export interface TokenBundle {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number; // epoch ms
}

export function buildAuthRequest(redirectUri: string) {
  return new AuthSession.AuthRequest({
    clientId: CLIENT_ID,
    scopes: SCOPES,
    redirectUri,
    responseType: AuthSession.ResponseType.Code,
    usePKCE: true,
    extraParams: {
      prompt: "select_account",
    },
  });
}

export async function exchangeCodeForToken(
  code: string,
  codeVerifier: string,
  redirectUri: string
): Promise<TokenBundle> {
  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    code_verifier: codeVerifier,
    scope: SCOPES.join(" "),
  });

  const res = await fetch(DISCOVERY.tokenEndpoint!, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params.toString(),
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Token exchange failed: ${err}`);
  }

  const data = await res.json();
  const bundle: TokenBundle = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };

  await SecureStore.setItemAsync(TOKEN_KEY, bundle.accessToken);
  if (bundle.refreshToken) {
    await SecureStore.setItemAsync(REFRESH_KEY, bundle.refreshToken);
  }

  return bundle;
}

export async function refreshAccessToken(): Promise<string | null> {
  const refreshToken = await SecureStore.getItemAsync(REFRESH_KEY);
  if (!refreshToken) return null;

  const params = new URLSearchParams({
    client_id: CLIENT_ID,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: SCOPES.join(" "),
  });

  try {
    const res = await fetch(DISCOVERY.tokenEndpoint!, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });

    if (!res.ok) return null;

    const data = await res.json();
    await SecureStore.setItemAsync(TOKEN_KEY, data.access_token);
    if (data.refresh_token) {
      await SecureStore.setItemAsync(REFRESH_KEY, data.refresh_token);
    }
    return data.access_token;
  } catch {
    return null;
  }
}

export async function getStoredToken(): Promise<string | null> {
  return SecureStore.getItemAsync(TOKEN_KEY);
}

export async function clearTokens(): Promise<void> {
  await SecureStore.deleteItemAsync(TOKEN_KEY);
  await SecureStore.deleteItemAsync(REFRESH_KEY);
  await SecureStore.deleteItemAsync(USER_KEY);
}

export async function fetchMSUser(accessToken: string): Promise<MSUser> {
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to fetch user profile");
  const data = await res.json();
  const user: MSUser = {
    id: data.id,
    displayName: data.displayName,
    mail: data.mail ?? data.userPrincipalName,
    jobTitle: data.jobTitle,
    userPrincipalName: data.userPrincipalName,
  };
  await SecureStore.setItemAsync(USER_KEY, JSON.stringify(user));
  return user;
}

export async function getStoredUser(): Promise<MSUser | null> {
  const raw = await SecureStore.getItemAsync(USER_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as MSUser;
  } catch {
    return null;
  }
}
