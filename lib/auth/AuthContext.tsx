import React, { createContext, useContext, useEffect, useState, useCallback } from "react";
import * as AuthSession from "expo-auth-session";
import * as WebBrowser from "expo-web-browser";
import { Platform } from "react-native";
import Constants from "expo-constants";
import {
  buildAuthRequest,
  exchangeCodeForToken,
  fetchMSUser,
  getStoredToken,
  getStoredUser,
  clearTokens,
  type MSUser,
} from "./microsoft";
import { checkIsSupervisor } from "@/lib/api/sharepoint";

WebBrowser.maybeCompleteAuthSession();

interface AuthState {
  isAuthenticated: boolean;
  isLoading: boolean;
  /** True when the role check is still in progress after sign-in */
  isRoleLoading: boolean;
  user: MSUser | null;
  accessToken: string | null;
  /**
   * True if the signed-in user is a member of the "Ausslope Job Control" M365 group.
   * Members of this group can approve day sheets and toggle mobilisation.
   */
  isSupervisor: boolean;
  signIn: () => Promise<void>;
  signOut: () => Promise<void>;
  error: string | null;
}

const AuthContext = createContext<AuthState>({
  isAuthenticated: false,
  isLoading: true,
  isRoleLoading: false,
  user: null,
  accessToken: null,
  isSupervisor: false,
  signIn: async () => {},
  signOut: async () => {},
  error: null,
});

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [isRoleLoading, setIsRoleLoading] = useState(false);
  const [user, setUser] = useState<MSUser | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isSupervisor, setIsSupervisor] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Build the correct redirect URI for each platform/environment:
  //   Web            → https://<host>/oauth/callback  (SPA flow)
  //   Expo Go        → exp://<hostUri>  (derived dynamically from the Metro manifest)
  //   Standalone app → manus20260626://auth  (custom scheme, registered in Azure iOS/macOS)
  const isExpoGo = Constants.appOwnership === "expo";

  // Derive the Expo Go redirect URI dynamically from the manifest hostUri so it
  // works regardless of which Manus sandbox is running.
  const hostUri: string =
    (Constants.expoConfig?.hostUri as string | undefined) ??
    (Constants.manifest2?.extra?.expoClient?.hostUri as string | undefined) ??
    "";
  // Strip any trailing port from the hostUri — Expo Go uses the bare host
  const expoGoHost = hostUri.replace(/:\d+$/, "");
  const expoGoRedirectUri = expoGoHost ? `exp://${expoGoHost}` : AuthSession.makeRedirectUri();

 const redirectUri = Platform.OS === "web"
    // Always use localhost for web — this is the URI registered in Entra ID.
    // The sandbox tunnel URL changes every session and cannot be permanently registered.
    ? "http://localhost:8081/oauth/callback"
    : isExpoGo
      ? expoGoRedirectUri
      : AuthSession.makeRedirectUri({ scheme: "manus20260626", path: "auth" });

  /** Check M365 group membership and cache the result */
  const loadRole = useCallback(async () => {
    setIsRoleLoading(true);
    try {
      const supervisor = await checkIsSupervisor();
      setIsSupervisor(supervisor);
    } catch {
      setIsSupervisor(false);
    } finally {
      setIsRoleLoading(false);
    }
  }, []);

  // Restore session on mount
  useEffect(() => {
    (async () => {
      try {
        const [token, storedUser] = await Promise.all([
          getStoredToken(),
          getStoredUser(),
        ]);
        if (token && storedUser) {
          setAccessToken(token);
          setUser(storedUser);
          setIsAuthenticated(true);
          // Load role in background — don't block the session restore
          loadRole();
        }
      } catch {
        // ignore
      } finally {
        setIsLoading(false);
      }
    })();
  }, [loadRole]);

  const signIn = useCallback(async () => {
    setError(null);
    setIsLoading(true);
    try {
      const request = buildAuthRequest(redirectUri);
      const tenantId = process.env.EXPO_PUBLIC_AZURE_TENANT_ID ?? "32bd37cb-b07f-4787-9d46-11184f6962d6";
      const authEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;

      if (Platform.OS === "web") {
        // On web: use full-page redirect to avoid popup blocker.
        // Store the code verifier in sessionStorage so the callback page can retrieve it.
        const authUrl = await request.makeAuthUrlAsync({ authorizationEndpoint: authEndpoint });
        if (typeof window !== "undefined" && request.codeVerifier) {
          sessionStorage.setItem("pkce_verifier", request.codeVerifier);
          sessionStorage.setItem("pkce_redirect_uri", redirectUri);
          window.location.href = authUrl;
        }
        return; // page will redirect away — no further state updates needed
      }

      // Native (iOS / Android / Expo Go): use in-app browser
      const result = await request.promptAsync({ authorizationEndpoint: authEndpoint });

      if (result.type === "success" && result.params.code) {
        const bundle = await exchangeCodeForToken(
          result.params.code,
          request.codeVerifier!,
          redirectUri
        );
        const msUser = await fetchMSUser(bundle.accessToken);
        setAccessToken(bundle.accessToken);
        setUser(msUser);
        setIsAuthenticated(true);
        loadRole();
      } else if (result.type === "error") {
        setError(result.error?.message ?? "Sign in failed");
      }
    } catch (e: any) {
      setError(e.message ?? "Sign in failed");
    } finally {
      setIsLoading(false);
    }
  }, [redirectUri, loadRole]);

  const signOut = useCallback(async () => {
    await clearTokens();
    setIsAuthenticated(false);
    setUser(null);
    setAccessToken(null);
    setIsSupervisor(false);
  }, []);

  return (
    <AuthContext.Provider
      value={{
        isAuthenticated,
        isLoading,
        isRoleLoading,
        user,
        accessToken,
        isSupervisor,
        signIn,
        signOut,
        error,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  return useContext(AuthContext);
}
