import React, { createContext, useContext, useEffect, useState, useCallback } from "react";
import * as AuthSession from "expo-auth-session";
import * as WebBrowser from "expo-web-browser";
import {
  buildAuthRequest,
  exchangeCodeForToken,
  fetchMSUser,
  getStoredToken,
  getStoredUser,
  clearTokens,
  type MSUser,
} from "./microsoft";

WebBrowser.maybeCompleteAuthSession();

interface AuthState {
  isAuthenticated: boolean;
  isLoading: boolean;
  user: MSUser | null;
  accessToken: string | null;
  signIn: () => Promise<void>;
  signOut: () => Promise<void>;
  error: string | null;
}

const AuthContext = createContext<AuthState>({
  isAuthenticated: false,
  isLoading: true,
  user: null,
  accessToken: null,
  signIn: async () => {},
  signOut: async () => {},
  error: null,
});

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [user, setUser] = useState<MSUser | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const redirectUri = AuthSession.makeRedirectUri({ scheme: "manus20260626" });

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
        }
      } catch {
        // ignore
      } finally {
        setIsLoading(false);
      }
    })();
  }, []);

  const signIn = useCallback(async () => {
    setError(null);
    setIsLoading(true);
    try {
      const request = buildAuthRequest(redirectUri);
      await request.makeAuthUrlAsync({ authorizationEndpoint: `https://login.microsoftonline.com/${process.env.EXPO_PUBLIC_AZURE_TENANT_ID ?? "common"}/oauth2/v2.0/authorize` });
      const result = await request.promptAsync({ authorizationEndpoint: `https://login.microsoftonline.com/${process.env.EXPO_PUBLIC_AZURE_TENANT_ID ?? "common"}/oauth2/v2.0/authorize` });

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
      } else if (result.type === "error") {
        setError(result.error?.message ?? "Sign in failed");
      }
    } catch (e: any) {
      setError(e.message ?? "Sign in failed");
    } finally {
      setIsLoading(false);
    }
  }, [redirectUri]);

  const signOut = useCallback(async () => {
    await clearTokens();
    setIsAuthenticated(false);
    setUser(null);
    setAccessToken(null);
  }, []);

  return (
    <AuthContext.Provider
      value={{ isAuthenticated, isLoading, user, accessToken, signIn, signOut, error }}
    >
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  return useContext(AuthContext);
}
