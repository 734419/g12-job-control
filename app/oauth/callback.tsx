/**
 * OAuth Callback Screen
 *
 * Handles two flows:
 * 1. Web redirect flow (PKCE): Microsoft redirects back here with ?code=...
 *    The PKCE verifier was stored in sessionStorage before the redirect.
 * 2. Native deep link flow: handled by expo-auth-session in AuthContext.
 */
import { useEffect, useState } from "react";
import { ActivityIndicator, Text, View } from "react-native";
import { useRouter } from "expo-router";
import { SafeAreaView } from "react-native-safe-area-context";
import { ThemedView } from "@/components/themed-view";
import { exchangeCodeForToken, fetchMSUser, storeUser } from "@/lib/auth/microsoft";

export default function OAuthCallback() {
  const router = useRouter();
  const [status, setStatus] = useState<"processing" | "success" | "error">("processing");
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  useEffect(() => {
    const handleCallback = async () => {
      try {
        // Only run on web — native flow is handled by AuthContext promptAsync
        if (typeof window === "undefined") {
          setStatus("error");
          setErrorMessage("Unexpected callback on native platform");
          return;
        }

        const params = new URLSearchParams(window.location.search);
        const code = params.get("code");
        const error = params.get("error");
        const errorDesc = params.get("error_description");

        if (error) {
          setStatus("error");
          setErrorMessage(errorDesc ?? error);
          return;
        }

        if (!code) {
          setStatus("error");
          setErrorMessage("No authorisation code received from Microsoft.");
          return;
        }

        // Retrieve PKCE verifier and redirect URI stored before the redirect
        const codeVerifier = sessionStorage.getItem("pkce_verifier");
        const redirectUri = sessionStorage.getItem("pkce_redirect_uri");

        if (!codeVerifier || !redirectUri) {
          setStatus("error");
          setErrorMessage("Session expired. Please try signing in again.");
          return;
        }

        // Clean up sessionStorage
        sessionStorage.removeItem("pkce_verifier");
        sessionStorage.removeItem("pkce_redirect_uri");

        // Exchange code for tokens
        const bundle = await exchangeCodeForToken(code, codeVerifier, redirectUri);
        const msUser = await fetchMSUser(bundle.accessToken);
        await storeUser(msUser);

        setStatus("success");
        setTimeout(() => {
          router.replace("/(tabs)");
        }, 800);
      } catch (e: any) {
        console.error("[OAuth callback]", e);
        setStatus("error");
        setErrorMessage(e.message ?? "Authentication failed");
      }
    };

    handleCallback();
  }, [router]);

  return (
    <SafeAreaView className="flex-1" edges={["top", "bottom", "left", "right"]}>
      <ThemedView className="flex-1 items-center justify-center gap-4 p-5">
        {status === "processing" && (
          <>
            <ActivityIndicator size="large" />
            <Text className="mt-4 text-base leading-6 text-center text-foreground">
              Completing sign in…
            </Text>
          </>
        )}
        {status === "success" && (
          <Text className="text-base leading-6 text-center text-foreground">
            Signed in successfully. Redirecting…
          </Text>
        )}
        {status === "error" && (
          <>
            <Text className="mb-2 text-xl font-bold leading-7 text-error">
              Sign in failed
            </Text>
            <Text className="text-base leading-6 text-center text-foreground">
              {errorMessage}
            </Text>
            <Text
              className="mt-4 text-primary text-base"
              onPress={() => router.replace("/login")}
            >
              Try again
            </Text>
          </>
        )}
      </ThemedView>
    </SafeAreaView>
  );
}
