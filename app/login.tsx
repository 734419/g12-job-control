import { View, Text, Pressable, ActivityIndicator, StyleSheet } from "react-native";
import { useRouter } from "expo-router";
import { useEffect } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { useAuth } from "@/lib/auth/AuthContext";
import { useColors } from "@/hooks/use-colors";

export default function LoginScreen() {
  const { signIn, isLoading, isAuthenticated, error } = useAuth();
  const colors = useColors();
  const router = useRouter();

  useEffect(() => {
    if (isAuthenticated) {
      router.replace("/(tabs)");
    }
  }, [isAuthenticated]);

  return (
    <ScreenContainer
      containerClassName="bg-navy"
      safeAreaClassName="justify-center items-center px-8"
    >
      {/* Logo / Brand */}
      <View style={styles.logoContainer}>
        <View style={[styles.logoCircle, { backgroundColor: "rgba(255,255,255,0.12)" }]}>
          <Text style={styles.logoText}>ASR</Text>
        </View>
        <Text style={styles.appTitle}>Job Control</Text>
        <Text style={styles.appSubtitle}>Field Operations Platform</Text>
      </View>

      {/* Sign in card */}
      <View style={[styles.card, { backgroundColor: colors.surface }]}>
        <Text style={[styles.cardTitle, { color: colors.foreground }]}>
          Sign in to continue
        </Text>
        <Text style={[styles.cardSubtitle, { color: colors.muted }]}>
          Use your Ausslope Microsoft 365 account
        </Text>

        {error ? (
          <View style={[styles.errorBox, { backgroundColor: "#FEE2E2" }]}>
            <Text style={styles.errorText}>{error}</Text>
          </View>
        ) : null}

        <Pressable
          onPress={signIn}
          disabled={isLoading}
          style={({ pressed }) => [
            styles.msButton,
            { opacity: pressed || isLoading ? 0.8 : 1 },
          ]}
        >
          {isLoading ? (
            <ActivityIndicator color="#fff" size="small" />
          ) : (
            <>
              <View style={styles.msLogo}>
                <View style={[styles.msSquare, { backgroundColor: "#F25022" }]} />
                <View style={[styles.msSquare, { backgroundColor: "#7FBA00" }]} />
                <View style={[styles.msSquare, { backgroundColor: "#00A4EF" }]} />
                <View style={[styles.msSquare, { backgroundColor: "#FFB900" }]} />
              </View>
              <Text style={styles.msButtonText}>Sign in with Microsoft 365</Text>
            </>
          )}
        </Pressable>
      </View>

      <Text style={styles.footer}>
        Ausslope · Australian Slope Retention · Secure access via Azure AD
      </Text>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  logoContainer: {
    alignItems: "center",
    marginBottom: 48,
  },
  logoCircle: {
    width: 88,
    height: 88,
    borderRadius: 44,
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 16,
    borderWidth: 2,
    borderColor: "rgba(255,255,255,0.3)",
  },
  logoText: {
    fontSize: 32,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
    letterSpacing: 2,
  },
  appTitle: {
    fontSize: 28,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
    letterSpacing: 0.5,
  },
  appSubtitle: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    color: "rgba(255,255,255,0.7)",
    marginTop: 4,
    letterSpacing: 0.5,
  },
  card: {
    width: "100%",
    borderRadius: 20,
    padding: 28,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 8 },
    shadowOpacity: 0.15,
    shadowRadius: 24,
    elevation: 8,
  },
  cardTitle: {
    fontSize: 20,
    fontFamily: "Montserrat_700Bold",
    marginBottom: 6,
  },
  cardSubtitle: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    marginBottom: 24,
    lineHeight: 20,
  },
  errorBox: {
    borderRadius: 10,
    padding: 12,
    marginBottom: 16,
  },
  errorText: {
    color: "#DC2626",
    fontSize: 13,
    fontFamily: "Montserrat_400Regular",
  },
  msButton: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: "#7F1F1F",
    borderRadius: 12,
    paddingVertical: 14,
    paddingHorizontal: 20,
    gap: 12,
  },
  msLogo: {
    width: 20,
    height: 20,
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 2,
  },
  msSquare: {
    width: 8,
    height: 8,
  },
  msButtonText: {
    color: "#FFFFFF",
    fontSize: 15,
    fontFamily: "Montserrat_600SemiBold",
    letterSpacing: 0.3,
  },
  footer: {
    marginTop: 32,
    fontSize: 12,
    color: "rgba(255,255,255,0.5)",
    fontFamily: "Montserrat_400Regular",
    textAlign: "center",
  },
});
