import { DarkTheme, DefaultTheme, ThemeProvider as NavThemeProvider } from "@react-navigation/native";
import { Stack } from "expo-router";
import { StatusBar } from "expo-status-bar";
import { useEffect } from "react";
import { GestureHandlerRootView } from "react-native-gesture-handler";
import {
  useFonts,
  Montserrat_400Regular,
  Montserrat_500Medium,
  Montserrat_600SemiBold,
  Montserrat_700Bold,
} from "@expo-google-fonts/montserrat";
import * as SplashScreen from "expo-splash-screen";
import "../global.css";
import { ThemeProvider, useThemeContext } from "@/lib/theme-provider";
import { AuthProvider } from "@/lib/auth/AuthContext";

SplashScreen.preventAutoHideAsync();

function RootLayoutInner() {
  const { colorScheme } = useThemeContext();
  return (
    <NavThemeProvider value={colorScheme === "dark" ? DarkTheme : DefaultTheme}>
      <Stack screenOptions={{ headerShown: false }}>
        <Stack.Screen name="(tabs)" options={{ headerShown: false }} />
        <Stack.Screen name="login" options={{ headerShown: false }} />
        <Stack.Screen name="job/[id]" options={{ headerShown: true, title: "Job Detail" }} />
        <Stack.Screen name="daysheet/new" options={{ headerShown: true, title: "New Day Sheet" }} />
        <Stack.Screen name="daysheet/[id]" options={{ headerShown: true, title: "Day Sheet" }} />
        <Stack.Screen name="subcontractor/[id]" options={{ headerShown: true, title: "Subcontractor" }} />
        <Stack.Screen name="subcontractor/new" options={{ headerShown: true, title: "Add Subcontractor" }} />
      </Stack>
      <StatusBar style={colorScheme === "dark" ? "light" : "dark"} />
    </NavThemeProvider>
  );
}

export default function RootLayout() {
  const [fontsLoaded] = useFonts({
    Montserrat_400Regular,
    Montserrat_500Medium,
    Montserrat_600SemiBold,
    Montserrat_700Bold,
  });

  useEffect(() => {
    if (fontsLoaded) {
      SplashScreen.hideAsync();
    }
  }, [fontsLoaded]);

  if (!fontsLoaded) return null;

  return (
    <GestureHandlerRootView style={{ flex: 1 }}>
      <ThemeProvider>
        <AuthProvider>
          <RootLayoutInner />
        </AuthProvider>
      </ThemeProvider>
    </GestureHandlerRootView>
  );
}
