import { Tabs, useRouter } from "expo-router";
import { useEffect } from "react";
import { useSafeAreaInsets } from "react-native-safe-area-context";
import { Platform } from "react-native";
import { HapticTab } from "@/components/haptic-tab";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";

export default function TabLayout() {
  const colors = useColors();
  const insets = useSafeAreaInsets();
  const { isAuthenticated, isLoading } = useAuth();
  const router = useRouter();

  const bottomPadding = Platform.OS === "web" ? 12 : Math.max(insets.bottom, 8);
  const tabBarHeight = 60 + bottomPadding;

  useEffect(() => {
    if (!isLoading && !isAuthenticated) {
      router.replace("/login");
    }
  }, [isAuthenticated, isLoading]);

  return (
    <Tabs
      screenOptions={{
        tabBarActiveTintColor: "#B91C1C",
        tabBarInactiveTintColor: colors.muted,
        headerShown: false,
        tabBarButton: HapticTab,
        tabBarLabelStyle: {
          fontSize: 10,
          fontFamily: "Montserrat_500Medium",
          marginTop: -2,
        },
        tabBarStyle: {
          paddingTop: 8,
          paddingBottom: bottomPadding,
          height: tabBarHeight,
          backgroundColor: colors.surface,
          borderTopColor: colors.border,
          borderTopWidth: 0.5,
        },
      }}
    >
      <Tabs.Screen
        name="index"
        options={{
          title: "Dashboard",
          tabBarIcon: ({ color }) => (
            <IconSymbol size={24} name="house.fill" color={color} />
          ),
        }}
      />
      <Tabs.Screen
        name="jobs"
        options={{
          title: "Jobs",
          tabBarIcon: ({ color }) => (
            <IconSymbol size={24} name="briefcase.fill" color={color} />
          ),
        }}
      />
      <Tabs.Screen
        name="daysheets"
        options={{
          title: "Day Sheets",
          tabBarIcon: ({ color }) => (
            <IconSymbol size={24} name="doc.text.fill" color={color} />
          ),
        }}
      />
      <Tabs.Screen
        name="compliance"
        options={{
          title: "Compliance",
          tabBarIcon: ({ color }) => (
            <IconSymbol size={24} name="shield.fill" color={color} />
          ),
        }}
      />
      <Tabs.Screen
        name="profile"
        options={{
          title: "Profile",
          tabBarIcon: ({ color }) => (
            <IconSymbol size={24} name="person.fill" color={color} />
          ),
        }}
      />
    </Tabs>
  );
}
