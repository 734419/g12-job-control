import {
  View,
  Text,
  ScrollView,
  Pressable,
  StyleSheet,
  Alert,
} from "react-native";
import { useState, useEffect } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";
import { getQueueCount, flushQueue } from "@/lib/offline/queue";

export default function ProfileScreen() {
  const colors = useColors();
  const { user, signOut } = useAuth();
  const [queueCount, setQueueCount] = useState(0);
  const [syncing, setSyncing] = useState(false);

  useEffect(() => {
    getQueueCount().then(setQueueCount);
  }, []);

  const handleSync = async () => {
    setSyncing(true);
    try {
      const result = await flushQueue();
      // Re-read the actual remaining count (failed items stay in the queue)
      const remaining = await getQueueCount();
      setQueueCount(remaining);
      Alert.alert(
        result.failed > 0 ? "Sync Partial" : "Sync Complete",
        `${result.success} item${result.success !== 1 ? "s" : ""} synced.${
          result.failed > 0
            ? ` ${result.failed} item${result.failed !== 1 ? "s" : ""} failed and will retry next time.`
            : ""
        }`
      );
    } catch {
      Alert.alert("Sync Failed", "Could not sync. Check your connection.");
    } finally {
      setSyncing(false);
    }
  };

  const handleSignOut = () => {
    Alert.alert("Sign Out", "Are you sure you want to sign out?", [
      { text: "Cancel", style: "cancel" },
      { text: "Sign Out", style: "destructive", onPress: signOut },
    ]);
  };

  const initials = user?.displayName
    ?.split(" ")
    .map((n) => n[0])
    .slice(0, 2)
    .join("")
    .toUpperCase() ?? "?";

  return (
    <ScreenContainer containerClassName="bg-background">
      <ScrollView showsVerticalScrollIndicator={false}>
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#1B2A4A" }]}>
          <View style={styles.avatar}>
            <Text style={styles.avatarText}>{initials}</Text>
          </View>
          <Text style={styles.displayName}>{user?.displayName ?? "Unknown User"}</Text>
          <Text style={styles.email}>{user?.mail ?? ""}</Text>
          {user?.jobTitle ? (
            <Text style={styles.jobTitle}>{user.jobTitle}</Text>
          ) : null}
        </View>

        {/* Sync status */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Sync Status</Text>
          <View style={styles.syncRow}>
            <View style={styles.syncInfo}>
              <IconSymbol
                name={queueCount > 0 ? "wifi.slash" : "wifi"}
                size={20}
                color={queueCount > 0 ? "#D97706" : "#16A34A"}
              />
              <View>
                <Text style={[styles.syncLabel, { color: colors.foreground }]}>
                  {queueCount > 0 ? `${queueCount} item${queueCount !== 1 ? "s" : ""} queued` : "All synced"}
                </Text>
                <Text style={[styles.syncSubLabel, { color: colors.muted }]}>
                  {queueCount > 0 ? "Pending upload to SharePoint" : "SharePoint is up to date"}
                </Text>
              </View>
            </View>
            {queueCount > 0 && (
              <Pressable
                onPress={handleSync}
                disabled={syncing}
                style={({ pressed }) => [
                  styles.syncBtn,
                  { opacity: pressed || syncing ? 0.8 : 1 },
                ]}
              >
                <Text style={styles.syncBtnText}>{syncing ? "Syncing…" : "Sync Now"}</Text>
              </Pressable>
            )}
          </View>
        </View>

        {/* App info */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>App Info</Text>
          <InfoRow label="Platform" value="G12 Job Control" colors={colors} />
          <InfoRow label="SharePoint Site" value="G12 Job Control" colors={colors} />
          <InfoRow label="Tenant" value="G12 Group" colors={colors} />
          <InfoRow label="Version" value="1.0.0" colors={colors} />
        </View>

        {/* Sign out */}
        <View style={{ padding: 16, paddingBottom: 40 }}>
          <Pressable
            onPress={handleSignOut}
            style={({ pressed }) => [
              styles.signOutBtn,
              { opacity: pressed ? 0.8 : 1 },
            ]}
          >
            <IconSymbol name="arrow.right.square" size={18} color="#DC2626" />
            <Text style={styles.signOutText}>Sign Out</Text>
          </Pressable>
        </View>
      </ScrollView>
    </ScreenContainer>
  );
}

function InfoRow({ label, value, colors }: { label: string; value: string; colors: any }) {
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
      <Text style={[styles.infoValue, { color: colors.foreground }]}>{value}</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  header: {
    paddingHorizontal: 24,
    paddingTop: 24,
    paddingBottom: 32,
    alignItems: "center",
    gap: 8,
  },
  avatar: {
    width: 80,
    height: 80,
    borderRadius: 40,
    backgroundColor: "rgba(255,255,255,0.15)",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 8,
    borderWidth: 2,
    borderColor: "rgba(255,255,255,0.3)",
  },
  avatarText: {
    fontSize: 28,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  displayName: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  email: {
    fontSize: 14,
    color: "rgba(255,255,255,0.7)",
    fontFamily: "Montserrat_400Regular",
  },
  jobTitle: {
    fontSize: 13,
    color: "rgba(255,255,255,0.5)",
    fontFamily: "Montserrat_400Regular",
  },
  card: {
    margin: 16,
    marginBottom: 0,
    borderRadius: 16,
    padding: 20,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.06,
    shadowRadius: 8,
    elevation: 2,
  },
  cardTitle: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    marginBottom: 16,
  },
  syncRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  syncInfo: {
    flexDirection: "row",
    alignItems: "center",
    gap: 12,
    flex: 1,
  },
  syncLabel: {
    fontSize: 15,
    fontFamily: "Montserrat_600SemiBold",
  },
  syncSubLabel: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    marginTop: 2,
  },
  syncBtn: {
    backgroundColor: "#1B2A4A",
    paddingHorizontal: 16,
    paddingVertical: 8,
    borderRadius: 10,
  },
  syncBtnText: {
    color: "#fff",
    fontSize: 13,
    fontFamily: "Montserrat_600SemiBold",
  },
  infoRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    paddingVertical: 10,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: "#E2E8F0",
  },
  infoLabel: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
  },
  infoValue: {
    fontSize: 14,
    fontFamily: "Montserrat_500Medium",
  },
  signOutBtn: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    backgroundColor: "#FEE2E2",
    padding: 16,
    borderRadius: 14,
  },
  signOutText: {
    color: "#DC2626",
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
  },
});
