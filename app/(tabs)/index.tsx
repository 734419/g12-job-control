import {
  View,
  Text,
  ScrollView,
  Pressable,
  StyleSheet,
  ActivityIndicator,
  RefreshControl,
} from "react-native";
import { useRouter } from "expo-router";
import { useEffect, useState, useCallback } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";
import { getJobs, getDaySheets, getSubcontractors } from "@/lib/api/sharepoint";
import { getQueueCount } from "@/lib/offline/queue";

interface DashboardStats {
  activeJobs: number;
  pendingApprovals: number;
  complianceAlerts: number;
  queuedItems: number;
}

export default function DashboardScreen() {
  const colors = useColors();
  const router = useRouter();
  const { user } = useAuth();
  const [stats, setStats] = useState<DashboardStats>({
    activeJobs: 0,
    pendingApprovals: 0,
    complianceAlerts: 0,
    queuedItems: 0,
  });
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [lastSync, setLastSync] = useState<string>("");
  const [error, setError] = useState<string | null>(null);

  const loadStats = useCallback(async () => {
    try {
      setError(null);
      const [jobs, sheets, subs, queued] = await Promise.all([
        getJobs().catch(() => []),
        getDaySheets().catch(() => []),
        getSubcontractors().catch(() => []),
        getQueueCount(),
      ]);

      setStats({
        activeJobs: jobs.filter((j) => j.status === "Active").length,
        pendingApprovals: sheets.filter((s) => s.approvalStatus === "Pending").length,
        complianceAlerts: subs.filter(
          (s) => s.complianceStatus === "Blocked" || s.complianceStatus === "Expiring Soon"
        ).length,
        queuedItems: queued,
      });
      setLastSync(new Date().toLocaleTimeString());
    } catch (e: any) {
      setError("Could not load data. Check your connection.");
    }
  }, []);

  useEffect(() => {
    loadStats().finally(() => setLoading(false));
  }, [loadStats]);

  const onRefresh = useCallback(async () => {
    setRefreshing(true);
    await loadStats();
    setRefreshing(false);
  }, [loadStats]);

  const greeting = () => {
    const hour = new Date().getHours();
    if (hour < 12) return "Good morning";
    if (hour < 17) return "Good afternoon";
    return "Good evening";
  };

  const firstName = user?.displayName?.split(" ")[0] ?? "there";

  return (
    <ScreenContainer containerClassName="bg-background">
      <ScrollView
        contentContainerStyle={{ flexGrow: 1 }}
        refreshControl={<RefreshControl refreshing={refreshing} onRefresh={onRefresh} />}
        showsVerticalScrollIndicator={false}
      >
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#1B2A4A" }]}>
          <View>
            <Text style={styles.greeting}>{greeting()}, {firstName}</Text>
            <Text style={styles.headerTitle}>G12 Job Control</Text>
          </View>
          <View style={styles.syncBadge}>
            <IconSymbol name="arrow.clockwise" size={12} color="rgba(255,255,255,0.7)" />
            <Text style={styles.syncText}>
              {lastSync ? `Synced ${lastSync}` : "Syncing…"}
            </Text>
          </View>
        </View>

        <View style={styles.content}>
          {/* Offline queue banner */}
          {stats.queuedItems > 0 && (
            <View style={[styles.offlineBanner, { backgroundColor: "#FEF3C7" }]}>
              <IconSymbol name="wifi.slash" size={16} color="#D97706" />
              <Text style={styles.offlineBannerText}>
                {stats.queuedItems} item{stats.queuedItems !== 1 ? "s" : ""} queued for sync
              </Text>
            </View>
          )}

          {error && (
            <View style={[styles.errorBanner, { backgroundColor: "#FEE2E2" }]}>
              <IconSymbol name="exclamationmark.triangle.fill" size={16} color="#DC2626" />
              <Text style={styles.errorBannerText}>{error}</Text>
            </View>
          )}

          {/* Stats cards */}
          {loading ? (
            <ActivityIndicator color="#1B2A4A" style={{ marginTop: 40 }} />
          ) : (
            <>
              <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Overview</Text>
              <View style={styles.statsGrid}>
                <StatCard
                  label="Active Jobs"
                  value={stats.activeJobs}
                  icon="briefcase.fill"
                  color="#2563EB"
                  onPress={() => router.push("/jobs")}
                  colors={colors}
                />
                <StatCard
                  label="Pending Approvals"
                  value={stats.pendingApprovals}
                  icon="clock.fill"
                  color="#D97706"
                  onPress={() => router.push("/daysheets")}
                  colors={colors}
                />
                <StatCard
                  label="Compliance Alerts"
                  value={stats.complianceAlerts}
                  icon="exclamationmark.triangle.fill"
                  color={stats.complianceAlerts > 0 ? "#DC2626" : "#16A34A"}
                  onPress={() => router.push("/compliance")}
                  colors={colors}
                />
                <StatCard
                  label="Queued Offline"
                  value={stats.queuedItems}
                  icon="tray.full.fill"
                  color="#6366F1"
                  colors={colors}
                />
              </View>

              {/* Quick actions */}
              <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Quick Actions</Text>
              <View style={styles.actionsRow}>
                <QuickAction
                  label="New Day Sheet"
                  icon="doc.badge.plus"
                  color="#2563EB"
                  onPress={() => router.push("/daysheet/new")}
                  colors={colors}
                />
                <QuickAction
                  label="View Jobs"
                  icon="briefcase.fill"
                  color="#1B2A4A"
                  onPress={() => router.push("/jobs")}
                  colors={colors}
                />
                <QuickAction
                  label="Compliance"
                  icon="shield.fill"
                  color="#16A34A"
                  onPress={() => router.push("/compliance")}
                  colors={colors}
                />
              </View>
            </>
          )}
        </View>
      </ScrollView>
    </ScreenContainer>
  );
}

function StatCard({
  label,
  value,
  icon,
  color,
  onPress,
  colors,
}: {
  label: string;
  value: number;
  icon: any;
  color: string;
  onPress?: () => void;
  colors: any;
}) {
  return (
    <Pressable
      onPress={onPress}
      style={({ pressed }) => [
        styles.statCard,
        { backgroundColor: colors.surface, opacity: pressed ? 0.85 : 1 },
      ]}
    >
      <View style={[styles.statIconBg, { backgroundColor: `${color}18` }]}>
        <IconSymbol name={icon} size={22} color={color} />
      </View>
      <Text style={[styles.statValue, { color: colors.foreground }]}>{value}</Text>
      <Text style={[styles.statLabel, { color: colors.muted }]}>{label}</Text>
    </Pressable>
  );
}

function QuickAction({
  label,
  icon,
  color,
  onPress,
  colors,
}: {
  label: string;
  icon: any;
  color: string;
  onPress: () => void;
  colors: any;
}) {
  return (
    <Pressable
      onPress={onPress}
      style={({ pressed }) => [
        styles.quickAction,
        { backgroundColor: colors.surface, opacity: pressed ? 0.85 : 1 },
      ]}
    >
      <View style={[styles.quickActionIcon, { backgroundColor: color }]}>
        <IconSymbol name={icon} size={20} color="#fff" />
      </View>
      <Text style={[styles.quickActionLabel, { color: colors.foreground }]}>{label}</Text>
    </Pressable>
  );
}

const styles = StyleSheet.create({
  header: {
    paddingHorizontal: 20,
    paddingTop: 20,
    paddingBottom: 28,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-start",
  },
  greeting: {
    fontSize: 13,
    color: "rgba(255,255,255,0.7)",
    fontFamily: "Montserrat_400Regular",
    marginBottom: 2,
  },
  headerTitle: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
    letterSpacing: 0.3,
  },
  syncBadge: {
    flexDirection: "row",
    alignItems: "center",
    gap: 4,
    marginTop: 4,
  },
  syncText: {
    fontSize: 11,
    color: "rgba(255,255,255,0.6)",
    fontFamily: "Montserrat_400Regular",
  },
  content: {
    padding: 16,
    gap: 8,
  },
  offlineBanner: {
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    padding: 12,
    borderRadius: 10,
    marginBottom: 4,
  },
  offlineBannerText: {
    fontSize: 13,
    color: "#92400E",
    fontFamily: "Montserrat_500Medium",
  },
  errorBanner: {
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    padding: 12,
    borderRadius: 10,
    marginBottom: 4,
  },
  errorBannerText: {
    fontSize: 13,
    color: "#DC2626",
    fontFamily: "Montserrat_400Regular",
    flex: 1,
  },
  sectionTitle: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    marginTop: 8,
    marginBottom: 4,
  },
  statsGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 12,
    marginBottom: 8,
  },
  statCard: {
    flex: 1,
    minWidth: "44%",
    borderRadius: 14,
    padding: 16,
    alignItems: "flex-start",
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.06,
    shadowRadius: 8,
    elevation: 2,
  },
  statIconBg: {
    width: 40,
    height: 40,
    borderRadius: 10,
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 10,
  },
  statValue: {
    fontSize: 28,
    fontFamily: "Montserrat_700Bold",
    lineHeight: 34,
  },
  statLabel: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    marginTop: 2,
    lineHeight: 16,
  },
  actionsRow: {
    flexDirection: "row",
    gap: 12,
  },
  quickAction: {
    flex: 1,
    borderRadius: 14,
    padding: 14,
    alignItems: "center",
    gap: 8,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.06,
    shadowRadius: 8,
    elevation: 2,
  },
  quickActionIcon: {
    width: 44,
    height: 44,
    borderRadius: 12,
    alignItems: "center",
    justifyContent: "center",
  },
  quickActionLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_600SemiBold",
    textAlign: "center",
    lineHeight: 15,
  },
});
