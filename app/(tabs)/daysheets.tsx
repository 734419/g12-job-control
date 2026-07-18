import {
  View,
  Text,
  FlatList,
  Pressable,
  StyleSheet,
  RefreshControl,
  ActivityIndicator,
  Image,
} from "react-native";
import { useRouter } from "expo-router";
import { useState, useEffect, useCallback } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getDaySheets, type DaySheet } from "@/lib/api/sharepoint";

const TABS = ["Pending", "Approved", "Exported"] as const;
type TabType = (typeof TABS)[number];

const STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  Pending: { bg: "#FEF3C7", text: "#D97706" },
  Approved: { bg: "#DCFCE7", text: "#16A34A" },
  Rejected: { bg: "#FEE2E2", text: "#DC2626" },
  Exported: { bg: "#EDE9FE", text: "#7C3AED" },
};

export default function DaySheetsScreen() {
  const colors = useColors();
  const router = useRouter();
  const [activeTab, setActiveTab] = useState<TabType>("Pending");
  const [sheets, setSheets] = useState<DaySheet[]>([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    try {
      setError(null);
      const data = await getDaySheets();
      setSheets(data);
    } catch {
      setError("Unable to load day sheets.");
    }
  }, []);

  useEffect(() => {
    load().finally(() => setLoading(false));
  }, [load]);

  const onRefresh = useCallback(async () => {
    setRefreshing(true);
    await load();
    setRefreshing(false);
  }, [load]);

  const filtered = sheets.filter((s) => {
    if (activeTab === "Exported") return s.payrollExportStatus === "Exported";
    if (activeTab === "Approved") return s.approvalStatus === "Approved" && s.payrollExportStatus !== "Exported";
    return s.approvalStatus === "Pending";
  });

  return (
    <ScreenContainer containerClassName="bg-background">
      {/* Header */}
      <View style={[styles.header, { backgroundColor: "#7F1F1F" }]}>
        <View style={styles.headerLeft}>
          <Image
            source={require("@/assets/images/asr-logo.png")}
            style={styles.logo}
            resizeMode="contain"
          />
          <Text style={styles.headerTitle}>Day Sheets</Text>
        </View>
        <View style={{ flexDirection: "row", gap: 8 }}>
          <Pressable
            style={styles.addBtn}
            onPress={() => router.push("/daysheet/export" as any)}
          >
            <IconSymbol name="arrow.up.doc.fill" size={18} color="#fff" />
            <Text style={styles.addBtnText}>Export</Text>
          </Pressable>
          <Pressable
            style={styles.addBtn}
            onPress={() => router.push("/daysheet/new")}
          >
            <IconSymbol name="plus.circle.fill" size={20} color="#fff" />
            <Text style={styles.addBtnText}>New</Text>
          </Pressable>
        </View>
      </View>

      {/* Tabs */}
      <View style={[styles.tabBar, { backgroundColor: colors.surface, borderBottomColor: colors.border }]}>
        {TABS.map((tab) => (
          <Pressable
            key={tab}
            style={[
              styles.tab,
              activeTab === tab && { borderBottomColor: "#B91C1C", borderBottomWidth: 2 },
            ]}
            onPress={() => setActiveTab(tab)}
          >
            <Text
              style={[
                styles.tabText,
                { color: activeTab === tab ? "#B91C1C" : colors.muted },
              ]}
            >
              {tab}
            </Text>
          </Pressable>
        ))}
      </View>

      {loading ? (
        <ActivityIndicator color="#7F1F1F" style={{ marginTop: 40 }} />
      ) : error ? (
        <View style={styles.emptyState}>
          <IconSymbol name="exclamationmark.triangle.fill" size={40} color="#DC2626" />
          <Text style={[styles.emptyTitle, { color: colors.foreground }]}>Connection Error</Text>
          <Text style={[styles.emptySubtitle, { color: colors.muted }]}>{error}</Text>
          <Pressable style={styles.retryBtn} onPress={onRefresh}>
            <Text style={styles.retryBtnText}>Retry</Text>
          </Pressable>
        </View>
      ) : (
        <FlatList
          data={filtered}
          keyExtractor={(item) => item.id}
          contentContainerStyle={{ padding: 16, gap: 10, paddingBottom: 100 }}
          refreshControl={<RefreshControl refreshing={refreshing} onRefresh={onRefresh} />}
          ListEmptyComponent={
            <View style={styles.emptyState}>
              <IconSymbol name="doc.text.fill" size={40} color={colors.muted} />
              <Text style={[styles.emptyTitle, { color: colors.foreground }]}>No Day Sheets</Text>
              <Text style={[styles.emptySubtitle, { color: colors.muted }]}>
                {activeTab === "Pending" ? "No pending approvals" : `No ${activeTab.toLowerCase()} day sheets`}
              </Text>
            </View>
          }
          renderItem={({ item }) => {
            const statusStyle = STATUS_COLORS[item.approvalStatus] ?? STATUS_COLORS.Pending;
            return (
              <Pressable
                style={({ pressed }) => [
                  styles.card,
                  { backgroundColor: colors.surface, opacity: pressed ? 0.85 : 1 },
                ]}
                onPress={() => router.push({ pathname: "/daysheet/[id]" as any, params: { id: item.id } })}
              >
                <View style={styles.cardTop}>
                  <View style={[styles.jobCodeBadge, { backgroundColor: "#7F1F1F" }]}>
                    <Text style={styles.jobCodeText}>{item.jobCode}</Text>
                  </View>
                  <View style={[styles.statusBadge, { backgroundColor: statusStyle.bg }]}>
                    <Text style={[styles.statusText, { color: statusStyle.text }]}>
                      {item.approvalStatus}
                    </Text>
                  </View>
                </View>
                <Text style={[styles.workerName, { color: colors.foreground }]}>{item.workerName}</Text>
                <View style={styles.cardMeta}>
                  <IconSymbol name="calendar" size={13} color={colors.muted} />
                  <Text style={[styles.metaText, { color: colors.muted }]}>{item.date}</Text>
                  <Text style={[styles.metaDot, { color: colors.muted }]}>·</Text>
                  <IconSymbol name="clock.fill" size={13} color={colors.muted} />
                  <Text style={[styles.metaText, { color: colors.muted }]}>
                    {item.ordinaryHours}h ordinary
                    {item.overtimeHours > 0 ? ` + ${item.overtimeHours}h OT` : ""}
                  </Text>
                </View>
                <View style={styles.cardFooter}>
                  {item.payrollExportStatus === "Exported" && (
                    <Text style={styles.exportedTag}>Exported to Payroll</Text>
                  )}
                  <IconSymbol name="chevron.right" size={16} color={colors.muted} style={{ marginLeft: "auto" }} />
                </View>
              </Pressable>
            );
          }}
        />
      )}
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  header: {
    paddingHorizontal: 20,
    paddingTop: 20,
    paddingBottom: 20,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  headerLeft: {
    flexDirection: "row",
    alignItems: "center",
    gap: 10,
  },
  logo: {
    width: 44,
    height: 44,
    tintColor: "#FFFFFF",
  },
  headerTitle: {
    fontSize: 20,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  addBtn: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    backgroundColor: "rgba(255,255,255,0.15)",
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 20,
  },
  addBtnText: {
    color: "#fff",
    fontFamily: "Montserrat_600SemiBold",
    fontSize: 14,
  },
  tabBar: {
    flexDirection: "row",
    borderBottomWidth: StyleSheet.hairlineWidth,
  },
  tab: {
    flex: 1,
    paddingVertical: 12,
    alignItems: "center",
  },
  tabText: {
    fontSize: 13,
    fontFamily: "Montserrat_600SemiBold",
  },
  card: {
    borderRadius: 14,
    padding: 16,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.06,
    shadowRadius: 8,
    elevation: 2,
    gap: 6,
  },
  cardTop: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: 4,
  },
  jobCodeBadge: {
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 6,
  },
  jobCodeText: {
    color: "#FFFFFF",
    fontSize: 12,
    fontFamily: "Montserrat_700Bold",
    letterSpacing: 0.5,
  },
  statusBadge: {
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 20,
  },
  statusText: {
    fontSize: 11,
    fontFamily: "Montserrat_600SemiBold",
  },
  workerName: {
    fontSize: 16,
    fontFamily: "Montserrat_600SemiBold",
  },
  cardMeta: {
    flexDirection: "row",
    alignItems: "center",
    gap: 4,
  },
  metaText: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
  },
  metaDot: {
    fontSize: 12,
  },
  cardFooter: {
    flexDirection: "row",
    alignItems: "center",
    marginTop: 4,
    paddingTop: 8,
    borderTopWidth: StyleSheet.hairlineWidth,
    borderTopColor: "#E2E8F0",
  },
  exportedTag: {
    fontSize: 11,
    fontFamily: "Montserrat_500Medium",
    color: "#7C3AED",
    backgroundColor: "#EDE9FE",
    paddingHorizontal: 8,
    paddingVertical: 3,
    borderRadius: 10,
  },
  emptyState: {
    alignItems: "center",
    paddingTop: 60,
    gap: 8,
  },
  emptyTitle: {
    fontSize: 18,
    fontFamily: "Montserrat_600SemiBold",
    marginTop: 8,
  },
  emptySubtitle: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    textAlign: "center",
    paddingHorizontal: 32,
  },
  retryBtn: {
    marginTop: 16,
    backgroundColor: "#7F1F1F",
    paddingHorizontal: 24,
    paddingVertical: 10,
    borderRadius: 10,
  },
  retryBtnText: {
    color: "#fff",
    fontFamily: "Montserrat_600SemiBold",
    fontSize: 14,
  },
});
