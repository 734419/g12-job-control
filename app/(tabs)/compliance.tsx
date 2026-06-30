import {
  View,
  Text,
  FlatList,
  Pressable,
  StyleSheet,
  TextInput,
  RefreshControl,
  ActivityIndicator,
  Image,
} from "react-native";
import { useRouter } from "expo-router";
import { useState, useEffect, useCallback } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getSubcontractors, type Subcontractor } from "@/lib/api/sharepoint";

const COMPLIANCE_COLORS: Record<string, { bg: string; text: string; icon: any }> = {
  Active: { bg: "#DCFCE7", text: "#16A34A", icon: "checkmark.circle.fill" },
  "Expiring Soon": { bg: "#FEF3C7", text: "#D97706", icon: "exclamationmark.triangle.fill" },
  Blocked: { bg: "#FEE2E2", text: "#DC2626", icon: "xmark.circle.fill" },
};

const FILTER_OPTIONS = ["All", "Active", "Expiring Soon", "Blocked"] as const;
type FilterType = (typeof FILTER_OPTIONS)[number];

export default function ComplianceScreen() {
  const colors = useColors();
  const router = useRouter();
  const [subs, setSubs] = useState<Subcontractor[]>([]);
  const [filtered, setFiltered] = useState<Subcontractor[]>([]);
  const [search, setSearch] = useState("");
  const [activeFilter, setActiveFilter] = useState<FilterType>("All");
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    try {
      setError(null);
      const data = await getSubcontractors();
      setSubs(data);
      setFiltered(data);
    } catch {
      setError("Unable to load compliance data.");
    }
  }, []);

  useEffect(() => {
    load().finally(() => setLoading(false));
  }, [load]);

  useEffect(() => {
    let result = subs;
    if (activeFilter !== "All") {
      result = result.filter((s) => s.complianceStatus === activeFilter);
    }
    if (search.trim()) {
      const q = search.toLowerCase();
      result = result.filter(
        (s) =>
          s.companyName.toLowerCase().includes(q) ||
          s.trade.toLowerCase().includes(q) ||
          s.abn.includes(q)
      );
    }
    setFiltered(result);
  }, [search, activeFilter, subs]);

  const onRefresh = useCallback(async () => {
    setRefreshing(true);
    await load();
    setRefreshing(false);
  }, [load]);

  const counts = {
    All: subs.length,
    Active: subs.filter((s) => s.complianceStatus === "Active").length,
    "Expiring Soon": subs.filter((s) => s.complianceStatus === "Expiring Soon").length,
    Blocked: subs.filter((s) => s.complianceStatus === "Blocked").length,
  };

  return (
    <ScreenContainer containerClassName="bg-background">
      {/* Header */}
      <View style={[styles.header, { backgroundColor: "#0F2A44" }]}>
        <View style={styles.headerLeft}>
          <Image
            source={require("@/assets/images/g12-logo.png")}
            style={styles.logo}
            resizeMode="contain"
          />
          <View>
            <Text style={styles.headerTitle}>Compliance</Text>
            <Text style={styles.headerSub}>{subs.length} subcontractor{subs.length !== 1 ? "s" : ""}</Text>
          </View>
        </View>
        <Pressable
          style={styles.addBtn}
          onPress={() => router.push("/subcontractor/new")}
        >
          <IconSymbol name="plus.circle.fill" size={20} color="#fff" />
          <Text style={styles.addBtnText}>Add</Text>
        </Pressable>
      </View>

      {/* Search */}
      <View style={[styles.searchContainer, { backgroundColor: colors.surface, borderColor: colors.border }]}>
        <IconSymbol name="magnifyingglass" size={16} color={colors.muted} />
        <TextInput
          style={[styles.searchInput, { color: colors.foreground }]}
          placeholder="Search by company, trade or ABN…"
          placeholderTextColor={colors.muted}
          value={search}
          onChangeText={setSearch}
          returnKeyType="search"
        />
        {search.length > 0 && (
          <Pressable onPress={() => setSearch("")}>
            <IconSymbol name="xmark.circle.fill" size={16} color={colors.muted} />
          </Pressable>
        )}
      </View>

      {/* Filter chips */}
      <View style={styles.filterRow}>
        {FILTER_OPTIONS.map((f) => (
          <Pressable
            key={f}
            onPress={() => setActiveFilter(f)}
            style={[
              styles.filterChip,
              activeFilter === f
                ? { backgroundColor: "#1B2A4A" }
                : { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1 },
            ]}
          >
            <Text
              style={[
                styles.filterChipText,
                { color: activeFilter === f ? "#fff" : colors.muted },
              ]}
            >
              {f} ({counts[f]})
            </Text>
          </Pressable>
        ))}
      </View>

      {loading ? (
        <ActivityIndicator color="#1B2A4A" style={{ marginTop: 40 }} />
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
              <IconSymbol name="shield.fill" size={40} color={colors.muted} />
              <Text style={[styles.emptyTitle, { color: colors.foreground }]}>No Subcontractors</Text>
              <Text style={[styles.emptySubtitle, { color: colors.muted }]}>
                {search ? "Try a different search" : "Add subcontractors to track compliance"}
              </Text>
            </View>
          }
          renderItem={({ item }) => {
            const compStyle = COMPLIANCE_COLORS[item.complianceStatus] ?? COMPLIANCE_COLORS.Active;
            return (
              <Pressable
                style={({ pressed }) => [
                  styles.card,
                  { backgroundColor: colors.surface, opacity: pressed ? 0.85 : 1 },
                ]}
                onPress={() => router.push({ pathname: "/subcontractor/[id]" as any, params: { id: item.id } })}
              >
                <View style={styles.cardTop}>
                  <Text style={[styles.companyName, { color: colors.foreground }]} numberOfLines={1}>
                    {item.companyName}
                  </Text>
                  <View style={[styles.complianceBadge, { backgroundColor: compStyle.bg }]}>
                    <IconSymbol name={compStyle.icon} size={12} color={compStyle.text} />
                    <Text style={[styles.complianceBadgeText, { color: compStyle.text }]}>
                      {item.complianceStatus}
                    </Text>
                  </View>
                </View>
                <Text style={[styles.trade, { color: colors.muted }]}>{item.trade}</Text>
                <View style={styles.expiryRow}>
                  <ExpiryTag label="Insurance" date={item.insuranceExpiry} />
                  <ExpiryTag label="Licence" date={item.licenceExpiry} />
                </View>
                <View style={styles.cardFooter}>
                  <View style={styles.statusRow}>
                    <StatusDot label="Induction" value={item.inductionStatus} />
                    <StatusDot label="SWMS" value={item.swmsStatus} />
                    <StatusDot label="Prequalification" value={item.prequalificationStatus} />
                  </View>
                  <IconSymbol name="chevron.right" size={16} color={colors.muted} />
                </View>
              </Pressable>
            );
          }}
        />
      )}
    </ScreenContainer>
  );
}

function ExpiryTag({ label, date }: { label: string; date: string }) {
  const expired = date && new Date(date) < new Date();
  const expiringSoon = !expired && date && (new Date(date).getTime() - Date.now()) < 30 * 86400000;
  const color = expired ? "#DC2626" : expiringSoon ? "#D97706" : "#16A34A";
  return (
    <View style={styles.expiryTag}>
      <Text style={[styles.expiryLabel, { color: "#64748B" }]}>{label}:</Text>
      <Text style={[styles.expiryDate, { color }]}>{date || "—"}</Text>
    </View>
  );
}

function StatusDot({ label, value }: { label: string; value: string }) {
  const isGood = value === "Complete" || value === "Approved";
  return (
    <View style={styles.statusDotRow}>
      <View style={[styles.dot, { backgroundColor: isGood ? "#16A34A" : "#D97706" }]} />
      <Text style={styles.statusDotLabel}>{label}</Text>
    </View>
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
  headerSub: {
    fontSize: 12,
    color: "rgba(255,255,255,0.6)",
    fontFamily: "Montserrat_400Regular",
    marginTop: 1,
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
  searchContainer: {
    flexDirection: "row",
    alignItems: "center",
    margin: 16,
    marginBottom: 8,
    borderRadius: 12,
    paddingHorizontal: 12,
    paddingVertical: 10,
    borderWidth: 1,
    gap: 8,
  },
  searchInput: {
    flex: 1,
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
  },
  filterRow: {
    flexDirection: "row",
    paddingHorizontal: 16,
    gap: 8,
    marginBottom: 4,
    flexWrap: "wrap",
  },
  filterChip: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 20,
  },
  filterChipText: {
    fontSize: 12,
    fontFamily: "Montserrat_500Medium",
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
    gap: 8,
  },
  companyName: {
    fontSize: 16,
    fontFamily: "Montserrat_600SemiBold",
    flex: 1,
  },
  complianceBadge: {
    flexDirection: "row",
    alignItems: "center",
    gap: 4,
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 20,
  },
  complianceBadgeText: {
    fontSize: 11,
    fontFamily: "Montserrat_600SemiBold",
  },
  trade: {
    fontSize: 13,
    fontFamily: "Montserrat_400Regular",
  },
  expiryRow: {
    flexDirection: "row",
    gap: 16,
    marginTop: 2,
  },
  expiryTag: {
    flexDirection: "row",
    gap: 4,
    alignItems: "center",
  },
  expiryLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_500Medium",
  },
  expiryDate: {
    fontSize: 11,
    fontFamily: "Montserrat_600SemiBold",
  },
  cardFooter: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    marginTop: 4,
    paddingTop: 8,
    borderTopWidth: StyleSheet.hairlineWidth,
    borderTopColor: "#E2E8F0",
  },
  statusRow: {
    flexDirection: "row",
    gap: 12,
  },
  statusDotRow: {
    flexDirection: "row",
    alignItems: "center",
    gap: 4,
  },
  dot: {
    width: 7,
    height: 7,
    borderRadius: 4,
  },
  statusDotLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_400Regular",
    color: "#64748B",
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
    backgroundColor: "#1B2A4A",
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
