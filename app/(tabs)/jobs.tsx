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
import { getJobs, type Job } from "@/lib/api/sharepoint";

const STATUS_COLORS: Record<string, string> = {
  Active: "#16A34A",
  Completed: "#6366F1",
  "On Hold": "#D97706",
  Cancelled: "#DC2626",
};

export default function JobsScreen() {
  const colors = useColors();
  const router = useRouter();
  const [jobs, setJobs] = useState<Job[]>([]);
  const [filtered, setFiltered] = useState<Job[]>([]);
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    try {
      setError(null);
      const data = await getJobs();
      setJobs(data);
      setFiltered(data);
    } catch {
      setError("Unable to load jobs. Check your connection.");
    }
  }, []);

  useEffect(() => {
    load().finally(() => setLoading(false));
  }, [load]);

  useEffect(() => {
    if (!search.trim()) {
      setFiltered(jobs);
    } else {
      const q = search.toLowerCase();
      setFiltered(
        jobs.filter(
          (j) =>
            j.jobCode.toLowerCase().includes(q) ||
            j.jobName.toLowerCase().includes(q) ||
            j.client.toLowerCase().includes(q)
        )
      );
    }
  }, [search, jobs]);

  const onRefresh = useCallback(async () => {
    setRefreshing(true);
    await load();
    setRefreshing(false);
  }, [load]);

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
            <Text style={styles.headerTitle}>Jobs</Text>
            <Text style={styles.headerSub}>{filtered.length} job{filtered.length !== 1 ? "s" : ""}</Text>
          </View>
        </View>
        <Pressable
          style={styles.newJobBtn}
          onPress={() => router.push("/job/new" as any)}
        >
          <IconSymbol name="plus.circle.fill" size={18} color="#fff" />
          <Text style={styles.newJobBtnText}>New Job</Text>
        </Pressable>
      </View>

      {/* Search */}
      <View style={[styles.searchContainer, { backgroundColor: colors.surface, borderColor: colors.border }]}>
        <IconSymbol name="magnifyingglass" size={16} color={colors.muted} />
        <TextInput
          style={[styles.searchInput, { color: colors.foreground }]}
          placeholder="Search by jobcode, name or client…"
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
              <IconSymbol name="briefcase.fill" size={40} color={colors.muted} />
              <Text style={[styles.emptyTitle, { color: colors.foreground }]}>No Jobs Found</Text>
              <Text style={[styles.emptySubtitle, { color: colors.muted }]}>
                {search ? "Try a different search term" : "No active jobs in SharePoint"}
              </Text>
            </View>
          }
          renderItem={({ item }) => (
            <Pressable
              style={({ pressed }) => [
                styles.jobCard,
                { backgroundColor: colors.surface, opacity: pressed ? 0.85 : 1 },
              ]}
              onPress={() => router.push({ pathname: "/job/[id]" as any, params: { id: item.id } })}
            >
              <View style={styles.jobCardTop}>
                <View style={[styles.jobCodeBadge, { backgroundColor: "#1B2A4A" }]}>
                  <Text style={styles.jobCodeText}>{item.jobCode}</Text>
                </View>
                <View style={[styles.statusBadge, { backgroundColor: `${STATUS_COLORS[item.status] ?? "#6366F1"}18` }]}>
                  <Text style={[styles.statusText, { color: STATUS_COLORS[item.status] ?? "#6366F1" }]}>
                    {item.status}
                  </Text>
                </View>
              </View>
              <Text style={[styles.jobName, { color: colors.foreground }]}>{item.jobName}</Text>
              <Text style={[styles.jobClient, { color: colors.muted }]}>{item.client}</Text>
              {item.siteAddress ? (
                <View style={styles.jobLocation}>
                  <IconSymbol name="location.fill" size={12} color={colors.muted} />
                  <Text style={[styles.jobLocationText, { color: colors.muted }]}>{item.siteAddress}</Text>
                </View>
              ) : null}
              <View style={styles.jobFooter}>
                <Text style={[styles.jobSupervisor, { color: colors.muted }]}>
                  PM: {item.projectManager || "—"}
                </Text>
                <IconSymbol name="chevron.right" size={16} color={colors.muted} />
              </View>
            </Pressable>
          )}
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
  headerSub: {
    fontSize: 12,
    color: "rgba(255,255,255,0.6)",
    fontFamily: "Montserrat_400Regular",
    marginTop: 1,
  },
  searchContainer: {
    flexDirection: "row",
    alignItems: "center",
    margin: 16,
    marginBottom: 4,
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
  jobCard: {
    borderRadius: 14,
    padding: 16,
    shadowColor: "#000",
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.06,
    shadowRadius: 8,
    elevation: 2,
    gap: 6,
  },
  jobCardTop: {
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
  jobName: {
    fontSize: 16,
    fontFamily: "Montserrat_600SemiBold",
    lineHeight: 22,
  },
  jobClient: {
    fontSize: 13,
    fontFamily: "Montserrat_400Regular",
  },
  jobLocation: {
    flexDirection: "row",
    alignItems: "center",
    gap: 4,
    marginTop: 2,
  },
  jobLocationText: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    flex: 1,
  },
  jobFooter: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    marginTop: 4,
    paddingTop: 8,
    borderTopWidth: StyleSheet.hairlineWidth,
    borderTopColor: "#E2E8F0",
  },
  jobSupervisor: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
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
  newJobBtn: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    backgroundColor: "rgba(255,255,255,0.15)",
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 20,
  },
  newJobBtnText: {
    color: "#fff",
    fontFamily: "Montserrat_600SemiBold",
    fontSize: 14,
  },
});
