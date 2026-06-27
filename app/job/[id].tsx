import {
  View,
  Text,
  ScrollView,
  Pressable,
  StyleSheet,
  ActivityIndicator,
} from "react-native";
import { useLocalSearchParams, useRouter } from "expo-router";
import { useEffect, useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getJob, type Job } from "@/lib/api/sharepoint";

const STATUS_COLORS: Record<string, string> = {
  Active: "#16A34A",
  Completed: "#6366F1",
  "On Hold": "#D97706",
  Cancelled: "#DC2626",
};

function InfoRow({ label, value, icon }: { label: string; value: string; icon?: any }) {
  const colors = useColors();
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
      <View style={styles.infoValueRow}>
        {icon && <IconSymbol name={icon} size={14} color={colors.muted} />}
        <Text style={[styles.infoValue, { color: colors.foreground }]}>{value || "—"}</Text>
      </View>
    </View>
  );
}

export default function JobDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const colors = useColors();
  const router = useRouter();
  const [job, setJob] = useState<Job | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (id) {
      getJob(id)
        .then(setJob)
        .finally(() => setLoading(false));
    }
  }, [id]);

  if (loading) {
    return (
      <ScreenContainer>
        <ActivityIndicator color="#1B2A4A" style={{ marginTop: 40 }} />
      </ScreenContainer>
    );
  }

  if (!job) {
    return (
      <ScreenContainer>
        <View style={styles.emptyState}>
          <Text style={[styles.emptyTitle, { color: colors.foreground }]}>Job not found</Text>
        </View>
      </ScreenContainer>
    );
  }

  const statusColor = STATUS_COLORS[job.status] ?? "#6366F1";

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView showsVerticalScrollIndicator={false}>
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#1B2A4A" }]}>
          <View style={[styles.jobCodeBadge, { backgroundColor: "rgba(255,255,255,0.15)" }]}>
            <Text style={styles.jobCodeText}>{job.jobCode}</Text>
          </View>
          <Text style={styles.jobName}>{job.jobName}</Text>
          <View style={[styles.statusBadge, { backgroundColor: `${statusColor}30` }]}>
            <Text style={[styles.statusText, { color: statusColor }]}>{job.status}</Text>
          </View>
        </View>

        {/* Details */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Job Details</Text>
          <InfoRow label="Client" value={job.client} icon="building.2.fill" />
          <InfoRow label="Site Address" value={job.siteAddress} icon="location.fill" />
          <InfoRow label="Start Date" value={job.startDate} icon="calendar" />
          <InfoRow label="Supervisor" value={job.supervisor} icon="person.fill" />
          {job.description ? (
            <InfoRow label="Description" value={job.description} />
          ) : null}
        </View>

        {/* Actions */}
        <View style={styles.actions}>
          <Pressable
            style={({ pressed }) => [styles.actionBtn, { backgroundColor: "#2563EB", opacity: pressed ? 0.85 : 1 }]}
            onPress={() => router.push({ pathname: "/daysheet/new" as any, params: { jobCode: job.jobCode } })}
          >
            <IconSymbol name="doc.badge.plus" size={18} color="#fff" />
            <Text style={styles.actionBtnText}>New Day Sheet</Text>
          </Pressable>
        </View>
      </ScrollView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  header: {
    padding: 24,
    gap: 10,
  },
  jobCodeBadge: {
    alignSelf: "flex-start",
    paddingHorizontal: 12,
    paddingVertical: 5,
    borderRadius: 8,
  },
  jobCodeText: {
    color: "#FFFFFF",
    fontSize: 13,
    fontFamily: "Montserrat_700Bold",
    letterSpacing: 1,
  },
  jobName: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
    lineHeight: 28,
  },
  statusBadge: {
    alignSelf: "flex-start",
    paddingHorizontal: 12,
    paddingVertical: 5,
    borderRadius: 20,
  },
  statusText: {
    fontSize: 12,
    fontFamily: "Montserrat_600SemiBold",
  },
  card: {
    margin: 16,
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
  infoRow: {
    paddingVertical: 10,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: "#E2E8F0",
    gap: 4,
  },
  infoLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_500Medium",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  infoValueRow: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    marginTop: 2,
  },
  infoValue: {
    fontSize: 15,
    fontFamily: "Montserrat_400Regular",
    flex: 1,
  },
  actions: {
    padding: 16,
    gap: 10,
  },
  actionBtn: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    padding: 14,
    borderRadius: 12,
  },
  actionBtnText: {
    color: "#fff",
    fontSize: 15,
    fontFamily: "Montserrat_600SemiBold",
  },
  emptyState: {
    alignItems: "center",
    paddingTop: 60,
  },
  emptyTitle: {
    fontSize: 18,
    fontFamily: "Montserrat_600SemiBold",
  },
});
