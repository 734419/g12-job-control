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
import { getJob, getDaySheets, type Job } from "@/lib/api/sharepoint";

const STATUS_COLORS: Record<string, string> = {
  Active: "#16A34A",
  Completed: "#6366F1",
  "On Hold": "#D97706",
  Cancelled: "#DC2626",
};

const PRIORITY_COLORS: Record<string, string> = {
  High: "#DC2626",
  Medium: "#D97706",
  Low: "#16A34A",
  Normal: "#6366F1",
};

function InfoRow({ label, value, icon }: { label: string; value: string; icon?: any }) {
  const colors = useColors();
  if (!value) return null;
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
      <View style={styles.infoValueRow}>
        {icon && <IconSymbol name={icon} size={14} color={colors.muted} />}
        <Text style={[styles.infoValue, { color: colors.foreground }]}>{value}</Text>
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
  const [linkedSheetCount, setLinkedSheetCount] = useState<number | null>(null);

  useEffect(() => {
    if (!id) return;
    getJob(id)
      .then((j) => {
        setJob(j);
        return j;
      })
      .then(async (j) => {
        if (!j) return;
        // Count day sheets linked to this job code
        try {
          const sheets = await getDaySheets();
          setLinkedSheetCount(sheets.filter((s) => s.jobCode === j.jobCode).length);
        } catch {
          setLinkedSheetCount(null);
        }
      })
      .finally(() => setLoading(false));
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
  const priorityColor = PRIORITY_COLORS[job.priority] ?? "#6366F1";

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView showsVerticalScrollIndicator={false}>
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#1B2A4A" }]}>
          <View style={styles.headerTopRow}>
            <Pressable onPress={() => router.back()} style={styles.backBtn}>
              <IconSymbol name="chevron.left" size={20} color="#fff" />
            </Pressable>
            <View style={[styles.jobCodeBadge, { backgroundColor: "rgba(255,255,255,0.15)" }]}>
              <Text style={styles.jobCodeText}>{job.jobCode}</Text>
            </View>
          </View>
          <Text style={styles.jobName}>{job.jobName}</Text>
          <View style={styles.headerBadgeRow}>
            <View style={[styles.statusBadge, { backgroundColor: `${statusColor}30` }]}>
              <Text style={[styles.statusText, { color: statusColor }]}>{job.status}</Text>
            </View>
            {job.priority && job.priority !== "Normal" && (
              <View style={[styles.statusBadge, { backgroundColor: `${priorityColor}20` }]}>
                <Text style={[styles.statusText, { color: priorityColor }]}>{job.priority} Priority</Text>
              </View>
            )}
            {job.jobType ? (
              <View style={[styles.statusBadge, { backgroundColor: "rgba(255,255,255,0.12)" }]}>
                <Text style={[styles.statusText, { color: "#CBD5E1" }]}>{job.jobType}</Text>
              </View>
            ) : null}
          </View>
        </View>

        {/* Day Sheets Summary Card */}
        {linkedSheetCount !== null && (
          <View style={[styles.summaryCard, { backgroundColor: "#EFF6FF", borderColor: "#BFDBFE" }]}>
            <View style={styles.summaryRow}>
              <View style={styles.summaryItem}>
                <IconSymbol name="doc.text.fill" size={20} color="#2563EB" />
                <Text style={styles.summaryValue}>{linkedSheetCount}</Text>
                <Text style={styles.summaryLabel}>Day Sheet{linkedSheetCount !== 1 ? "s" : ""}</Text>
              </View>
            </View>
          </View>
        )}

        {/* Job Details Card */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Job Details</Text>
          <InfoRow label="Client" value={job.client} icon="building.2.fill" />
          <InfoRow label="Site Address" value={job.siteAddress} icon="location.fill" />
          <InfoRow label="Start Date" value={job.startDate} icon="calendar" />
          <InfoRow label="Completion Date" value={job.completionDate} icon="calendar.badge.checkmark" />
          <InfoRow label="Contract Value" value={job.contractValue} icon="dollarsign.circle.fill" />
        </View>

        {/* Team Card */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Team</Text>
          <InfoRow label="Project Manager" value={job.projectManager} icon="person.fill" />
          <InfoRow label="Superintendent" value={job.superintendent} icon="person.badge.shield.checkmark.fill" />
        </View>

        {/* Description Card */}
        {job.description ? (
          <View style={[styles.card, { backgroundColor: colors.surface }]}>
            <Text style={[styles.cardTitle, { color: colors.foreground }]}>Description</Text>
            <Text style={[styles.descriptionText, { color: colors.foreground }]}>{job.description}</Text>
          </View>
        ) : null}

        {/* Actions */}
        <View style={styles.actions}>
          <Pressable
            style={({ pressed }) => [styles.actionBtn, { backgroundColor: "#2563EB", opacity: pressed ? 0.85 : 1 }]}
            onPress={() => router.push({ pathname: "/daysheet/new" as any, params: { jobCode: job.jobCode } })}
          >
            <IconSymbol name="doc.badge.plus" size={18} color="#fff" />
            <Text style={styles.actionBtnText}>New Day Sheet</Text>
          </Pressable>

          {linkedSheetCount !== null && linkedSheetCount > 0 && (
            <Pressable
              style={({ pressed }) => [
                styles.actionBtn,
                { backgroundColor: colors.surface, borderWidth: 1.5, borderColor: "#1B2A4A", opacity: pressed ? 0.85 : 1 },
              ]}
              onPress={() => router.push("/(tabs)/daysheets" as any)}
            >
              <IconSymbol name="doc.text.fill" size={18} color="#1B2A4A" />
              <Text style={[styles.actionBtnText, { color: "#1B2A4A" }]}>
                View {linkedSheetCount} Day Sheet{linkedSheetCount !== 1 ? "s" : ""}
              </Text>
            </Pressable>
          )}
        </View>
      </ScrollView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  header: {
    padding: 20,
    paddingTop: 16,
    gap: 10,
  },
  headerTopRow: {
    flexDirection: "row",
    alignItems: "center",
    gap: 12,
  },
  backBtn: {
    width: 36,
    height: 36,
    alignItems: "center",
    justifyContent: "center",
    borderRadius: 18,
    backgroundColor: "rgba(255,255,255,0.12)",
  },
  jobCodeBadge: {
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
  headerBadgeRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  statusBadge: {
    paddingHorizontal: 12,
    paddingVertical: 5,
    borderRadius: 20,
  },
  statusText: {
    fontSize: 12,
    fontFamily: "Montserrat_600SemiBold",
  },
  summaryCard: {
    margin: 16,
    marginBottom: 4,
    borderRadius: 14,
    padding: 16,
    borderWidth: 1,
  },
  summaryRow: {
    flexDirection: "row",
    justifyContent: "center",
    alignItems: "center",
    gap: 8,
  },
  summaryItem: {
    alignItems: "center",
    gap: 4,
  },
  summaryValue: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#1E40AF",
  },
  summaryLabel: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    color: "#3B82F6",
  },
  card: {
    margin: 16,
    marginBottom: 4,
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
  descriptionText: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    lineHeight: 22,
  },
  actions: {
    padding: 16,
    paddingTop: 12,
    gap: 10,
    paddingBottom: 32,
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
