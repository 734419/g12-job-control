import {
  View,
  Text,
  ScrollView,
  Pressable,
  StyleSheet,
  Alert,
  ActivityIndicator,
} from "react-native";
import { useLocalSearchParams, useRouter } from "expo-router";
import { useEffect, useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getDaySheets, approveDaySheet, rejectDaySheet, type DaySheet } from "@/lib/api/sharepoint";
import { useAuth } from "@/lib/auth/AuthContext";

const STATUS_COLORS: Record<string, { bg: string; text: string }> = {
  Pending: { bg: "#FEF3C7", text: "#D97706" },
  Approved: { bg: "#DCFCE7", text: "#16A34A" },
  Rejected: { bg: "#FEE2E2", text: "#DC2626" },
};

function InfoRow({ label, value }: { label: string; value: string }) {
  const colors = useColors();
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
      <Text style={[styles.infoValue, { color: colors.foreground }]}>{value || "—"}</Text>
    </View>
  );
}

export default function DaySheetDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const colors = useColors();
  const router = useRouter();
  const { user, isSupervisor } = useAuth();
  const [sheet, setSheet] = useState<DaySheet | null>(null);
  const [loading, setLoading] = useState(true);
  const [actioning, setActioning] = useState(false);

  useEffect(() => {
    if (id) {
      getDaySheets()
        .then((sheets) => setSheet(sheets.find((s) => s.id === id) ?? null))
        .finally(() => setLoading(false));
    }
  }, [id]);

  const handleApprove = async () => {
    if (!sheet || !user) return;
    setActioning(true);
    try {
      const today = new Date().toISOString().split("T")[0];
      await approveDaySheet(sheet.id, user.displayName);
      setSheet((prev) =>
        prev
          ? { ...prev, approvalStatus: "Approved", approvedBy: user.displayName, approvedDate: today }
          : prev
      );
      Alert.alert("Approved", "Day sheet has been approved.");
    } catch {
      Alert.alert("Error", "Could not approve. Check your connection.");
    } finally {
      setActioning(false);
    }
  };

  const handleReject = async () => {
    if (!sheet || !user) return;
    setActioning(true);
    try {
      const today = new Date().toISOString().split("T")[0];
      await rejectDaySheet(sheet.id, user.displayName);
      setSheet((prev) =>
        prev
          ? { ...prev, approvalStatus: "Rejected", approvedBy: user.displayName, approvedDate: today }
          : prev
      );
      Alert.alert("Rejected", "Day sheet has been rejected.");
    } catch {
      Alert.alert("Error", "Could not reject. Check your connection.");
    } finally {
      setActioning(false);
    }
  };

  if (loading) {
    return (
      <ScreenContainer>
        <ActivityIndicator color="#1B2A4A" style={{ marginTop: 40 }} />
      </ScreenContainer>
    );
  }

  if (!sheet) {
    return (
      <ScreenContainer>
        <View style={styles.emptyState}>
          <Text style={[styles.emptyTitle, { color: colors.foreground }]}>Day sheet not found</Text>
        </View>
      </ScreenContainer>
    );
  }

  const statusStyle = STATUS_COLORS[sheet.approvalStatus] ?? STATUS_COLORS.Pending;

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView showsVerticalScrollIndicator={false}>
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#1B2A4A" }]}>
          <Text style={styles.headerWorker}>{sheet.workerName}</Text>
          <Text style={styles.headerDate}>{sheet.date} · {sheet.jobCode}</Text>
          <View style={[styles.statusBadge, { backgroundColor: statusStyle.bg }]}>
            <Text style={[styles.statusText, { color: statusStyle.text }]}>
              {sheet.approvalStatus}
            </Text>
          </View>
        </View>

        {/* Hours card */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Hours</Text>
          <View style={styles.hoursGrid}>
            <HoursBox label="Start" value={sheet.startTime} />
            <HoursBox label="Finish" value={sheet.finishTime} />
            <HoursBox label="Break" value={`${sheet.breakMinutes} min`} />
            <HoursBox label="Ordinary" value={`${sheet.ordinaryHours} hrs`} highlight />
            <HoursBox label="Overtime" value={`${sheet.overtimeHours} hrs`} highlight={sheet.overtimeHours > 0} />
          </View>
        </View>

        {/* Details */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Details</Text>
          <InfoRow label="Worker Email" value={sheet.workerEmail} />
          {sheet.allowances ? (
            <InfoRow label="Allowances" value={sheet.allowances} />
          ) : null}
          {sheet.notes && <InfoRow label="Notes" value={sheet.notes} />}
        </View>

        {/* Payroll */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Payroll</Text>
          <InfoRow label="Export Status" value={sheet.payrollExportStatus} />
          {sheet.xeroReference && <InfoRow label="Xero Reference" value={sheet.xeroReference} />}
          {sheet.approvedBy && <InfoRow label="Approved By" value={sheet.approvedBy} />}
          {sheet.approvedDate && <InfoRow label="Approved Date" value={sheet.approvedDate} />}
        </View>

        {/* Approval actions — supervisor only */}
        {sheet.approvalStatus === "Pending" && isSupervisor && (
          <View style={styles.actions}>
            <Pressable
              onPress={handleApprove}
              disabled={actioning}
              style={({ pressed }) => [
                styles.approveBtn,
                { opacity: pressed || actioning ? 0.8 : 1 },
              ]}
            >
              {actioning ? (
                <ActivityIndicator color="#fff" />
              ) : (
                <>
                  <IconSymbol name="checkmark.circle.fill" size={18} color="#fff" />
                  <Text style={styles.actionBtnText}>Approve</Text>
                </>
              )}
            </Pressable>
            <Pressable
              onPress={handleReject}
              disabled={actioning}
              style={({ pressed }) => [
                styles.rejectBtn,
                { opacity: pressed || actioning ? 0.8 : 1 },
              ]}
            >
              <IconSymbol name="xmark.circle.fill" size={18} color="#DC2626" />
              <Text style={[styles.actionBtnText, { color: "#DC2626" }]}>Reject</Text>
            </Pressable>
          </View>
        )}
      </ScrollView>
    </ScreenContainer>
  );
}

function HoursBox({ label, value, highlight }: { label: string; value: string; highlight?: boolean }) {
  return (
    <View style={[styles.hoursBox, highlight && { backgroundColor: "#EFF6FF" }]}>
      <Text style={[styles.hoursBoxLabel, highlight && { color: "#1E40AF" }]}>{label}</Text>
      <Text style={[styles.hoursBoxValue, highlight && { color: "#1E40AF" }]}>{value}</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  header: {
    padding: 24,
    gap: 8,
  },
  headerWorker: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  headerDate: {
    fontSize: 14,
    color: "rgba(255,255,255,0.7)",
    fontFamily: "Montserrat_400Regular",
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
  hoursGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  hoursBox: {
    flex: 1,
    minWidth: "28%",
    backgroundColor: "#F8FAFC",
    borderRadius: 10,
    padding: 12,
    alignItems: "center",
  },
  hoursBoxLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_500Medium",
    color: "#64748B",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  hoursBoxValue: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    color: "#1B2A4A",
    marginTop: 4,
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
  infoValue: {
    fontSize: 15,
    fontFamily: "Montserrat_400Regular",
  },
  actions: {
    padding: 16,
    gap: 10,
    marginTop: 8,
  },
  approveBtn: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    backgroundColor: "#16A34A",
    padding: 14,
    borderRadius: 12,
  },
  rejectBtn: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: 8,
    backgroundColor: "#FEE2E2",
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
