import {
  View,
  Text,
  ScrollView,
  StyleSheet,
  Alert,
  ActivityIndicator,
  Switch,
  Linking,
  TouchableOpacity,
} from "react-native";
import { useLocalSearchParams } from "expo-router";
import { useEffect, useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getSubcontractors, updateSubcontractorMobilisation, type Subcontractor } from "@/lib/api/sharepoint";
import { useAuth } from "@/lib/auth/AuthContext";

const COMPLIANCE_COLORS: Record<string, { bg: string; text: string; icon: any }> = {
  Active: { bg: "#DCFCE7", text: "#16A34A", icon: "checkmark.circle.fill" },
  "Expiring Soon": { bg: "#FEF3C7", text: "#D97706", icon: "exclamationmark.triangle.fill" },
  Blocked: { bg: "#FEE2E2", text: "#DC2626", icon: "xmark.circle.fill" },
};

function InfoRow({ label, value, valueColor }: { label: string; value: string; valueColor?: string }) {
  const colors = useColors();
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
      <Text style={[styles.infoValue, { color: valueColor ?? colors.foreground }]}>{value || "—"}</Text>
    </View>
  );
}

function DocRow({ label, url }: { label: string; url?: string }) {
  const colors = useColors();
  const hasDoc = !!url;

  const handleOpen = () => {
    if (!url) return;
    Linking.openURL(url).catch(() =>
      Alert.alert("Cannot Open", "Could not open the document. Please check your connection.")
    );
  };

  return (
    <View style={[styles.infoRow, styles.docRow]}>
      <View style={{ flex: 1 }}>
        <Text style={[styles.infoLabel, { color: colors.muted }]}>{label}</Text>
        <Text
          style={[
            styles.infoValue,
            { color: hasDoc ? "#7F1F1F" : colors.muted },
          ]}
        >
          {hasDoc ? "Document available" : "Not uploaded"}
        </Text>
      </View>
      {hasDoc && (
        <TouchableOpacity
          onPress={handleOpen}
          style={styles.docButton}
          activeOpacity={0.75}
        >
          <IconSymbol name="arrow.up.right.square" size={14} color="#fff" />
          <Text style={styles.docButtonText}>Open</Text>
        </TouchableOpacity>
      )}
    </View>
  );
}

function StatusBadge({ label, value }: { label: string; value: string }) {
  const statusColors: Record<string, { bg: string; text: string }> = {
    Complete: { bg: "#DCFCE7", text: "#16A34A" },
    Approved: { bg: "#DCFCE7", text: "#16A34A" },
    Pending: { bg: "#FEF3C7", text: "#D97706" },
    "Not Started": { bg: "#F1F5F9", text: "#64748B" },
    "Not Submitted": { bg: "#F1F5F9", text: "#64748B" },
  };
  const style = statusColors[value] ?? { bg: "#F1F5F9", text: "#64748B" };
  const themeColors = useColors();
  return (
    <View style={styles.infoRow}>
      <Text style={[styles.infoLabel, { color: themeColors.muted }]}>{label}</Text>
      <View style={[styles.statusPill, { backgroundColor: style.bg }]}>
        <Text style={[styles.statusPillText, { color: style.text }]}>{value}</Text>
      </View>
    </View>
  );
}

export default function SubcontractorDetailScreen() {
  const { id } = useLocalSearchParams<{ id: string }>();
  const colors = useColors();
  const { isSupervisor } = useAuth();
  const [sub, setSub] = useState<Subcontractor | null>(null);
  const [loading, setLoading] = useState(true);
  const [toggling, setToggling] = useState(false);

  useEffect(() => {
    if (id) {
      getSubcontractors()
        .then((subs) => setSub(subs.find((s) => s.id === id) ?? null))
        .finally(() => setLoading(false));
    }
  }, [id]);

  const handleToggleMobilisation = async (value: boolean) => {
    if (!sub) return;
    if (sub.complianceStatus === "Blocked" && value) {
      Alert.alert("Blocked", "This subcontractor has expired compliance. Resolve before approving mobilisation.");
      return;
    }
    setToggling(true);
    try {
      await updateSubcontractorMobilisation(sub.id, value);
      setSub((prev) => prev ? { ...prev, mobilisationApproved: value } : prev);
    } catch {
      Alert.alert("Error", "Could not update mobilisation status.");
    } finally {
      setToggling(false);
    }
  };

  if (loading) {
    return (
      <ScreenContainer>
        <ActivityIndicator color="#7F1F1F" style={{ marginTop: 40 }} />
      </ScreenContainer>
    );
  }

  if (!sub) {
    return (
      <ScreenContainer>
        <View style={styles.emptyState}>
          <Text style={[styles.emptyTitle, { color: colors.foreground }]}>Subcontractor not found</Text>
        </View>
      </ScreenContainer>
    );
  }

  const compStyle = COMPLIANCE_COLORS[sub.complianceStatus] ?? COMPLIANCE_COLORS.Active;

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView showsVerticalScrollIndicator={false}>
        {/* Header */}
        <View style={[styles.header, { backgroundColor: "#7F1F1F" }]}>
          <Text style={styles.companyName}>{sub.companyName}</Text>
          <Text style={styles.trade}>{sub.trade}</Text>
          <View style={[styles.complianceBadge, { backgroundColor: compStyle.bg }]}>
            <IconSymbol name={compStyle.icon} size={14} color={compStyle.text} />
            <Text style={[styles.complianceBadgeText, { color: compStyle.text }]}>
              {sub.complianceStatus}
            </Text>
          </View>
        </View>

        {/* Mobilisation toggle */}
        <View style={[styles.mobilisationCard, { backgroundColor: sub.mobilisationApproved ? "#DCFCE7" : "#FEF3C7" }]}>
          <View style={{ flex: 1 }}>
            <Text style={[styles.mobilisationTitle, { color: sub.mobilisationApproved ? "#16A34A" : "#92400E" }]}>
              Mobilisation {sub.mobilisationApproved ? "Approved" : "Not Approved"}
            </Text>
            <Text style={styles.mobilisationSubtitle}>
              {!isSupervisor
                ? "Supervisor access required"
                : sub.complianceStatus === "Blocked"
                ? "Resolve compliance issues first"
                : "Toggle to approve/revoke mobilisation"}
            </Text>
          </View>
          <Switch
            value={sub.mobilisationApproved}
            onValueChange={handleToggleMobilisation}
            disabled={toggling || sub.complianceStatus === "Blocked" || !isSupervisor}
            trackColor={{ false: "#E2E8F0", true: "#16A34A" }}
            thumbColor={sub.mobilisationApproved ? "#fff" : "#94A3B8"}
          />
        </View>

        {/* Contact */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Contact</Text>
          <InfoRow label="Company" value={sub.companyName} />
          <InfoRow label="ABN" value={sub.abn} />
          <InfoRow label="Contact Name" value={sub.contactName} />
          <InfoRow label="Phone" value={sub.contactPhone} />
          <InfoRow label="Email" value={sub.contactEmail} />
        </View>

        {/* Compliance */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Compliance</Text>
          <InfoRow
            label="Insurance Expiry"
            value={sub.insuranceExpiry}
            valueColor={
              new Date(sub.insuranceExpiry) < new Date() ? "#DC2626" :
              (new Date(sub.insuranceExpiry).getTime() - Date.now()) < 30 * 86400000 ? "#D97706" : undefined
            }
          />
          <InfoRow
            label="Licence Expiry"
            value={sub.licenceExpiry}
            valueColor={
              new Date(sub.licenceExpiry) < new Date() ? "#DC2626" :
              (new Date(sub.licenceExpiry).getTime() - Date.now()) < 30 * 86400000 ? "#D97706" : undefined
            }
          />
          <StatusBadge label="Induction" value={sub.inductionStatus} />
          <StatusBadge label="SWMS" value={sub.swmsStatus} />
          <StatusBadge label="Prequalification" value={sub.prequalificationStatus} />
        </View>

        {/* Documents */}
        <View style={[styles.card, { backgroundColor: colors.surface }]}>
          <Text style={[styles.cardTitle, { color: colors.foreground }]}>Documents</Text>
          <DocRow label="Insurance Certificate" url={sub.insuranceDocUrl} />
          <DocRow label="Contractor Licence" url={sub.licenceDocUrl} />
          <DocRow label="SWMS" url={sub.swmsDocUrl} />
        </View>

        {/* Active Jobcodes */}
        {sub.activeJobCodes.length > 0 && (
          <View style={[styles.card, { backgroundColor: colors.surface }]}>
            <Text style={[styles.cardTitle, { color: colors.foreground }]}>Active Jobcodes</Text>
            <View style={styles.jobCodesWrap}>
              {sub.activeJobCodes.map((jc) => (
                <View key={jc} style={styles.jobCodeChip}>
                  <Text style={styles.jobCodeChipText}>{jc}</Text>
                </View>
              ))}
            </View>
          </View>
        )}

        {sub.notes ? (
          <View style={[styles.card, { backgroundColor: colors.surface }]}>
            <Text style={[styles.cardTitle, { color: colors.foreground }]}>Notes</Text>
            <Text style={[styles.notes, { color: colors.foreground }]}>{sub.notes}</Text>
          </View>
        ) : null}

        <View style={{ height: 24 }} />
      </ScrollView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  header: {
    padding: 24,
    gap: 8,
  },
  companyName: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  trade: {
    fontSize: 14,
    color: "rgba(255,255,255,0.7)",
    fontFamily: "Montserrat_400Regular",
  },
  complianceBadge: {
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
    alignSelf: "flex-start",
    paddingHorizontal: 12,
    paddingVertical: 5,
    borderRadius: 20,
  },
  complianceBadgeText: {
    fontSize: 12,
    fontFamily: "Montserrat_600SemiBold",
  },
  mobilisationCard: {
    margin: 16,
    marginBottom: 0,
    borderRadius: 14,
    padding: 16,
    flexDirection: "row",
    alignItems: "center",
    gap: 12,
  },
  mobilisationTitle: {
    fontSize: 15,
    fontFamily: "Montserrat_700Bold",
  },
  mobilisationSubtitle: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    color: "#64748B",
    marginTop: 2,
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
  infoRow: {
    paddingVertical: 10,
    borderBottomWidth: StyleSheet.hairlineWidth,
    borderBottomColor: "#E2E8F0",
    gap: 4,
  },
  docRow: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
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
  docButton: {
    flexDirection: "row",
    alignItems: "center",
    gap: 5,
    backgroundColor: "#7F1F1F",
    paddingHorizontal: 12,
    paddingVertical: 7,
    borderRadius: 8,
    marginLeft: 12,
  },
  docButtonText: {
    color: "#fff",
    fontSize: 12,
    fontFamily: "Montserrat_600SemiBold",
  },
  statusPill: {
    alignSelf: "flex-start",
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 20,
  },
  statusPillText: {
    fontSize: 12,
    fontFamily: "Montserrat_600SemiBold",
  },
  jobCodesWrap: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  jobCodeChip: {
    backgroundColor: "#7F1F1F",
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 8,
  },
  jobCodeChipText: {
    color: "#fff",
    fontSize: 12,
    fontFamily: "Montserrat_700Bold",
    letterSpacing: 0.5,
  },
  notes: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    lineHeight: 22,
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
