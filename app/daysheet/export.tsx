import {
  View,
  Text,
  FlatList,
  Pressable,
  StyleSheet,
  Alert,
  ActivityIndicator,
  Share,
  Platform,
} from "react-native";
import { useRouter } from "expo-router";
import { useState, useEffect, useCallback } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { getDaySheets, markDaySheetsExported, type DaySheet } from "@/lib/api/sharepoint";

function buildXeroCSV(sheets: DaySheet[]): string {
  const headers = [
    "Employee Name",
    "Employee Email",
    "Date",
    "Job Code",
    "Job Name",
    "Ordinary Hours",
    "Overtime Hours",
    "Total Hours",
    "Allowances",
    "Notes",
    "Approved By",
  ];

  const rows = sheets.map((s) => [
    `"${s.workerName}"`,
    `"${s.workerEmail}"`,
    `"${s.date}"`,
    `"${s.jobCode}"`,
    `"${s.jobName}"`,
    s.ordinaryHours.toFixed(2),
    s.overtimeHours.toFixed(2),
    (s.ordinaryHours + s.overtimeHours).toFixed(2),
    `"${s.allowances}"`,
    `"${s.notes.replace(/"/g, "'")}"`,
    `"${s.approvedBy}"`,
  ]);

  return [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
}

export default function ExportScreen() {
  const colors = useColors();
  const router = useRouter();
  const [sheets, setSheets] = useState<DaySheet[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [loading, setLoading] = useState(true);
  const [exporting, setExporting] = useState(false);

  const load = useCallback(async () => {
    const data = await getDaySheets();
    // Only show approved, not yet exported
    const exportable = data.filter(
      (s) => s.approvalStatus === "Approved" && s.payrollExportStatus !== "Exported"
    );
    setSheets(exportable);
    // Pre-select all
    setSelected(new Set(exportable.map((s) => s.id)));
  }, []);

  useEffect(() => {
    load().finally(() => setLoading(false));
  }, [load]);

  const toggleSelect = (id: string) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleAll = () => {
    if (selected.size === sheets.length) {
      setSelected(new Set());
    } else {
      setSelected(new Set(sheets.map((s) => s.id)));
    }
  };

  const handleExport = async () => {
    const toExport = sheets.filter((s) => selected.has(s.id));
    if (toExport.length === 0) {
      Alert.alert("Nothing Selected", "Please select at least one day sheet to export.");
      return;
    }

    setExporting(true);
    try {
      const csv = buildXeroCSV(toExport);
      const batchDate = new Date().toISOString().split("T")[0];
      const filename = `ASR-Payroll-${batchDate}.csv`;
      const xeroRef = `XERO-${batchDate}`;

      if (Platform.OS === "web") {
        // Web: trigger download via data URI
        const blob = new Blob([csv], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
      } else {
        // Mobile: share the CSV text
        await Share.share({
          title: filename,
          message: csv,
        });
      }

      // Write export status back to SharePoint so these sheets no longer appear
      // in the export queue on the next load.
      const ids = toExport.map((s) => s.id);
      const writeBack = await markDaySheetsExported(ids, xeroRef);

      // Remove exported sheets from the local list immediately
      setSheets((prev) => prev.filter((s) => !selected.has(s.id)));
      setSelected(new Set());

      const failNote = writeBack.failed > 0
        ? `\n(${writeBack.failed} item${writeBack.failed !== 1 ? "s" : ""} could not be marked in SharePoint — they will reappear next time.)`
        : "";
      Alert.alert(
        "Exported",
        `${toExport.length} day sheet${toExport.length !== 1 ? "s" : ""} exported as ${filename}.${failNote}`
      );
    } catch {
      Alert.alert("Export Failed", "Could not generate the export file.");
    } finally {
      setExporting(false);
    }
  };

  // Summary totals
  const selectedSheets = sheets.filter((s) => selected.has(s.id));
  const totalOrdinary = selectedSheets.reduce((sum, s) => sum + s.ordinaryHours, 0);
  const totalOvertime = selectedSheets.reduce((sum, s) => sum + s.overtimeHours, 0);
  const uniqueWorkers = new Set(selectedSheets.map((s) => s.workerName)).size;

  return (
    <ScreenContainer containerClassName="bg-background">
      {/* Header */}
      <View style={[styles.header, { backgroundColor: "#7F1F1F" }]}>
        <Pressable onPress={() => router.back()} style={styles.backBtn}>
          <IconSymbol name="chevron.left" size={20} color="#fff" />
        </Pressable>
        <Text style={styles.headerTitle}>Payroll Export</Text>
        <View style={{ width: 44 }} />
      </View>

      {loading ? (
        <ActivityIndicator color="#7F1F1F" style={{ marginTop: 40 }} />
      ) : (
        <>
          {/* Summary Card */}
          <View style={[styles.summaryCard, { backgroundColor: "#FFF1F1", borderColor: "#FECACA" }]}>
            <Text style={styles.summaryTitle}>Export Summary</Text>
            <View style={styles.summaryRow}>
              <View style={styles.summaryItem}>
                <Text style={styles.summaryValue}>{selected.size}</Text>
                <Text style={styles.summaryLabel}>Sheets</Text>
              </View>
              <View style={styles.summaryDivider} />
              <View style={styles.summaryItem}>
                <Text style={styles.summaryValue}>{uniqueWorkers}</Text>
                <Text style={styles.summaryLabel}>Workers</Text>
              </View>
              <View style={styles.summaryDivider} />
              <View style={styles.summaryItem}>
                <Text style={styles.summaryValue}>{totalOrdinary.toFixed(1)}</Text>
                <Text style={styles.summaryLabel}>Ord. Hrs</Text>
              </View>
              <View style={styles.summaryDivider} />
              <View style={styles.summaryItem}>
                <Text style={[styles.summaryValue, totalOvertime > 0 ? { color: "#D97706" } : {}]}>
                  {totalOvertime.toFixed(1)}
                </Text>
                <Text style={styles.summaryLabel}>OT Hrs</Text>
              </View>
            </View>
          </View>

          {/* Select all toggle */}
          {sheets.length > 0 && (
            <Pressable
              onPress={toggleAll}
              style={({ pressed }) => [
                styles.selectAllBtn,
                { borderColor: colors.border, backgroundColor: colors.surface, opacity: pressed ? 0.8 : 1 },
              ]}
            >
              <IconSymbol
                name={selected.size === sheets.length ? "checkmark.square.fill" : "square"}
                size={20}
                color="#7F1F1F"
              />
              <Text style={[styles.selectAllText, { color: colors.foreground }]}>
                {selected.size === sheets.length ? "Deselect All" : "Select All"}
              </Text>
              <Text style={[styles.selectAllCount, { color: colors.muted }]}>
                {selected.size} of {sheets.length} selected
              </Text>
            </Pressable>
          )}

          <FlatList
            data={sheets}
            keyExtractor={(item) => item.id}
            contentContainerStyle={{ padding: 16, gap: 10, paddingBottom: 120 }}
            ListEmptyComponent={
              <View style={styles.emptyContainer}>
                <IconSymbol name="checkmark.circle.fill" size={48} color="#22C55E" />
                <Text style={[styles.emptyTitle, { color: colors.foreground }]}>All caught up!</Text>
                <Text style={[styles.emptySubtitle, { color: colors.muted }]}>
                  No approved day sheets pending export.
                </Text>
              </View>
            }
            renderItem={({ item }) => {
              const isSelected = selected.has(item.id);
              return (
                <Pressable
                  onPress={() => toggleSelect(item.id)}
                  style={({ pressed }) => [
                    styles.sheetCard,
                    {
                      backgroundColor: isSelected ? "#F0F9FF" : colors.surface,
                      borderColor: isSelected ? "#7F1F1F" : colors.border,
                      opacity: pressed ? 0.85 : 1,
                    },
                  ]}
                >
                  <View style={styles.sheetCardRow}>
                    <IconSymbol
                      name={isSelected ? "checkmark.square.fill" : "square"}
                      size={22}
                      color={isSelected ? "#7F1F1F" : colors.muted}
                    />
                    <View style={{ flex: 1, marginLeft: 12 }}>
                      <Text style={[styles.workerName, { color: colors.foreground }]}>{item.workerName}</Text>
                      <Text style={[styles.sheetMeta, { color: colors.muted }]}>
                        {item.date} · {item.jobCode}
                      </Text>
                    </View>
                    <View style={styles.hoursCol}>
                      <Text style={[styles.hoursMain, { color: "#7F1F1F" }]}>
                        {(item.ordinaryHours + item.overtimeHours).toFixed(2)} hrs
                      </Text>
                      {item.overtimeHours > 0 && (
                        <Text style={styles.hoursOT}>+{item.overtimeHours.toFixed(2)} OT</Text>
                      )}
                    </View>
                  </View>
                  {item.allowances ? (
                    <Text style={[styles.allowanceText, { color: colors.muted }]}>
                      Allowances: {item.allowances}
                    </Text>
                  ) : null}
                </Pressable>
              );
            }}
          />

          {/* Export Button */}
          {sheets.length > 0 && (
            <View style={[styles.exportFooter, { backgroundColor: colors.background, borderTopColor: colors.border }]}>
              <Pressable
                onPress={handleExport}
                disabled={exporting || selected.size === 0}
                style={({ pressed }) => [
                  styles.exportBtn,
                  {
                    backgroundColor: selected.size === 0 ? colors.muted : "#7F1F1F",
                    opacity: pressed || exporting ? 0.8 : 1,
                  },
                ]}
              >
                {exporting ? (
                  <ActivityIndicator color="#fff" />
                ) : (
                  <>
                    <IconSymbol name="arrow.up.doc.fill" size={18} color="#fff" />
                    <Text style={styles.exportBtnText}>
                      Export {selected.size} Sheet{selected.size !== 1 ? "s" : ""} to Xero CSV
                    </Text>
                  </>
                )}
              </Pressable>
            </View>
          )}
        </>
      )}
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  header: {
    flexDirection: "row",
    alignItems: "center",
    paddingHorizontal: 16,
    paddingVertical: 14,
    justifyContent: "space-between",
  },
  backBtn: {
    width: 44,
    height: 44,
    alignItems: "center",
    justifyContent: "center",
  },
  headerTitle: {
    fontSize: 18,
    fontFamily: "Montserrat_700Bold",
    color: "#FFFFFF",
  },
  summaryCard: {
    margin: 16,
    marginBottom: 8,
    borderRadius: 14,
    padding: 16,
    borderWidth: 1,
  },
  summaryTitle: {
    fontSize: 13,
    fontFamily: "Montserrat_600SemiBold",
    color: "#7F1F1F",
    marginBottom: 12,
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  summaryRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  summaryItem: {
    flex: 1,
    alignItems: "center",
  },
  summaryValue: {
    fontSize: 22,
    fontFamily: "Montserrat_700Bold",
    color: "#7F1F1F",
  },
  summaryLabel: {
    fontSize: 11,
    fontFamily: "Montserrat_400Regular",
    color: "#EF4444",
    marginTop: 2,
  },
  summaryDivider: {
    width: 1,
    height: 36,
    backgroundColor: "#FECACA",
  },
  selectAllBtn: {
    flexDirection: "row",
    alignItems: "center",
    gap: 10,
    marginHorizontal: 16,
    marginBottom: 4,
    paddingHorizontal: 14,
    paddingVertical: 10,
    borderRadius: 10,
    borderWidth: 1,
  },
  selectAllText: {
    fontSize: 14,
    fontFamily: "Montserrat_500Medium",
    flex: 1,
  },
  selectAllCount: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
  },
  sheetCard: {
    borderRadius: 12,
    padding: 14,
    borderWidth: 1.5,
    gap: 6,
  },
  sheetCardRow: {
    flexDirection: "row",
    alignItems: "center",
  },
  workerName: {
    fontSize: 15,
    fontFamily: "Montserrat_600SemiBold",
  },
  sheetMeta: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    marginTop: 2,
  },
  hoursCol: {
    alignItems: "flex-end",
  },
  hoursMain: {
    fontSize: 15,
    fontFamily: "Montserrat_700Bold",
  },
  hoursOT: {
    fontSize: 11,
    fontFamily: "Montserrat_500Medium",
    color: "#D97706",
    marginTop: 2,
  },
  allowanceText: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    marginLeft: 34,
  },
  emptyContainer: {
    alignItems: "center",
    paddingTop: 60,
    gap: 12,
  },
  emptyTitle: {
    fontSize: 18,
    fontFamily: "Montserrat_700Bold",
  },
  emptySubtitle: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    textAlign: "center",
  },
  exportFooter: {
    position: "absolute",
    bottom: 0,
    left: 0,
    right: 0,
    padding: 16,
    borderTopWidth: 1,
  },
  exportBtn: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "center",
    gap: 10,
    paddingVertical: 16,
    borderRadius: 14,
  },
  exportBtnText: {
    color: "#fff",
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
  },
});
