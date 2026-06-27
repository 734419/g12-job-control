import {
  View,
  Text,
  ScrollView,
  TextInput,
  Pressable,
  StyleSheet,
  Alert,
  ActivityIndicator,
} from "react-native";
import { useRouter, useLocalSearchParams } from "expo-router";
import { useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";
import { createDaySheet } from "@/lib/api/sharepoint";
import { enqueue } from "@/lib/offline/queue";

const ALLOWANCES = [
  "Travel",
  "Meal",
  "Tool",
  "Height",
  "Confined Space",
  "First Aid",
  "Site",
];

function calculateHours(start: string, finish: string, breakMins: number): { ordinary: number; overtime: number } {
  const parseTime = (t: string) => {
    const [h, m] = t.split(":").map(Number);
    return isNaN(h) || isNaN(m) ? null : h * 60 + m;
  };
  const s = parseTime(start);
  const f = parseTime(finish);
  if (s === null || f === null) return { ordinary: 0, overtime: 0 };
  const totalMins = Math.max(0, f - s - breakMins);
  const totalHours = totalMins / 60;
  const ordinary = Math.min(totalHours, 8);
  const overtime = Math.max(0, totalHours - 8);
  return {
    ordinary: Math.round(ordinary * 100) / 100,
    overtime: Math.round(overtime * 100) / 100,
  };
}

export default function NewDaySheetScreen() {
  const colors = useColors();
  const router = useRouter();
  const { user } = useAuth();
  const params = useLocalSearchParams<{ jobCode?: string }>();

  const [jobCode, setJobCode] = useState(params.jobCode ?? "");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);
  const [startTime, setStartTime] = useState("07:00");
  const [finishTime, setFinishTime] = useState("15:30");
  const [breakMins, setBreakMins] = useState("30");
  const [selectedAllowances, setSelectedAllowances] = useState<string[]>([]);
  const [notes, setNotes] = useState("");
  const [saving, setSaving] = useState(false);

  const { ordinary, overtime } = calculateHours(startTime, finishTime, Number(breakMins) || 0);

  const toggleAllowance = (a: string) => {
    setSelectedAllowances((prev) =>
      prev.includes(a) ? prev.filter((x) => x !== a) : [...prev, a]
    );
  };

  const handleSubmit = async () => {
    if (!jobCode.trim()) {
      Alert.alert("Validation", "Please enter a Job Code.");
      return;
    }
    setSaving(true);
    const sheet = {
      jobCode: jobCode.trim(),
      workerName: user?.displayName ?? "Unknown",
      workerEmail: user?.mail ?? "",
      date,
      startTime,
      finishTime,
      breakMinutes: Number(breakMins) || 0,
      ordinaryHours: ordinary,
      overtimeHours: overtime,
      allowances: selectedAllowances,
      notes,
      approvalStatus: "Pending" as const,
      payrollExportStatus: "Not Exported" as const,
    };

    try {
      await createDaySheet(sheet);
      Alert.alert("Submitted", "Day sheet saved to SharePoint.", [
        { text: "OK", onPress: () => router.back() },
      ]);
    } catch {
      // Save offline
      await enqueue({ type: "CREATE_DAYSHEET", payload: sheet });
      Alert.alert(
        "Saved Offline",
        "No connection. Day sheet queued and will sync when online.",
        [{ text: "OK", onPress: () => router.back() }]
      );
    } finally {
      setSaving(false);
    }
  };

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView contentContainerStyle={{ padding: 16, gap: 16, paddingBottom: 40 }} showsVerticalScrollIndicator={false}>
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Job Details</Text>

        <FormField label="Job Code *" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
            value={jobCode}
            onChangeText={setJobCode}
            placeholder="e.g. G12-001"
            placeholderTextColor={colors.muted}
            autoCapitalize="characters"
            returnKeyType="done"
          />
        </FormField>

        <FormField label="Date" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
            value={date}
            onChangeText={setDate}
            placeholder="YYYY-MM-DD"
            placeholderTextColor={colors.muted}
            returnKeyType="done"
          />
        </FormField>

        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Hours</Text>

        <View style={styles.row}>
          <FormField label="Start Time" colors={colors} style={{ flex: 1 }}>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
              value={startTime}
              onChangeText={setStartTime}
              placeholder="07:00"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </FormField>
          <FormField label="Finish Time" colors={colors} style={{ flex: 1 }}>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
              value={finishTime}
              onChangeText={setFinishTime}
              placeholder="15:30"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </FormField>
        </View>

        <FormField label="Break (minutes)" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
            value={breakMins}
            onChangeText={setBreakMins}
            placeholder="30"
            placeholderTextColor={colors.muted}
            keyboardType="numeric"
            returnKeyType="done"
          />
        </FormField>

        {/* Calculated hours */}
        <View style={[styles.hoursCard, { backgroundColor: "#EFF6FF" }]}>
          <View style={styles.hoursRow}>
            <Text style={styles.hoursLabel}>Ordinary Hours</Text>
            <Text style={styles.hoursValue}>{ordinary.toFixed(2)} hrs</Text>
          </View>
          <View style={styles.hoursRow}>
            <Text style={styles.hoursLabel}>Overtime Hours</Text>
            <Text style={[styles.hoursValue, overtime > 0 && { color: "#D97706" }]}>
              {overtime.toFixed(2)} hrs
            </Text>
          </View>
        </View>

        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Allowances</Text>
        <View style={styles.allowancesGrid}>
          {ALLOWANCES.map((a) => (
            <Pressable
              key={a}
              onPress={() => toggleAllowance(a)}
              style={[
                styles.allowanceChip,
                selectedAllowances.includes(a)
                  ? { backgroundColor: "#1B2A4A" }
                  : { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1 },
              ]}
            >
              <Text
                style={[
                  styles.allowanceText,
                  { color: selectedAllowances.includes(a) ? "#fff" : colors.foreground },
                ]}
              >
                {a}
              </Text>
            </Pressable>
          ))}
        </View>

        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Notes</Text>
        <TextInput
          style={[styles.notesInput, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
          value={notes}
          onChangeText={setNotes}
          placeholder="Any additional notes…"
          placeholderTextColor={colors.muted}
          multiline
          numberOfLines={4}
          textAlignVertical="top"
        />

        <Pressable
          onPress={handleSubmit}
          disabled={saving}
          style={({ pressed }) => [
            styles.submitBtn,
            { backgroundColor: "#1B2A4A", opacity: pressed || saving ? 0.8 : 1 },
          ]}
        >
          {saving ? (
            <ActivityIndicator color="#fff" />
          ) : (
            <Text style={styles.submitBtnText}>Submit Day Sheet</Text>
          )}
        </Pressable>
      </ScrollView>
    </ScreenContainer>
  );
}

function FormField({
  label,
  children,
  colors,
  style,
}: {
  label: string;
  children: React.ReactNode;
  colors: any;
  style?: any;
}) {
  return (
    <View style={[{ gap: 6 }, style]}>
      <Text style={[styles.fieldLabel, { color: colors.muted }]}>{label}</Text>
      {children}
    </View>
  );
}

const styles = StyleSheet.create({
  sectionTitle: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    marginTop: 4,
  },
  row: {
    flexDirection: "row",
    gap: 12,
  },
  fieldLabel: {
    fontSize: 12,
    fontFamily: "Montserrat_500Medium",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  input: {
    borderWidth: 1,
    borderRadius: 10,
    paddingHorizontal: 14,
    paddingVertical: 12,
    fontSize: 15,
    fontFamily: "Montserrat_400Regular",
  },
  hoursCard: {
    borderRadius: 12,
    padding: 16,
    gap: 8,
  },
  hoursRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  hoursLabel: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    color: "#1E40AF",
  },
  hoursValue: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    color: "#1E40AF",
  },
  allowancesGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  allowanceChip: {
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 20,
  },
  allowanceText: {
    fontSize: 13,
    fontFamily: "Montserrat_500Medium",
  },
  notesInput: {
    borderWidth: 1,
    borderRadius: 10,
    padding: 14,
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    minHeight: 100,
  },
  submitBtn: {
    padding: 16,
    borderRadius: 14,
    alignItems: "center",
    marginTop: 8,
  },
  submitBtnText: {
    color: "#fff",
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
  },
});
