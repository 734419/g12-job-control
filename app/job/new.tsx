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
import { useRouter } from "expo-router";
import { useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { createJob } from "@/lib/api/sharepoint";

const JOB_TYPES = ["Civil", "Structural", "Earthworks", "Demolition", "Fitout", "Other"];
const PRIORITIES = ["High", "Medium", "Low"];

export default function NewJobScreen() {
  const colors = useColors();
  const router = useRouter();

  const [jobCode, setJobCode] = useState("");
  const [jobName, setJobName] = useState("");
  const [client, setClient] = useState("");
  const [siteAddress, setSiteAddress] = useState("");
  const [startDate, setStartDate] = useState(new Date().toISOString().split("T")[0]);
  const [completionDate, setCompletionDate] = useState("");
  const [contractValue, setContractValue] = useState("");
  const [projectManager, setProjectManager] = useState("");
  const [superintendent, setSuperintendent] = useState("");
  const [jobType, setJobType] = useState("Civil");
  const [priority, setPriority] = useState("Medium");
  const [description, setDescription] = useState("");
  const [saving, setSaving] = useState(false);

  const handleSubmit = async () => {
    if (!jobCode.trim()) {
      Alert.alert("Validation", "Job Code is required.");
      return;
    }
    if (!jobName.trim()) {
      Alert.alert("Validation", "Job Name is required.");
      return;
    }
    if (!client.trim()) {
      Alert.alert("Validation", "Client is required.");
      return;
    }

    setSaving(true);
    try {
      await createJob({
        jobNumber: jobCode.trim(),
        jobCode: jobCode.trim(),
        jobName: jobName.trim(),
        client: client.trim(),
        siteAddress: siteAddress.trim(),
        status: "Active",
        startDate,
        completionDate,
        contractValue: contractValue.trim(),
        projectManager: projectManager.trim(),
        superintendent: superintendent.trim(),
        jobType,
        description: description.trim(),
        priority,
      });
      Alert.alert("Job Created", `${jobCode} — ${jobName} has been added.`, [
        { text: "OK", onPress: () => router.back() },
      ]);
    } catch {
      Alert.alert("Error", "Could not create job. Check your connection and try again.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      {/* Header */}
      <View style={[styles.header, { backgroundColor: "#7F1F1F" }]}>
        <Pressable onPress={() => router.back()} style={styles.backBtn}>
          <IconSymbol name="chevron.left" size={20} color="#fff" />
        </Pressable>
        <Text style={styles.headerTitle}>New Job</Text>
        <View style={{ width: 44 }} />
      </View>

      <ScrollView contentContainerStyle={{ padding: 16, gap: 14, paddingBottom: 40 }} showsVerticalScrollIndicator={false}>

        {/* Job Code & Name */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Job Details</Text>

        <Field label="Job Code *" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={jobCode}
            onChangeText={setJobCode}
            placeholder="e.g. G12-025"
            placeholderTextColor={colors.muted}
            autoCapitalize="characters"
            returnKeyType="next"
          />
        </Field>

        <Field label="Job Name *" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={jobName}
            onChangeText={setJobName}
            placeholder="e.g. Kellyville Ridge Retaining Wall"
            placeholderTextColor={colors.muted}
            returnKeyType="next"
          />
        </Field>

        <Field label="Client *" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={client}
            onChangeText={setClient}
            placeholder="Client or developer name"
            placeholderTextColor={colors.muted}
            returnKeyType="next"
          />
        </Field>

        <Field label="Site Address" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={siteAddress}
            onChangeText={setSiteAddress}
            placeholder="Full site address"
            placeholderTextColor={colors.muted}
            returnKeyType="next"
          />
        </Field>

        {/* Dates */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Schedule</Text>

        <View style={styles.row}>
          <Field label="Start Date" colors={colors} style={{ flex: 1 }}>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
              value={startDate}
              onChangeText={setStartDate}
              placeholder="YYYY-MM-DD"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </Field>
          <Field label="Completion Date" colors={colors} style={{ flex: 1 }}>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
              value={completionDate}
              onChangeText={setCompletionDate}
              placeholder="YYYY-MM-DD"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </Field>
        </View>

        {/* Financials */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Contract</Text>

        <Field label="Contract Value ($)" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={contractValue}
            onChangeText={setContractValue}
            placeholder="e.g. 450000"
            placeholderTextColor={colors.muted}
            keyboardType="numeric"
            returnKeyType="done"
          />
        </Field>

        {/* Team */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Team</Text>

        <Field label="Project Manager" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={projectManager}
            onChangeText={setProjectManager}
            placeholder="Full name"
            placeholderTextColor={colors.muted}
            returnKeyType="next"
          />
        </Field>

        <Field label="Superintendent" colors={colors}>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={superintendent}
            onChangeText={setSuperintendent}
            placeholder="Full name"
            placeholderTextColor={colors.muted}
            returnKeyType="done"
          />
        </Field>

        {/* Job Type */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Classification</Text>

        <Field label="Job Type" colors={colors}>
          <View style={styles.chipRow}>
            {JOB_TYPES.map((t) => (
              <Pressable
                key={t}
                onPress={() => setJobType(t)}
                style={[
                  styles.chip,
                  jobType === t
                    ? { backgroundColor: "#7F1F1F" }
                    : { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1 },
                ]}
              >
                <Text style={[styles.chipText, { color: jobType === t ? "#fff" : colors.foreground }]}>{t}</Text>
              </Pressable>
            ))}
          </View>
        </Field>

        <Field label="Priority" colors={colors}>
          <View style={styles.chipRow}>
            {PRIORITIES.map((p) => {
              const activeColor = p === "High" ? "#DC2626" : p === "Medium" ? "#D97706" : "#16A34A";
              return (
                <Pressable
                  key={p}
                  onPress={() => setPriority(p)}
                  style={[
                    styles.chip,
                    priority === p
                      ? { backgroundColor: activeColor }
                      : { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1 },
                  ]}
                >
                  <Text style={[styles.chipText, { color: priority === p ? "#fff" : colors.foreground }]}>{p}</Text>
                </Pressable>
              );
            })}
          </View>
        </Field>

        {/* Description */}
        <Field label="Description / Scope" colors={colors}>
          <TextInput
            style={[styles.notesInput, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={description}
            onChangeText={setDescription}
            placeholder="Brief description of work scope…"
            placeholderTextColor={colors.muted}
            multiline
            numberOfLines={4}
            textAlignVertical="top"
          />
        </Field>

        {/* Submit */}
        <Pressable
          onPress={handleSubmit}
          disabled={saving}
          style={({ pressed }) => [
            styles.submitBtn,
            { backgroundColor: "#7F1F1F", opacity: pressed || saving ? 0.8 : 1 },
          ]}
        >
          {saving ? (
            <View style={{ flexDirection: "row", alignItems: "center", gap: 8 }}>
              <ActivityIndicator color="#fff" />
              <Text style={styles.submitBtnText}>Creating Job…</Text>
            </View>
          ) : (
            <Text style={styles.submitBtnText}>Create Job</Text>
          )}
        </Pressable>
      </ScrollView>
    </ScreenContainer>
  );
}

function Field({ label, children, colors, style }: { label: string; children: React.ReactNode; colors: any; style?: any }) {
  return (
    <View style={[{ gap: 6 }, style]}>
      <Text style={[styles.fieldLabel, { color: colors.muted }]}>{label}</Text>
      {children}
    </View>
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
  sectionTitle: {
    fontSize: 15,
    fontFamily: "Montserrat_700Bold",
    marginTop: 4,
  },
  fieldLabel: {
    fontSize: 11,
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
  row: {
    flexDirection: "row",
    gap: 12,
  },
  chipRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  chip: {
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 20,
  },
  chipText: {
    fontSize: 13,
    fontFamily: "Montserrat_500Medium",
  },
  notesInput: {
    borderWidth: 1,
    borderRadius: 10,
    paddingHorizontal: 14,
    paddingVertical: 12,
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    minHeight: 100,
  },
  submitBtn: {
    paddingVertical: 16,
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
