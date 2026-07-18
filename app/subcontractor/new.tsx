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
import { useColors } from "@/hooks/use-colors";
import { createSubcontractor } from "@/lib/api/sharepoint";

// ─── Option picker helpers ────────────────────────────────────────────────────

const PREQUAL_OPTIONS = ["Not Started", "Pending", "Approved", "Rejected"] as const;
const INDUCTION_OPTIONS = ["Not Started", "Pending", "Complete"] as const;
const SWMS_OPTIONS = ["Not Submitted", "Pending", "Approved"] as const;

type PrequalStatus = (typeof PREQUAL_OPTIONS)[number];
type InductionStatus = (typeof INDUCTION_OPTIONS)[number];
type SwmsStatus = (typeof SWMS_OPTIONS)[number];

function OptionPicker<T extends string>({
  label,
  options,
  value,
  onChange,
}: {
  label: string;
  options: readonly T[];
  value: T;
  onChange: (v: T) => void;
}) {
  const colors = useColors();
  return (
    <View style={{ gap: 6 }}>
      <Text style={[styles.label, { color: colors.muted }]}>{label}</Text>
      <View style={styles.optionRow}>
        {options.map((opt) => {
          const active = opt === value;
          return (
            <Pressable
              key={opt}
              onPress={() => onChange(opt)}
              style={({ pressed }) => [
                styles.optionChip,
                {
                  backgroundColor: active ? "#1B2A4A" : colors.surface,
                  borderColor: active ? "#1B2A4A" : colors.border,
                  opacity: pressed ? 0.8 : 1,
                },
              ]}
            >
              <Text
                style={[
                  styles.optionChipText,
                  { color: active ? "#fff" : colors.foreground },
                ]}
              >
                {opt}
              </Text>
            </Pressable>
          );
        })}
      </View>
    </View>
  );
}

// ─── Screen ───────────────────────────────────────────────────────────────────

export default function NewSubcontractorScreen() {
  const colors = useColors();
  const router = useRouter();
  const [saving, setSaving] = useState(false);

  const [form, setForm] = useState({
    companyName: "",
    abn: "",
    trade: "",
    contactName: "",
    contactPhone: "",
    contactEmail: "",
    insuranceExpiry: "",
    licenceExpiry: "",
    licenceNumber: "",
    notes: "",
  });

  const [prequalificationStatus, setPrequalificationStatus] =
    useState<PrequalStatus>("Not Started");
  const [inductionStatus, setInductionStatus] =
    useState<InductionStatus>("Not Started");
  const [swmsStatus, setSwmsStatus] = useState<SwmsStatus>("Not Submitted");

  const update = (key: keyof typeof form) => (value: string) =>
    setForm((prev) => ({ ...prev, [key]: value }));

  const handleSubmit = async () => {
    if (!form.companyName.trim()) {
      Alert.alert("Validation", "Company name is required.");
      return;
    }
    setSaving(true);
    try {
      await createSubcontractor({
        ...form,
        prequalificationStatus,
        inductionStatus,
        swmsStatus,
        mobilisationApproved: false,
        activeJobCodes: [],
      });
      Alert.alert("Saved", "Subcontractor added to SharePoint.", [
        { text: "OK", onPress: () => router.back() },
      ]);
    } catch {
      Alert.alert("Error", "Could not save. Check your connection.");
    } finally {
      setSaving(false);
    }
  };

  const textFields: { key: keyof typeof form; label: string; placeholder: string; keyboard?: any }[] = [
    { key: "companyName", label: "Company Name *", placeholder: "e.g. Smith Electrical Pty Ltd" },
    { key: "abn", label: "ABN", placeholder: "12 345 678 901", keyboard: "numeric" },
    { key: "trade", label: "Trade", placeholder: "e.g. Electrician" },
    { key: "contactName", label: "Contact Name", placeholder: "Full name" },
    { key: "contactPhone", label: "Contact Phone", placeholder: "+61 4xx xxx xxx", keyboard: "phone-pad" },
    { key: "contactEmail", label: "Contact Email", placeholder: "email@example.com", keyboard: "email-address" },
    { key: "insuranceExpiry", label: "Insurance Expiry (YYYY-MM-DD)", placeholder: "2026-12-31" },
    { key: "licenceExpiry", label: "Licence Expiry (YYYY-MM-DD)", placeholder: "2026-12-31" },
    { key: "licenceNumber", label: "Licence Number", placeholder: "e.g. CL-48291" },
  ];

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView contentContainerStyle={{ padding: 16, gap: 14, paddingBottom: 40 }}>

        {/* Section: Company & Contact */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Company & Contact</Text>
        {textFields.map(({ key, label, placeholder, keyboard }) => (
          <View key={key} style={{ gap: 6 }}>
            <Text style={[styles.label, { color: colors.muted }]}>{label}</Text>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
              value={form[key]}
              onChangeText={update(key)}
              placeholder={placeholder}
              placeholderTextColor={colors.muted}
              keyboardType={keyboard ?? "default"}
              returnKeyType="done"
              autoCapitalize={key === "contactEmail" ? "none" : "sentences"}
            />
          </View>
        ))}

        {/* Section: Compliance Status */}
        <Text style={[styles.sectionTitle, { color: colors.foreground, marginTop: 8 }]}>Compliance Status</Text>

        <OptionPicker
          label="Prequalification Status"
          options={PREQUAL_OPTIONS}
          value={prequalificationStatus}
          onChange={setPrequalificationStatus}
        />

        <OptionPicker
          label="Induction Status"
          options={INDUCTION_OPTIONS}
          value={inductionStatus}
          onChange={setInductionStatus}
        />

        <OptionPicker
          label="SWMS Status"
          options={SWMS_OPTIONS}
          value={swmsStatus}
          onChange={setSwmsStatus}
        />

        {/* Notes */}
        <View style={{ gap: 6 }}>
          <Text style={[styles.label, { color: colors.muted }]}>Notes</Text>
          <TextInput
            style={[
              styles.input,
              styles.textArea,
              { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface },
            ]}
            value={form.notes}
            onChangeText={update("notes")}
            placeholder="Any additional notes about this subcontractor…"
            placeholderTextColor={colors.muted}
            multiline
            numberOfLines={4}
            textAlignVertical="top"
          />
        </View>

        <Pressable
          onPress={handleSubmit}
          disabled={saving}
          style={({ pressed }) => [styles.submitBtn, { opacity: pressed || saving ? 0.8 : 1 }]}
        >
          {saving ? (
            <ActivityIndicator color="#fff" />
          ) : (
            <Text style={styles.submitBtnText}>Add Subcontractor</Text>
          )}
        </Pressable>
      </ScrollView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
  sectionTitle: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    marginBottom: 2,
  },
  label: {
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
  textArea: {
    minHeight: 96,
    paddingTop: 12,
  },
  optionRow: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  optionChip: {
    paddingHorizontal: 14,
    paddingVertical: 8,
    borderRadius: 20,
    borderWidth: 1.5,
  },
  optionChipText: {
    fontSize: 13,
    fontFamily: "Montserrat_500Medium",
  },
  submitBtn: {
    backgroundColor: "#1B2A4A",
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
