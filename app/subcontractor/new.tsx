import { View, Text, ScrollView, TextInput, Pressable, StyleSheet, Alert, ActivityIndicator } from "react-native";
import { useRouter } from "expo-router";
import { useState } from "react";
import { ScreenContainer } from "@/components/screen-container";
import { useColors } from "@/hooks/use-colors";
import { createSubcontractor } from "@/lib/api/sharepoint";

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
  });

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
        licenceNumber: form.licenceNumber ?? "",
        inductionStatus: "Not Started",
        swmsStatus: "Not Submitted",
        prequalificationStatus: "Not Started",
        mobilisationApproved: false,
        activeJobCodes: [],
        notes: "",
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

  return (
    <ScreenContainer containerClassName="bg-background" edges={["left", "right", "bottom"]}>
      <ScrollView contentContainerStyle={{ padding: 16, gap: 14, paddingBottom: 40 }}>
        {[
          { key: "companyName", label: "Company Name *", placeholder: "e.g. Smith Electrical Pty Ltd" },
          { key: "abn", label: "ABN", placeholder: "12 345 678 901" },
          { key: "trade", label: "Trade", placeholder: "e.g. Electrician" },
          { key: "contactName", label: "Contact Name", placeholder: "Full name" },
          { key: "contactPhone", label: "Contact Phone", placeholder: "+61 4xx xxx xxx" },
          { key: "contactEmail", label: "Contact Email", placeholder: "email@example.com" },
          { key: "insuranceExpiry", label: "Insurance Expiry (YYYY-MM-DD)", placeholder: "2026-12-31" },
          { key: "licenceExpiry", label: "Licence Expiry (YYYY-MM-DD)", placeholder: "2026-12-31" },
        ].map(({ key, label, placeholder }) => (
          <View key={key} style={{ gap: 6 }}>
            <Text style={[styles.label, { color: colors.muted }]}>{label}</Text>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border }]}
              value={form[key as keyof typeof form]}
              onChangeText={update(key as keyof typeof form)}
              placeholder={placeholder}
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </View>
        ))}

        <Pressable
          onPress={handleSubmit}
          disabled={saving}
          style={({ pressed }) => [styles.submitBtn, { opacity: pressed || saving ? 0.8 : 1 }]}
        >
          {saving ? <ActivityIndicator color="#fff" /> : <Text style={styles.submitBtnText}>Add Subcontractor</Text>}
        </Pressable>
      </ScrollView>
    </ScreenContainer>
  );
}

const styles = StyleSheet.create({
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
