import {
  View,
  Text,
  ScrollView,
  TextInput,
  Pressable,
  StyleSheet,
  Alert,
  ActivityIndicator,
  Image,
  Platform,
} from "react-native";
import { useRouter, useLocalSearchParams } from "expo-router";
import { useState } from "react";
import * as ImagePicker from "expo-image-picker";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";
import { createDaySheet, uploadSitePhoto } from "@/lib/api/sharepoint";
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
  const [photos, setPhotos] = useState<Array<{ uri: string; name: string }>>([]);
  const [saving, setSaving] = useState(false);
  const [uploadingPhoto, setUploadingPhoto] = useState(false);

  const { ordinary, overtime } = calculateHours(startTime, finishTime, Number(breakMins) || 0);

  const toggleAllowance = (a: string) => {
    setSelectedAllowances((prev) =>
      prev.includes(a) ? prev.filter((x) => x !== a) : [...prev, a]
    );
  };

  const takePhoto = async () => {
    if (Platform.OS === "web") {
      Alert.alert("Camera", "Camera capture is only available on iOS and Android.");
      return;
    }
    const { status } = await ImagePicker.requestCameraPermissionsAsync();
    if (status !== "granted") {
      Alert.alert("Permission Required", "Camera access is needed to take site photos.");
      return;
    }
    const result = await ImagePicker.launchCameraAsync({
      allowsEditing: false,
      quality: 0.7,
      mediaTypes: ImagePicker.MediaTypeOptions.Images,
    });
    if (!result.canceled && result.assets[0]) {
      const asset = result.assets[0];
      const name = `site-photo-${Date.now()}.jpg`;
      setPhotos((prev) => [...prev, { uri: asset.uri, name }]);
    }
  };

  const pickFromLibrary = async () => {
    const result = await ImagePicker.launchImageLibraryAsync({
      allowsEditing: false,
      quality: 0.7,
      mediaTypes: ImagePicker.MediaTypeOptions.Images,
      allowsMultipleSelection: true,
    });
    if (!result.canceled) {
      const newPhotos = result.assets.map((a, i) => ({
        uri: a.uri,
        name: `site-photo-${Date.now()}-${i}.jpg`,
      }));
      setPhotos((prev) => [...prev, ...newPhotos]);
    }
  };

  const removePhoto = (index: number) => {
    setPhotos((prev) => prev.filter((_, i) => i !== index));
  };

  const handleSubmit = async () => {
    if (!jobCode.trim()) {
      Alert.alert("Validation", "Please enter a Job Code.");
      return;
    }
    setSaving(true);
    const sheet = {
      jobCode: jobCode.trim(),
      jobName: "",
      workerName: user?.displayName ?? "Unknown",
      workerEmail: user?.mail ?? "",
      date,
      startTime,
      finishTime,
      breakMinutes: Number(breakMins) || 0,
      ordinaryHours: ordinary,
      overtimeHours: overtime,
      allowances: selectedAllowances.join("; "),
      notes,
      approvalStatus: "Pending" as const,
      approvedBy: "",
      approvedDate: "",
      payrollExportStatus: "Not Exported" as const,
      xeroReference: "",
      site: "",
      trade: "",
    };

    try {
      await createDaySheet(sheet);

      // Upload photos to SharePoint Site Photos library
      if (photos.length > 0) {
        setUploadingPhoto(true);
        for (const photo of photos) {
          try {
            await uploadSitePhoto(
              photo.uri,
              jobCode.trim(),
              user?.displayName ?? "Unknown",
              date
            );
          } catch {
            // Non-fatal — photos may fail individually
          }
        }
        setUploadingPhoto(false);
      }

      Alert.alert(
        "Submitted",
        `Day sheet saved to SharePoint${photos.length > 0 ? ` with ${photos.length} photo(s)` : ""}.`,
        [{ text: "OK", onPress: () => router.back() }]
      );
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
      setUploadingPhoto(false);
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
            <Text style={[styles.hoursValue, overtime > 0 ? { color: "#D97706" } : {}]}>
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

        {/* Site Photos */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Site Photos</Text>
        <Text style={[styles.photoSubtitle, { color: colors.muted }]}>
          Attach site photos or SWMS documents. Photos are uploaded to the SharePoint Site Photos library.
        </Text>

        {/* Photo thumbnails */}
        {photos.length > 0 && (
          <View style={styles.photoGrid}>
            {photos.map((photo, index) => (
              <View key={index} style={styles.photoThumb}>
                <Image source={{ uri: photo.uri }} style={styles.photoImage} />
                <Pressable
                  onPress={() => removePhoto(index)}
                  style={styles.photoRemoveBtn}
                >
                  <Text style={styles.photoRemoveText}>✕</Text>
                </Pressable>
              </View>
            ))}
          </View>
        )}

        {/* Photo action buttons */}
        <View style={styles.photoActions}>
          {Platform.OS !== "web" && (
            <Pressable
              onPress={takePhoto}
              style={({ pressed }) => [
                styles.photoBtn,
                { backgroundColor: "#1B2A4A", opacity: pressed ? 0.8 : 1 },
              ]}
            >
              <IconSymbol name="camera.fill" size={18} color="#fff" />
              <Text style={styles.photoBtnText}>Take Photo</Text>
            </Pressable>
          )}
          <Pressable
            onPress={pickFromLibrary}
            style={({ pressed }) => [
              styles.photoBtn,
              { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1, opacity: pressed ? 0.8 : 1 },
            ]}
          >
            <IconSymbol name="photo.fill" size={18} color={colors.foreground} />
            <Text style={[styles.photoBtnText, { color: colors.foreground }]}>Choose from Library</Text>
          </Pressable>
        </View>

        <Pressable
          onPress={handleSubmit}
          disabled={saving || uploadingPhoto}
          style={({ pressed }) => [
            styles.submitBtn,
            { backgroundColor: "#1B2A4A", opacity: pressed || saving || uploadingPhoto ? 0.8 : 1 },
          ]}
        >
          {saving || uploadingPhoto ? (
            <View style={{ flexDirection: "row", alignItems: "center", gap: 8 }}>
              <ActivityIndicator color="#fff" />
              <Text style={styles.submitBtnText}>
                {uploadingPhoto ? "Uploading photos…" : "Saving…"}
              </Text>
            </View>
          ) : (
            <Text style={styles.submitBtnText}>
              Submit Day Sheet{photos.length > 0 ? ` + ${photos.length} Photo${photos.length > 1 ? "s" : ""}` : ""}
            </Text>
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
  photoSubtitle: {
    fontSize: 13,
    fontFamily: "Montserrat_400Regular",
    lineHeight: 18,
    marginTop: -8,
  },
  photoGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 10,
  },
  photoThumb: {
    width: 90,
    height: 90,
    borderRadius: 10,
    overflow: "hidden",
    position: "relative",
  },
  photoImage: {
    width: "100%",
    height: "100%",
    resizeMode: "cover",
  },
  photoRemoveBtn: {
    position: "absolute",
    top: 4,
    right: 4,
    backgroundColor: "rgba(0,0,0,0.6)",
    borderRadius: 10,
    width: 20,
    height: 20,
    alignItems: "center",
    justifyContent: "center",
  },
  photoRemoveText: {
    color: "#fff",
    fontSize: 11,
    fontFamily: "Montserrat_700Bold",
  },
  photoActions: {
    flexDirection: "row",
    gap: 10,
    flexWrap: "wrap",
  },
  photoBtn: {
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    paddingHorizontal: 16,
    paddingVertical: 12,
    borderRadius: 12,
  },
  photoBtnText: {
    fontSize: 14,
    fontFamily: "Montserrat_600SemiBold",
    color: "#fff",
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
