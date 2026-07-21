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
  Modal,
  FlatList,
} from "react-native";
import { useRouter, useLocalSearchParams } from "expo-router";
import { useState, useEffect } from "react";
import * as ImagePicker from "expo-image-picker";
import { ScreenContainer } from "@/components/screen-container";
import { IconSymbol } from "@/components/ui/icon-symbol";
import { useColors } from "@/hooks/use-colors";
import { useAuth } from "@/lib/auth/AuthContext";
import { createDaySheet, uploadSitePhoto, getJobs, getSubcontractorList, getCostCodes, type Job } from "@/lib/api/sharepoint";
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

  const [jobs, setJobs] = useState<Job[]>([]);
  const [selectedJob, setSelectedJob] = useState<Job | null>(null);
  const [jobPickerVisible, setJobPickerVisible] = useState(false);
  // Subcontractor / Crew picker
  const [subcontractors, setSubcontractors] = useState<{ id: string; label: string }[]>([]);
  const [selectedSubcontractor, setSelectedSubcontractor] = useState<{ id: string; label: string } | null>(null);
  const [subPickerVisible, setSubPickerVisible] = useState(false);
  // Cost Code picker
  const [costCodes, setCostCodes] = useState<string[]>([]);
  const [selectedCostCode, setSelectedCostCode] = useState("");
  const [costCodePickerVisible, setCostCodePickerVisible] = useState(false);
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

  useEffect(() => {
    getJobs().then((data) => {
      const active = data.filter((j) => j.status === "Active");
      setJobs(active);
      // Pre-select if jobCode param was passed
      if (params.jobCode) {
        const match = active.find((j) => j.jobCode === params.jobCode);
        if (match) setSelectedJob(match);
      }
    });
    getSubcontractorList().then(setSubcontractors);
    getCostCodes().then(setCostCodes);
  }, [params.jobCode]);

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
    if (!selectedJob) {
      Alert.alert("Validation", "Please select a Job.");
      return;
    }
    if (!selectedSubcontractor) {
      Alert.alert("Validation", "Please select a Subcontractor / Crew.");
      return;
    }
    setSaving(true);
    const sheet = {
      jobCode: selectedJob.jobCode,
      jobName: selectedJob.jobName,
      workerName: selectedSubcontractor.label,
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
      site: selectedJob.siteAddress ?? "",
      trade: selectedCostCode,
    };

    try {
      await createDaySheet(sheet);

      if (photos.length > 0) {
        setUploadingPhoto(true);
        for (const photo of photos) {
          try {
            await uploadSitePhoto(photo.uri, selectedJob.jobCode, workerName.trim(), date);
          } catch {
            // Non-fatal
          }
        }
        setUploadingPhoto(false);
      }

      Alert.alert(
        "Submitted",
        `Day sheet saved${photos.length > 0 ? ` with ${photos.length} photo(s)` : ""}.`,
        [{ text: "OK", onPress: () => router.back() }]
      );
    } catch {
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
      {/* Header */}
      <View style={[styles.header, { backgroundColor: "#7F1F1F" }]}>
        <Pressable onPress={() => router.back()} style={styles.backBtn}>
          <IconSymbol name="chevron.left" size={20} color="#fff" />
        </Pressable>
        <Text style={styles.headerTitle}>New Day Sheet</Text>
        <View style={{ width: 44 }} />
      </View>

      <ScrollView contentContainerStyle={{ padding: 16, gap: 16, paddingBottom: 40 }} showsVerticalScrollIndicator={false}>

        {/* Job Picker */}
        <View style={{ gap: 6 }}>
          <Text style={[styles.fieldLabel, { color: colors.muted }]}>Job *</Text>
          <Pressable
            onPress={() => setJobPickerVisible(true)}
            style={[styles.pickerBtn, { borderColor: selectedJob ? "#7F1F1F" : colors.border, backgroundColor: colors.surface }]}
          >
            {selectedJob ? (
              <View style={{ flex: 1 }}>
                <Text style={[styles.pickerBtnText, { color: colors.foreground }]} numberOfLines={1}>
                  {selectedJob.jobCode} — {selectedJob.jobName}
                </Text>
                <Text style={[styles.pickerBtnSub, { color: colors.muted }]} numberOfLines={1}>
                  {selectedJob.client}
                </Text>
              </View>
            ) : (
              <Text style={[styles.pickerBtnText, { color: colors.muted }]}>Select a job…</Text>
            )}
            <IconSymbol name="chevron.right" size={16} color={colors.muted} />
          </Pressable>
        </View>

        {/* Worker Name */}
        {/* Subcontractor / Crew Picker */}
        <View style={{ gap: 6 }}>
          <Text style={[styles.fieldLabel, { color: colors.muted }]}>Subcontractor / Crew *</Text>
          <Pressable
            onPress={() => setSubPickerVisible(true)}
            style={[styles.pickerBtn, { borderColor: selectedSubcontractor ? "#7F1F1F" : colors.border, backgroundColor: colors.surface }]}
          >
            {selectedSubcontractor ? (
              <Text style={[styles.pickerBtnText, { color: colors.foreground }]} numberOfLines={1}>
                {selectedSubcontractor.label}
              </Text>
            ) : (
              <Text style={[styles.pickerBtnText, { color: colors.muted }]}>Select subcontractor or crew…</Text>
            )}
            <IconSymbol name="chevron.right" size={16} color={colors.muted} />
          </Pressable>
        </View>

        {/* Date */}
        <View style={{ gap: 6 }}>
          <Text style={[styles.fieldLabel, { color: colors.muted }]}>Date</Text>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={date}
            onChangeText={setDate}
            placeholder="YYYY-MM-DD"
            placeholderTextColor={colors.muted}
            returnKeyType="done"
          />
        </View>

        {/* Hours */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Hours</Text>
        <View style={styles.row}>
          <View style={[{ flex: 1, gap: 6 }]}>
            <Text style={[styles.fieldLabel, { color: colors.muted }]}>Start Time</Text>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
              value={startTime}
              onChangeText={setStartTime}
              placeholder="07:00"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </View>
          <View style={[{ flex: 1, gap: 6 }]}>
            <Text style={[styles.fieldLabel, { color: colors.muted }]}>Finish Time</Text>
            <TextInput
              style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
              value={finishTime}
              onChangeText={setFinishTime}
              placeholder="15:30"
              placeholderTextColor={colors.muted}
              returnKeyType="done"
            />
          </View>
        </View>

        <View style={{ gap: 6 }}>
          <Text style={[styles.fieldLabel, { color: colors.muted }]}>Break (minutes)</Text>
          <TextInput
            style={[styles.input, { color: colors.foreground, borderColor: colors.border, backgroundColor: colors.surface }]}
            value={breakMins}
            onChangeText={setBreakMins}
            placeholder="30"
            placeholderTextColor={colors.muted}
            keyboardType="numeric"
            returnKeyType="done"
          />
        </View>

        {/* Calculated hours summary */}
        <View style={[styles.hoursCard, { backgroundColor: "#FFF1F1" }]}>
          <View style={styles.hoursRow}>
            <Text style={styles.hoursLabel}>Ordinary Hours</Text>
            <Text style={styles.hoursValue}>{ordinary.toFixed(2)} hrs</Text>
          </View>
          <View style={[styles.hoursRow, { marginTop: 6 }]}>
            <Text style={styles.hoursLabel}>Overtime Hours</Text>
            <Text style={[styles.hoursValue, overtime > 0 ? { color: "#D97706" } : {}]}>
              {overtime.toFixed(2)} hrs
            </Text>
          </View>
          <View style={[styles.hoursRow, { marginTop: 6, paddingTop: 8, borderTopWidth: 1, borderTopColor: "#FECACA" }]}>
            <Text style={[styles.hoursLabel, { fontFamily: "Montserrat_700Bold" }]}>Total Hours</Text>
            <Text style={[styles.hoursValue, { fontFamily: "Montserrat_700Bold" }]}>
              {(ordinary + overtime).toFixed(2)} hrs
            </Text>
          </View>
        </View>

        {/* Allowances */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Allowances</Text>
        <View style={styles.allowancesGrid}>
          {ALLOWANCES.map((a) => (
            <Pressable
              key={a}
              onPress={() => toggleAllowance(a)}
              style={[
                styles.allowanceChip,
                selectedAllowances.includes(a)
                  ? { backgroundColor: "#7F1F1F" }
                  : { backgroundColor: colors.surface, borderColor: colors.border, borderWidth: 1 },
              ]}
            >
              <Text style={[styles.allowanceText, { color: selectedAllowances.includes(a) ? "#fff" : colors.foreground }]}>
                {a}
              </Text>
            </Pressable>
          ))}
        </View>

        {/* Notes */}
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

        {/* Cost Code Picker */}
        <View style={{ gap: 6 }}>
          <Text style={[styles.fieldLabel, { color: colors.muted }]}>Cost Code</Text>
          <Pressable
            onPress={() => setCostCodePickerVisible(true)}
            style={[styles.pickerBtn, { borderColor: selectedCostCode ? "#7F1F1F" : colors.border, backgroundColor: colors.surface }]}
          >
            {selectedCostCode ? (
              <Text style={[styles.pickerBtnText, { color: colors.foreground }]} numberOfLines={1}>
                {selectedCostCode}
              </Text>
            ) : (
              <Text style={[styles.pickerBtnText, { color: colors.muted }]}>Select cost code…</Text>
            )}
            <IconSymbol name="chevron.right" size={16} color={colors.muted} />
          </Pressable>
        </View>

        {/* Site Photos */}
        <Text style={[styles.sectionTitle, { color: colors.foreground }]}>Site Photos</Text>
        <Text style={[styles.photoSubtitle, { color: colors.muted }]}>
          Attach site photos. Uploaded to the SharePoint Site Photos library.
        </Text>

        {photos.length > 0 && (
          <View style={styles.photoGrid}>
            {photos.map((photo, index) => (
              <View key={index} style={styles.photoThumb}>
                <Image source={{ uri: photo.uri }} style={styles.photoImage} />
                <Pressable onPress={() => removePhoto(index)} style={styles.photoRemoveBtn}>
                  <Text style={styles.photoRemoveText}>✕</Text>
                </Pressable>
              </View>
            ))}
          </View>
        )}

        <View style={styles.photoActions}>
          {Platform.OS !== "web" && (
            <Pressable
              onPress={takePhoto}
              style={({ pressed }) => [styles.photoBtn, { backgroundColor: "#7F1F1F", opacity: pressed ? 0.8 : 1 }]}
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

        {/* Submit */}
        <Pressable
          onPress={handleSubmit}
          disabled={saving || uploadingPhoto}
          style={({ pressed }) => [
            styles.submitBtn,
            { backgroundColor: "#7F1F1F", opacity: pressed || saving || uploadingPhoto ? 0.8 : 1 },
          ]}
        >
          {saving || uploadingPhoto ? (
            <View style={{ flexDirection: "row", alignItems: "center", gap: 8 }}>
              <ActivityIndicator color="#fff" />
              <Text style={styles.submitBtnText}>{uploadingPhoto ? "Uploading photos…" : "Saving…"}</Text>
            </View>
          ) : (
            <Text style={styles.submitBtnText}>
              Submit Day Sheet{photos.length > 0 ? ` + ${photos.length} Photo${photos.length > 1 ? "s" : ""}` : ""}
            </Text>
          )}
        </Pressable>
      </ScrollView>

      {/* Job Picker Modal */}
        <Modal visible={jobPickerVisible} animationType="slide" presentationStyle="pageSheet">
        {/* Subcontractor Picker Modal */}
        <Modal visible={subPickerVisible} animationType="slide" presentationStyle="pageSheet">
          <View style={[styles.modal, { backgroundColor: colors.background }]}>
            <View style={styles.modalHeader}>
              <Text style={[styles.modalTitle, { color: colors.foreground }]}>Select Subcontractor / Crew</Text>
              <Pressable onPress={() => setSubPickerVisible(false)}>
                <Text style={{ color: "#7F1F1F", fontSize: 16 }}>Cancel</Text>
              </Pressable>
            </View>
            <FlatList
              data={subcontractors}
              keyExtractor={(item) => item.id}
              renderItem={({ item }) => (
                <Pressable
                  onPress={() => { setSelectedSubcontractor(item); setSubPickerVisible(false); }}
                  style={[styles.modalItem, { borderBottomColor: colors.border }]}
                >
                  <Text style={[styles.modalItemText, { color: colors.foreground }]}>{item.label}</Text>
                  {selectedSubcontractor?.id === item.id && (
                    <IconSymbol name="checkmark" size={16} color="#7F1F1F" />
                  )}
                </Pressable>
              )}
            />
          </View>
        </Modal>

        {/* Cost Code Picker Modal */}
        <Modal visible={costCodePickerVisible} animationType="slide" presentationStyle="pageSheet">
          <View style={[styles.modal, { backgroundColor: colors.background }]}>
            <View style={styles.modalHeader}>
              <Text style={[styles.modalTitle, { color: colors.foreground }]}>Select Cost Code</Text>
              <Pressable onPress={() => setCostCodePickerVisible(false)}>
                <Text style={{ color: "#7F1F1F", fontSize: 16 }}>Cancel</Text>
              </Pressable>
            </View>
            <FlatList
              data={costCodes}
              keyExtractor={(item) => item}
              renderItem={({ item }) => (
                <Pressable
                  onPress={() => { setSelectedCostCode(item); setCostCodePickerVisible(false); }}
                  style={[styles.modalItem, { borderBottomColor: colors.border }]}
                >
                  <Text style={[styles.modalItemText, { color: colors.foreground }]}>{item}</Text>
                  {selectedCostCode === item && (
                    <IconSymbol name="checkmark" size={16} color="#7F1F1F" />
                  )}
                </Pressable>
              )}
            />
          </View>
        </Modal>
        <View style={[styles.modalContainer, { backgroundColor: colors.background }]}>
          <View style={[styles.modalHeader, { backgroundColor: "#7F1F1F" }]}>
            <Text style={styles.modalTitle}>Select Job</Text>
            <Pressable onPress={() => setJobPickerVisible(false)} style={styles.modalCloseBtn}>
              <IconSymbol name="xmark.circle.fill" size={24} color="rgba(255,255,255,0.7)" />
            </Pressable>
          </View>
          <FlatList
            data={jobs}
            keyExtractor={(item) => item.id}
            contentContainerStyle={{ padding: 16, gap: 10 }}
            ListEmptyComponent={
              <Text style={[styles.emptyText, { color: colors.muted }]}>No active jobs found.</Text>
            }
            renderItem={({ item }) => (
              <Pressable
                onPress={() => {
                  setSelectedJob(item);
                  setJobPickerVisible(false);
                }}
                style={({ pressed }) => [
                  styles.jobPickerItem,
                  {
                    backgroundColor: selectedJob?.id === item.id ? "#FFF1F1" : colors.surface,
                    borderColor: selectedJob?.id === item.id ? "#7F1F1F" : colors.border,
                    opacity: pressed ? 0.85 : 1,
                  },
                ]}
              >
                <View style={[styles.jobCodeBadge, { backgroundColor: "#7F1F1F" }]}>
                  <Text style={styles.jobCodeText}>{item.jobCode}</Text>
                </View>
                <View style={{ flex: 1, marginTop: 6, gap: 2 }}>
                  <Text style={[styles.jobPickerName, { color: colors.foreground }]} numberOfLines={2}>
                    {item.jobName}
                  </Text>
                  <Text style={[styles.jobPickerClient, { color: colors.muted }]} numberOfLines={1}>
                    {item.client}
                  </Text>
                </View>
                {selectedJob?.id === item.id && (
                  <IconSymbol name="checkmark.circle.fill" size={20} color="#7F1F1F" />
                )}
              </Pressable>
            )}
          />
        </View>
      </Modal>
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
  pickerBtn: {
    borderWidth: 1.5,
    borderRadius: 10,
    paddingHorizontal: 14,
    paddingVertical: 12,
    flexDirection: "row",
    alignItems: "center",
    gap: 8,
    minHeight: 52,
  },
  pickerBtnText: {
    fontSize: 15,
    fontFamily: "Montserrat_500Medium",
    flex: 1,
  },
  pickerBtnSub: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
    marginTop: 2,
  },
  row: {
    flexDirection: "row",
    gap: 12,
  },
  hoursCard: {
    borderRadius: 12,
    padding: 16,
  },
  hoursRow: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
  },
  hoursLabel: {
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    color: "#7F1F1F",
  },
  hoursValue: {
    fontSize: 16,
    fontFamily: "Montserrat_700Bold",
    color: "#7F1F1F",
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
    paddingHorizontal: 14,
    paddingVertical: 12,
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
    minHeight: 100,
  },
  photoSubtitle: {
    fontSize: 13,
    fontFamily: "Montserrat_400Regular",
    marginTop: -8,
  },
  photoGrid: {
    flexDirection: "row",
    flexWrap: "wrap",
    gap: 8,
  },
  photoThumb: {
    width: 80,
    height: 80,
    borderRadius: 8,
    overflow: "hidden",
    position: "relative",
  },
  photoImage: {
    width: "100%",
    height: "100%",
  },
  photoRemoveBtn: {
    position: "absolute",
    top: 2,
    right: 2,
    backgroundColor: "rgba(0,0,0,0.6)",
    borderRadius: 10,
    width: 20,
    height: 20,
    alignItems: "center",
    justifyContent: "center",
  },
  photoRemoveText: {
    color: "#fff",
    fontSize: 10,
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
    paddingVertical: 10,
    borderRadius: 10,
  },
  photoBtnText: {
    fontSize: 14,
    fontFamily: "Montserrat_500Medium",
    color: "#fff",
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
  // Modal styles
  modalContainer: {
    flex: 1,
  },
  modalHeader: {
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    paddingHorizontal: 20,
    paddingVertical: 16,
  },
  modalTitle: {
    fontSize: 18,
    fontFamily: "Montserrat_700Bold",
    color: "#fff",
  },
  modalCloseBtn: {
    padding: 4,
  },
  jobPickerItem: {
    borderRadius: 12,
    padding: 14,
    borderWidth: 1.5,
    gap: 4,
    flexDirection: "column",
  },
  jobCodeBadge: {
    alignSelf: "flex-start",
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 6,
  },
  jobCodeText: {
    color: "#fff",
    fontSize: 11,
    fontFamily: "Montserrat_700Bold",
    letterSpacing: 0.5,
  },
  jobPickerName: {
    fontSize: 15,
    fontFamily: "Montserrat_600SemiBold",
    lineHeight: 20,
  },
  jobPickerClient: {
    fontSize: 12,
    fontFamily: "Montserrat_400Regular",
  },
  emptyText: {
    textAlign: "center",
    marginTop: 40,
    fontSize: 14,
    fontFamily: "Montserrat_400Regular",
  },
});
