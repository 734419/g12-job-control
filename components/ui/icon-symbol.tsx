import MaterialIcons from "@expo/vector-icons/MaterialIcons";
import { SymbolWeight, SymbolViewProps } from "expo-symbols";
import { ComponentProps } from "react";
import { OpaqueColorValue, type StyleProp, type TextStyle } from "react-native";

type IconMapping = Record<SymbolViewProps["name"], ComponentProps<typeof MaterialIcons>["name"]>;
type IconSymbolName = keyof typeof MAPPING;

const MAPPING = {
  // Default template icons
  "house.fill": "home",
  "paperplane.fill": "send",
  "chevron.left.forwardslash.chevron.right": "code",
  "chevron.right": "chevron-right",
  // G12 Job Control icons
  "briefcase.fill": "work",
  "doc.text.fill": "description",
  "person.2.fill": "group",
  "person.fill": "person",
  "checkmark.circle.fill": "check-circle",
  "xmark.circle.fill": "cancel",
  "clock.fill": "schedule",
  "exclamationmark.triangle.fill": "warning",
  "plus.circle.fill": "add-circle",
  "arrow.clockwise": "refresh",
  "wifi.slash": "wifi-off",
  "wifi": "wifi",
  "magnifyingglass": "search",
  "line.3.horizontal.decrease.circle": "filter-list",
  "chevron.left": "chevron-left",
  "ellipsis": "more-horiz",
  "square.and.arrow.up": "share",
  "bell.fill": "notifications",
  "gear": "settings",
  "arrow.right.square": "logout",
  "calendar": "calendar-today",
  "location.fill": "location-on",
  "phone.fill": "phone",
  "envelope.fill": "email",
  "building.2.fill": "business",
  "shield.fill": "security",
  "doc.badge.plus": "note-add",
  "tray.full.fill": "inbox",
  "chart.bar.fill": "bar-chart",
  "camera.fill": "camera-alt",
  "photo.fill": "photo-library",
  "trash.fill": "delete",
  "arrow.up.doc.fill": "upload-file",
  "checkmark.square.fill": "check-box",
  "square": "check-box-outline-blank",
} as IconMapping;

export function IconSymbol({
  name,
  size = 24,
  color,
  style,
}: {
  name: IconSymbolName;
  size?: number;
  color: string | OpaqueColorValue;
  style?: StyleProp<TextStyle>;
  weight?: SymbolWeight;
}) {
  return <MaterialIcons color={color} size={size} name={MAPPING[name]} style={style} />;
}
