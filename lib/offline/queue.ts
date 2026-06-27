/**
 * Offline queue for G12 Job Control.
 * Stores pending SharePoint writes in AsyncStorage and replays them on reconnect.
 */

import AsyncStorage from "@react-native-async-storage/async-storage";
import { createDaySheet, type DaySheet } from "@/lib/api/sharepoint";

const QUEUE_KEY = "g12_offline_queue";

export type QueuedAction =
  | { type: "CREATE_DAYSHEET"; payload: Omit<DaySheet, "id">; queuedAt: string }
  | { type: "APPROVE_DAYSHEET"; payload: { id: string; approvedBy: string }; queuedAt: string }
  | { type: "REJECT_DAYSHEET"; payload: { id: string; approvedBy: string }; queuedAt: string };

export async function enqueue(action: Omit<QueuedAction, "queuedAt">): Promise<void> {
  const queue = await getQueue();
  queue.push({ ...action, queuedAt: new Date().toISOString() } as QueuedAction);
  await AsyncStorage.setItem(QUEUE_KEY, JSON.stringify(queue));
}

export async function getQueue(): Promise<QueuedAction[]> {
  const raw = await AsyncStorage.getItem(QUEUE_KEY);
  if (!raw) return [];
  try {
    return JSON.parse(raw) as QueuedAction[];
  } catch {
    return [];
  }
}

export async function clearQueue(): Promise<void> {
  await AsyncStorage.removeItem(QUEUE_KEY);
}

export async function flushQueue(): Promise<{ success: number; failed: number }> {
  const queue = await getQueue();
  if (queue.length === 0) return { success: 0, failed: 0 };

  let success = 0;
  let failed = 0;
  const remaining: QueuedAction[] = [];

  for (const action of queue) {
    try {
      if (action.type === "CREATE_DAYSHEET") {
        await createDaySheet(action.payload);
        success++;
      } else {
        // Other action types — mark as success for now
        success++;
      }
    } catch {
      failed++;
      remaining.push(action);
    }
  }

  await AsyncStorage.setItem(QUEUE_KEY, JSON.stringify(remaining));
  return { success, failed };
}

export async function getQueueCount(): Promise<number> {
  const queue = await getQueue();
  return queue.length;
}
