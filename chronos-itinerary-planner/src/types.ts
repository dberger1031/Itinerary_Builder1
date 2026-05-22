/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ItinerarySubItem {
  id: string;
  description: string;
  duration: number; // minutes
  isDurationLocked: boolean;
}

export interface ItineraryItem {
  id: string;
  description: string;
  startTime: string; // "0900"
  duration: number; // minutes
  isDurationLocked: boolean;
  isTimelineLocked: boolean;
  notes: string;
  subItems: ItinerarySubItem[];
}

export interface DayData {
  id: string;
  date: string; // ISO string
  items: ItineraryItem[];
}

export interface ItineraryData {
  id: string;
  title: string;
  days: DayData[];
}

export interface TemplateItem {
  id: string;
  name: string;
  description?: string;
  duration: number;
  isDurationLocked: boolean;
  notes: string;
  subItems: ItinerarySubItem[];
}

export type UndoAction = 
  | { type: 'day'; data: DayData; index: number }
  | { type: 'item'; data: ItineraryItem; index: number; dayId: string };
