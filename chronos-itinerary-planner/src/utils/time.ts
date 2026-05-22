/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export function militaryToMinutes(military: string): number {
  if (!military) return 0;
  const clean = military.padStart(4, "0");
  const h = parseInt(clean.substring(0, 2)) || 0;
  const m = parseInt(clean.substring(2, 4)) || 0;
  return h * 60 + m;
}

export function minutesToMilitary(minutes: number): string {
  const h = Math.floor(minutes / 60) % 24;
  const m = minutes % 60;
  return String(h).padStart(2, '0') + String(m).padStart(2, '0');
}

export function addMinutesToMilitary(military: string, duration: number): string {
  const mins = militaryToMinutes(military);
  return minutesToMilitary(mins + duration);
}

export function formatMilitary(military: string): string {
  if (!military || military.length < 4) return military;
  const hh = parseInt(military.substring(0, 2)) || 0;
  const mm = parseInt(military.substring(2, 4)) || 0;
  const ampm = hh >= 12 ? 'PM' : 'AM';
  const displayH = hh % 12 === 0 ? 12 : hh % 12;
  return `${displayH}:${String(mm).padStart(2, '0')} ${ampm}`;
}

export function checkOverlap(start1: string, duration1: number, start2: string, duration2: number): boolean {
  const s1 = militaryToMinutes(start1);
  const e1 = s1 + duration1;
  const s2 = militaryToMinutes(start2);
  const e2 = s2 + duration2;
  return Math.max(s1, s2) < Math.min(e1, e2);
}

export function getDuration(start: string, end: string): number {
  let s = militaryToMinutes(start);
  let e = militaryToMinutes(end);
  if (e < s) {
    e += 1440; // overflow past midnight
  }
  return e - s;
}
