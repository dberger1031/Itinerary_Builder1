/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useCallback, useRef } from "react";
import { 
  Plus, 
  Trash2, 
  Lock, 
  Unlock, 
  GripVertical, 
  ChevronDown, 
  ChevronUp, 
  ChevronLeft,
  ChevronRight,
  Upload,
  Calendar,
  Clock,
  FileText,
  AlertCircle,
  FileDown,
  Archive,
  X,
  Copy,
  RotateCcw,
  Layers
} from "lucide-react";
import { Reorder, motion, AnimatePresence } from "motion/react";
import { format, parseISO, startOfDay, addDays } from "date-fns";
import { Document, Packer, Paragraph, HeadingLevel, Table, TableRow, TableCell, WidthType, TextRun, AlignmentType } from "docx";
import mammoth from "mammoth";
import { 
  ItineraryData, 
  DayData, 
  ItineraryItem, 
  ItinerarySubItem,
  TemplateItem,
  UndoAction
} from "./types";
import { 
  addMinutesToMilitary, 
  formatMilitary, 
  militaryToMinutes, 
  getDuration
} from "./utils/time";
import { cn } from "./lib/utils";

const STORAGE_KEY = "chronos_itinerary_data";

const DEFAULT_ITINERARY_ID = "default-trip";

const createDefaultItinerary = (id = DEFAULT_ITINERARY_ID, title = "Trip Itinerary"): ItineraryData => ({
  id,
  title,
  days: [
    {
      id: `day-${Date.now()}`,
      date: new Date().toISOString(),
      items: [
        {
          id: "item-1",
          description: "Morning Meeting",
          startTime: "0900",
          duration: 60,
          isDurationLocked: true,
          isTimelineLocked: false,
          notes: "Conference Room A",
          subItems: []
        }
      ]
    }
  ]
});

const normalizeItineraries = (data: ItineraryData[]): ItineraryData[] => {
  if (!Array.isArray(data)) return [];
  return data.map(itin => ({
    ...itin,
    days: (itin.days || []).map(day => ({
      ...day,
      items: (day.items || []).map(item => ({
        ...item,
        duration: Number(item.duration) || 0,
        subItems: (item.subItems || []).map(sub => ({
          ...sub,
          duration: Number(sub.duration) || 0
        }))
      }))
    }))
  }));
};

const normalizeTemplates = (data: TemplateItem[]): TemplateItem[] => {
  if (!Array.isArray(data)) return [];
  return data.map(tpl => ({
    ...tpl,
    duration: Number(tpl.duration) || 0,
    subItems: (tpl.subItems || []).map(sub => ({
      ...sub,
      duration: Number(sub.duration) || 0
    }))
  }));
};

export default function App() {
  const [itineraries, setItineraries] = useState<ItineraryData[]>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Migration for old single-itinerary format
        if (!Array.isArray(parsed.itineraries)) {
          const legacy = parsed as any;
          return normalizeItineraries([createDefaultItinerary("legacy", legacy.title || "My Trip")]);
        }
        return normalizeItineraries(parsed.itineraries);
      } catch (e) {
        console.error("Failed to parse saved itinerary data. Resetting...", e);
        return normalizeItineraries([createDefaultItinerary()]);
      }
    }
    return normalizeItineraries([createDefaultItinerary()]);
  });

  const [activeItineraryId, setActiveItineraryId] = useState<string>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return parsed.activeItineraryId || itineraries[0]?.id || DEFAULT_ITINERARY_ID;
      } catch (e) {
        return DEFAULT_ITINERARY_ID;
      }
    }
    return DEFAULT_ITINERARY_ID;
  });

  const [templates, setTemplates] = useState<TemplateItem[]>(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        return normalizeTemplates(parsed.templates || []);
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  const [isToolboxOpen, setIsToolboxOpen] = useState(false);
  const [undoAction, setUndoAction] = useState<UndoAction | null>(null);
  const [savedToast, setSavedToast] = useState<string | null>(null);

  // Custom calendar states
  const [isCalendarOpen, setIsCalendarOpen] = useState(false);
  const [viewDate, setViewDate] = useState<Date>(new Date());
  const calendarRef = useRef<HTMLDivElement>(null);

  // Click outside listener for custom calendar
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (calendarRef.current && !calendarRef.current.contains(event.target as Node)) {
        setIsCalendarOpen(false);
      }
    }
    if (isCalendarOpen) {
      document.removeEventListener("mousedown", handleClickOutside);
      document.addEventListener("mousedown", handleClickOutside);
    }
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, [isCalendarOpen]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Auto-clear undo after 6 seconds
  useEffect(() => {
    if (undoAction) {
      const timer = setTimeout(() => setUndoAction(null), 6000);
      return () => clearTimeout(timer);
    }
  }, [undoAction]);

  // Auto-clear saved toast after 3 seconds
  useEffect(() => {
    if (savedToast) {
      const timer = setTimeout(() => setSavedToast(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [savedToast]);

  const activeItinerary = itineraries.find(i => i.id === activeItineraryId) || itineraries[0] || createDefaultItinerary();

  useEffect(() => {
    const appData = {
      itineraries,
      activeItineraryId,
      templates
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
  }, [itineraries, activeItineraryId, templates]);

  const updateActiveItinerary = (updates: Partial<ItineraryData>) => {
    setItineraries(prev => prev.map(i => i.id === activeItineraryId ? { ...i, ...updates } : i));
  };

  // Custom calendar helpers
  const currentYear = new Date().getFullYear();
  const YEAR_OPTIONS = Array.from({ length: 41 }, (_, i) => currentYear - 20 + i);
  const MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June", 
    "July", "August", "September", "October", "November", "December"
  ];

  const getDaysInMonth = (year: number, month: number) => {
    return new Date(year, month + 1, 0).getDate();
  };

  const getFirstDayOfWeek = (year: number, month: number) => {
    return new Date(year, month, 1).getDay();
  };

  const selectDate = (date: Date) => {
    const newStart = startOfDay(date);
    updateActiveItinerary({
      days: activeItinerary.days.map((d, i) => ({
        ...d,
        date: addDays(newStart, i).toISOString()
      }))
    });
    setIsCalendarOpen(false);
  };

  const renderCalendarDays = () => {
    const year = viewDate.getFullYear();
    const month = viewDate.getMonth();
    const totalDays = getDaysInMonth(year, month);
    const firstDayIndex = getFirstDayOfWeek(year, month);
    
    const activeStartDate = activeItinerary.days[0] ? parseISO(activeItinerary.days[0].date) : null;
    
    const days = [];
    for (let i = 0; i < firstDayIndex; i++) {
      days.push(
        <div key={`empty-${i}`} className="h-8 w-8" />
      );
    }
    
    for (let day = 1; day <= totalDays; day++) {
      const cellDate = new Date(year, month, day);
      const isSelected = activeStartDate && 
        cellDate.getFullYear() === activeStartDate.getFullYear() &&
        cellDate.getMonth() === activeStartDate.getMonth() &&
        cellDate.getDate() === activeStartDate.getDate();
        
      const isToday = (() => {
        const today = new Date();
        return cellDate.getFullYear() === today.getFullYear() &&
               cellDate.getMonth() === today.getMonth() &&
               cellDate.getDate() === today.getDate();
      })();
      
      days.push(
        <button
          key={`day-${day}`}
          onClick={() => selectDate(cellDate)}
          className={cn(
            "h-8 w-8 rounded-lg flex items-center justify-center font-bold text-xs transition-all cursor-pointer",
            isSelected 
              ? "bg-indigo-600 text-white font-extrabold shadow-sm shadow-indigo-500/30 scale-105" 
              : isToday 
                ? "bg-zinc-800 text-indigo-400 border border-indigo-500/30 font-extrabold" 
                : "text-zinc-300 hover:bg-zinc-800 hover:text-white"
          )}
        >
          {day}
        </button>
      );
    }
    return days;
  };

  const createNewItinerary = () => {
    const newId = `trip-${Date.now()}`;
    const newTrip = createDefaultItinerary(newId, "New Trip");
    setItineraries(prev => [...prev, newTrip]);
    setActiveItineraryId(newId);
  };

  const removeItinerary = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (itineraries.length <= 1) return;
    
    if (window.confirm("Delete this itinerary entire voyage?")) {
      setItineraries(prev => prev.filter(i => i.id !== id));
      if (activeItineraryId === id) {
        setActiveItineraryId(itineraries.find(i => i.id !== id)?.id || "");
      }
    }
  };

  const updateDay = (dayId: string, updates: Partial<DayData>) => {
    updateActiveItinerary({
      days: activeItinerary.days.map(d => d.id === dayId ? { ...d, ...updates } : d)
    });
  };

  const addDay = () => {
    const lastDay = activeItinerary.days[activeItinerary.days.length - 1];
    const newDate = lastDay ? addDays(parseISO(lastDay.date), 1) : new Date();
    const newDay: DayData = {
      id: `day-${Date.now()}`,
      date: newDate.toISOString(),
      items: []
    };
    updateActiveItinerary({ days: [...activeItinerary.days, newDay] });
  };

  const removeDay = (dayId: string) => {
    const index = activeItinerary.days.findIndex(d => d.id === dayId);
    if (index === -1) return;
    
    setUndoAction({
      type: 'day',
      data: activeItinerary.days[index],
      index
    });
    
    updateActiveItinerary({ days: activeItinerary.days.filter(d => d.id !== dayId) });
  };

  const duplicateDay = (dayId: string) => {
    const index = activeItinerary.days.findIndex(d => d.id === dayId);
    if (index === -1) return;
    
    const dayToCopy = activeItinerary.days[index];
    const newDayId = `day-${Date.now()}`;
    
    // Deep copy and generate fresh IDs for items and sub-items to avoid clashes
    const clonedItems = (dayToCopy.items || []).map((item, idx) => ({
      ...item,
      id: `item-${Date.now()}-${idx}-${Math.floor(Math.random() * 1000)}`,
      subItems: (item.subItems || []).map((sub, sIdx) => ({
        ...sub,
        id: `sub-${Date.now()}-${idx}-${sIdx}-${Math.floor(Math.random() * 1000)}`
      }))
    }));

    // Find a proper sequential date (e.g. +1 day from the last day in activeItinerary)
    const lastDay = activeItinerary.days[activeItinerary.days.length - 1];
    const newDate = lastDay ? addDays(parseISO(lastDay.date), 1) : new Date();

    const newDay: DayData = {
      id: newDayId,
      date: newDate.toISOString(),
      items: clonedItems
    };

    updateActiveItinerary({ days: [...activeItinerary.days, newDay] });
    setSavedToast(`Duplicated Day schedule`);
  };

  const undoLastDelete = () => {
    if (!undoAction) return;

    if (undoAction.type === 'day') {
        updateActiveItinerary({
            days: [
              ...activeItinerary.days.slice(0, undoAction.index),
              undoAction.data,
              ...activeItinerary.days.slice(undoAction.index)
            ]
        });
    } else if (undoAction.type === 'item') {
        updateActiveItinerary({
            days: activeItinerary.days.map(day => {
                if (day.id === undoAction.dayId) {
                    const newItems = [...day.items];
                    newItems.splice(undoAction.index, 0, undoAction.data);
                    return { ...day, items: newItems };
                }
                return day;
            })
        });
    }
    setUndoAction(null);
  };

  const saveToLibrary = (item: ItineraryItem) => {
    const template: TemplateItem = {
      id: `tpl-${Date.now()}`,
      name: item.description || "Unnamed Template",
      description: item.description,
      duration: item.duration,
      isDurationLocked: item.isDurationLocked,
      notes: item.notes,
      subItems: JSON.parse(JSON.stringify(item.subItems || []))
    };

    setTemplates(prev => [...prev, template]);
    setIsToolboxOpen(true);
    setSavedToast(item.description || "Item");
  };

  const removeTemplate = (id: string) => {
    setTemplates(prev => prev.filter(t => t.id !== id));
  };

  const exportToWord = async () => {
    // Generate the rows for the single global table
    const tableRows = [
      // 1. Header Row
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({
            width: { size: 15, type: WidthType.PERCENTAGE },
            shading: { fill: "EAEAEA" },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: "Date", bold: true, size: 22 })],
                alignment: AlignmentType.CENTER
              })
            ]
          }),
          new TableCell({
            width: { size: 55, type: WidthType.PERCENTAGE },
            shading: { fill: "EAEAEA" },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: "Timeline (All Times Local)", bold: true, size: 22 })],
                alignment: AlignmentType.CENTER
              })
            ]
          }),
          new TableCell({
            width: { size: 30, type: WidthType.PERCENTAGE },
            shading: { fill: "EAEAEA" },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },
            children: [
              new Paragraph({
                children: [new TextRun({ text: "Notes", bold: true, size: 22 })],
                alignment: AlignmentType.CENTER
              })
            ]
          })
        ]
      }),

      // 2. Day Rows
      ...activeItinerary.days.map(day => {
        const dateObj = parseISO(day.date);
        const dayStr = format(dateObj, "dd MMM"); // e.g. "22 May" / "DD M"
        const weekdayStr = format(dateObj, "EEEE"); // e.g. "Monday"

        // Build list of paragraphs for Column 2 (Timeline (All Times Local))
        const timelineParagraphs: Paragraph[] = [];

        day.items.forEach(item => {
          const startTime = item.startTime;
          const endTime = addMinutesToMilitary(item.startTime, item.duration);
          const timeRangeStr = `${startTime}-${endTime}`;

          const descLower = item.description.toLowerCase();
          const isRoutine = descLower.includes("transition") ||
                            descLower.includes("lodging") ||
                            descLower.includes("executive time") ||
                            descLower.includes("lunch") ||
                            descLower.includes("dinner") ||
                            descLower.includes("airport ops") ||
                            descLower.includes("flight") ||
                            descLower.includes("drive") ||
                            descLower.includes("travel");

          // Bold if it has sub-items or is not a routine transition/routine task
          const shouldBold = item.subItems.length > 0 || !isRoutine;

          timelineParagraphs.push(new Paragraph({
            children: [
              new TextRun({
                text: `${timeRangeStr}: `,
                bold: shouldBold,
                size: 20
              }),
              new TextRun({
                text: item.description,
                bold: shouldBold,
                size: 20
              })
            ],
            spacing: { before: 40, after: 40 }
          }));

          // Render nested sub-items if present
          if (item.subItems && item.subItems.length > 0) {
            item.subItems.forEach((sub, subIdx) => {
              const subStartTime = calculateSubItemTime(item.startTime, item.subItems, subIdx);
              const subEndTime = addMinutesToMilitary(subStartTime, sub.duration);
              const subTimeRangeStr = `${subStartTime}-${subEndTime}`;

              timelineParagraphs.push(new Paragraph({
                children: [
                  new TextRun({
                    text: `\u2022    ${subTimeRangeStr}: `,
                    bold: false,
                    size: 20
                  }),
                  new TextRun({
                    text: sub.description,
                    bold: false,
                    size: 20
                  })
                ],
                indent: { left: 400 }, // Clean indentation for step bullet points
                spacing: { before: 20, after: 20 }
              }));
            });
          }
        });

        if (timelineParagraphs.length === 0) {
          timelineParagraphs.push(new Paragraph({
            children: [new TextRun({ text: "No scheduled events", size: 20 })],
            spacing: { before: 40, after: 40 }
          }));
        }

        // Build list of paragraphs for Column 3 (Notes)
        const notesParagraphs: Paragraph[] = [];
        let hasAnyNotes = false;

        day.items.forEach(item => {
          if (item.notes && item.notes.trim()) {
            hasAnyNotes = true;
            notesParagraphs.push(new Paragraph({
              children: [
                new TextRun({
                  text: item.description,
                  bold: true,
                  size: 18
                }),
                new TextRun({
                  text: ": " + item.notes,
                  size: 18
                })
              ],
              spacing: { before: 40, after: 40 }
            }));
          }
        });

        // Ensure there is at least one paragraph in the cell
        if (!hasAnyNotes) {
          notesParagraphs.push(new Paragraph({
            children: [new TextRun({ text: "", size: 18 })]
          }));
        }

        return new TableRow({
          children: [
            // Date cell (Column 1)
            new TableCell({
              width: { size: 15, type: WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 120, right: 120 },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: dayStr, bold: true, size: 20 })],
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 40, after: 20 }
                }),
                new Paragraph({
                  children: [new TextRun({ text: weekdayStr, bold: true, size: 20 })],
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 20, after: 40 }
                })
              ]
            }),

            // Timeline cell (Column 2)
            new TableCell({
              width: { size: 55, type: WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 120, right: 120 },
              children: timelineParagraphs
            }),

            // Notes cell (Column 3)
            new TableCell({
              width: { size: 30, type: WidthType.PERCENTAGE },
              margins: { top: 120, bottom: 120, left: 120, right: 120 },
              children: notesParagraphs
            })
          ]
        });
      })
    ];

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            text: activeItinerary.title,
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 400 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: tableRows
          })
        ]
      }]
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${activeItinerary.title.replace(/\s+/g, '_')}.docx`;
    a.click();
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      const arrayBuffer = event.target?.result as ArrayBuffer;
      try {
        // Try importing as structured table HTML first for perfect alignment
        const htmlResult = await mammoth.convertToHtml({ arrayBuffer });
        const parser = new DOMParser();
        const doc = parser.parseFromString(htmlResult.value, "text/html");
        const table = doc.querySelector("table");

        let importedDays: DayData[] = [];
        const fileTitle = file.name.replace(/\.[^/.]+$/, "").replace(/[_-]/g, " ");

        if (table) {
          const rows = Array.from(table.querySelectorAll("tr"));
          // Try to skip header row if there is one
          let startRowIdx = 0;
          if (rows.length > 0) {
            const firstRowCells = Array.from(rows[0].querySelectorAll("td"));
            const firstCellText = firstRowCells[0]?.textContent?.trim().toLowerCase() || "";
            const secondCellText = firstRowCells[1]?.textContent?.trim().toLowerCase() || "";
            if (firstCellText.includes("date") || secondCellText.includes("timeline")) {
              startRowIdx = 1;
            }
          }

          let lastParsedDate = activeItinerary.days[0] 
            ? parseISO(activeItinerary.days[0].date) 
            : new Date();

          for (let rIdx = startRowIdx; rIdx < rows.length; rIdx++) {
            const cells = Array.from(rows[rIdx].querySelectorAll("td"));
            if (cells.length < 2) continue; // Need at least Date and Timeline columns

            const dateCellText = cells[0]?.textContent || "";
            const timelineCell = cells[1];
            const notesCell = cells[2];

            // Parse Date
            let dayDate = lastParsedDate;
            if (rIdx > startRowIdx) {
              dayDate = addDays(lastParsedDate, 1);
            }
            if (dateCellText.trim()) {
              dayDate = parseDateCell(dateCellText, dayDate);
            }
            lastParsedDate = dayDate;

            // Parse Timeline Cell to extract parent items and subitems
            const dayItems: ItineraryItem[] = [];
            let currentParentItem: ItineraryItem | null = null;

            if (timelineCell) {
              const pTags = Array.from(timelineCell.querySelectorAll("p, li"));
              pTags.forEach((p, pIdx) => {
                const text = p.textContent || "";
                if (!text.trim()) return;

                const isSubItem = text.startsWith('\u2022') || text.startsWith('•') || p.tagName.toLowerCase() === 'li';
                const cleanText = text.replace(/^[\u2022•\s\-\*]+/, "").trim();

                // regex matches e.g. "0700-0900: Event name" or "0700 - 0900: Event name"
                const timeRegex = /^(\d{4})\s*-\s*(\d{4})\s*:\s*(.*)$/;
                const match = cleanText.match(timeRegex);

                if (isSubItem && currentParentItem) {
                  if (match) {
                    const subStart = match[1];
                    const subEnd = match[2];
                    const subDesc = match[3];
                    const subDur = getDuration(subStart, subEnd);

                    const subItem: ItinerarySubItem = {
                      id: `sub-${Date.now()}-${rIdx}-${pIdx}`,
                      description: subDesc || "Sub activity",
                      duration: subDur,
                      isDurationLocked: false
                    };
                    currentParentItem.subItems.push(subItem);
                  } else {
                    const subItem: ItinerarySubItem = {
                      id: `sub-${Date.now()}-${rIdx}-${pIdx}`,
                      description: cleanText,
                      duration: 15,
                      isDurationLocked: false
                    };
                    currentParentItem.subItems.push(subItem);
                  }
                } else {
                  // It's a Parent Itinerary Item!
                  if (match) {
                    const itemStart = match[1];
                    const itemEnd = match[2];
                    const itemDesc = match[3];
                    const itemDur = getDuration(itemStart, itemEnd);

                    currentParentItem = {
                      id: `item-${Date.now()}-${rIdx}-${pIdx}`,
                      description: itemDesc || "Activity",
                      startTime: itemStart,
                      duration: itemDur,
                      isDurationLocked: false,
                      isTimelineLocked: false,
                      notes: "",
                      subItems: []
                    };
                    dayItems.push(currentParentItem);
                  } else {
                    if (cleanText !== "No scheduled events") {
                      currentParentItem = {
                        id: `item-${Date.now()}-${rIdx}-${pIdx}`,
                        description: cleanText,
                        startTime: "0900",
                        duration: 30,
                        isDurationLocked: false,
                        isTimelineLocked: false,
                        notes: "",
                        subItems: []
                      };
                      dayItems.push(currentParentItem);
                    }
                  }
                }
              });
            }

            // Parse Notes Cell
            if (notesCell && dayItems.length > 0) {
              const notesRows = Array.from(notesCell.querySelectorAll("p, li"));
              notesRows.forEach(nRow => {
                const nText = nRow.textContent?.trim() || "";
                if (!nText) return;

                const colonIdx = nText.indexOf(":");
                if (colonIdx !== -1) {
                  const keyDesc = nText.slice(0, colonIdx).trim().toLowerCase();
                  const valNotes = nText.slice(colonIdx + 1).trim();

                  // Find exact or substring matching item description
                  const matchedItem = dayItems.find(item => 
                    item.description.toLowerCase() === keyDesc || 
                    item.description.toLowerCase().includes(keyDesc) || 
                    keyDesc.includes(item.description.toLowerCase())
                  );
                  if (matchedItem) {
                    // Prepend or set notes
                    matchedItem.notes = matchedItem.notes 
                      ? matchedItem.notes + "; " + valNotes 
                      : valNotes;
                  }
                }
              });
            }

            importedDays.push({
              id: `day-${Date.now()}-${rIdx}`,
              date: dayDate.toISOString(),
              items: dayItems
            });
          }
        }

        // Fallback or if no table found (standard paragraph format)
        if (importedDays.length === 0) {
          const rawResult = await mammoth.extractRawText({ arrayBuffer });
          const rawLines = rawResult.value.split('\n').filter(l => l.trim().length > 0);
          const newItems: ItineraryItem[] = [];

          rawLines.forEach((l, i) => {
            const clean = l.trim();
            const timeRegex = /^(\d{4})\s*-\s*(\d{4})\s*:\s*(.*)$/;
            const match = clean.match(timeRegex);

            if (match) {
              const itemStart = match[1];
              const itemEnd = match[2];
              const itemDesc = match[3];
              const itemDur = getDuration(itemStart, itemEnd);

              newItems.push({
                id: `imported-${Date.now()}-${i}`,
                description: itemDesc,
                startTime: itemStart,
                duration: itemDur,
                isDurationLocked: false,
                isTimelineLocked: false,
                notes: "",
                subItems: []
              });
            } else {
              newItems.push({
                id: `imported-${Date.now()}-${i}`,
                description: clean,
                startTime: "0900",
                duration: 30,
                isDurationLocked: false,
                isTimelineLocked: false,
                notes: "",
                subItems: []
              });
            }
          });

          if (newItems.length > 0) {
            importedDays.push({
              id: `day-${Date.now()}`,
              date: new Date().toISOString(),
              items: newItems
            });
          }
        }

        if (importedDays.length > 0) {
          // Create as a seamless, standalone saved itinerary in the list of itineraries
          const newId = `trip-imported-${Date.now()}`;
          const newItinerary: ItineraryData = {
            id: newId,
            title: fileTitle,
            days: importedDays
          };

          setItineraries(prev => [...prev, newItinerary]);
          setActiveItineraryId(newId);
          setSavedToast(`Imported '${fileTitle}' successfully!`);
        } else {
          setSavedToast("Could not find any schedule structure to import");
        }
      } catch (err) {
        console.error("Failed to parse .docx file via mammoth", err);
        setSavedToast("Failed to parse .docx file. Ensure it is a valid document.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="h-screen bg-[#09090B] font-sans text-zinc-100 flex flex-col overflow-hidden">
      {/* Undo Toast */}
      <AnimatePresence>
        {undoAction && (
          <motion.div 
            initial={{ opacity: 0, y: 50, scale: 0.9 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 20, scale: 0.95 }}
            className="fixed bottom-8 left-1/2 -translate-x-1/2 z-[100] bg-[#18181B] text-zinc-100 px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-4 border border-zinc-800 backdrop-blur-md bg-opacity-95 min-w-[320px] justify-between"
          >
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 bg-zinc-800 rounded-full flex items-center justify-center text-indigo-400">
                <Trash2 size={16} />
              </div>
              <span className="text-sm font-medium">
                {undoAction.type === 'day' ? 'Day' : 'Itinerary item'} deleted
              </span>
            </div>
            <button 
              onClick={undoLastDelete}
              className="flex items-center gap-2 px-4 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-xs font-bold rounded-lg transition-colors group cursor-pointer"
            >
              <RotateCcw size={14} className="group-active:rotate-[-90deg] transition-transform" />
              UNDO
            </button>
          </motion.div>
        )}
      </AnimatePresence>

      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Tabs Bar */}
        <div className="bg-[#09090B] border-b border-zinc-800/80 px-4 md:px-8 py-2 flex items-center gap-2 overflow-x-auto no-scrollbar scroll-smooth">
          {itineraries.map(trip => (
            <div 
              key={trip.id}
              onClick={() => setActiveItineraryId(trip.id)}
              className={cn(
                "group flex items-center gap-3 px-3 py-1.5 rounded-xl cursor-pointer transition-all border whitespace-nowrap",
                activeItineraryId === trip.id 
                  ? "bg-indigo-600/10 border-indigo-500/30 text-indigo-400 shadow-sm" 
                  : "bg-zinc-900 border-zinc-800/80 text-zinc-400 hover:bg-zinc-800/50 hover:border-zinc-700 hover:text-zinc-200"
              )}
            >
              <FileText size={14} className={cn(activeItineraryId === trip.id ? "text-indigo-400" : "text-zinc-500")} />
              <span className="text-xs font-bold leading-none truncate max-w-[120px]">
                {trip.title || "Untitled Trip"}
              </span>
              {itineraries.length > 1 && (
                <button 
                  onClick={(e) => removeItinerary(trip.id, e)}
                  className="p-1 hover:bg-zinc-800 rounded-md opacity-0 group-hover:opacity-100 transition-opacity text-zinc-500 hover:text-rose-450 ml-1 cursor-pointer"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          ))}
          <button 
            onClick={createNewItinerary}
            className="flex items-center gap-2 px-3 py-1.5 rounded-xl text-zinc-500 hover:bg-zinc-900 hover:text-zinc-300 transition-all border border-dashed border-zinc-800 hover:border-zinc-750 text-xs font-bold shrink-0 cursor-pointer"
          >
            <Plus size={14} />
            <span>New Trip</span>
          </button>
        </div>

        {/* Header */}
        <header className="bg-[#09090B] border-b border-zinc-800/85 p-4 md:px-6 flex flex-col md:flex-row justify-between items-center gap-4 z-20 shadow-xs transition-all duration-500 ease-in-out">
          <div className="flex-1 w-full max-w-4xl mx-auto flex flex-col md:flex-row items-center gap-4">
            <div className="flex items-center gap-3 flex-1 w-full">
              <input 
                type="text" 
                className="text-xl font-bold bg-transparent border-none outline-none focus:ring-1 focus:ring-zinc-800 rounded-lg px-2 w-full transition-all text-zinc-100 placeholder:text-zinc-650"
                value={activeItinerary.title}
                onChange={(e) => updateActiveItinerary({ title: e.target.value })}
                placeholder="Enter Itinerary Title..."
                id="itinerary-title"
              />
              <div 
                onClick={() => setSavedToast("Itinerary saved for future edits!")}
                className="flex items-center gap-1.5 px-2.5 py-1 bg-zinc-900/80 hover:bg-zinc-800/80 text-zinc-450 hover:text-zinc-200 border border-zinc-800/80 rounded-full text-[10px] font-bold cursor-pointer select-none transition-all active:scale-95 shrink-0"
                title="Your edits are saved automatically. Click to ensure instant manual backup."
              >
                <span className="relative flex h-1.5 w-1.5">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-500"></span>
                </span>
                <span>Auto-Saved</span>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-3 p-1.5 bg-zinc-900/40 rounded-xl border border-zinc-800/80 shrink-0">
              <div className="flex items-center gap-2 relative">
                <Calendar size={14} className="text-zinc-500" />
                <div 
                  className="flex flex-col cursor-pointer select-none" 
                  onClick={() => {
                    setIsCalendarOpen(!isCalendarOpen);
                    if (activeItinerary.days[0]) {
                      setViewDate(parseISO(activeItinerary.days[0].date));
                    } else {
                      setViewDate(new Date());
                    }
                  }}
                >
                  <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-tight selection:bg-transparent">Start Date</span>
                  <span className="text-sm font-semibold text-zinc-200 hover:text-indigo-400 transition-colors">
                    {activeItinerary.days[0] ? format(parseISO(activeItinerary.days[0].date), "MMM dd, yyyy") : "Select Date"}
                  </span>
                </div>
                
                {/* Popover Calendar */}
                <AnimatePresence>
                  {isCalendarOpen && (
                    <motion.div
                      ref={calendarRef}
                      initial={{ opacity: 0, y: 10, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: 10, scale: 0.95 }}
                      transition={{ duration: 0.15 }}
                      className="absolute top-full left-0 mt-2 z-50 bg-[#121214] border border-zinc-800/90 p-3.5 rounded-2xl shadow-2xl w-[280px]"
                    >
                      {/* Month & Year Navigation / Selectors */}
                      <div className="flex items-center justify-between mb-3 gap-1">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            const nextD = new Date(viewDate);
                            nextD.setMonth(viewDate.getMonth() - 1);
                            setViewDate(nextD);
                          }}
                          className="p-1.5 text-zinc-400 hover:text-white hover:bg-zinc-800 rounded-lg transition-colors cursor-pointer"
                        >
                          <ChevronLeft size={14} />
                        </button>
                        
                        <div className="flex items-center gap-1 flex-1 justify-center">
                          {/* Month Selector */}
                          <select
                            value={viewDate.getMonth()}
                            onChange={(e) => {
                              const newD = new Date(viewDate);
                              newD.setMonth(parseInt(e.target.value));
                              setViewDate(newD);
                            }}
                            className="bg-[#1c1c1f] border border-zinc-800/80 text-[11px] font-bold px-1.5 py-0.5 rounded-md text-zinc-200 focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                          >
                            {MONTH_NAMES.map((m, idx) => (
                              <option key={m} value={idx}>{m}</option>
                            ))}
                          </select>
                          
                          {/* Year Selector */}
                          <select
                            value={viewDate.getFullYear()}
                            onChange={(e) => {
                              const newD = new Date(viewDate);
                              newD.setFullYear(parseInt(e.target.value));
                              setViewDate(newD);
                            }}
                            className="bg-[#1c1c1f] border border-zinc-800/80 text-[11px] font-bold px-1.5 py-0.5 rounded-md text-zinc-200 focus:outline-none focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                          >
                            {YEAR_OPTIONS.map(y => (
                              <option key={y} value={y}>{y}</option>
                            ))}
                          </select>
                        </div>
                        
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            const nextD = new Date(viewDate);
                            nextD.setMonth(viewDate.getMonth() + 1);
                            setViewDate(nextD);
                          }}
                          className="p-1.5 text-zinc-400 hover:text-white hover:bg-zinc-800 rounded-lg transition-colors cursor-pointer"
                        >
                          <ChevronRight size={14} />
                        </button>
                      </div>

                      {/* Day Labels */}
                      <div className="grid grid-cols-7 text-center text-zinc-500 font-bold text-[9px] mb-2 uppercase tracking-widest leading-none">
                        <span>Su</span>
                        <span>Mo</span>
                        <span>Tu</span>
                        <span>We</span>
                        <span>Th</span>
                        <span>Fr</span>
                        <span>Sa</span>
                      </div>

                      {/* Days Grid */}
                      <div className="grid grid-cols-7 gap-1">
                        {renderCalendarDays()}
                      </div>
                      
                      {/* Shortcut buttons */}
                      <div className="mt-3.5 pt-2 border-t border-zinc-800/60 flex justify-between items-center text-[10px]">
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            const today = new Date();
                            setViewDate(today);
                            selectDate(today);
                          }}
                          className="text-indigo-400 hover:text-indigo-300 font-bold cursor-pointer"
                        >
                          Today
                        </button>
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            setIsCalendarOpen(false);
                          }}
                          className="text-zinc-500 hover:text-zinc-300 font-bold cursor-pointer"
                        >
                          Close
                        </button>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
              <div className="w-px h-8 bg-zinc-800 hidden sm:block" />
              <div className="flex items-center gap-2">
                <Clock size={16} className="text-zinc-500" />
                <div className="flex flex-col">
                  <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-tight">Total Days</span>
                  <input 
                    type="number"
                    min="1"
                    max="31"
                    className="bg-transparent border-none outline-none text-sm font-bold w-12 p-0 focus:ring-0 text-zinc-200"
                    value={activeItinerary.days.length}
                    onChange={(e) => {
                      const count = Math.max(1, parseInt(e.target.value) || 1);
                      const currentDays = [...activeItinerary.days];
                      if (count > currentDays.length) {
                        const extraDays = count - currentDays.length;
                        const lastDate = parseISO(currentDays[currentDays.length - 1]?.date || new Date().toISOString());
                        const newDays = Array.from({ length: extraDays }, (_, i) => ({
                          id: `day-${Date.now()}-${i}`,
                          date: addDays(lastDate, i + 1).toISOString(),
                          items: []
                        }));
                        updateActiveItinerary({ days: [...currentDays, ...newDays] });
                      } else {
                        updateActiveItinerary({ days: currentDays.slice(0, count) });
                      }
                    }}
                  />
                </div>
              </div>
            </div>
            <div className="flex items-center gap-2 grow-0 shrink-0">
              <button 
                onClick={() => setIsToolboxOpen(!isToolboxOpen)}
                className={cn(
                  "flex items-center gap-2 px-3 py-1.5 rounded-xl transition-all cursor-pointer border text-xs font-bold",
                  isToolboxOpen ? "bg-indigo-600 border-indigo-500 text-white shadow-lg" : "text-zinc-300 border-zinc-800 hover:bg-zinc-900"
                )}
                title="Open Templates Toolbox"
                id="btn-toolbox"
              >
                <Archive size={16} />
                <span className="hidden sm:inline">Templates</span>
              </button>
              <div className="w-px h-6 bg-zinc-800" />
              <button 
                onClick={() => fileInputRef.current?.click()}
                className="flex items-center gap-2 px-3 py-1.5 text-zinc-350 border border-transparent hover:bg-zinc-900 hover:border-zinc-800 rounded-xl transition-all text-xs font-bold cursor-pointer"
                title="Import from Word"
                id="btn-import"
              >
                <Upload size={16} />
                <span className="hidden sm:inline">Import</span>
              </button>
              <input 
                type="file" 
                ref={fileInputRef} 
                className="hidden" 
                accept=".docx" 
                onChange={handleImport}
              />
              <button 
                onClick={() => setSavedToast("Itinerary saved for future edits!")}
                className="flex items-center gap-2 px-3 py-1.5 text-emerald-450 border border-emerald-500/20 bg-emerald-500/5 hover:bg-emerald-500/12 hover:text-emerald-300 rounded-xl transition-all text-xs font-bold cursor-pointer"
                title="Save Itinerary for Future Edits"
                id="btn-manual-save"
              >
                <Layers size={16} />
                <span className="hidden sm:inline">Save</span>
              </button>
              <button 
                onClick={exportToWord}
                className="flex items-center gap-2 px-3 py-1.5 bg-indigo-600 text-white hover:bg-indigo-550 rounded-xl transition-colors shadow-sm cursor-pointer text-xs font-bold"
                title="Export to Word"
                id="btn-export"
              >
                <FileDown size={16} />
                <span className="hidden sm:inline">Word</span>
              </button>
            </div>
          </div>
        </header>

        {/* Sidebar Layout */}
        <div className="flex-1 flex overflow-hidden">
          <main className={cn(
            "flex-1 overflow-y-auto px-4 md:px-8 py-4 transition-all duration-500 ease-in-out",
            isToolboxOpen ? "md:mr-80" : "mr-0"
          )}>
            <div className="max-w-4xl mx-auto space-y-8 pb-16" id="itinerary-days">
              {activeItinerary.days.length > 1 && (
                <div className="sticky top-0 z-30 bg-[#09090B]/85 backdrop-blur-md py-2 -mx-4 px-4 sm:-mx-8 sm:px-8 flex items-center gap-1.5 border-b border-zinc-800/60 overflow-x-auto no-scrollbar shadow-xs">
                  <span className="text-[10px] font-bold uppercase tracking-wider text-zinc-500 mr-2 shrink-0">Days Jump Bar:</span>
                  {activeItinerary.days.map((day, idx) => {
                    const parsedDate = parseISO(day.date);
                    return (
                      <button
                        key={day.id}
                        onClick={() => {
                          document.getElementById(`section-${day.id}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        }}
                        className="px-3 py-1 bg-zinc-900 hover:bg-zinc-800 hover:border-zinc-700 border border-zinc-800 rounded-full text-xs font-bold text-zinc-300 shadow-xs flex items-center gap-1.5 transition-all shrink-0 active:scale-95 cursor-pointer"
                      >
                        <span className="w-1.5 h-1.5 rounded-full bg-indigo-500 animate-pulse" />
                        Day {idx + 1} ({format(parsedDate, "MMM d")})
                      </button>
                    );
                  })}
                </div>
              )}
              {activeItinerary.days.map((day) => (
                <DaySection 
                  key={day.id} 
                  day={day} 
                  onUpdate={(updates) => updateDay(day.id, updates)}
                  onRemove={() => removeDay(day.id)}
                  onDuplicate={() => duplicateDay(day.id)}
                  onSaveToLibrary={saveToLibrary}
                  onSetUndo={setUndoAction}
                />
              ))}
              
              <button 
                onClick={addDay}
                className="w-full py-5 bg-[#18181B] border-2 border-dashed border-zinc-800 rounded-2xl text-zinc-400 hover:border-indigo-500/40 hover:text-indigo-400 transition-all flex flex-col items-center justify-center gap-1.5 group cursor-pointer"
                id="btn-add-day"
              >
                <div className="p-2 bg-zinc-900 rounded-full group-hover:bg-indigo-950/40 transition-colors">
                  <Plus size={20} className="text-zinc-500 group-hover:text-indigo-400" />
                </div>
                <span className="font-bold text-sm">Add Another Day</span>
              </button>
            </div>
          </main>

          {/* Toolbox Sidebar */}
          <motion.div 
            initial={false}
            animate={{ 
              x: isToolboxOpen ? 0 : "100%",
              opacity: isToolboxOpen ? 1 : 0
            }}
            transition={{ type: "spring", damping: 25, stiffness: 200 }}
            className={cn(
              "fixed right-0 top-0 h-full w-80 bg-[#18181B] z-50 flex flex-col border-l border-zinc-800 transition-shadow",
              isToolboxOpen ? "shadow-2xl shadow-black/80" : "shadow-none"
            )}
          >
            <div className="p-6 border-b border-zinc-800/80 flex items-center justify-between bg-zinc-900/30">
              <div className="flex items-center gap-2 font-bold text-zinc-150">
                <Archive size={20} className="text-indigo-400" />
                Toolbox
              </div>
              <button onClick={() => setIsToolboxOpen(false)} className="p-2 hover:bg-zinc-800 rounded-lg text-zinc-500 hover:text-zinc-300 transition-colors cursor-pointer">
                <X size={20} />
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-4 space-y-4">
              {templates.length === 0 ? (
                <div className="text-center py-12 px-4 space-y-3">
                  <div className="mx-auto w-12 h-12 bg-zinc-900 border border-zinc-800/60 rounded-full flex items-center justify-center text-zinc-600">
                    <Copy size={24} />
                  </div>
                  <p className="text-xs text-zinc-500 font-medium leading-relaxed">
                    Your library is empty. Click the archive icon on any itinerary item to save it here.
                  </p>
                </div>
              ) : (
                templates.map((tpl) => (
                  <motion.div 
                    key={tpl.id}
                    layout
                    draggable
                    onDragStart={(e) => {
                      e.dataTransfer.setData("application/chronos-template", JSON.stringify(tpl));
                      e.dataTransfer.effectAllowed = "copy";
                    }}
                    onClick={() => {
                      // Add to the first day by default on click
                      if (activeItinerary.days.length > 0) {
                        const item: ItineraryItem = {
                          id: `item-${Date.now()}`,
                          description: tpl.description || "",
                          duration: tpl.duration,
                          startTime: "0900", // Will be recalculated by DaySection
                          isDurationLocked: tpl.isDurationLocked,
                          isTimelineLocked: false,
                          notes: tpl.notes,
                          subItems: tpl.subItems ? JSON.parse(JSON.stringify(tpl.subItems)) : []
                        };
                        const firstDayId = activeItinerary.days[0].id;
                        updateDay(firstDayId, { items: [...(activeItinerary.days[0].items || []), item] });
                        setSavedToast(`Added ${tpl.name}`);
                      }
                    }}
                    className="group p-4 bg-zinc-900 border border-zinc-800/80 rounded-xl hover:border-indigo-500/40 hover:shadow-lg transition-all cursor-grab active:cursor-grabbing relative overflow-hidden active:scale-95"
                  >
                    <div className="absolute top-0 left-0 w-1 h-full bg-indigo-500 opacity-0 group-hover:opacity-100 transition-opacity" />
                    <div className="flex justify-between items-start mb-2">
                      <div className="font-bold text-sm text-zinc-200 truncate pr-6 group-hover:text-indigo-400 transition-colors">{tpl.name}</div>
                      <button 
                        onClick={(e) => {
                          e.stopPropagation();
                          removeTemplate(tpl.id);
                        }}
                        className="p-1 text-zinc-500 hover:text-rose-400 hover:bg-zinc-800 rounded transition-colors cursor-pointer"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                    <div className="flex items-center gap-3 text-[10px] font-bold text-zinc-500 uppercase tracking-widest">
                      <span className="flex items-center gap-1">
                        <Clock size={10} className="text-indigo-400" />
                        {tpl.duration}m
                      </span>
                      {tpl.subItems.length > 0 && (
                        <span className="flex items-center gap-1">
                          <Plus size={10} className="text-emerald-400" />
                          {tpl.subItems.length} steps
                        </span>
                      )}
                    </div>
                  </motion.div>
                ))
              )}
            </div>
            <div className="p-4 bg-zinc-900/60 text-[10px] text-zinc-500 italic text-center border-t border-zinc-800 font-bold uppercase tracking-wider">
              Click to add to Day 1, or drag into any timeline area.
            </div>
          </motion.div>
        </div>

      {/* Saved Toast */}
      <AnimatePresence>
        {savedToast && (
          <motion.div 
            initial={{ opacity: 0, scale: 0.5, y: -20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.5, y: -40 }}
            className="fixed top-8 left-1/2 -translate-x-1/2 z-[100] bg-zinc-900 border border-indigo-500/40 text-zinc-100 px-6 py-4 rounded-2xl shadow-2xl flex items-center gap-4 backdrop-blur-md bg-opacity-95"
          >
            <div className="w-10 h-10 bg-indigo-650/20 rounded-full flex items-center justify-center shadow-inner text-indigo-400 border border-indigo-550/20">
               <Archive size={20} />
            </div>
            <div>
              <div className="text-[10px] font-bold uppercase tracking-widest text-indigo-400 mb-0.5">Library Saved</div>
              <div className="text-sm font-bold truncate max-w-[200px] text-zinc-200">
                {savedToast}
              </div>
            </div>
            <motion.div 
              className="ml-2 w-2 h-2 rounded-full bg-indigo-500"
              animate={{ scale: [1, 1.5, 1] }}
              transition={{ repeat: Infinity, duration: 1 }}
            />
          </motion.div>
        )}
      </AnimatePresence>
    </div>
    </div>
  );
}

interface DaySectionProps {
  key?: React.Key;
  day: DayData;
  onUpdate: (updates: Partial<DayData>) => void;
  onRemove: () => void;
  onDuplicate: () => void;
  onSaveToLibrary: (item: ItineraryItem) => void;
  onSetUndo: (action: UndoAction) => void;
}

function DaySection({ day, onUpdate, onRemove, onDuplicate, onSaveToLibrary, onSetUndo }: DaySectionProps) {
  const [items, setItems] = useState(day.items || []);

  // Sync internal state with external updates
  useEffect(() => {
    setItems(day.items || []);
  }, [day.items]);

  const recalculate = useCallback((newItems: ItineraryItem[]) => {
    if (!newItems || newItems.length === 0) return [];
    
    const sorted = [...newItems];
    // Anchor to the first item's time
    let currentTime = militaryToMinutes(sorted[0].startTime);

    return sorted.map((item, idx) => {
      let startMins: number;

      if (idx === 0) {
        startMins = currentTime;
      } else if (item.isTimelineLocked) {
        startMins = militaryToMinutes(item.startTime);
      } else {
        startMins = currentTime;
      }

      const h = Math.floor(startMins / 60) % 24;
      const m = startMins % 60;
      const startStr = String(h).padStart(2, '0') + String(m).padStart(2, '0');
      
      currentTime = startMins + item.duration;
      
      return { ...item, startTime: startStr };
    });
  }, []);

  const handleReorder = (newOrder: ItineraryItem[]) => {
    // When reordering, we want to shift times based on the new sequence
    const recalculated = recalculate(newOrder);
    setItems(recalculated);
    onUpdate({ items: recalculated });
  };

  const addItem = () => {
    const newItem: ItineraryItem = {
      id: `item-${Date.now()}`,
      description: "",
      startTime: items.length > 0 ? addMinutesToMilitary(items[items.length - 1].startTime, items[items.length - 1].duration) : "0900",
      duration: 30,
      isDurationLocked: false,
      isTimelineLocked: false,
      notes: "",
      subItems: []
    };
    onUpdate({ items: [...items, newItem] });
  };

  const updateItem = (itemId: string, updates: Partial<ItineraryItem>) => {
    // Find the item and the index
    const index = items.findIndex(it => it.id === itemId);
    if (index === -1) return;

    const newItems = [...items];
    const oldItem = items[index];

    // Enforce 24-hour limit (1440 minutes max, 1 minute min) for parent event
    if (updates.duration !== undefined) {
      updates.duration = Math.min(1440, Math.max(1, updates.duration));
    }

    // Auto-lock when start time is edited directly unless explicitly toggled otherwise
    if (updates.startTime !== undefined && updates.isTimelineLocked === undefined) {
      updates.isTimelineLocked = true;
    }

    let updatedItem = { ...oldItem, ...updates };

    // Scaling logic for sub-items
    const subItemsList = updatedItem.subItems || [];
    if (updates.duration !== undefined && updates.subItems === undefined && subItemsList.length > 0) {
      const oldDuration = oldItem.duration;
      const newDuration = updates.duration;
      const subItems = [...subItemsList];
      
      const lockedDuration = subItems.filter(s => s.isDurationLocked).reduce((sum, s) => sum + s.duration, 0);
      const oldUnlockedDuration = oldDuration - lockedDuration;
      const newUnlockedDuration = newDuration - lockedDuration;
      
      if (oldUnlockedDuration > 0 && newUnlockedDuration > 0) {
        const ratio = newUnlockedDuration / oldUnlockedDuration;
        updatedItem.subItems = subItems.map(s => {
          if (s.isDurationLocked) return s;
          const scaledVal = Math.round(s.duration * ratio);
          const alignedVal = Math.round(scaledVal / 5) * 5;
          return { ...s, duration: Math.min(300, Math.max(5, alignedVal)) };
        });
        
        // Final adjustment to match exactly
        const currentSum = updatedItem.subItems.reduce((sum, s) => sum + s.duration, 0);
        if (currentSum !== newDuration) {
           const lastUnlockedIdx = [...updatedItem.subItems].reverse().findIndex(s => !s.isDurationLocked);
           if (lastUnlockedIdx !== -1) {
             const actualIdx = updatedItem.subItems.length - 1 - lastUnlockedIdx;
             const idealVal = updatedItem.subItems[actualIdx].duration + (newDuration - currentSum);
             const alignedIdeal = Math.round(idealVal / 5) * 5;
             updatedItem.subItems[actualIdx].duration = Math.min(300, Math.max(5, alignedIdeal));
           }
        }
      }
    }

    newItems[index] = updatedItem;

    // Recalculate using the standard helper which respects locking and propagates times correctly
    const result = recalculate(newItems);

    setItems(result);
    onUpdate({ items: result });
  };

  const removeItem = (itemId: string) => {
    const index = items.findIndex(it => it.id === itemId);
    if (index !== -1) {
      onSetUndo({
        type: 'item',
        data: items[index],
        index,
        dayId: day.id
      });
    }

    const updated = items.filter(it => it.id !== itemId);
    const recalculated = recalculate(updated);
    onUpdate({ items: recalculated });
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const dataStr = e.dataTransfer.getData("application/chronos-template");
    if (!dataStr) return;

    try {
      const template = JSON.parse(dataStr) as TemplateItem;
      const newItem: ItineraryItem = {
        ...template,
        id: `item-${Date.now()}`,
        description: template.description || "",
        startTime: items.length > 0 ? addMinutesToMilitary(items[items.length - 1].startTime, items[items.length - 1].duration) : "0900",
        isTimelineLocked: false,
        subItems: template.subItems.map(s => ({ ...s, id: `sub-${Date.now()}-${Math.random()}` }))
      };
      
      const updated = [...items, newItem];
      onUpdate({ items: recalculate(updated) });
    } catch (err) {
      console.error("Failed to parse template data", err);
    }
  };

  // Issue Detection (Overlap & Gap)
  const issues = new Map<string, 'overlap' | 'gap'>();
  for (let i = 1; i < items.length; i++) {
    const prev = items[i - 1];
    const curr = items[i];
    const prevEnd = militaryToMinutes(prev.startTime) + prev.duration;
    const currStart = militaryToMinutes(curr.startTime);

    if (currStart < prevEnd) {
      issues.set(curr.id, 'overlap');
    } else if (currStart > prevEnd) {
      issues.set(curr.id, 'gap');
    }
  }

  return (
    <section className="space-y-4" id={`section-${day.id}`}>
      {/* 3 Columns Header */}
      <div className="grid grid-cols-12 gap-4 items-start">
        {/* Column 1: Date Info */}
        <div className="col-span-12 lg:col-span-2 bg-[#18181B] rounded-2xl border border-zinc-800/80 p-4 shadow-md sticky top-4">
          <div className="space-y-0.5">
            <h3 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider flex items-center gap-1.5">
              <Calendar size={12} className="text-indigo-400" />
              Date
            </h3>
            <div className="text-xl font-extrabold text-[#FAFAFA]">{format(parseISO(day.date), "MMM d")}</div>
            <div className="text-zinc-400 text-xs font-semibold">{format(parseISO(day.date), "EEEE")}</div>
          </div>
          <div className="mt-3 pt-3 border-t border-zinc-800/80 flex flex-col gap-1.5">
            <button 
              onClick={onDuplicate}
              className="text-zinc-400 hover:text-indigo-400 p-1.5 hover:bg-indigo-500/10 rounded-lg transition-colors flex items-center justify-center gap-1.5 text-xs font-bold cursor-pointer"
              id={`duplicate-day-${day.id}`}
            >
              <Copy size={14} />
              Duplicate
            </button>
            <button 
              onClick={onRemove}
              className="text-rose-400 hover:text-rose-300 p-1.5 hover:bg-rose-500/10 rounded-lg transition-colors flex items-center justify-center gap-1.5 text-xs font-bold cursor-pointer"
              id={`remove-day-${day.id}`}
            >
              <Trash2 size={14} />
              Delete Day
            </button>
          </div>
        </div>

        {/* Column 2: Timeline items */}
        <div 
          className="col-span-12 lg:col-span-10 space-y-4"
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
        >
          <div className="grid grid-cols-1 md:grid-cols-[1fr,240px] gap-4">
            <div className="space-y-3">
              {/* Day Timeline Map */}
              {items.length > 0 && (
                <div className="bg-[#18181B] rounded-2xl border border-zinc-800/80 p-3 shadow-md space-y-2">
                  <div className="flex items-center justify-between text-[10px] font-bold text-zinc-500 uppercase tracking-wider px-1">
                    <span className="flex items-center gap-1.5">
                      <Clock size={12} className="text-indigo-400" />
                      Day Timeline Map
                    </span>
                    <span className="text-zinc-400 text-[10px]">
                      {items.length} events • {Math.floor(items.reduce((sum, s) => sum + s.duration, 0) / 60)}h {items.reduce((sum, s) => sum + s.duration, 0) % 60}m total
                    </span>
                  </div>
                  <div className="flex w-full h-8 rounded-xl overflow-hidden border border-zinc-800 bg-[#09090B] relative p-0.5 gap-0.5">
                    {items.map((item, idx) => {
                      const totalMins = items.reduce((sum, s) => sum + s.duration, 0);
                      const pct = totalMins > 0 ? (item.duration / totalMins) * 105 : 0;
                      const colorClass = [
                        "bg-indigo-950/45 text-indigo-300 border-indigo-900/40 hover:bg-indigo-900/30 hover:text-indigo-200",
                        "bg-violet-950/45 text-violet-300 border-violet-900/40 hover:bg-violet-900/30 hover:text-violet-200",
                        "bg-purple-950/45 text-purple-300 border-purple-900/40 hover:bg-purple-900/30 hover:text-purple-200",
                        "bg-emerald-950/45 text-emerald-300 border-emerald-900/40 hover:bg-emerald-900/30 hover:text-emerald-200",
                        "bg-amber-950/45 text-amber-300 border-amber-900/40 hover:bg-amber-900/30 hover:text-amber-200",
                        "bg-rose-950/45 text-rose-300 border-rose-900/40 hover:bg-rose-900/30 hover:text-rose-200",
                      ][idx % 6];

                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            const containerEl = document.getElementById(`item-container-${item.id}`);
                            if (containerEl) {
                              containerEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
                              containerEl.classList.add('ring-4', 'ring-indigo-500/20', 'border-indigo-500');
                              setTimeout(() => {
                                containerEl.classList.remove('ring-4', 'ring-indigo-500/20', 'border-indigo-500');
                              }, 2000);
                            }
                          }}
                          style={{ width: `${pct}%` }}
                          className={cn(
                            "h-full rounded-lg border text-[9px] font-bold flex items-center justify-center px-1.5 cursor-pointer transition-all truncate hover:-translate-y-0.5 hover:shadow-md select-none pr-2",
                            colorClass
                          )}
                          title={`${formatMilitary(item.startTime)} - ${item.description || "Unnamed"} (${item.duration} min)`}
                        >
                          <span className="truncate">{item.description || formatMilitary(item.startTime)}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              <Reorder.Group axis="y" values={items} onReorder={handleReorder} className="space-y-2">
                <AnimatePresence initial={false}>
                  {items.map((item, idx) => (
                    <ReorderItem 
                      key={item.id} 
                      item={item} 
                      issue={issues.get(item.id)}
                      prevItem={idx > 0 ? items[idx - 1] : undefined}
                      onUpdate={(updates) => updateItem(item.id, updates)}
                      onRemove={() => removeItem(item.id)}
                      onSaveToLibrary={() => onSaveToLibrary(item)}
                    />
                  ))}
                </AnimatePresence>
              </Reorder.Group>
              
              <button 
                onClick={addItem}
                className="w-full h-11 flex items-center justify-center border-2 border-dashed border-zinc-800 bg-[#18181B] rounded-xl text-zinc-400 hover:border-indigo-500/40 hover:text-indigo-400 transition-all font-bold text-xs gap-2 cursor-pointer"
                id={`add-item-${day.id}`}
              >
                <Plus size={16} />
                Add Item
              </button>
            </div>

            {/* Column 3: Day Summary / Notes space */}
            <div className="hidden md:block">
              <div className="bg-zinc-900/25 rounded-2xl p-4 h-full border border-zinc-800/80 border-dashed sticky top-4">
                <div className="flex items-center gap-2 text-zinc-500 mb-3 font-bold uppercase text-[10px] tracking-widest">
                  <FileText size={14} className="text-zinc-650" />
                  Workspace
                </div>
                <div className="space-y-2">
                  <div className="text-[10px] text-zinc-500 leading-relaxed italic font-medium">
                    Drag items to reorder. Locked items stay fixed.
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

interface ReorderItemProps {
  key?: React.Key;
  item: ItineraryItem;
  issue?: 'overlap' | 'gap';
  prevItem?: ItineraryItem;
  onUpdate: (updates: Partial<ItineraryItem>) => void;
  onRemove: () => void;
  onSaveToLibrary: () => void;
}

function ReorderItem({ item, issue, prevItem, onUpdate, onRemove, onSaveToLibrary }: ReorderItemProps) {
  const [showSubItems, setShowSubItems] = useState(false);
  const [localTime, setLocalTime] = useState(item.startTime);
  const [localEndTime, setLocalEndTime] = useState(addMinutesToMilitary(item.startTime, item.duration));

  const isValidTime = (time: string) => {
    if (time.length === 0) return true;
    if (time.length < 4) return false;
    const hh = parseInt(time.substring(0, 2));
    const mm = parseInt(time.substring(2, 4));
    if (hh === 24 && mm === 0) return true;
    return hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59;
  };

  const totalSubItemDuration = (item.subItems || []).reduce((acc, sub) => acc + Number(sub.duration), 0);
  const hasDurationMismatch = (item.subItems || []).length > 0 && Number(totalSubItemDuration) !== Number(item.duration);
  const discrepancy = Number(item.duration) - Number(totalSubItemDuration);

  const isInvalid = localTime.length > 0 && !isValidTime(localTime.padStart(4, "0"));

  // Sync local time if global time changes (e.g. shifted by others)
  useEffect(() => {
    setLocalTime(item.startTime);
    setLocalEndTime(addMinutesToMilitary(item.startTime, item.duration));
  }, [item.startTime, item.duration]);

  const handleTimeBlur = () => {
    const padded = localTime.padStart(4, "0");
    if (isValidTime(padded)) {
      setLocalTime(padded);
      onUpdate({ startTime: padded });
    } else {
      setLocalTime(item.startTime); // Revert on invalid
    }
  };

  const handleEndTimeBlur = () => {
    const padded = localEndTime.padStart(4, "0");
    if (isValidTime(padded)) {
      setLocalEndTime(padded);
      const newDuration = getDuration(item.startTime, padded);
      // Enforce 24-hour limit (1440 minutes max)
      onUpdate({ duration: Math.min(1440, newDuration) });
    } else {
      setLocalEndTime(addMinutesToMilitary(item.startTime, item.duration));
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent, blurHandler: () => void) => {
    if (e.key === "Enter") {
      blurHandler();
      (e.target as HTMLInputElement).blur();
    }
  };

  const addSubItem = () => {
    const newSub: ItinerarySubItem = { id: `sub-${Date.now()}`, description: "", duration: 15, isDurationLocked: false };
    const currentSubItems = item.subItems || [];
    const newSubItems = [...currentSubItems, newSub];
    const newTotalDuration = newSubItems.reduce((acc, s) => acc + Number(s.duration), 0);
    onUpdate({ 
      subItems: newSubItems,
      ...(item.isDurationLocked ? {} : { duration: newTotalDuration })
    });
    setShowSubItems(true);
  };

  const syncParentToSubItems = () => {
    onUpdate({ duration: totalSubItemDuration });
  };

  const distributeDiscrepancyToLastSubItem = () => {
    const currentSubs = item.subItems || [];
    if (currentSubs.length === 0) return;
    const newSubs = [...currentSubs];
    const lastUnlockedIdx = [...newSubs].reverse().findIndex(s => !s.isDurationLocked);
    const targetIdx = lastUnlockedIdx !== -1 ? (newSubs.length - 1 - lastUnlockedIdx) : (newSubs.length - 1);
    
    const newDur = Math.max(1, Number(newSubs[targetIdx].duration) + discrepancy);
    newSubs[targetIdx] = { ...newSubs[targetIdx], duration: newDur };
    
    const newTotal = newSubs.reduce((acc, s) => acc + Number(s.duration), 0);
    onUpdate({ subItems: newSubs, duration: newTotal });
  };

  return (
    <Reorder.Item 
      value={item}
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, scale: 0.95 }}
      whileDrag={{ scale: 1.02, boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.7)" }}
      className={cn(
        "relative group bg-[#18181B] rounded-2xl border transition-all duration-300 shadow-md overflow-hidden",
        issue === 'overlap' 
          ? "border-rose-500/50 bg-rose-950/15 shadow-rose-950/50" 
          : issue === 'gap'
            ? "border-amber-500/50 bg-amber-950/15 shadow-amber-950/50"
            : "border-zinc-800 hover:border-zinc-700 shadow-black/20"
      )}
      id={`item-container-${item.id}`}
    >
      <div className="flex items-stretch overflow-hidden">
        {/* Left Drag Handle & Overlap Indicator */}
        <div className={cn(
          "w-10 flex flex-col items-center justify-center border-r transition-colors shrink-0",
          issue === 'overlap'
            ? "border-rose-950/50 bg-rose-950/20"
            : issue === 'gap'
              ? "border-amber-950/50 bg-amber-950/20"
              : "border-zinc-800/60 bg-zinc-900/40 group-hover:bg-zinc-850/40"
        )}>
          <GripVertical size={16} className={cn(
            "cursor-grab active:cursor-grabbing transition-colors",
            issue === 'overlap'
              ? "text-rose-400 group-hover:text-rose-300"
              : issue === 'gap'
                ? "text-amber-400 group-hover:text-amber-300"
                : "text-zinc-650 group-hover:text-indigo-400"
          )} />
          {issue && (
            <div className={cn(
              "mt-1.5 transition-colors",
              issue === 'overlap' ? "text-rose-455 animate-pulse" : "text-amber-455"
            )} title={issue === 'overlap' ? "Schedule Overlap Detected" : "Timeline Gap Detected"}>
              <AlertCircle size={11} />
            </div>
          )}
        </div>

        {/* Content */}
        <div className="flex-1 p-3 grid grid-cols-1 md:grid-cols-12 gap-3">
          {/* Time & Duration */}
          <div className="md:col-span-3 flex flex-col justify-center gap-2 border-r md:border-none border-zinc-800/50 pr-3 md:pr-0">
            <div className="flex items-center gap-1.5">
              <div className="flex items-center gap-1 bg-[#09090B] p-1 rounded-xl border border-zinc-800 shadow-inner shrink-0">
                <input 
                  type="text"
                  maxLength={4}
                  className={cn(
                    "w-12 text-[10px] font-mono font-bold border px-1 py-0.5 rounded outline-none transition-all text-center",
                    isInvalid ? "border-rose-500 ring-2 ring-rose-900/20 bg-rose-950/40 text-rose-200" : (
                      item.isTimelineLocked 
                        ? "bg-amber-500/10 text-amber-300 border-amber-500/30 focus:ring-1 focus:ring-amber-400" 
                        : "bg-[#09090B] text-zinc-100 border-zinc-850/80 focus:ring-1 focus:ring-indigo-500"
                    )
                  )}
                  value={localTime}
                  onChange={(e) => {
                    const val = e.target.value.replace(/\D/g, "");
                    if (val.length <= 4) setLocalTime(val);
                  }}
                  onBlur={handleTimeBlur}
                  onKeyDown={(e) => handleKeyDown(e, handleTimeBlur)}
                />
                <span className="text-zinc-650 text-[10px] font-bold">-</span>
                <input 
                  type="text"
                  maxLength={4}
                  className={cn(
                    "w-12 text-[10px] font-mono font-bold border px-1 py-0.5 rounded outline-none transition-all text-center",
                    "bg-zinc-800/40 text-zinc-300 border-zinc-850/80 focus:ring-1 focus:ring-indigo-500 focus:bg-zinc-900 focus:text-white"
                  )}
                  value={localEndTime}
                  onChange={(e) => {
                    const val = e.target.value.replace(/\D/g, "");
                    if (val.length <= 4) setLocalEndTime(val);
                  }}
                  onBlur={handleEndTimeBlur}
                  onKeyDown={(e) => handleKeyDown(e, handleEndTimeBlur)}
                />
              </div>
              <button 
                onClick={() => onUpdate({ isTimelineLocked: !item.isTimelineLocked })}
                className={cn(
                  "p-1 rounded-lg transition-all shrink-0 cursor-pointer",
                  item.isTimelineLocked ? "bg-amber-500/15 text-amber-400" : "text-zinc-500 hover:text-zinc-300 hover:bg-zinc-800"
                )}
                title={item.isTimelineLocked ? "Locked to timeline" : "Click to lock to timeline"}
              >
                {item.isTimelineLocked ? <Lock size={12} /> : <Unlock size={12} />}
              </button>
            </div>

            {issue && prevItem && (
              <div className={cn(
                "flex flex-col gap-1.5 p-2 rounded-xl border text-[8px] leading-tight shadow-md w-full max-w-[200px] animate-in slide-in-from-top-1 duration-250",
                issue === "overlap" ? "bg-rose-950/80 border-rose-900/40 text-rose-200" : "bg-amber-950/80 border-amber-900/40 text-amber-200"
              )}>
                <div className={cn(
                  "font-black uppercase tracking-wider flex items-center gap-1 text-[9px]",
                  issue === "overlap" ? "text-rose-450" : "text-amber-450"
                )}>
                  <AlertCircle size={10} className="shrink-0 animate-bounce" />
                  <span>Time {issue === "overlap" ? "Overlap" : "Gap"}</span>
                </div>
                <div className="flex flex-col gap-1">
                  <button
                    onClick={() => onUpdate({ 
                      startTime: addMinutesToMilitary(prevItem.startTime, prevItem.duration),
                      isTimelineLocked: true
                    })}
                    className={cn(
                      "w-full text-left px-2 py-1 rounded-lg font-bold transition-all text-[8px] active:scale-95 cursor-pointer truncate text-white shadow-xs",
                      issue === "overlap" ? "bg-rose-650 hover:bg-rose-700" : "bg-amber-650 hover:bg-amber-700"
                    )}
                    title={`Snap start time to exactly follow the previous item ending at ${formatMilitary(addMinutesToMilitary(prevItem.startTime, prevItem.duration))}`}
                  >
                    Snap to {formatMilitary(addMinutesToMilitary(prevItem.startTime, prevItem.duration))}
                  </button>
                  <button
                    onClick={() => onUpdate({ isTimelineLocked: false })}
                    className={cn(
                      "w-full text-left px-2 py-1 bg-zinc-900 border border-zinc-800 rounded-lg font-bold transition-all text-[8px] active:scale-95 cursor-pointer truncate shadow-xs text-zinc-300 hover:bg-zinc-800 hover:text-white"
                    )}
                    title="Unlock start time to let it flow automatically from the previous event"
                  >
                    Auto-Flow Time
                  </button>
                </div>
              </div>
            )}
            
            <div className="flex items-center gap-1.5">
              <div className={cn(
                "flex items-center border rounded-lg overflow-hidden transition-all duration-300 shrink-0",
                hasDurationMismatch 
                  ? "border-rose-400 bg-rose-950/20 ring-2 ring-rose-900/20" 
                  : (item.isDurationLocked ? "bg-zinc-800 border-zinc-700/60" : "bg-zinc-900 border-zinc-800")
              )}>
                <input 
                  type="number" 
                  min="1"
                  max="1440"
                  className={cn(
                    "w-10 text-center text-[10px] font-bold bg-transparent outline-none py-0.5",
                    hasDurationMismatch ? "text-rose-400" : "text-zinc-200"
                  )}
                  value={item.duration}
                  disabled={item.isDurationLocked}
                  onChange={(e) => {
                    const typedVal = parseInt(e.target.value) || 0;
                    onUpdate({ duration: Math.min(1440, Math.max(1, typedVal)) });
                  }}
                />
                <span className={cn(
                  "text-[8px] uppercase font-bold px-1 flex items-center h-full transition-colors",
                  hasDurationMismatch ? "bg-rose-900/40 text-rose-300" : "bg-zinc-800 text-zinc-500"
                )}>min</span>
              </div>
              {hasDurationMismatch && (
                <button 
                  onClick={() => setShowSubItems(true)}
                  className="flex items-center gap-0.5 text-[9px] font-black text-rose-450 bg-rose-950/30 border border-rose-800/40 rounded px-1.5 py-0.5 cursor-pointer hover:bg-rose-900/40 transition-colors shrink-0"
                  title={`Sub-items sum to ${totalSubItemDuration}m, but overall event is ${item.duration}m. Click to expand steps.`}
                >
                  <AlertCircle size={10} className="shrink-0 animate-bounce" />
                  <span>{discrepancy > 0 ? `+${discrepancy}m` : `${discrepancy}m`}</span>
                </button>
              )}
              <button 
                onClick={() => onUpdate({ isDurationLocked: !item.isDurationLocked })}
                className={cn(
                   "p-0.5 rounded-md transition-all text-[10px] shrink-0 cursor-pointer",
                   item.isDurationLocked ? "text-indigo-400 bg-indigo-950/40" : "text-zinc-500 hover:text-zinc-300 font-bold"
                )}
                title={item.isDurationLocked ? "Duration Locked" : "Click to lock duration"}
              >
                {item.isDurationLocked ? <Lock size={10} /> : <Unlock size={10} />}
              </button>
            </div>

            {hasDurationMismatch && (
              <div className="flex flex-col gap-1.5 p-2 bg-rose-950/80 border border-rose-900/40 rounded-xl text-[8px] leading-tight shadow-md w-full max-w-[200px] animate-in slide-in-from-top-1 duration-250">
                <div className="text-rose-450 font-extrabold uppercase tracking-wider flex items-center gap-1 text-[9px]">
                  <AlertCircle size={10} className="shrink-0 animate-pulse" />
                  <span>Mismatch Fix</span>
                </div>
                <div className="flex flex-col gap-1">
                  <button
                    onClick={syncParentToSubItems}
                    disabled={item.isDurationLocked}
                    className={cn(
                      "w-full text-left px-2 py-1 rounded-lg font-bold transition-all text-[8px] active:scale-95 cursor-pointer truncate shadow-xs",
                      item.isDurationLocked 
                        ? "bg-zinc-800 text-zinc-500 cursor-not-allowed border border-zinc-700/60" 
                        : "bg-rose-650 hover:bg-rose-700 text-white"
                    )}
                    title={item.isDurationLocked ? "Unlock overall duration to sync event time" : `Adjust parent event duration to exactly ${totalSubItemDuration} minutes`}
                  >
                    Set Event: {totalSubItemDuration}m
                  </button>
                  {item.subItems.length > 0 && (
                    <button
                      onClick={distributeDiscrepancyToLastSubItem}
                      className="w-full text-left px-2 py-1 bg-zinc-900 text-rose-350 border border-rose-900/50 hover:bg-rose-900/20 rounded-lg font-bold transition-all text-[8px] active:scale-95 cursor-pointer truncate shadow-xs"
                      title={`Adjust the last step by ${discrepancy > 0 ? `+${discrepancy}` : discrepancy} to align perfectly`}
                    >
                      Set Steps: {item.duration}m
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Description & Notes */}
          <div className="md:col-span-12 lg:col-span-7 flex flex-col gap-1 pt-1 md:pt-0">
            <input 
              className="font-bold text-base bg-transparent border-none outline-none placeholder:text-zinc-650 focus:ring-0 leading-tight text-[#FAFAFA]"
              placeholder="Agenda item name..."
              value={item.description}
              onChange={(e) => onUpdate({ description: e.target.value })}
            />
            <textarea 
              className="text-xs text-zinc-455 bg-transparent border-none outline-none placeholder:text-zinc-650 resize-none h-8 min-h-[32px] focus:ring-0 leading-snug"
              placeholder="Notes, details, locations..."
              value={item.notes}
              onChange={(e) => onUpdate({ notes: e.target.value })}
            />
            {!showSubItems && (item.subItems || []).length > 0 && (
              <motion.div 
                initial={{ opacity: 0, y: -5 }}
                animate={{ opacity: 1, y: 0 }}
                className={cn(
                  "mt-1.5 flex items-center gap-1.5 text-[9px] font-bold w-fit px-2.5 py-0.5 rounded-full border cursor-pointer transition-colors hover:scale-102 dynamic-discrepancy-badge",
                  hasDurationMismatch 
                    ? "text-rose-450 bg-rose-950/35 border-rose-900/40 hover:bg-rose-900/45" 
                    : "text-indigo-450 bg-indigo-950/35 border-indigo-900/40 hover:bg-indigo-900/45"
                )}
                onClick={() => setShowSubItems(true)}
              >
                <Layers size={8} />
                <span>{item.subItems.length} COLLAPSED STEPS</span>
                {hasDurationMismatch && (
                  <span className="bg-rose-900/40 text-rose-300 px-1.5 py-0.2 rounded text-[8px] font-black ml-1">
                    ({discrepancy > 0 ? `+${discrepancy}m leftover` : `${discrepancy}m over`})
                  </span>
                )}
              </motion.div>
            )}
          </div>

          {/* Actions */}
          <div className="md:col-span-2 flex md:flex-col items-center justify-center gap-1 border-t md:border-none border-zinc-800/50 pt-2 md:pt-0">
             <button 
               onClick={onSaveToLibrary}
               className="p-1 text-zinc-500 hover:text-indigo-400 hover:bg-indigo-950/30 rounded-lg transition-all cursor-pointer"
               title="Save structure to Toolbox"
             >
               <Archive size={16} />
             </button>
             <button 
               onClick={onRemove}
               className="p-1 text-zinc-500 hover:text-rose-400 hover:bg-rose-950/30 rounded-lg transition-all cursor-pointer"
               title="Remove Item"
             >
               <Trash2 size={16} />
             </button>
             <button 
               onClick={addSubItem}
               className="p-1 text-zinc-500 hover:text-indigo-400 hover:bg-indigo-950/30 rounded-lg transition-all cursor-pointer"
               title="Add Sub-item"
             >
               <Plus size={16} />
             </button>
             <button 
               onClick={() => setShowSubItems(!showSubItems)}
               className={cn(
                 "p-1 text-zinc-500 hover:text-zinc-300 hover:bg-zinc-800 rounded-lg transition-all relative cursor-pointer", 
                 showSubItems && "rotate-180"
               )}
             >
               <ChevronDown size={16} />
               {!showSubItems && (item.subItems || []).length > 0 && (
                 <span className="absolute -top-0.5 -right-0.5 flex h-3 w-3 items-center justify-center rounded-full bg-indigo-550 text-[6px] font-bold text-white shadow-sm ring-1 ring-[#18181B] animate-in zoom-in">
                   {item.subItems.length}
                 </span>
               )}
             </button>
          </div>
        </div>
      </div>

      {/* Sub-items */}
      <AnimatePresence>
        {showSubItems && (item.subItems || []).length > 0 && (
          <motion.div 
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            className="border-t border-zinc-800 bg-zinc-900/10 overflow-hidden"
          >
            <div className="p-2 pl-12 space-y-1">
              <div className={cn(
                "text-[9px] font-bold uppercase tracking-widest flex flex-wrap items-center gap-2 mb-2 pb-1.5 border-b border-dashed border-zinc-805/60 transition-colors",
                hasDurationMismatch ? "text-rose-450" : "text-zinc-500"
              )}>
                <div className="flex items-center gap-1.5 mr-auto">
                  <div className={cn("w-1.5 h-1.5 rounded-full", hasDurationMismatch ? "bg-rose-500 animate-ping" : "bg-zinc-700")} />
                  {hasDurationMismatch ? "Duration Mismatch" : "Linked Steps"}
                  {hasDurationMismatch && (
                    <span className="bg-rose-950/50 text-rose-300 px-1.5 py-0.5 rounded border border-rose-900/40 font-bold ml-1 text-[8px] tracking-wide normal-case inline-block">
                      Steps total is {totalSubItemDuration}m, parent is {item.duration}m
                    </span>
                  )}
                </div>
                {hasDurationMismatch && (
                  <div className="flex items-center gap-1.5 py-0.5 px-1 bg-rose-950/40 rounded-lg border border-rose-900/30">
                    <span className="text-[8px] text-rose-450 font-bold lowercase">Fix:</span>
                    <button
                      onClick={syncParentToSubItems}
                      disabled={item.isDurationLocked}
                      className={cn(
                        "px-2 py-0.5 rounded font-extrabold transition-all active:scale-95 cursor-pointer text-[8px] tracking-tight uppercase shadow-xs shrink-0",
                        item.isDurationLocked
                          ? "bg-zinc-800 text-zinc-500 cursor-not-allowed border border-zinc-700/60"
                          : "bg-rose-650 hover:bg-rose-700 text-white"
                      )}
                      title={item.isDurationLocked ? "Unlock overall duration to sync event time" : `Adjust parent item duration to exactly ${totalSubItemDuration} minutes`}
                    >
                      Make event {totalSubItemDuration}m
                    </button>
                    {item.subItems.length > 0 && (
                      <button
                        onClick={distributeDiscrepancyToLastSubItem}
                        className="px-2 py-0.5 bg-zinc-900 hover:bg-zinc-850 text-rose-350 border border-rose-900/50 rounded font-extrabold transition-all active:scale-95 cursor-pointer text-[8px] tracking-tight uppercase shadow-xs shrink-0"
                        title={`Adjust the last step by ${discrepancy > 0 ? `+${discrepancy}` : discrepancy} to align perfectly`}
                      >
                        Adjust steps to {item.duration}m
                      </button>
                    )}
                  </div>
                )}
              </div>
              {(item.subItems || []).map((sub, idx) => (
                <div key={sub.id} className="flex items-center gap-4 bg-[#18181B] rounded-xl border border-zinc-800 p-2 shadow-sm group/sub">
                   <div className="text-xs font-mono font-bold text-indigo-400 bg-indigo-950/45 px-2 py-0.5 rounded border border-indigo-900/40">
                      {formatMilitary(calculateSubItemTime(item.startTime, item.subItems || [], idx))}
                   </div>
                   <input 
                     className="flex-1 text-sm bg-transparent border-none outline-none focus:ring-0 text-[#FAFAFA] placeholder:text-zinc-600"
                     value={sub.description}
                     placeholder="Sub-item description..."
                     onChange={(e) => {
                       const currentSubs = item.subItems || [];
                       const newSubs = [...currentSubs];
                       newSubs[idx] = { ...newSubs[idx], description: e.target.value };
                       onUpdate({ subItems: newSubs });
                     }}
                   />
                    <div className="flex items-center gap-1.5 opacity-0 group-hover/sub:opacity-100 transition-opacity">
                      <div className="flex items-center bg-[#09090B] border border-zinc-800 rounded-lg overflow-hidden h-7">
                        <button 
                          onClick={() => {
                            const currentSubs = item.subItems || [];
                            const newSubs = [...currentSubs];
                            const newSubDuration = Math.min(300, Math.max(5, (Math.ceil(sub.duration / 5) - 1) * 5));
                            newSubs[idx] = { ...newSubs[idx], duration: newSubDuration };
                            const newTotalDuration = newSubs.reduce((acc, s) => acc + s.duration, 0);
                            onUpdate({ 
                              subItems: newSubs, 
                              ...(item.isDurationLocked ? {} : { duration: newTotalDuration })
                            });
                          }}
                          disabled={sub.isDurationLocked}
                          className="px-1 h-full text-zinc-500 hover:text-indigo-400 hover:bg-zinc-800 disabled:opacity-30 transition-colors cursor-pointer"
                        >
                          <ChevronDown size={14} />
                        </button>
                        <input 
                          type="number"
                          min="5"
                          max="300"
                          step="5"
                          className="w-10 text-[10px] font-bold bg-transparent text-center outline-none text-zinc-200 disabled:text-zinc-650"
                          value={sub.duration}
                          disabled={sub.isDurationLocked}
                          onChange={(e) => {
                            const currentSubs = item.subItems || [];
                            const newSubs = [...currentSubs];
                            const parsedValue = parseInt(e.target.value) || 5;
                            const alignedValue = Math.round(Math.min(300, Math.max(5, parsedValue)) / 5) * 5;
                            const finalValue = Math.min(300, Math.max(5, alignedValue));
                            newSubs[idx] = { ...newSubs[idx], duration: finalValue };
                            const newTotalDuration = newSubs.reduce((acc, s) => acc + s.duration, 0);
                            onUpdate({ 
                              subItems: newSubs, 
                              ...(item.isDurationLocked ? {} : { duration: newTotalDuration })
                            });
                          }}
                        />
                        <button 
                          onClick={() => {
                            const currentSubs = item.subItems || [];
                            const newSubs = [...currentSubs];
                            const newSubDuration = Math.min(300, Math.max(5, (Math.floor(sub.duration / 5) + 1) * 5));
                            newSubs[idx] = { ...newSubs[idx], duration: newSubDuration };
                            const newTotalDuration = newSubs.reduce((acc, s) => acc + s.duration, 0);
                            onUpdate({ 
                              subItems: newSubs, 
                              ...(item.isDurationLocked ? {} : { duration: newTotalDuration })
                            });
                          }}
                          disabled={sub.isDurationLocked}
                          className="px-1 h-full text-zinc-500 hover:text-indigo-400 hover:bg-zinc-800 disabled:opacity-30 transition-colors cursor-pointer"
                        >
                          <ChevronUp size={14} />
                        </button>
                      </div>
                      <span className="text-[8px] text-zinc-500 font-bold">m</span>
                      <button 
                        onClick={() => {
                          const currentSubs = item.subItems || [];
                          const newSubs = [...currentSubs];
                          newSubs[idx] = { ...newSubs[idx], isDurationLocked: !sub.isDurationLocked };
                          onUpdate({ subItems: newSubs });
                        }}
                        className={cn(
                          "p-1 rounded-md transition-all cursor-pointer",
                          sub.isDurationLocked ? "text-indigo-400 bg-indigo-950/40" : "text-zinc-500 hover:text-zinc-300"
                        )}
                        title={sub.isDurationLocked ? "Unlock sub-item duration" : "Lock sub-item duration"}
                      >
                        {sub.isDurationLocked ? <Lock size={12} /> : <Unlock size={12} />}
                      </button>
                     <button 
                        onClick={() => {
                          const currentSubs = item.subItems || [];
                          const newSubs = currentSubs.filter(s => s.id !== sub.id);
                          const newTotalDuration = newSubs.reduce((acc, s) => acc + s.duration, 0);
                          onUpdate({ 
                            subItems: newSubs,
                            ...(item.isDurationLocked 
                              ? {} 
                              : { duration: newSubs.length > 0 ? newTotalDuration : item.duration })
                          });
                        }}
                        className="text-zinc-500 hover:text-rose-450 ml-1 cursor-pointer"
                      >
                        <Trash2 size={12} />
                      </button>
                    </div>
                </div>
              ))}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </Reorder.Item>
  );
}

function calculateSubItemTime(parentStart: string, subItems: ItinerarySubItem[], index: number): string {
  let currentTime = parentStart;
  for (let i = 0; i < index; i++) {
    currentTime = addMinutesToMilitary(currentTime, subItems[i].duration);
  }
  return currentTime;
}

function parseDateCell(dateText: string, fallbackDate: Date): Date {
  const clean = dateText.replace(/\s+/g, " ").trim();
  const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const words = clean.toLowerCase().split(/[\s,]+/);
  let dayNum = NaN;
  let monthIdx = NaN;
  const yearNum = fallbackDate.getFullYear();

  for (const word of words) {
    const matchMonth = months.findIndex(m => word.startsWith(m));
    if (matchMonth !== -1) {
      monthIdx = matchMonth;
    } else {
      const parsedNum = parseInt(word, 10);
      if (!isNaN(parsedNum) && parsedNum > 0 && parsedNum <= 31) {
        dayNum = parsedNum;
      }
    }
  }

  if (!isNaN(dayNum) && !isNaN(monthIdx)) {
    return new Date(yearNum, monthIdx, dayNum);
  }
  return fallbackDate;
}
