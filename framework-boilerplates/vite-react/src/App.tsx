import React, { useEffect, useMemo, useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { Separator } from "@/components/ui/separator";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import {
  Trash2,
  Plus,
  CalendarDays,
  BookOpen,
  AlertTriangle,
  Download,
  Upload,
  Save,
  GripVertical,
  ChevronUp,
  ChevronDown,
  RefreshCw,
  Lock,
  LockOpen,
} from "lucide-react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  DndContext,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
} from "@dnd-kit/core";
import {
  SortableContext,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";

const STORAGE_KEY = "progression-pedagogique-stmg-v3";

const VACATION_PRESETS = {
  "2026-2027": {
    A: [
      { label: "Vacances de Toussaint", start: "2026-10-17", end: "2026-11-02", type: "vacances" },
      { label: "Vacances de Noël", start: "2026-12-19", end: "2027-01-04", type: "vacances" },
      { label: "Vacances d'hiver", start: "2027-02-06", end: "2027-02-22", type: "vacances" },
      { label: "Vacances de printemps", start: "2027-04-10", end: "2027-04-26", type: "vacances" }
    ],
    B: [
      { label: "Vacances de Toussaint", start: "2026-10-17", end: "2026-11-02", type: "vacances" },
      { label: "Vacances de Noël", start: "2026-12-19", end: "2027-01-04", type: "vacances" },
      { label: "Vacances d'hiver", start: "2027-02-13", end: "2027-03-01", type: "vacances" },
      { label: "Vacances de printemps", start: "2027-04-03", end: "2027-04-19", type: "vacances" }
    ],
    C: [
      { label: "Vacances de Toussaint", start: "2026-10-17", end: "2026-11-02", type: "vacances" },
      { label: "Vacances de Noël", start: "2026-12-19", end: "2027-01-04", type: "vacances" },
      { label: "Vacances d'hiver", start: "2027-02-20", end: "2027-03-08", type: "vacances" },
      { label: "Vacances de printemps", start: "2027-04-17", end: "2027-05-03", type: "vacances" }
    ]
  }
};

const DAYS = [
  { key: "monday", label: "Lundi" },
  { key: "tuesday", label: "Mardi" },
  { key: "wednesday", label: "Mercredi" },
  { key: "thursday", label: "Jeudi" },
  { key: "friday", label: "Vendredi" },
  { key: "saturday", label: "Samedi" },
  { key: "sunday", label: "Dimanche" },
];

const TYPE_LABELS = {
  course: "Cours",
  assessment: "Évaluation",
  correction: "Correction",
  buffer: "Tampon",
};

const uid = () =>
  typeof crypto !== "undefined" && crypto.randomUUID
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

const dayKeyFromDate = (date) => {
  const jsDay = date.getDay();
  const map = {
    0: "sunday",
    1: "monday",
    2: "tuesday",
    3: "wednesday",
    4: "thursday",
    5: "friday",
    6: "saturday",
  };
  return map[jsDay];
};

const formatDate = (dateStr) => {
  if (!dateStr) return "";
  const date = new Date(`${dateStr}T00:00:00`);
  return date.toLocaleDateString("fr-FR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
};

const addDays = (date, days) => {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
};

const toISODate = (date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

const hoursToSlots = (hours) => Math.round(Number(hours || 0) * 2);
const slotsToHoursNumber = (slots) => Number((slots / 2).toFixed(1));
const slotsToHoursLabel = (slots) => String((slots / 2).toFixed(1)).replace(".0", "");

const createDefaultClass = () => ({
  id: uid(),
  meta: {
    className: "1re STMG A",
    subject: "Sciences de gestion et numérique",
    schoolYear: "2026-2027",
    teacher: "",
    zone: "A"
  },
  program: [
    {
      id: uid(),
      title: "Thème 1 - Les organisations et leurs acteurs",
      sequences: [
        {
          id: uid(),
          title: "Chapitre 1 - Identifier les caractéristiques d’une organisation",
          estimatedHours: 4,
          competencies: [
            "Identifier les caractéristiques d’une organisation",
            "Mobiliser le vocabulaire disciplinaire"
          ],
          includeAssessment: true,
          assessmentHours: 1,
          correctionHours: 1,
          locked: false
        }
      ]
    }
  ],
  calendar: {
    schoolYearStart: "2026-09-01",
    schoolYearEnd: "2027-06-30",
    weeklySlots: {
      monday: 2,
      tuesday: 0,
      wednesday: 1,
      thursday: 2,
      friday: 0,
      saturday: 0,
      sunday: 0
    },
    blockedPeriods: [
      {
        id: uid(),
        label: "Vacances d'automne",
        start: "2026-10-17",
        end: "2026-11-02",
        type: "vacances"
      }
    ],
    isolatedBlockedDates: [
      {
        id: uid(),
        label: "Jour férié",
        date: "2026-11-11",
        type: "ferie"
      }
    ],
    bufferHours: 6,
    taughtUntilDate: ""
  }
});

const defaultData = {
  classes: [createDefaultClass()],
  activeClassId: null
};

function downloadText(filename, text, mime = "application/json;charset=utf-8") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function computeAvailableHours(calendar) {
  const start = new Date(`${calendar.schoolYearStart}T00:00:00`);
  const end = new Date(`${calendar.schoolYearEnd}T00:00:00`);
  let totalSlots = 0;

  for (let current = new Date(start); current <= end; current = addDays(current, 1)) {
    const isoDate = toISODate(current);
    const dayKey = dayKeyFromDate(current);
    const blockedPeriod = (calendar.blockedPeriods || []).some(
      (period) => isoDate >= period.start && isoDate <= period.end
    );
    const blockedDate = (calendar.isolatedBlockedDates || []).some((d) => d.date === isoDate);
    if (blockedPeriod || blockedDate) continue;
    totalSlots += hoursToSlots(calendar.weeklySlots?.[dayKey] || 0);
  }

  return slotsToHoursNumber(totalSlots);
}

function flattenProgram(program, includeBufferHours = 0) {
  const items = [];

  program.forEach((theme, themeIndex) => {
    theme.sequences.forEach((sequence, sequenceIndex) => {
      items.push({
        uid: uid(),
        type: "course",
        themeId: theme.id,
        themeTitle: theme.title,
        sequenceId: sequence.id,
        sequenceOrder: sequenceIndex,
        themeOrder: themeIndex,
        title: sequence.title,
        hours: Number(sequence.estimatedHours || 0),
        competencies: sequence.competencies || [],
        locked: !!sequence.locked,
      });

      if (sequence.includeAssessment && Number(sequence.assessmentHours || 0) > 0) {
        items.push({
          uid: uid(),
          type: "assessment",
          themeId: theme.id,
          themeTitle: theme.title,
          sequenceId: sequence.id,
          sequenceOrder: sequenceIndex,
          themeOrder: themeIndex,
          title: `Évaluation - ${sequence.title}`,
          hours: Number(sequence.assessmentHours || 0),
          competencies: [],
          locked: !!sequence.locked,
        });
      }

      if (sequence.includeAssessment && Number(sequence.correctionHours || 0) > 0) {
        items.push({
          uid: uid(),
          type: "correction",
          themeId: theme.id,
          themeTitle: theme.title,
          sequenceId: sequence.id,
          sequenceOrder: sequenceIndex,
          themeOrder: themeIndex,
          title: `Correction - ${sequence.title}`,
          hours: Number(sequence.correctionHours || 0),
          competencies: [],
          locked: !!sequence.locked,
        });
      }
    });
  });

  if (Number(includeBufferHours || 0) > 0) {
    items.push({
      uid: uid(),
      type: "buffer",
      themeId: "buffer",
      themeTitle: "Marge de manœuvre",
      sequenceId: "buffer",
      sequenceOrder: 9999,
      themeOrder: 9999,
      title: "Heures tampon",
      hours: Number(includeBufferHours),
      competencies: [],
      locked: false,
    });
  }

  return items;
}

function generateSchedule(program, calendar) {
  const flatItems = flattenProgram(program, calendar.bufferHours || 0);
  const items = flatItems.map((item) => ({
    ...item,
    remainingSlots: hoursToSlots(item.hours),
    totalSlots: hoursToSlots(item.hours),
  }));

  const isBlockedRange = (isoDate) =>
    (calendar.blockedPeriods || []).some((period) => isoDate >= period.start && isoDate <= period.end);

  const isBlockedDate = (isoDate) =>
    (calendar.isolatedBlockedDates || []).some((d) => d.date === isoDate);

  const start = new Date(`${calendar.schoolYearStart}T00:00:00`);
  const end = new Date(`${calendar.schoolYearEnd}T00:00:00`);

  const sessions = [];
  let cursor = 0;

  for (let current = new Date(start); current <= end; current = addDays(current, 1)) {
    const isoDate = toISODate(current);
    const dayKey = dayKeyFromDate(current);
    const dayLabel = DAYS.find((d) => d.key === dayKey)?.label;
    const availableSlots = hoursToSlots(calendar.weeklySlots?.[dayKey] || 0);

    if (availableSlots <= 0) continue;
    if (isBlockedRange(isoDate) || isBlockedDate(isoDate)) continue;

    let remainingDaySlots = availableSlots;

    while (remainingDaySlots > 0 && cursor < items.length) {
      const currentItem = items[cursor];
      const allocated = Math.min(remainingDaySlots, currentItem.remainingSlots);

      sessions.push({
        id: uid(),
        date: isoDate,
        dayLabel,
        themeId: currentItem.themeId,
        themeTitle: currentItem.themeTitle,
        sequenceId: currentItem.sequenceId,
        title: currentItem.title,
        type: currentItem.type,
        allocatedHours: slotsToHoursNumber(allocated),
        competencies: currentItem.competencies,
        locked: currentItem.locked,
      });

      currentItem.remainingSlots -= allocated;
      remainingDaySlots -= allocated;

      if (currentItem.remainingSlots <= 0) cursor += 1;
    }
  }

  const unscheduledHours = items.slice(cursor).reduce((sum, item) => sum + slotsToHoursNumber(item.remainingSlots), 0);
  const totalRequiredHours = flatItems.reduce((sum, item) => sum + Number(item.hours || 0), 0);
  const totalAvailableHours = computeAvailableHours(calendar);
  const hoursWithoutBuffer = totalRequiredHours - Number(calendar.bufferHours || 0);
  const overloadBeforeBuffer = hoursWithoutBuffer > totalAvailableHours;
  const overloadWithBuffer = totalRequiredHours > totalAvailableHours;

  return {
    sessions,
    summary: {
      totalRequiredHours,
      totalAvailableHours,
      unscheduledHours,
      bufferHours: Number(calendar.bufferHours || 0),
      overloadBeforeBuffer,
      overloadWithBuffer,
      freeHoursAfterPlanned: Number((totalAvailableHours - totalRequiredHours).toFixed(1)),
      lockedSessions: sessions.filter((s) => s.locked).length,
    },
  };
}

function SortableSequenceCard({ children, id }) {
  const { attributes, listeners, setNodeRef, transform, transition } = useSortable({ id });
  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  return (
    <div ref={setNodeRef} style={style} className="relative">
      <div className="absolute left-2 top-3 z-10 rounded-md border bg-white p-1 text-slate-500" {...attributes} {...listeners}>
        <GripVertical className="h-4 w-4" />
      </div>
      <div className="pl-12">{children}</div>
    </div>
  );
}

export default function ProgressionPedagogiqueV3() {
  const [data, setData] = useState(() => {
    const base = {
      ...defaultData,
      activeClassId: defaultData.classes[0]?.id || null,
    };
    return base;
  });
  const [message, setMessage] = useState("Prototype V3 prêt.");

  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 6 } }));

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        setData({
          ...parsed,
          activeClassId: parsed.activeClassId || parsed.classes?.[0]?.id || null,
        });
        setMessage("Dernière sauvegarde locale restaurée.");
      }
    } catch {
      setMessage("Impossible de restaurer la sauvegarde locale.");
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch {
      setMessage("Sauvegarde locale impossible sur ce navigateur.");
    }
  }, [data]);

  const activeClass = useMemo(() => {
    return data.classes.find((c) => c.id === data.activeClassId) || data.classes[0] || null;
  }, [data]);

  const generated = useMemo(() => {
    if (!activeClass) {
      return {
        sessions: [],
        summary: {
          totalRequiredHours: 0,
          totalAvailableHours: 0,
          unscheduledHours: 0,
          bufferHours: 0,
          overloadBeforeBuffer: false,
          overloadWithBuffer: false,
          freeHoursAfterPlanned: 0,
          lockedSessions: 0,
        },
      };
    }
    return generateSchedule(activeClass.program, activeClass.calendar);
  }, [activeClass]);

  const totalPlannedWithoutBuffer = useMemo(() => {
    if (!activeClass) return 0;
    return activeClass.program.reduce(
      (sum, theme) =>
        sum +
        theme.sequences.reduce(
          (inner, seq) =>
            inner +
            Number(seq.estimatedHours || 0) +
            Number(seq.includeAssessment ? seq.assessmentHours || 0 : 0) +
            Number(seq.includeAssessment ? seq.correctionHours || 0 : 0),
          0
        ),
      0
    );
  }, [activeClass]);

  const availableHours = useMemo(() => {
    return activeClass ? computeAvailableHours(activeClass.calendar) : 0;
  }, [activeClass]);

  const groupedSessions = useMemo(() => {
    const groups = {};
    (generated.sessions || []).forEach((session) => {
      if (!groups[session.date]) groups[session.date] = [];
      groups[session.date].push(session);
    });
    return Object.entries(groups).sort(([a], [b]) => a.localeCompare(b));
  }, [generated]);

  const monthlySessions = useMemo(() => {
    const groups = {};
    (generated.sessions || []).forEach((session) => {
      const monthKey = session.date.slice(0, 7);
      if (!groups[monthKey]) groups[monthKey] = [];
      groups[monthKey].push(session);
    });
    return Object.entries(groups).sort(([a], [b]) => a.localeCompare(b));
  }, [generated]);

  const updateActiveClass = (updater) => {
    setData((prev) => ({
      ...prev,
      classes: prev.classes.map((cls) => (cls.id === prev.activeClassId ? updater(cls) : cls)),
    }));
  };

  const addClass = () => {
    const newClass = createDefaultClass();
    setData((prev) => ({
      ...prev,
      classes: [...prev.classes, newClass],
      activeClassId: newClass.id,
    }));
    setMessage("Nouvelle classe ajoutée.");
  };

  const duplicateClass = () => {
    if (!activeClass) return;
    const clone = JSON.parse(JSON.stringify(activeClass));
    clone.id = uid();
    clone.meta.className = `${clone.meta.className} - copie`;
    clone.program = clone.program.map((theme) => ({
      ...theme,
      id: uid(),
      sequences: theme.sequences.map((seq) => ({ ...seq, id: uid() })),
    }));
    clone.calendar.blockedPeriods = clone.calendar.blockedPeriods.map((p) => ({ ...p, id: uid() }));
    clone.calendar.isolatedBlockedDates = clone.calendar.isolatedBlockedDates.map((d) => ({ ...d, id: uid() }));
    setData((prev) => ({
      ...prev,
      classes: [...prev.classes, clone],
      activeClassId: clone.id,
    }));
    setMessage("Classe dupliquée.");
  };

  const deleteClass = () => {
    if (!activeClass || data.classes.length <= 1) return;
    const nextClasses = data.classes.filter((c) => c.id !== activeClass.id);
    setData((prev) => ({
      ...prev,
      classes: nextClasses,
      activeClassId: nextClasses[0]?.id || null,
    }));
    setMessage("Classe supprimée.");
  };

  const updateMeta = (field, value) => {
    updateActiveClass((cls) => ({ ...cls, meta: { ...cls.meta, [field]: value } }));
  };

  const updateCalendarField = (field, value) => {
    updateActiveClass((cls) => ({ ...cls, calendar: { ...cls.calendar, [field]: value } }));
  };

  const applyVacationPreset = () => {
    if (!activeClass) return;
    const year = activeClass.meta.schoolYear;
    const zone = activeClass.meta.zone;
    const preset = VACATION_PRESETS?.[year]?.[zone];
    if (!preset) {
      setMessage("Aucun jeu de vacances prédéfini pour cette année et cette zone.");
      return;
    }
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        blockedPeriods: preset.map((p) => ({ ...p, id: uid() })),
      },
    }));
    setMessage(`Vacances de la zone ${zone} importées.`);
  };

  const addTheme = () => {
    updateActiveClass((cls) => ({
      ...cls,
      program: [...cls.program, { id: uid(), title: `Thème ${cls.program.length + 1}`, sequences: [] }],
    }));
  };

  const updateTheme = (themeId, field, value) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) => (theme.id === themeId ? { ...theme, [field]: value } : theme)),
    }));
  };

  const deleteTheme = (themeId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.filter((theme) => theme.id !== themeId),
    }));
  };

  const addSequence = (themeId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? {
              ...theme,
              sequences: [
                ...theme.sequences,
                {
                  id: uid(),
                  title: "Nouveau chapitre",
                  estimatedHours: 2,
                  competencies: [],
                  includeAssessment: false,
                  assessmentHours: 1,
                  correctionHours: 1,
                  locked: false,
                },
              ],
            }
          : theme
      ),
    }));
  };

  const updateSequence = (themeId, sequenceId, field, value) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? {
              ...theme,
              sequences: theme.sequences.map((seq) =>
                seq.id === sequenceId ? { ...seq, [field]: value } : seq
              ),
            }
          : theme
      ),
    }));
  };

  const deleteSequence = (themeId, sequenceId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? { ...theme, sequences: theme.sequences.filter((seq) => seq.id !== sequenceId) }
          : theme
      ),
    }));
  };

  const moveSequence = (themeId, sequenceId, direction) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) => {
        if (theme.id !== themeId) return theme;
        const index = theme.sequences.findIndex((s) => s.id === sequenceId);
        const newIndex = direction === "up" ? index - 1 : index + 1;
        if (index < 0 || newIndex < 0 || newIndex >= theme.sequences.length) return theme;
        return { ...theme, sequences: arrayMove(theme.sequences, index, newIndex) };
      }),
    }));
  };

  const handleDragEnd = (themeId, event) => {
    const { active, over } = event;
    if (!over || active.id === over.id) return;
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) => {
        if (theme.id !== themeId) return theme;
        const oldIndex = theme.sequences.findIndex((s) => s.id === active.id);
        const newIndex = theme.sequences.findIndex((s) => s.id === over.id);
        if (oldIndex < 0 || newIndex < 0) return theme;
        return { ...theme, sequences: arrayMove(theme.sequences, oldIndex, newIndex) };
      }),
    }));
  };

  const addBlockedPeriod = () => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        blockedPeriods: [
          ...cls.calendar.blockedPeriods,
          {
            id: uid(),
            label: "Nouvelle période bloquée",
            start: cls.calendar.schoolYearStart,
            end: cls.calendar.schoolYearStart,
            type: "vacances",
          },
        ],
      },
    }));
  };

  const updateBlockedPeriod = (id, field, value) => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        blockedPeriods: cls.calendar.blockedPeriods.map((period) =>
          period.id === id ? { ...period, [field]: value } : period
        ),
      },
    }));
  };

  const deleteBlockedPeriod = (id) => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        blockedPeriods: cls.calendar.blockedPeriods.filter((period) => period.id !== id),
      },
    }));
  };

  const addBlockedDate = () => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        isolatedBlockedDates: [
          ...cls.calendar.isolatedBlockedDates,
          {
            id: uid(),
            label: "Nouvelle date bloquée",
            date: cls.calendar.schoolYearStart,
            type: "ferie",
          },
        ],
      },
    }));
  };

  const updateBlockedDate = (id, field, value) => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        isolatedBlockedDates: cls.calendar.isolatedBlockedDates.map((d) =>
          d.id === id ? { ...d, [field]: value } : d
        ),
      },
    }));
  };

  const deleteBlockedDate = (id) => {
    updateActiveClass((cls) => ({
      ...cls,
      calendar: {
        ...cls.calendar,
        isolatedBlockedDates: cls.calendar.isolatedBlockedDates.filter((d) => d.id !== id),
      },
    }));
  };

  const addOneHourToSequence = (themeId, sequenceId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? {
              ...theme,
              sequences: theme.sequences.map((seq) =>
                seq.id === sequenceId ? { ...seq, estimatedHours: Number(seq.estimatedHours || 0) + 1 } : seq
              ),
            }
          : theme
      ),
    }));
    setMessage("1 heure ajoutée. La progression a été recalculée.");
  };

  const removeOneHourToSequence = (themeId, sequenceId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? {
              ...theme,
              sequences: theme.sequences.map((seq) =>
                seq.id === sequenceId
                  ? { ...seq, estimatedHours: Math.max(0.5, Number(seq.estimatedHours || 0) - 1) }
                  : seq
              ),
            }
          : theme
      ),
    }));
    setMessage("1 heure retirée. La progression a été recalculée.");
  };

  const toggleLockSequence = (themeId, sequenceId) => {
    updateActiveClass((cls) => ({
      ...cls,
      program: cls.program.map((theme) =>
        theme.id === themeId
          ? {
              ...theme,
              sequences: theme.sequences.map((seq) =>
                seq.id === sequenceId ? { ...seq, locked: !seq.locked } : seq
              ),
            }
          : theme
      ),
    }));
  };

  const saveNow = () => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    setMessage("Sauvegarde locale effectuée.");
  };

  const resetAll = () => {
    const fresh = { classes: [createDefaultClass()], activeClassId: null };
    fresh.activeClassId = fresh.classes[0].id;
    setData(fresh);
    setMessage("Application réinitialisée.");
  };

  const exportJson = () => {
    downloadText(`progressions-${activeClass?.meta.className || "classes"}.json`, JSON.stringify(data, null, 2));
    setMessage("Export JSON téléchargé.");
  };

  const importJson = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const text = await file.text();
    try {
      const parsed = JSON.parse(text);
      setData({
        ...parsed,
        activeClassId: parsed.activeClassId || parsed.classes?.[0]?.id || null,
      });
      setMessage("Import JSON réussi.");
    } catch {
      setMessage("Le fichier JSON n'est pas valide.");
    }
    event.target.value = "";
  };

  const exportExcel = () => {
    if (!activeClass) return;
    const wb = XLSX.utils.book_new();

    const wsClasses = XLSX.utils.json_to_sheet(
      data.classes.map((cls) => ({
        Classe: cls.meta.className,
        Matiere: cls.meta.subject,
        Annee: cls.meta.schoolYear,
        Zone: cls.meta.zone,
        Debut: cls.calendar.schoolYearStart,
        Fin: cls.calendar.schoolYearEnd,
      }))
    );

    const wsProgram = XLSX.utils.json_to_sheet(
      activeClass.program.flatMap((theme, themeIndex) =>
        theme.sequences.map((seq, seqIndex) => ({
          Theme: themeIndex + 1,
          Titre_Theme: theme.title,
          Chapitre: seqIndex + 1,
          Titre_Chapitre: seq.title,
          Heures_Cours: Number(seq.estimatedHours || 0),
          Evaluation: seq.includeAssessment ? "Oui" : "Non",
          Heures_Evaluation: seq.includeAssessment ? Number(seq.assessmentHours || 0) : 0,
          Heures_Correction: seq.includeAssessment ? Number(seq.correctionHours || 0) : 0,
          Competences: (seq.competencies || []).join(" | "),
          Verrouille: seq.locked ? "Oui" : "Non",
        }))
      )
    );

    const wsProgression = XLSX.utils.json_to_sheet(
      generated.sessions.map((s) => ({
        Date: formatDate(s.date),
        Jour: s.dayLabel,
        Theme: s.themeTitle,
        Intitule: s.title,
        Type: TYPE_LABELS[s.type],
        Duree_h: s.allocatedHours,
        Verrouillee: s.locked ? "Oui" : "Non",
      }))
    );

    XLSX.utils.book_append_sheet(wb, wsClasses, "Classes");
    XLSX.utils.book_append_sheet(wb, wsProgram, "Programme");
    XLSX.utils.book_append_sheet(wb, wsProgression, "Progression");
    XLSX.writeFile(wb, `progression-${activeClass.meta.className || "classe"}.xlsx`);
    setMessage("Export Excel téléchargé.");
  };

  const exportPdf = () => {
    if (!activeClass) return;
    const doc = new jsPDF({ orientation: "landscape" });
    doc.setFontSize(16);
    doc.text(`Progression annuelle - ${activeClass.meta.className}`, 14, 14);
    doc.setFontSize(10);
    doc.text(`${activeClass.meta.subject} - ${activeClass.meta.schoolYear} - Zone ${activeClass.meta.zone}`, 14, 21);

    autoTable(doc, {
      startY: 28,
      head: [["Date", "Jour", "Thème", "Séance", "Type", "Durée", "État"]],
      body: generated.sessions.map((s) => [
        formatDate(s.date),
        s.dayLabel,
        s.themeTitle,
        s.title,
        TYPE_LABELS[s.type],
        `${s.allocatedHours} h`,
        s.locked ? "Fait / verrouillé" : "À venir",
      ]),
      styles: { fontSize: 8 },
      headStyles: { fillColor: [60, 60, 60] },
    });

    doc.save(`progression-${activeClass.meta.className || "classe"}.pdf`);
    setMessage("Export PDF téléchargé.");
  };

  if (!activeClass) {
    return <div className="p-6">Aucune classe disponible.</div>;
  }

  return (
    <div className="min-h-screen bg-slate-50 p-6">
      <div className="mx-auto max-w-7xl space-y-6">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
          <div>
            <h1 className="text-3xl font-bold tracking-tight">Générateur de progression pédagogique STMG</h1>
            <p className="text-sm text-slate-600">
              V3 multi-classes avec vacances par zone, verrouillage des séances passées et vue agenda mensuelle.
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <Button variant="outline" onClick={saveNow}><Save className="mr-2 h-4 w-4" />Sauvegarder</Button>
            <Button variant="outline" onClick={exportJson}><Download className="mr-2 h-4 w-4" />JSON</Button>
            <Button variant="outline" onClick={exportExcel}><Download className="mr-2 h-4 w-4" />Excel</Button>
            <Button variant="outline" onClick={exportPdf}><Download className="mr-2 h-4 w-4" />PDF</Button>
            <label className="inline-flex cursor-pointer items-center rounded-md border border-input bg-background px-4 py-2 text-sm font-medium shadow-sm hover:bg-accent hover:text-accent-foreground">
              <Upload className="mr-2 h-4 w-4" />Importer
              <input type="file" accept="application/json" className="hidden" onChange={importJson} />
            </label>
            <Button variant="ghost" onClick={resetAll}><RefreshCw className="mr-2 h-4 w-4" />Réinitialiser</Button>
          </div>
        </div>

        <Alert className="rounded-2xl">
          <AlertTitle>État</AlertTitle>
          <AlertDescription>{message}</AlertDescription>
        </Alert>

        <Card className="rounded-2xl shadow-sm">
          <CardHeader className="flex flex-row items-center justify-between space-y-0">
            <CardTitle>Gestion des classes</CardTitle>
            <div className="flex gap-2">
              <Button variant="outline" onClick={addClass}><Plus className="mr-2 h-4 w-4" />Ajouter</Button>
              <Button variant="outline" onClick={duplicateClass}>Dupliquer</Button>
              <Button variant="ghost" onClick={deleteClass}><Trash2 className="mr-2 h-4 w-4" />Supprimer</Button>
            </div>
          </CardHeader>
          <CardContent className="grid gap-4 md:grid-cols-3">
            <div>
              <Label>Classe active</Label>
              <Select value={data.activeClassId || ""} onValueChange={(value) => setData((prev) => ({ ...prev, activeClassId: value }))}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  {data.classes.map((cls) => (
                    <SelectItem key={cls.id} value={cls.id}>{cls.meta.className}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <div>
              <Label>Nom de la classe</Label>
              <Input value={activeClass.meta.className} onChange={(e) => updateMeta("className", e.target.value)} />
            </div>
            <div>
              <Label>Matière</Label>
              <Input value={activeClass.meta.subject} onChange={(e) => updateMeta("subject", e.target.value)} />
            </div>
          </CardContent>
        </Card>

        <div className="grid gap-4 md:grid-cols-5">
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Heures prévues</div><div className="mt-2 text-3xl font-semibold">{totalPlannedWithoutBuffer} h</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Heures disponibles</div><div className="mt-2 text-3xl font-semibold">{availableHours} h</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Tampon</div><div className="mt-2 text-3xl font-semibold">{activeClass.calendar.bufferHours} h</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Non placé</div><div className="mt-2 text-3xl font-semibold">{generated.summary.unscheduledHours} h</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Séances verrouillées</div><div className="mt-2 text-3xl font-semibold">{generated.summary.lockedSessions}</div></CardContent></Card>
        </div>

        {(generated.summary.overloadBeforeBuffer || generated.summary.overloadWithBuffer) && (
          <Alert className="rounded-2xl border-amber-300 bg-amber-50 text-amber-900">
            <AlertTriangle className="h-4 w-4" />
            <AlertTitle>Alerte de capacité</AlertTitle>
            <AlertDescription>
              {generated.summary.overloadBeforeBuffer
                ? "Le programme dépasse le temps annuel réellement disponible."
                : "Le programme consomme tout ou partie de votre marge tampon."}
            </AlertDescription>
          </Alert>
        )}

        <Tabs defaultValue="programme" className="space-y-4">
          <TabsList className="grid w-full grid-cols-5 rounded-2xl">
            <TabsTrigger value="programme">Programme</TabsTrigger>
            <TabsTrigger value="calendrier">Calendrier</TabsTrigger>
            <TabsTrigger value="resultat">Tableau</TabsTrigger>
            <TabsTrigger value="agenda">Agenda mensuel</TabsTrigger>
            <TabsTrigger value="pilotage">Pilotage</TabsTrigger>
          </TabsList>

          <TabsContent value="programme" className="space-y-4">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader><CardTitle>Informations pédagogiques</CardTitle></CardHeader>
              <CardContent className="grid gap-4 md:grid-cols-4">
                <div><Label>Année scolaire</Label><Input value={activeClass.meta.schoolYear} onChange={(e) => updateMeta("schoolYear", e.target.value)} /></div>
                <div><Label>Enseignant</Label><Input value={activeClass.meta.teacher} onChange={(e) => updateMeta("teacher", e.target.value)} /></div>
                <div>
                  <Label>Zone</Label>
                  <Select value={activeClass.meta.zone} onValueChange={(value) => updateMeta("zone", value)}>
                    <SelectTrigger><SelectValue /></SelectTrigger>
                    <SelectContent>
                      <SelectItem value="A">Zone A</SelectItem>
                      <SelectItem value="B">Zone B</SelectItem>
                      <SelectItem value="C">Zone C</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
                <div className="flex items-end"><Button variant="outline" onClick={applyVacationPreset}>Importer vacances de la zone</Button></div>
              </CardContent>
            </Card>

            <div className="flex justify-end"><Button onClick={addTheme}><Plus className="mr-2 h-4 w-4" />Ajouter un thème</Button></div>

            {activeClass.program.map((theme, themeIndex) => (
              <Card key={theme.id} className="rounded-2xl shadow-sm">
                <CardHeader className="flex flex-row items-center justify-between space-y-0">
                  <CardTitle>Thème {themeIndex + 1}</CardTitle>
                  <Button variant="ghost" size="icon" onClick={() => deleteTheme(theme.id)}><Trash2 className="h-4 w-4" /></Button>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div><Label>Titre du thème</Label><Input value={theme.title} onChange={(e) => updateTheme(theme.id, "title", e.target.value)} /></div>
                  <div className="flex items-center justify-between"><div className="font-medium">Chapitres du thème</div><Button variant="outline" onClick={() => addSequence(theme.id)}><Plus className="mr-2 h-4 w-4" />Ajouter un chapitre</Button></div>

                  <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={(event) => handleDragEnd(theme.id, event)}>
                    <SortableContext items={theme.sequences.map((s) => s.id)} strategy={verticalListSortingStrategy}>
                      <div className="space-y-3">
                        {theme.sequences.map((sequence, seqIndex) => (
                          <SortableSequenceCard key={sequence.id} id={sequence.id}>
                            <Card className="rounded-2xl border-slate-200">
                              <CardContent className="space-y-4 p-4">
                                <div className="flex flex-wrap items-center justify-between gap-2">
                                  <div className="font-medium">Chapitre {seqIndex + 1}</div>
                                  <div className="flex flex-wrap gap-2">
                                    <Button variant="ghost" size="icon" onClick={() => moveSequence(theme.id, sequence.id, "up")}><ChevronUp className="h-4 w-4" /></Button>
                                    <Button variant="ghost" size="icon" onClick={() => moveSequence(theme.id, sequence.id, "down")}><ChevronDown className="h-4 w-4" /></Button>
                                    <Button variant="ghost" size="icon" onClick={() => toggleLockSequence(theme.id, sequence.id)}>{sequence.locked ? <Lock className="h-4 w-4" /> : <LockOpen className="h-4 w-4" />}</Button>
                                    <Button variant="ghost" size="icon" onClick={() => deleteSequence(theme.id, sequence.id)}><Trash2 className="h-4 w-4" /></Button>
                                  </div>
                                </div>

                                <div className="grid gap-4 md:grid-cols-2">
                                  <div><Label>Titre</Label><Input value={sequence.title} onChange={(e) => updateSequence(theme.id, sequence.id, "title", e.target.value)} /></div>
                                  <div><Label>Heures estimées</Label><Input type="number" min="0.5" step="0.5" value={sequence.estimatedHours} onChange={(e) => updateSequence(theme.id, sequence.id, "estimatedHours", e.target.value)} /></div>
                                </div>

                                <div className="flex flex-wrap gap-2">
                                  <Button variant="outline" size="sm" onClick={() => addOneHourToSequence(theme.id, sequence.id)}>+ 1 h</Button>
                                  <Button variant="outline" size="sm" onClick={() => removeOneHourToSequence(theme.id, sequence.id)}>- 1 h</Button>
                                  <Badge variant="secondary">{sequence.locked ? "Verrouillé" : "Recalculable"}</Badge>
                                </div>

                                <div>
                                  <Label>Compétences visées (une par ligne)</Label>
                                  <Textarea value={(sequence.competencies || []).join("
")} onChange={(e) => updateSequence(theme.id, sequence.id, "competencies", e.target.value.split("
").map((x) => x.trim()).filter(Boolean))} />
                                  <div className="mt-2 flex flex-wrap gap-2">
                                    {(sequence.competencies || []).map((comp) => <Badge key={comp} variant="outline">{comp}</Badge>)}
                                  </div>
                                </div>

                                <div className="rounded-xl border p-4">
                                  <div className="mb-3 flex items-center justify-between">
                                    <Label>Prévoir une évaluation</Label>
                                    <Switch checked={sequence.includeAssessment} onCheckedChange={(checked) => updateSequence(theme.id, sequence.id, "includeAssessment", checked)} />
                                  </div>
                                  {sequence.includeAssessment && (
                                    <div className="grid gap-4 md:grid-cols-2">
                                      <div><Label>Durée évaluation</Label><Input type="number" min="0" step="0.5" value={sequence.assessmentHours} onChange={(e) => updateSequence(theme.id, sequence.id, "assessmentHours", e.target.value)} /></div>
                                      <div><Label>Durée correction</Label><Input type="number" min="0" step="0.5" value={sequence.correctionHours} onChange={(e) => updateSequence(theme.id, sequence.id, "correctionHours", e.target.value)} /></div>
                                    </div>
                                  )}
                                </div>
                              </CardContent>
                            </Card>
                          </SortableSequenceCard>
                        ))}
                      </div>
                    </SortableContext>
                  </DndContext>
                </CardContent>
              </Card>
            ))}
          </TabsContent>

          <TabsContent value="calendrier" className="space-y-4">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader><CardTitle>Calendrier annuel</CardTitle></CardHeader>
              <CardContent className="space-y-6">
                <div className="grid gap-4 md:grid-cols-4">
                  <div><Label>Date de rentrée</Label><Input type="date" value={activeClass.calendar.schoolYearStart} onChange={(e) => updateCalendarField("schoolYearStart", e.target.value)} /></div>
                  <div><Label>Date de fin d'année</Label><Input type="date" value={activeClass.calendar.schoolYearEnd} onChange={(e) => updateCalendarField("schoolYearEnd", e.target.value)} /></div>
                  <div><Label>Heures tampon</Label><Input type="number" min="0" step="0.5" value={activeClass.calendar.bufferHours} onChange={(e) => updateCalendarField("bufferHours", e.target.value)} /></div>
                  <div><Label>Séances faites jusqu'au</Label><Input type="date" value={activeClass.calendar.taughtUntilDate || ""} onChange={(e) => updateCalendarField("taughtUntilDate", e.target.value)} /></div>
                </div>

                <Separator />

                <div>
                  <div className="mb-3 font-medium">Créneaux de cours par semaine</div>
                  <div className="grid gap-4 md:grid-cols-4">
                    {DAYS.map((day) => (
                      <div key={day.key} className="rounded-xl border p-3">
                        <Label>{day.label}</Label>
                        <Input
                          type="number"
                          min="0"
                          step="0.5"
                          value={activeClass.calendar.weeklySlots[day.key]}
                          onChange={(e) =>
                            updateActiveClass((cls) => ({
                              ...cls,
                              calendar: {
                                ...cls.calendar,
                                weeklySlots: { ...cls.calendar.weeklySlots, [day.key]: e.target.value },
                              },
                            }))
                          }
                        />
                      </div>
                    ))}
                  </div>
                </div>
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader className="flex flex-row items-center justify-between space-y-0">
                <CardTitle>Périodes bloquées</CardTitle>
                <Button variant="outline" onClick={addBlockedPeriod}><Plus className="mr-2 h-4 w-4" />Ajouter</Button>
              </CardHeader>
              <CardContent className="space-y-3">
                {activeClass.calendar.blockedPeriods.map((period) => (
                  <div key={period.id} className="grid gap-3 rounded-xl border p-4 md:grid-cols-5">
                    <Input value={period.label} onChange={(e) => updateBlockedPeriod(period.id, "label", e.target.value)} />
                    <Input type="date" value={period.start} onChange={(e) => updateBlockedPeriod(period.id, "start", e.target.value)} />
                    <Input type="date" value={period.end} onChange={(e) => updateBlockedPeriod(period.id, "end", e.target.value)} />
                    <Select value={period.type} onValueChange={(value) => updateBlockedPeriod(period.id, "type", value)}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="vacances">Vacances</SelectItem>
                        <SelectItem value="examens">Examens blancs</SelectItem>
                        <SelectItem value="sortie">Projet / sortie</SelectItem>
                      </SelectContent>
                    </Select>
                    <Button variant="ghost" onClick={() => deleteBlockedPeriod(period.id)}><Trash2 className="h-4 w-4" /></Button>
                  </div>
                ))}
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader className="flex flex-row items-center justify-between space-y-0">
                <CardTitle>Dates isolées bloquées</CardTitle>
                <Button variant="outline" onClick={addBlockedDate}><Plus className="mr-2 h-4 w-4" />Ajouter</Button>
              </CardHeader>
              <CardContent className="space-y-3">
                {activeClass.calendar.isolatedBlockedDates.map((d) => (
                  <div key={d.id} className="grid gap-3 rounded-xl border p-4 md:grid-cols-4">
                    <Input value={d.label} onChange={(e) => updateBlockedDate(d.id, "label", e.target.value)} />
                    <Input type="date" value={d.date} onChange={(e) => updateBlockedDate(d.id, "date", e.target.value)} />
                    <Select value={d.type} onValueChange={(value) => updateBlockedDate(d.id, "type", value)}>
                      <SelectTrigger><SelectValue /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="ferie">Jour férié</SelectItem>
                        <SelectItem value="banalise">Journée banalisée</SelectItem>
                        <SelectItem value="reunion">Réunion / conseil</SelectItem>
                      </SelectContent>
                    </Select>
                    <Button variant="ghost" onClick={() => deleteBlockedDate(d.id)}><Trash2 className="h-4 w-4" /></Button>
                  </div>
                ))}
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="resultat" className="space-y-4">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader><CardTitle>Tableau de progression</CardTitle></CardHeader>
              <CardContent>
                <div className="overflow-x-auto">
                  <table className="min-w-full border-collapse text-sm">
                    <thead>
                      <tr className="border-b bg-slate-50 text-left">
                        <th className="p-3">Date</th>
                        <th className="p-3">Jour</th>
                        <th className="p-3">Thème</th>
                        <th className="p-3">Séance</th>
                        <th className="p-3">Type</th>
                        <th className="p-3">Durée</th>
                        <th className="p-3">État</th>
                      </tr>
                    </thead>
                    <tbody>
                      {generated.sessions.map((session) => (
                        <tr key={session.id} className="border-b hover:bg-slate-50">
                          <td className="p-3">{formatDate(session.date)}</td>
                          <td className="p-3">{session.dayLabel}</td>
                          <td className="p-3">{session.themeTitle}</td>
                          <td className="p-3">{session.title}</td>
                          <td className="p-3"><Badge variant={session.type === "assessment" ? "destructive" : session.type === "correction" ? "secondary" : "default"}>{TYPE_LABELS[session.type]}</Badge></td>
                          <td className="p-3">{session.allocatedHours} h</td>
                          <td className="p-3">{session.locked ? <Badge variant="outline">Fait / verrouillé</Badge> : <Badge variant="secondary">À venir</Badge>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="agenda" className="space-y-4">
            {monthlySessions.length === 0 ? (
              <Card className="rounded-2xl shadow-sm"><CardContent className="p-6 text-sm text-slate-500">Aucune séance générée.</CardContent></Card>
            ) : (
              monthlySessions.map(([month, sessions]) => (
                <Card key={month} className="rounded-2xl shadow-sm">
                  <CardHeader><CardTitle>{new Date(`${month}-01T00:00:00`).toLocaleDateString("fr-FR", { month: "long", year: "numeric" })}</CardTitle></CardHeader>
                  <CardContent>
                    <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
                      {sessions.map((session) => (
                        <div key={session.id} className="rounded-xl border bg-slate-50 p-4">
                          <div className="text-xs uppercase tracking-wide text-slate-500">{formatDate(session.date)} • {session.dayLabel}</div>
                          <div className="mt-2 font-semibold">{session.title}</div>
                          <div className="mt-1 text-sm text-slate-600">{session.themeTitle}</div>
                          <div className="mt-3 flex flex-wrap gap-2">
                            <Badge variant={session.type === "assessment" ? "destructive" : session.type === "correction" ? "secondary" : "default"}>{TYPE_LABELS[session.type]}</Badge>
                            <Badge variant="outline">{session.allocatedHours} h</Badge>
                            {session.locked && <Badge variant="outline">Verrouillée</Badge>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </CardContent>
                </Card>
              ))
            )}
          </TabsContent>

          <TabsContent value="pilotage" className="space-y-4">
            <Card className="rounded-2xl shadow-sm">
              <CardHeader><CardTitle>Ajustement terrain</CardTitle></CardHeader>
              <CardContent className="space-y-4">
                <p className="text-sm text-slate-600">
                  Utilisez +1 h ou -1 h pour absorber un retard. Les séances situées avant la date “séances faites jusqu'au” sont marquées comme verrouillées.
                </p>
                <div className="grid gap-4 md:grid-cols-2">
                  {activeClass.program.map((theme) => (
                    <div key={theme.id} className="rounded-xl border p-4">
                      <div className="mb-3 font-semibold">{theme.title}</div>
                      <div className="space-y-2">
                        {theme.sequences.map((seq) => (
                          <div key={seq.id} className="flex items-center justify-between rounded-lg bg-slate-50 p-3">
                            <div>
                              <div className="font-medium">{seq.title}</div>
                              <div className="text-sm text-slate-500">{seq.estimatedHours} h prévues</div>
                            </div>
                            <div className="flex gap-2">
                              <Button variant="outline" size="sm" onClick={() => removeOneHourToSequence(theme.id, seq.id)}>-1 h</Button>
                              <Button variant="outline" size="sm" onClick={() => addOneHourToSequence(theme.id, seq.id)}>+1 h</Button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader><CardTitle>Vue chronologique</CardTitle></CardHeader>
              <CardContent className="space-y-4">
                {groupedSessions.map(([date, sessions]) => (
                  <div key={date} className="rounded-2xl border p-4">
                    <div className="mb-3 font-semibold">{formatDate(date)}</div>
                    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                      {sessions.map((session) => (
                        <div key={session.id} className="rounded-xl bg-slate-50 p-3">
                          <div className="text-xs uppercase tracking-wide text-slate-500">{session.themeTitle}</div>
                          <div className="mt-1 font-medium">{session.title}</div>
                          <div className="mt-2 text-sm text-slate-600">{session.allocatedHours} h • {TYPE_LABELS[session.type]}</div>
                          {session.locked && <div className="mt-2"><Badge variant="outline">Déjà fait / verrouillé</Badge></div>}
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}
