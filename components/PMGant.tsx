import React, { useEffect, useMemo, useState } from 'react';
import {
  Calendar,
  CheckCircle2,
  Download,
  Edit3,
  Eye,
  FileBarChart,
  GanttChartSquare,
  Mail,
  Plus,
  Save,
  Sparkles,
  Trash2,
  Wand2
} from 'lucide-react';
import { LLMConfig, PMGanttItem, PMGanttPriority, PMGanttStatus, Project, Team, User } from '../types';
import { generateId } from '../services/storage';
import { extractPMGanttItemFromText, generatePMGanttNarrative } from '../services/llmService';

interface PMGantProps {
  teams: Team[];
  users: User[];
  currentUser: User;
  llmConfig: LLMConfig;
  gantItems: PMGanttItem[];
  onSaveItem: (item: PMGanttItem) => void;
  onDeleteItem: (id: string) => void;
}

type PMGantView = 'workspace' | 'preview';
type PMGantProject = Project & { teamName: string };

interface PMGantFormState {
  projectId: string;
  title: string;
  description: string;
  owner: string;
  startDate: string;
  endDate: string;
  progressPct: number;
  status: PMGanttStatus;
  priority: PMGanttPriority;
  isMilestone: boolean;
  notes: string;
}

type PMGantNarrative = {
  executiveSummary: string;
  projectSummaries: { projectId: string; summary: string; keyMilestones: string[] }[];
};

const STATUS_VALUES: PMGanttStatus[] = ['Planned', 'In Progress', 'Done', 'Blocked'];
const PRIORITY_VALUES: PMGanttPriority[] = ['Low', 'Medium', 'High', 'Critical'];

const STATUS_COLORS: Record<PMGanttStatus, string> = {
  Planned: '#3b82f6',
  'In Progress': '#0ea5e9',
  Done: '#10b981',
  Blocked: '#ef4444',
};

const priorityBadgeClass = (priority: PMGanttPriority) => {
  if (priority === 'Critical') return 'bg-red-100 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800';
  if (priority === 'High') return 'bg-amber-100 text-amber-700 border-amber-200 dark:bg-amber-900/20 dark:text-amber-300 dark:border-amber-800';
  if (priority === 'Medium') return 'bg-blue-100 text-blue-700 border-blue-200 dark:bg-blue-900/20 dark:text-blue-300 dark:border-blue-800';
  return 'bg-slate-100 text-slate-700 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700';
};

const statusBadgeClass = (status: PMGanttStatus) => {
  if (status === 'Done') return 'bg-emerald-100 text-emerald-700 border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-300 dark:border-emerald-800';
  if (status === 'In Progress') return 'bg-sky-100 text-sky-700 border-sky-200 dark:bg-sky-900/20 dark:text-sky-300 dark:border-sky-800';
  if (status === 'Blocked') return 'bg-red-100 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800';
  return 'bg-indigo-100 text-indigo-700 border-indigo-200 dark:bg-indigo-900/20 dark:text-indigo-300 dark:border-indigo-800';
};

const toISODate = (date: Date): string => date.toISOString().split('T')[0];
const todayISO = (): string => toISODate(new Date());
const plusDaysISO = (days: number): string => {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return toISODate(d);
};

const clampPct = (value: number): number => Math.max(0, Math.min(100, Math.round(value)));

const parseDate = (value: string): Date | null => {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
};

const addDays = (date: Date, days: number): Date => {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
};

const formatShortDate = (value: string): string => {
  const date = parseDate(value);
  return date ? date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
};

const escapeHTML = (value: string): string =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const buildDefaultNarrative = (projects: Array<{ project: PMGantProject; items: PMGanttItem[] }>): PMGantNarrative => {
  const summary = `Roadmap generated for ${projects.length} project${projects.length > 1 ? 's' : ''}. Timeline reflects current planning inputs and milestone readiness.`;
  return {
    executiveSummary: summary,
    projectSummaries: projects.map(({ project, items }) => ({
      projectId: project.id,
      summary: `${items.length} roadmap item${items.length > 1 ? 's' : ''} planned for ${project.name}.`,
      keyMilestones: items.filter(item => item.isMilestone).slice(0, 4).map(item => item.title),
    })),
  };
};

const buildPMGantDeckHTML = (
  projects: Array<{ project: PMGantProject; items: PMGanttItem[] }>,
  narrative: PMGantNarrative
): string => {
  const now = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  const summaryMap = new Map(narrative.projectSummaries.map(item => [item.projectId, item]));
  const tickPercents = [0, 25, 50, 75, 100];

  const css = `
  <style>
    .deck{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:#0f172a;max-width:1380px;margin:0 auto}
    .deck *{box-sizing:border-box}
    .deck-header{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:4px solid #00915A;padding-bottom:14px;margin-bottom:22px}
    .deck-title{font-size:28px;font-weight:800;background:linear-gradient(135deg,#00915A,#00A86B);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
    .deck-sub{font-size:12px;color:#64748b;margin-top:4px}
    .deck-conf{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#94a3b8;font-weight:700}
    .exec-summary{background:#e9f8f2;border-left:4px solid #00915A;border-radius:12px;padding:14px 16px;margin-bottom:20px;font-size:13px;line-height:1.6;color:#334155}
    .project-block{margin-bottom:20px;page-break-inside:avoid;break-inside:avoid-page}
    .project-banner{background:linear-gradient(135deg,#00915A 0%,#00A86B 52%,#007A4C 100%);color:#fff;border-radius:12px 12px 0 0;padding:14px 18px;display:flex;justify-content:space-between;align-items:center}
    .project-banner h2{margin:0;font-size:19px;line-height:1.2}
    .project-meta{font-size:11px;opacity:.85;margin-top:3px}
    .project-body{border:1px solid #dbe5ef;border-top:none;border-radius:0 0 12px 12px;background:#fff;padding:14px 16px}
    .project-summary{background:#f8fafc;border-left:4px solid #00915A;border-radius:10px;padding:10px 12px;font-size:12px;line-height:1.5;color:#334155;margin-bottom:12px}
    .milestone-list{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 12px 0}
    .milestone-tag{font-size:10px;font-weight:700;color:#065f46;background:#d1fae5;border:1px solid #a7f3d0;padding:3px 8px;border-radius:999px}
    .timeline-head{display:grid;grid-template-columns:260px 100px 90px 130px 70px 1fr;gap:10px;font-size:10px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.5px;padding:6px 0;border-bottom:1px solid #e2e8f0}
    .gantt-row{display:grid;grid-template-columns:260px 100px 90px 130px 70px 1fr;gap:10px;align-items:center;padding:8px 0;border-bottom:1px solid #f1f5f9;page-break-inside:avoid;break-inside:avoid-page}
    .gantt-row:last-child{border-bottom:none}
    .task-cell{font-size:12px;color:#0f172a}
    .task-cell strong{display:block;font-size:12px}
    .task-desc{font-size:10px;color:#64748b;margin-top:2px;line-height:1.4}
    .owner-cell,.status-cell,.date-cell,.prog-cell{font-size:11px;color:#334155}
    .status-chip{display:inline-flex;padding:2px 8px;border-radius:999px;font-weight:700;color:#fff;font-size:10px;white-space:nowrap}
    .progress-pill{display:inline-flex;min-width:42px;justify-content:center;padding:2px 8px;border-radius:999px;font-weight:700;font-size:10px;background:#e2e8f0;color:#0f172a}
    .timeline-wrap{position:relative}
    .timeline-axis{position:relative;height:18px;margin-bottom:6px}
    .axis-line{position:absolute;top:0;bottom:0;border-left:1px dashed #cbd5e1}
    .axis-label{position:absolute;top:0;transform:translateX(-50%);font-size:9px;color:#64748b;white-space:nowrap;background:#fff;padding:0 3px}
    .timeline-track{position:relative;height:22px;border-radius:7px;background:#f8fafc;border:1px solid #dbe5ef;overflow:visible}
    .timeline-grid-line{position:absolute;top:0;bottom:0;border-left:1px dotted #e2e8f0}
    .task-bar{position:absolute;top:4px;height:12px;border-radius:8px;box-shadow:0 1px 4px rgba(15,23,42,.2)}
    .task-milestone{position:absolute;top:4px;width:12px;height:12px;transform:translateX(-50%) rotate(45deg);border-radius:2px;box-shadow:0 1px 4px rgba(15,23,42,.25)}
    .empty-state{padding:14px;border:1px dashed #cbd5e1;border-radius:10px;font-size:12px;color:#64748b;background:#f8fafc}
    @media print{
      .deck{max-width:none}
      .project-block{margin-bottom:14px}
      .project-body{padding:10px 12px}
      .timeline-head{font-size:9px}
      .gantt-row{padding:6px 0}
    }
  </style>`;

  const projectSections = projects.map(({ project, items }) => {
    const sortedItems = [...items].sort((a, b) => {
      const aStart = parseDate(a.startDate)?.getTime() || 0;
      const bStart = parseDate(b.startDate)?.getTime() || 0;
      if (aStart !== bStart) return aStart - bStart;
      return a.title.localeCompare(b.title);
    });

    const summaryEntry = summaryMap.get(project.id);
    const summaryText = summaryEntry?.summary || `${sortedItems.length} roadmap item${sortedItems.length > 1 ? 's' : ''} prepared for execution tracking.`;
    const milestoneTags = (summaryEntry?.keyMilestones && summaryEntry.keyMilestones.length > 0)
      ? summaryEntry.keyMilestones
      : sortedItems.filter(item => item.isMilestone).map(item => item.title).slice(0, 4);

    const validDates = sortedItems
      .flatMap(item => [parseDate(item.startDate), parseDate(item.endDate)])
      .filter((date): date is Date => date instanceof Date);

    const minDate = validDates.length > 0 ? new Date(Math.min(...validDates.map(date => date.getTime()))) : new Date(`${todayISO()}T00:00:00`);
    let maxDate = validDates.length > 0 ? new Date(Math.max(...validDates.map(date => date.getTime()))) : addDays(minDate, 30);
    if (maxDate.getTime() <= minDate.getTime()) maxDate = addDays(minDate, 14);
    const totalMs = Math.max(24 * 3600 * 1000, maxDate.getTime() - minDate.getTime());

    const axisLabels = tickPercents.map(percent => {
      const d = new Date(minDate.getTime() + (totalMs * percent) / 100);
      return {
        percent,
        label: d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
      };
    });

    const rowsHTML = sortedItems.map(item => {
      const start = parseDate(item.startDate) || minDate;
      const rawEnd = parseDate(item.endDate) || start;
      const end = rawEnd.getTime() < start.getTime() ? start : rawEnd;
      const leftPct = ((start.getTime() - minDate.getTime()) / totalMs) * 100;
      const rightPct = ((end.getTime() - minDate.getTime()) / totalMs) * 100;
      const widthPct = Math.max(item.isMilestone ? 0 : 2, rightPct - leftPct);
      const color = STATUS_COLORS[item.status] || '#64748b';

      return `
      <div class="gantt-row">
        <div class="task-cell">
          <strong>${escapeHTML(item.title || 'Untitled')}</strong>
          ${item.description ? `<div class="task-desc">${escapeHTML(item.description)}</div>` : ''}
        </div>
        <div class="owner-cell">${escapeHTML(item.owner || '—')}</div>
        <div class="status-cell"><span class="status-chip" style="background:${color}">${escapeHTML(item.status)}</span></div>
        <div class="date-cell">${escapeHTML(formatShortDate(item.startDate))} → ${escapeHTML(formatShortDate(item.endDate))}</div>
        <div class="prog-cell"><span class="progress-pill">${clampPct(item.progressPct)}%</span></div>
        <div class="timeline-wrap">
          <div class="timeline-track">
            ${tickPercents.map(percent => `<span class="timeline-grid-line" style="left:${percent}%"></span>`).join('')}
            ${item.isMilestone
              ? `<div class="task-milestone" style="left:${leftPct}%;background:${color}"></div>`
              : `<div class="task-bar" style="left:${leftPct}%;width:${widthPct}%;background:${color}"></div>`}
          </div>
        </div>
      </div>`;
    }).join('');

    return `
    <section class="project-block">
      <div class="project-banner">
        <div>
          <h2>${escapeHTML(project.name)}</h2>
          <div class="project-meta">${escapeHTML(project.teamName)} • ${escapeHTML(project.status)} • Deadline: ${escapeHTML(project.deadline)}</div>
        </div>
        <div style="font-size:11px;font-weight:700">Items: ${sortedItems.length}</div>
      </div>
      <div class="project-body">
        <div class="project-summary">${escapeHTML(summaryText)}</div>
        ${milestoneTags.length > 0
          ? `<div class="milestone-list">${milestoneTags.map(m => `<span class="milestone-tag">${escapeHTML(m)}</span>`).join('')}</div>`
          : ''}
        ${sortedItems.length > 0 ? `
          <div class="timeline-head">
            <div>Roadmap Item</div>
            <div>Owner</div>
            <div>Status</div>
            <div>Dates</div>
            <div>Progress</div>
            <div>Timeline</div>
          </div>
          <div class="timeline-axis">
            ${axisLabels.map(t => `
              <span class="axis-line" style="left:${t.percent}%"></span>
              <span class="axis-label" style="left:${t.percent}%">${escapeHTML(t.label)}</span>
            `).join('')}
          </div>
          ${rowsHTML}
        ` : `<div class="empty-state">No roadmap items defined yet for this project.</div>`}
      </div>
    </section>`;
  }).join('');

  return `${css}
  <div class="deck">
    <header class="deck-header">
      <div>
        <div class="deck-title">DOINg - PM Gant Roadmap</div>
        <div class="deck-sub">Generated ${escapeHTML(now)} • ${projects.length} project${projects.length > 1 ? 's' : ''}</div>
      </div>
      <div class="deck-conf">Professional Roadmap Export</div>
    </header>
    <section class="exec-summary">${escapeHTML(narrative.executiveSummary || 'Roadmap generated from selected projects and manually curated PM inputs.')}</section>
    ${projectSections}
  </div>`;
};

const PMGant: React.FC<PMGantProps> = ({
  teams,
  users,
  currentUser,
  llmConfig,
  gantItems,
  onSaveItem,
  onDeleteItem,
}) => {
  const [view, setView] = useState<PMGantView>('workspace');
  const [selectedProjectIds, setSelectedProjectIds] = useState<string[]>([]);
  const [generatedHTML, setGeneratedHTML] = useState('');
  const [isGeneratingDoc, setIsGeneratingDoc] = useState(false);
  const [isAIAutofilling, setIsAIAutofilling] = useState(false);
  const [aiInput, setAiInput] = useState('');
  const [aiError, setAiError] = useState('');
  const [editingItemId, setEditingItemId] = useState<string | null>(null);

  const allProjects = useMemo<PMGantProject[]>(() => {
    const list: PMGantProject[] = [];
    teams.forEach(team => {
      (team.projects || []).forEach(project => {
        if (!project.isArchived) {
          list.push({ ...project, teamName: team.name });
        }
      });
    });
    return list;
  }, [teams]);

  const projectById = useMemo(() => {
    const map = new Map<string, PMGantProject>();
    allProjects.forEach(project => map.set(project.id, project));
    return map;
  }, [allProjects]);

  useEffect(() => {
    setSelectedProjectIds(prev => prev.filter(id => projectById.has(id)));
  }, [projectById]);

  useEffect(() => {
    if (selectedProjectIds.length === 0 && allProjects.length > 0) {
      setSelectedProjectIds([allProjects[0].id]);
    }
  }, [allProjects, selectedProjectIds.length]);

  const defaultProjectId = selectedProjectIds[0] || allProjects[0]?.id || '';
  const [form, setForm] = useState<PMGantFormState>({
    projectId: defaultProjectId,
    title: '',
    description: '',
    owner: `${currentUser.firstName} ${currentUser.lastName}`.trim(),
    startDate: todayISO(),
    endDate: plusDaysISO(14),
    progressPct: 0,
    status: 'Planned',
    priority: 'Medium',
    isMilestone: false,
    notes: '',
  });

  useEffect(() => {
    if (!form.projectId && defaultProjectId) {
      setForm(prev => ({ ...prev, projectId: defaultProjectId }));
    }
  }, [defaultProjectId, form.projectId]);

  const selectedProjects = useMemo(() => {
    return allProjects.filter(project => selectedProjectIds.includes(project.id));
  }, [allProjects, selectedProjectIds]);

  const selectedItems = useMemo(() => {
    return gantItems
      .filter(item => selectedProjectIds.includes(item.projectId))
      .sort((a, b) => {
        const aStart = parseDate(a.startDate)?.getTime() || 0;
        const bStart = parseDate(b.startDate)?.getTime() || 0;
        if (aStart !== bStart) return aStart - bStart;
        return a.title.localeCompare(b.title);
      });
  }, [gantItems, selectedProjectIds]);

  const resetForm = () => {
    setEditingItemId(null);
    setForm({
      projectId: defaultProjectId,
      title: '',
      description: '',
      owner: `${currentUser.firstName} ${currentUser.lastName}`.trim(),
      startDate: todayISO(),
      endDate: plusDaysISO(14),
      progressPct: 0,
      status: 'Planned',
      priority: 'Medium',
      isMilestone: false,
      notes: '',
    });
    setAiInput('');
    setAiError('');
  };

  const toggleProject = (projectId: string) => {
    setSelectedProjectIds(prev =>
      prev.includes(projectId) ? prev.filter(id => id !== projectId) : [...prev, projectId]
    );
  };

  const handleSelectAll = () => setSelectedProjectIds(allProjects.map(project => project.id));
  const handleClearSelection = () => setSelectedProjectIds([]);

  const handleEditItem = (item: PMGanttItem) => {
    setEditingItemId(item.id);
    setForm({
      projectId: item.projectId,
      title: item.title,
      description: item.description,
      owner: item.owner,
      startDate: item.startDate,
      endDate: item.endDate,
      progressPct: clampPct(item.progressPct),
      status: item.status,
      priority: item.priority,
      isMilestone: item.isMilestone === true,
      notes: item.notes || '',
    });
    if (!selectedProjectIds.includes(item.projectId)) {
      setSelectedProjectIds(prev => [...prev, item.projectId]);
    }
  };

  const handleDeleteItem = (itemId: string) => {
    if (!window.confirm('Delete this roadmap item?')) return;
    onDeleteItem(itemId);
    if (editingItemId === itemId) resetForm();
  };

  const handleSaveRoadmapItem = () => {
    if (!form.projectId) return alert('Please select a project.');
    if (!form.title.trim()) return alert('Roadmap item title is required.');
    if (!form.startDate || !form.endDate) return alert('Start and end dates are required.');

    const start = parseDate(form.startDate);
    const end = parseDate(form.endDate);
    if (!start || !end) return alert('Invalid dates.');
    if (end.getTime() < start.getTime()) return alert('End date must be on or after start date.');

    const now = new Date().toISOString();
    const existing = editingItemId ? gantItems.find(item => item.id === editingItemId) : null;
    const nextItem: PMGanttItem = {
      id: editingItemId || generateId(),
      projectId: form.projectId,
      createdByUserId: existing?.createdByUserId || currentUser.id,
      createdAt: existing?.createdAt || now,
      updatedAt: now,
      title: form.title.trim(),
      description: form.description.trim(),
      owner: form.owner.trim(),
      startDate: form.startDate,
      endDate: form.endDate,
      progressPct: clampPct(form.progressPct),
      status: form.status,
      priority: form.priority,
      isMilestone: form.isMilestone,
      notes: form.notes.trim(),
    };

    onSaveItem(nextItem);
    resetForm();
  };

  const handleAIAutofill = async () => {
    if (!aiInput.trim()) {
      setAiError('Paste text first to use AI autofill.');
      return;
    }
    setAiError('');
    setIsAIAutofilling(true);
    try {
      const extracted = await extractPMGanttItemFromText(aiInput, llmConfig);
      setForm(prev => ({
        ...prev,
        projectId: prev.projectId || defaultProjectId,
        title: extracted.title || prev.title,
        description: extracted.description || prev.description,
        owner: extracted.owner || prev.owner,
        startDate: extracted.startDate || prev.startDate,
        endDate: extracted.endDate || prev.endDate,
        status: extracted.status || prev.status,
        priority: extracted.priority || prev.priority,
        progressPct: extracted.progressPct != null ? clampPct(extracted.progressPct) : prev.progressPct,
        isMilestone: extracted.isMilestone != null ? extracted.isMilestone : prev.isMilestone,
        notes: extracted.notes || prev.notes,
      }));
    } catch (e: any) {
      setAiError(e?.message || 'AI autofill failed.');
    } finally {
      setIsAIAutofilling(false);
    }
  };

  const handleGenerateRoadmap = async () => {
    if (selectedProjects.length === 0) {
      alert('Select at least one project.');
      return;
    }

    const bundles = selectedProjects.map(project => ({
      project,
      items: gantItems.filter(item => item.projectId === project.id),
    }));
    const hasAtLeastOneItem = bundles.some(bundle => bundle.items.length > 0);
    if (!hasAtLeastOneItem) {
      alert('Add roadmap items to selected projects before generating the document.');
      return;
    }

    setIsGeneratingDoc(true);
    try {
      const fallback = buildDefaultNarrative(bundles);
      let narrative: PMGantNarrative = fallback;

      try {
        const aiNarrative = await generatePMGanttNarrative(
          bundles.map(bundle => ({
            id: bundle.project.id,
            name: bundle.project.name,
            teamName: bundle.project.teamName,
            status: bundle.project.status,
            deadline: bundle.project.deadline,
            items: bundle.items,
          })),
          llmConfig
        );
        narrative = {
          executiveSummary: aiNarrative.executiveSummary || fallback.executiveSummary,
          projectSummaries: aiNarrative.projectSummaries?.length ? aiNarrative.projectSummaries : fallback.projectSummaries,
        };
      } catch (aiError) {
        console.warn('PM Gant narrative AI failed, using deterministic fallback', aiError);
      }

      const html = buildPMGantDeckHTML(bundles, narrative);
      setGeneratedHTML(html);
      setView('preview');
    } finally {
      setIsGeneratingDoc(false);
    }
  };

  const buildExportDocumentHTML = (withPrintControls: boolean) => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>DOINg - PM Gant Roadmap</title>
<style>@page{size:landscape;margin:9mm}@media print{body{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;font-size:11px;line-height:1.35}.no-print{display:none!important}.deck{max-width:none}.project-block,.gantt-row{break-inside:avoid-page;page-break-inside:avoid}}</style>
</head><body style="background:#fff;margin:0;padding:18px;font-family:'Segoe UI',system-ui,-apple-system,sans-serif">
${withPrintControls ? `
<div class="no-print" style="text-align:center;margin-bottom:20px;padding:14px;background:#f8fafc;border-radius:12px;border:1px solid #e2e8f0">
  <button onclick="window.print()" style="background:linear-gradient(135deg,#00915A,#00A86B);color:#fff;border:none;padding:12px 34px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;box-shadow:0 4px 12px rgba(0,145,90,.28)">Print / Save as PDF</button>
  <button onclick="window.close()" style="margin-left:10px;background:#fff;color:#334155;border:1px solid #cbd5e1;padding:12px 24px;border-radius:10px;font-size:14px;cursor:pointer">Cancel</button>
</div>` : ''}
${generatedHTML}</body></html>`;

  const handleExportPDF = () => {
    const w = window.open('', '_blank');
    if (!w) return;
    w.document.write(buildExportDocumentHTML(true));
    w.document.close();
  };

  const handleExportEmail = () => {
    if (!generatedHTML.trim()) return;
    const fileDate = todayISO();
    const subjectDate = new Date().toLocaleDateString('en-GB');
    const htmlBody = buildExportDocumentHTML(false);
    const emlContent = [
      'X-Unsent: 1',
      'To: ',
      `Subject: DOINg - PM Gant Roadmap - ${subjectDate}`,
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      htmlBody
    ].join('\r\n');

    const blob = new Blob([emlContent], { type: 'message/rfc822;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `doing-pm-gant-roadmap-${fileDate}.eml`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  if (view === 'preview') {
    return (
      <div className="max-w-7xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h2 className="text-xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
              <Eye className="w-5 h-5 text-indigo-500" /> PM Gant Preview
            </h2>
            <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
              Professional roadmap document generated from selected projects and roadmap items
            </p>
          </div>
          <div className="flex gap-2">
            <button
              onClick={() => setView('workspace')}
              className="px-4 py-2 text-sm font-medium text-gray-600 bg-gray-100 dark:bg-gray-800 dark:text-gray-300 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-700 transition-colors"
            >
              Back
            </button>
            <button
              onClick={handleExportEmail}
              className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-700 rounded-lg hover:bg-emerald-100 dark:hover:bg-emerald-900/40 transition-colors"
            >
              <Mail className="w-4 h-4" /> Export Email (.eml)
            </button>
            <button
              onClick={handleExportPDF}
              className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-white bg-gradient-to-r from-indigo-600 to-sky-600 rounded-lg hover:from-indigo-700 hover:to-sky-700 transition-all shadow-md"
            >
              <Download className="w-4 h-4" /> Export PDF
            </button>
          </div>
        </div>

        <div className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-6 shadow-sm overflow-auto" style={{ maxHeight: '78vh' }}>
          <div dangerouslySetInnerHTML={{ __html: generatedHTML }} />
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-extrabold text-gray-900 dark:text-white flex items-center gap-3">
            <span className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center text-white shadow-lg">
              <GanttChartSquare className="w-5 h-5" />
            </span>
            PM Gant
          </h2>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-2">
            Build a board-ready roadmap workflow: select projects, curate milestones, generate a polished Gantt document with local AI.
          </p>
        </div>
        <button
          onClick={handleGenerateRoadmap}
          disabled={isGeneratingDoc || selectedProjectIds.length === 0}
          className={`flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-bold transition-all shadow-md ${
            selectedProjectIds.length === 0
              ? 'bg-gray-200 text-gray-400 cursor-not-allowed dark:bg-gray-800 dark:text-gray-600'
              : 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white hover:from-emerald-700 hover:to-teal-700'
          }`}
        >
          {isGeneratingDoc ? (
            <>
              <Sparkles className="w-4 h-4 animate-spin" /> Generating...
            </>
          ) : (
            <>
              <FileBarChart className="w-4 h-4" /> Generate PM Gant Document
            </>
          )}
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <section className="lg:col-span-4 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-gray-900 dark:text-white">1. Select Projects</h3>
            <div className="flex items-center gap-2">
              <button onClick={handleSelectAll} className="text-xs font-semibold text-indigo-600 hover:underline">All</button>
              <button onClick={handleClearSelection} className="text-xs font-semibold text-gray-500 hover:underline">Clear</button>
            </div>
          </div>
          <div className="space-y-2 max-h-[360px] overflow-y-auto pr-1">
            {allProjects.map(project => {
              const selected = selectedProjectIds.includes(project.id);
              const itemCount = gantItems.filter(item => item.projectId === project.id).length;
              return (
                <button
                  key={project.id}
                  onClick={() => toggleProject(project.id)}
                  className={`w-full text-left p-3 rounded-lg border transition-colors ${
                    selected
                      ? 'border-emerald-400 bg-emerald-50 dark:bg-emerald-900/20'
                      : 'border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40 hover:bg-gray-100 dark:hover:bg-gray-800'
                  }`}
                >
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="text-sm font-semibold text-gray-900 dark:text-white">{project.name}</p>
                      <p className="text-xs text-gray-500 dark:text-gray-400">{project.teamName}</p>
                    </div>
                    {selected ? <CheckCircle2 className="w-4 h-4 text-emerald-500 mt-0.5" /> : <span className="w-4 h-4 rounded-full border border-gray-300 dark:border-gray-600 mt-0.5" />}
                  </div>
                  <div className="mt-2 flex items-center gap-3 text-[11px] text-gray-500 dark:text-gray-400">
                    <span className="flex items-center gap-1"><Calendar className="w-3 h-3" /> {project.deadline || 'No deadline'}</span>
                    <span>{itemCount} item{itemCount > 1 ? 's' : ''}</span>
                  </div>
                </button>
              );
            })}
            {allProjects.length === 0 && (
              <p className="text-sm text-gray-400 dark:text-gray-500 py-8 text-center">No accessible projects.</p>
            )}
          </div>
        </section>

        <section className="lg:col-span-8 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-sm font-bold text-gray-900 dark:text-white">
              2. Roadmap Item Form {editingItemId ? '(Edit Mode)' : '(Create Mode)'}
            </h3>
            {editingItemId && (
              <button onClick={resetForm} className="text-xs font-semibold text-gray-500 hover:underline">Cancel edit</button>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Project</label>
              <select
                value={form.projectId}
                onChange={e => setForm(prev => ({ ...prev, projectId: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              >
                <option value="">Select project...</option>
                {allProjects.map(project => (
                  <option key={project.id} value={project.id}>{project.teamName} • {project.name}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Owner</label>
              <input
                value={form.owner}
                onChange={e => setForm(prev => ({ ...prev, owner: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                placeholder="Owner name"
              />
            </div>
            <div className="md:col-span-2">
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Title</label>
              <input
                value={form.title}
                onChange={e => setForm(prev => ({ ...prev, title: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                placeholder="Milestone or workstream title"
              />
            </div>
            <div className="md:col-span-2">
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Description</label>
              <textarea
                value={form.description}
                onChange={e => setForm(prev => ({ ...prev, description: e.target.value }))}
                rows={3}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 resize-none"
                placeholder="Concise description for stakeholders"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Start Date</label>
              <input
                type="date"
                value={form.startDate}
                onChange={e => setForm(prev => ({ ...prev, startDate: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">End Date</label>
              <input
                type="date"
                value={form.endDate}
                onChange={e => setForm(prev => ({ ...prev, endDate: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Status</label>
              <select
                value={form.status}
                onChange={e => setForm(prev => ({ ...prev, status: e.target.value as PMGanttStatus }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              >
                {STATUS_VALUES.map(status => <option key={status} value={status}>{status}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Priority</label>
              <select
                value={form.priority}
                onChange={e => setForm(prev => ({ ...prev, priority: e.target.value as PMGanttPriority }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              >
                {PRIORITY_VALUES.map(priority => <option key={priority} value={priority}>{priority}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Progress {clampPct(form.progressPct)}%</label>
              <input
                type="range"
                min={0}
                max={100}
                value={form.progressPct}
                onChange={e => setForm(prev => ({ ...prev, progressPct: clampPct(parseInt(e.target.value, 10) || 0) }))}
                className="w-full accent-emerald-600"
              />
            </div>
            <div className="flex items-end">
              <label className="inline-flex items-center gap-2 text-sm font-medium text-gray-700 dark:text-gray-300 cursor-pointer">
                <input
                  type="checkbox"
                  checked={form.isMilestone}
                  onChange={e => setForm(prev => ({ ...prev, isMilestone: e.target.checked }))}
                  className="w-4 h-4 rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                />
                Milestone item
              </label>
            </div>
            <div className="md:col-span-2">
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Notes (optional)</label>
              <textarea
                value={form.notes}
                onChange={e => setForm(prev => ({ ...prev, notes: e.target.value }))}
                rows={2}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 resize-none"
                placeholder="Dependencies, external constraints, assumptions..."
              />
            </div>
          </div>

          <div className="flex gap-2 mt-4">
            <button
              onClick={handleSaveRoadmapItem}
              className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-indigo-600 rounded-lg hover:bg-indigo-700 transition-colors"
            >
              <Save className="w-4 h-4" /> {editingItemId ? 'Update Item' : 'Add Item'}
            </button>
            <button
              onClick={resetForm}
              className="px-4 py-2 text-sm font-medium text-gray-600 bg-gray-100 dark:bg-gray-800 dark:text-gray-300 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-700 transition-colors"
            >
              Reset
            </button>
          </div>

          <div className="mt-5 p-4 rounded-lg border border-indigo-100 dark:border-indigo-800 bg-indigo-50/60 dark:bg-indigo-900/20">
            <div className="flex items-center justify-between mb-2">
              <h4 className="text-sm font-bold text-indigo-900 dark:text-indigo-100 flex items-center gap-2">
                <Wand2 className="w-4 h-4" /> AI Autofill
              </h4>
              <button
                onClick={handleAIAutofill}
                disabled={isAIAutofilling || !aiInput.trim()}
                className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:opacity-50 transition-colors"
              >
                <Sparkles className={`w-3.5 h-3.5 ${isAIAutofilling ? 'animate-spin' : ''}`} />
                Autofill from text
              </button>
            </div>
            <p className="text-xs text-indigo-700 dark:text-indigo-300 mb-2">
              Paste rough notes, an email excerpt, or a planning paragraph. Local LLM will structure title/dates/status/milestone fields.
            </p>
            {aiError && (
              <p className="text-xs text-red-600 dark:text-red-400 mb-2">{aiError}</p>
            )}
            <textarea
              value={aiInput}
              onChange={e => setAiInput(e.target.value)}
              rows={3}
              className="w-full px-3 py-2 text-sm border border-indigo-200 dark:border-indigo-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 resize-none"
              placeholder="Example: Final UAT validation from March 10 to March 24 led by Claire, high priority, 20% complete, milestone delivery by March 24..."
            />
          </div>
        </section>
      </div>

      <section className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-sm font-bold text-gray-900 dark:text-white">3. Curated Roadmap Items ({selectedItems.length})</h3>
          <p className="text-xs text-gray-500 dark:text-gray-400">Only items from selected projects are shown</p>
        </div>
        {selectedItems.length === 0 ? (
          <div className="text-sm text-gray-400 dark:text-gray-500 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg p-6 text-center">
            No roadmap items for current selection.
          </div>
        ) : (
          <div className="space-y-2 max-h-[420px] overflow-y-auto pr-1">
            {selectedItems.map(item => {
              const project = projectById.get(item.projectId);
              return (
                <div key={item.id} className="p-3 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40">
                  <div className="flex items-start justify-between gap-3">
                    <div className="min-w-0">
                      <p className="text-sm font-semibold text-gray-900 dark:text-white truncate">{item.title}</p>
                      <p className="text-xs text-gray-500 dark:text-gray-400 truncate">{project?.teamName} • {project?.name}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => handleEditItem(item)}
                        className="inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold text-indigo-600 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-700 rounded-md hover:bg-indigo-50 dark:hover:bg-indigo-900/20"
                      >
                        <Edit3 className="w-3 h-3" /> Edit
                      </button>
                      <button
                        onClick={() => handleDeleteItem(item.id)}
                        className="inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold text-red-600 dark:text-red-300 border border-red-200 dark:border-red-700 rounded-md hover:bg-red-50 dark:hover:bg-red-900/20"
                      >
                        <Trash2 className="w-3 h-3" /> Delete
                      </button>
                    </div>
                  </div>
                  <div className="mt-2 flex flex-wrap items-center gap-2 text-[11px]">
                    <span className={`px-2 py-0.5 rounded-full border font-bold ${statusBadgeClass(item.status)}`}>{item.status}</span>
                    <span className={`px-2 py-0.5 rounded-full border font-bold ${priorityBadgeClass(item.priority)}`}>{item.priority}</span>
                    <span className="px-2 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700">{item.progressPct}%</span>
                    {item.isMilestone && <span className="px-2 py-0.5 rounded-full bg-emerald-100 dark:bg-emerald-900/20 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-700 font-bold">Milestone</span>}
                    <span className="text-gray-500 dark:text-gray-400">{formatShortDate(item.startDate)} → {formatShortDate(item.endDate)}</span>
                  </div>
                  {item.description && <p className="mt-2 text-xs text-gray-600 dark:text-gray-300">{item.description}</p>}
                </div>
              );
            })}
          </div>
        )}
      </section>
    </div>
  );
};

export default PMGant;
