import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  Calendar,
  CheckCircle2,
  Download,
  Edit3,
  Eye,
  FileBarChart,
  GanttChartSquare,
  Link2,
  Mail,
  Plus,
  Search,
  Save,
  Sparkles,
  Star,
  Trash2,
  Wand2
} from 'lucide-react';
import { LLMConfig, PMGanttItem, PMGanttPriority, PMGanttStatus, Project, Team, User } from '../types';
import { generateId } from '../services/storage';
import {
  extractPMGanttItemFromText,
  extractPMGanttItemsFromText,
  generatePMGanttMilestonesFromProject,
  generatePMGanttNarrative
} from '../services/llmService';

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
export type PMGantProject = Project & { teamName: string };

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
  isMajorDelivery: boolean;
  dependsOnIds: string[];
  notes: string;
}

export type PMGantNarrative = {
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
const normalizeTitle = (value: string): string => value.trim().toLowerCase().replace(/\s+/g, ' ');

const isMajorDeliveryItem = (item: PMGanttItem): boolean =>
  item.isMajorDelivery === true || (item.isMilestone === true && item.priority === 'Critical');

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

export const buildDefaultNarrative = (projects: Array<{ project: PMGantProject; items: PMGanttItem[] }>): PMGantNarrative => {
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

export const buildPMGantDeckHTML = (
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
    .project-body{
      border:1px solid #dbe5ef;border-top:none;border-radius:0 0 12px 12px;background:#fff;padding:14px 16px;
      --gantt-col-item:250px;
      --gantt-col-owner:100px;
      --gantt-col-status:90px;
      --gantt-col-dates:130px;
      --gantt-col-signals:90px;
      --gantt-col-timeline:minmax(380px,1fr);
      --gantt-col-gap:10px;
      --gantt-grid:var(--gantt-col-item) var(--gantt-col-owner) var(--gantt-col-status) var(--gantt-col-dates) var(--gantt-col-signals) var(--gantt-col-timeline);
    }
    .project-summary{background:#f8fafc;border-left:4px solid #00915A;border-radius:10px;padding:10px 12px;font-size:12px;line-height:1.5;color:#334155;margin-bottom:12px}
    .milestone-list{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 12px 0}
    .milestone-tag{font-size:10px;font-weight:700;color:#065f46;background:#d1fae5;border:1px solid #a7f3d0;padding:3px 8px;border-radius:999px}
    .dependency-overview{background:#fffbea;border:1px solid #fde68a;border-radius:10px;padding:8px 10px;margin-bottom:10px}
    .dependency-overview-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:#92400e;margin-bottom:4px}
    .dependency-item{font-size:10px;color:#78350f;line-height:1.4}
    .timeline-head{display:grid;grid-template-columns:var(--gantt-grid);gap:var(--gantt-col-gap);font-size:10px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.5px;padding:6px 0;border-bottom:1px solid #e2e8f0}
    .gantt-row{display:grid;grid-template-columns:var(--gantt-grid);gap:var(--gantt-col-gap);align-items:center;padding:8px 0;border-bottom:1px solid #f1f5f9;page-break-inside:avoid;break-inside:avoid-page}
    .gantt-row:last-child{border-bottom:none}
    .task-cell{font-size:12px;color:#0f172a}
    .task-cell strong{display:block;font-size:12px}
    .task-desc{font-size:10px;color:#64748b;margin-top:2px;line-height:1.4}
    .task-meta-badges{display:flex;flex-wrap:wrap;gap:4px;margin-top:4px}
    .meta-pill{font-size:9px;font-weight:700;padding:2px 6px;border-radius:999px;border:1px solid #cbd5e1;background:#f8fafc;color:#334155}
    .meta-pill.parallel{background:#eef2ff;border-color:#c7d2fe;color:#3730a3}
    .meta-pill.dep{background:#eff6ff;border-color:#bfdbfe;color:#1e40af}
    .meta-pill.major{background:#fff7ed;border-color:#fdba74;color:#9a3412}
    .owner-cell,.status-cell,.date-cell,.prog-cell{font-size:11px;color:#334155}
    .status-chip{display:inline-flex;padding:2px 8px;border-radius:999px;font-weight:700;color:#fff;font-size:10px;white-space:nowrap}
    .progress-pill{display:inline-flex;min-width:42px;justify-content:center;padding:2px 8px;border-radius:999px;font-weight:700;font-size:10px;background:#e2e8f0;color:#0f172a}
    .prog-flags{margin-top:3px;font-size:9px;color:#64748b;line-height:1.3}
    .timeline-wrap{position:relative}
    .timeline-axis-row{display:grid;grid-template-columns:var(--gantt-grid);gap:var(--gantt-col-gap);padding:5px 0 7px 0}
    .timeline-axis-spacer{grid-column:1 / span 5}
    .timeline-axis{grid-column:6;position:relative;height:20px;margin:0}
    .axis-line{position:absolute;top:0;bottom:0;border-left:1px dashed #cbd5e1}
    .axis-label{position:absolute;top:0;transform:translateX(-50%);font-size:9px;color:#64748b;white-space:nowrap;background:#fff;padding:0 3px}
    .axis-label-start{transform:none}
    .axis-label-end{transform:translateX(-100%)}
    .today-axis-line{position:absolute;top:0;bottom:0;border-left:1px solid #ef4444;opacity:.75}
    .today-axis-label{position:absolute;top:10px;transform:translateX(-50%);font-size:9px;font-weight:800;color:#b91c1c;background:#fee2e2;border:1px solid #fca5a5;border-radius:999px;padding:0 6px}
    .timeline-track{position:relative;height:24px;border-radius:7px;background:#f8fafc;border:1px solid #dbe5ef;overflow:visible}
    .timeline-grid-line{position:absolute;top:0;bottom:0;border-left:1px dotted #e2e8f0}
    .timeline-today-line{position:absolute;top:0;bottom:0;border-left:1px solid #ef4444;opacity:.75}
    .dependency-link{position:absolute;top:11px;height:2px;background:#2563eb;opacity:.7;border-radius:2px}
    .dependency-link.overlap{background:#f59e0b}
    .task-bar{position:absolute;top:5px;height:12px;border-radius:8px;box-shadow:0 1px 4px rgba(15,23,42,.2)}
    .task-bar.milestone-bar{background:linear-gradient(90deg,#64748b,#0ea5e9)}
    .task-edge{position:absolute;top:4px;width:10px;height:10px;border-radius:999px;border:2px solid #fff;box-shadow:0 1px 3px rgba(15,23,42,.2)}
    .task-edge.start{transform:translateX(-50%)}
    .task-edge.end{transform:translateX(-50%)}
    .task-done-dot{position:absolute;top:8px;width:8px;height:8px;border-radius:999px;background:#16a34a;border:1px solid #fff;transform:translateX(-50%);box-shadow:0 0 0 1px rgba(22,163,74,.28)}
    .task-major-star{position:absolute;top:-3px;transform:translateX(-50%);font-size:12px;color:#d4a017;text-shadow:0 1px 2px rgba(0,0,0,.2)}
    .empty-state{padding:14px;border:1px dashed #cbd5e1;border-radius:10px;font-size:12px;color:#64748b;background:#f8fafc}
    @media print{
      .deck{max-width:none}
      .project-block{margin-bottom:14px}
      .project-body{padding:10px 12px}
      .timeline-head{font-size:9px}
      .gantt-row{padding:6px 0}
      .today-axis-label{font-size:8px}
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

    const timelineSource = sortedItems.filter(item => item.isMilestone).length > 0
      ? sortedItems.filter(item => item.isMilestone)
      : sortedItems;

    const validDates = timelineSource
      .flatMap(item => [parseDate(item.startDate), parseDate(item.endDate)])
      .filter((date): date is Date => date instanceof Date);

    const minRawDate = validDates.length > 0 ? new Date(Math.min(...validDates.map(date => date.getTime()))) : new Date(`${todayISO()}T00:00:00`);
    let maxRawDate = validDates.length > 0 ? new Date(Math.max(...validDates.map(date => date.getTime()))) : addDays(minRawDate, 30);
    if (maxRawDate.getTime() <= minRawDate.getTime()) maxRawDate = addDays(minRawDate, 14);
    const minDate = addDays(minRawDate, -2);
    let maxDate = addDays(maxRawDate, 2);
    if (maxDate.getTime() <= minDate.getTime()) maxDate = addDays(minDate, 14);
    const totalMs = Math.max(24 * 3600 * 1000, maxDate.getTime() - minDate.getTime());
    const toPct = (date: Date): number => Math.max(0, Math.min(100, ((date.getTime() - minDate.getTime()) / totalMs) * 100));

    const today = parseDate(todayISO()) || new Date(`${todayISO()}T00:00:00`);
    const todayPct = toPct(today);

    const axisLabels = tickPercents.map(percent => {
      const d = new Date(minDate.getTime() + (totalMs * percent) / 100);
      return {
        percent,
        label: d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }),
      };
    });

    const itemById = new Map(sortedItems.map(item => [item.id, item]));
    const positionById = new Map<string, { start: Date; end: Date; startPct: number; endPct: number }>();

    sortedItems.forEach(item => {
      const start = parseDate(item.startDate) || minDate;
      const rawEnd = parseDate(item.endDate) || start;
      const end = rawEnd.getTime() < start.getTime() ? start : rawEnd;
      positionById.set(item.id, {
        start,
        end,
        startPct: toPct(start),
        endPct: toPct(end),
      });
    });

    const overlapCountById = new Map<string, number>();
    sortedItems.forEach(item => {
      const pos = positionById.get(item.id);
      if (!pos) {
        overlapCountById.set(item.id, 0);
        return;
      }
      const count = sortedItems.filter(other => {
        if (other.id === item.id) return false;
        const otherPos = positionById.get(other.id);
        if (!otherPos) return false;
        return pos.start.getTime() <= otherPos.end.getTime() && otherPos.start.getTime() <= pos.end.getTime();
      }).length;
      overlapCountById.set(item.id, count);
    });

    const dependencyChains = sortedItems.flatMap(item =>
      (item.dependsOnIds || [])
        .map(depId => {
          const dep = itemById.get(depId);
          if (!dep) return null;
          return { from: dep.title, to: item.title };
        })
        .filter((entry): entry is { from: string; to: string } => Boolean(entry))
    );

    const rowsHTML = sortedItems.map(item => {
      const pos = positionById.get(item.id);
      if (!pos) return '';
      const leftPct = pos.startPct;
      const rightPct = pos.endPct;
      const widthPct = Math.max(item.isMilestone ? 1 : 2, rightPct - leftPct);
      const color = STATUS_COLORS[item.status] || '#64748b';
      const overlapCount = overlapCountById.get(item.id) || 0;
      const dependencyTitles = (item.dependsOnIds || [])
        .map(depId => itemById.get(depId)?.title)
        .filter((title): title is string => Boolean(title));
      const dependencyEnds = (item.dependsOnIds || [])
        .map(depId => positionById.get(depId)?.endPct)
        .filter((pct): pct is number => Number.isFinite(pct));
      const dependencySignalLabel =
        dependencyTitles.length === 0
          ? 'No dependency'
          : dependencyTitles.length === 1
            ? '1 dependency'
            : `${dependencyTitles.length} dependencies`;

      let dependencyLinkHTML = '';
      if (dependencyEnds.length > 0) {
        const anchorPct = Math.max(...dependencyEnds);
        const left = Math.min(anchorPct, leftPct);
        const width = Math.max(0.6, Math.abs(leftPct - anchorPct));
        const overlapClass = anchorPct > leftPct ? ' overlap' : '';
        dependencyLinkHTML = `<span class="dependency-link${overlapClass}" style="left:${left}%;width:${width}%"></span>`;
      }

      const doneMarker = (item.status === 'Done' || item.progressPct >= 100)
        ? `<span class="task-done-dot" style="left:${rightPct}%"></span>`
        : '';
      const majorDelivery = isMajorDeliveryItem(item);
      const majorMarker = majorDelivery
        ? `<span class="task-major-star" style="left:${rightPct}%">★</span>`
        : '';

      return `
      <div class="gantt-row">
        <div class="task-cell">
          <strong>${escapeHTML(item.title || 'Untitled')}</strong>
          ${item.description ? `<div class="task-desc">${escapeHTML(item.description)}</div>` : ''}
          ${item.isMilestone ? `<div class="task-desc">Milestone window: ${escapeHTML(formatShortDate(item.startDate))} → ${escapeHTML(formatShortDate(item.endDate))}</div>` : ''}
          <div class="task-meta-badges">
            ${overlapCount > 0 ? `<span class="meta-pill parallel">Parallel with ${overlapCount} item${overlapCount > 1 ? 's' : ''}</span>` : ''}
            ${dependencyTitles.length > 0 ? `<span class="meta-pill dep">Depends on: ${escapeHTML(dependencyTitles.join(' • '))}</span>` : ''}
            ${majorDelivery ? `<span class="meta-pill major">Major Delivery ★</span>` : ''}
          </div>
        </div>
        <div class="owner-cell">${escapeHTML(item.owner || '—')}</div>
        <div class="status-cell"><span class="status-chip" style="background:${color}">${escapeHTML(item.status)}</span></div>
        <div class="date-cell">${escapeHTML(formatShortDate(item.startDate))} → ${escapeHTML(formatShortDate(item.endDate))}</div>
        <div class="prog-cell">
          <span class="progress-pill">${clampPct(item.progressPct)}%</span>
          <div class="prog-flags">${dependencySignalLabel}</div>
        </div>
        <div class="timeline-wrap">
          <div class="timeline-track">
            ${tickPercents.map(percent => `<span class="timeline-grid-line" style="left:${percent}%"></span>`).join('')}
            <span class="timeline-today-line" style="left:${todayPct}%"></span>
            ${dependencyLinkHTML}
            ${item.isMilestone
              ? `
                <div class="task-bar milestone-bar" style="left:${leftPct}%;width:${widthPct}%;background:${color}"></div>
                <span class="task-edge start" style="left:${leftPct}%;background:${color}"></span>
                <span class="task-edge end" style="left:${rightPct}%;background:${color}"></span>
              `
              : `<div class="task-bar" style="left:${leftPct}%;width:${widthPct}%;background:${color}"></div>`}
            ${doneMarker}
            ${majorMarker}
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
        ${dependencyChains.length > 0 ? `
          <div class="dependency-overview">
            <div class="dependency-overview-title">Dependency Chains</div>
            ${dependencyChains.slice(0, 8).map(dep => `<div class="dependency-item">${escapeHTML(dep.from)} → ${escapeHTML(dep.to)}</div>`).join('')}
          </div>
        ` : ''}
        ${sortedItems.length > 0 ? `
          <div class="timeline-head">
            <div>Roadmap Item</div>
            <div>Owner</div>
            <div>Status</div>
            <div>Dates</div>
            <div>Signals</div>
            <div>Timeline</div>
          </div>
          <div class="timeline-axis-row">
            <div class="timeline-axis-spacer"></div>
            <div class="timeline-axis">
              ${axisLabels.map(t => {
                const anchorClass = t.percent === 0 ? ' axis-label-start' : (t.percent === 100 ? ' axis-label-end' : '');
                return `
                  <span class="axis-line" style="left:${t.percent}%"></span>
                  <span class="axis-label${anchorClass}" style="left:${t.percent}%">${escapeHTML(t.label)}</span>
                `;
              }).join('')}
              <span class="today-axis-line" style="left:${todayPct}%"></span>
              <span class="today-axis-label" style="left:${todayPct}%">Today</span>
            </div>
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

export const buildPMGantHTMLFromSelection = (
  projects: PMGantProject[],
  gantItems: PMGanttItem[]
): string => {
  const bundles = projects
    .map(project => ({
      project,
      items: gantItems
        .filter(item => item.projectId === project.id)
        .sort((a, b) => {
          const aStart = parseDate(a.startDate)?.getTime() || 0;
          const bStart = parseDate(b.startDate)?.getTime() || 0;
          if (aStart !== bStart) return aStart - bStart;
          return a.title.localeCompare(b.title);
        })
    }))
    .filter(bundle => bundle.items.length > 0)
    .sort((a, b) => {
      const teamCompare = a.project.teamName.localeCompare(b.project.teamName);
      if (teamCompare !== 0) return teamCompare;
      return a.project.name.localeCompare(b.project.name);
    });

  const narrative = buildDefaultNarrative(bundles);
  return buildPMGantDeckHTML(bundles, narrative);
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
  const [isAIBulkCreating, setIsAIBulkCreating] = useState(false);
  const [isAIGeneratingProjectMilestones, setIsAIGeneratingProjectMilestones] = useState(false);
  const [aiInput, setAiInput] = useState('');
  const [aiError, setAiError] = useState('');
  const [aiInfo, setAiInfo] = useState('');
  const [projectSearch, setProjectSearch] = useState('');
  const [editingItemId, setEditingItemId] = useState<string | null>(null);
  const didAutoSelectInitialProject = useRef(false);

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

  const ownerSuggestions = useMemo(
    () => Array.from(new Set(users.map(user => `${user.firstName} ${user.lastName}`.trim()).filter(Boolean))),
    [users]
  );

  useEffect(() => {
    setSelectedProjectIds(prev => prev.filter(id => projectById.has(id)));
  }, [projectById]);

  useEffect(() => {
    if (!didAutoSelectInitialProject.current && selectedProjectIds.length === 0 && allProjects.length > 0) {
      setSelectedProjectIds([allProjects[0].id]);
      didAutoSelectInitialProject.current = true;
    }
  }, [allProjects, selectedProjectIds.length]);

  const defaultProjectId = selectedProjectIds[0] || '';
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
    isMajorDelivery: false,
    dependsOnIds: [],
    notes: '',
  });

  useEffect(() => {
    setForm(prev => {
      if (selectedProjectIds.length === 0) {
        if (!prev.projectId && prev.dependsOnIds.length === 0) return prev;
        return { ...prev, projectId: '', dependsOnIds: [] };
      }

      if (selectedProjectIds.length === 1) {
        const onlyProjectId = selectedProjectIds[0];
        if (prev.projectId === onlyProjectId) return prev;
        return { ...prev, projectId: onlyProjectId, dependsOnIds: [] };
      }

      if (prev.projectId && selectedProjectIds.includes(prev.projectId)) return prev;
      return { ...prev, projectId: selectedProjectIds[0], dependsOnIds: [] };
    });
  }, [selectedProjectIds]);

  const selectedProjects = useMemo(() => {
    return allProjects.filter(project => selectedProjectIds.includes(project.id));
  }, [allProjects, selectedProjectIds]);

  const filteredProjects = useMemo(() => {
    const query = projectSearch.trim().toLowerCase();
    if (!query) return allProjects;
    return allProjects.filter(project => {
      const name = project.name.toLowerCase();
      const team = (project.teamName || '').toLowerCase();
      const status = (project.status || '').toLowerCase();
      return name.includes(query) || team.includes(query) || status.includes(query);
    });
  }, [allProjects, projectSearch]);

  const canUseForm = selectedProjects.length > 0;
  const hasSingleSelectedProject = selectedProjects.length === 1;
  const activeFormProjectId = form.projectId || (hasSingleSelectedProject ? selectedProjects[0].id : '');

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

  const dependencyOptions = useMemo(() => {
    return gantItems
      .filter(item => item.projectId === activeFormProjectId && item.id !== editingItemId)
      .sort((a, b) => {
        const aStart = parseDate(a.startDate)?.getTime() || 0;
        const bStart = parseDate(b.startDate)?.getTime() || 0;
        if (aStart !== bStart) return aStart - bStart;
        return a.title.localeCompare(b.title);
      });
  }, [gantItems, activeFormProjectId, editingItemId]);

  useEffect(() => {
    setForm(prev => {
      const validIds = prev.dependsOnIds.filter(depId => dependencyOptions.some(option => option.id === depId));
      if (validIds.length === prev.dependsOnIds.length) return prev;
      return { ...prev, dependsOnIds: validIds };
    });
  }, [dependencyOptions]);

  const gantItemById = useMemo(() => {
    const map = new Map<string, PMGanttItem>();
    gantItems.forEach(item => map.set(item.id, item));
    return map;
  }, [gantItems]);

  const saveDraftItems = (
    projectId: string,
    drafts: Array<{
      title: string;
      description?: string;
      owner?: string;
      startDate?: string;
      endDate?: string;
      status?: PMGanttStatus | '';
      priority?: PMGanttPriority | '';
      progressPct?: number | null;
      notes?: string;
      isMilestone?: boolean | null;
    }>
  ): number => {
    const project = projectById.get(projectId);
    if (!project) return 0;

    const existingTitles = new Set(
      gantItems
        .filter(item => item.projectId === projectId)
        .map(item => normalizeTitle(item.title))
    );

    let created = 0;
    drafts.forEach(draft => {
      const title = (draft.title || '').trim();
      if (!title) return;
      const normalized = normalizeTitle(title);
      if (existingTitles.has(normalized)) return;

      const startDate = draft.startDate && parseDate(draft.startDate) ? draft.startDate : todayISO();
      const proposedEnd = draft.endDate && parseDate(draft.endDate) ? draft.endDate : plusDaysISO(14);
      const start = parseDate(startDate) || new Date(`${todayISO()}T00:00:00`);
      const end = parseDate(proposedEnd) || addDays(start, 14);
      const endDate = end.getTime() < start.getTime() ? toISODate(addDays(start, 14)) : proposedEnd;

      const now = new Date().toISOString();
      const nextItem: PMGanttItem = {
        id: generateId(),
        projectId,
        createdByUserId: currentUser.id,
        createdAt: now,
        updatedAt: now,
        title,
        description: (draft.description || '').trim(),
        owner: (draft.owner || '').trim() || project.owner || `${currentUser.firstName} ${currentUser.lastName}`.trim(),
        startDate,
        endDate,
        progressPct: clampPct(draft.progressPct == null ? 0 : draft.progressPct),
        status: (draft.status && STATUS_VALUES.includes(draft.status as PMGanttStatus)) ? (draft.status as PMGanttStatus) : 'Planned',
        priority: (draft.priority && PRIORITY_VALUES.includes(draft.priority as PMGanttPriority)) ? (draft.priority as PMGanttPriority) : 'Medium',
        isMilestone: draft.isMilestone !== false,
        isMajorDelivery: (draft.isMilestone !== false) && draft.priority === 'Critical',
        dependsOnIds: [],
        notes: (draft.notes || '').trim(),
      };
      onSaveItem(nextItem);
      existingTitles.add(normalized);
      created += 1;
    });

    return created;
  };

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
      isMajorDelivery: false,
      dependsOnIds: [],
      notes: '',
    });
    setAiInput('');
    setAiError('');
    setAiInfo('');
  };

  const toggleProject = (projectId: string) => {
    setSelectedProjectIds(prev =>
      prev.includes(projectId) ? prev.filter(id => id !== projectId) : [...prev, projectId]
    );
  };

  const handleSelectAll = () => {
    if (filteredProjects.length === 0) return;
    const filteredIds = new Set(filteredProjects.map(project => project.id));
    setSelectedProjectIds(prev => Array.from(new Set([...prev, ...filteredIds])));
  };
  const handleClearSelection = () => setSelectedProjectIds([]);

  const toggleDependency = (dependencyId: string) => {
    setForm(prev => ({
      ...prev,
      dependsOnIds: prev.dependsOnIds.includes(dependencyId)
        ? prev.dependsOnIds.filter(id => id !== dependencyId)
        : [...prev.dependsOnIds, dependencyId],
    }));
  };

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
      isMajorDelivery: item.isMajorDelivery === true,
      dependsOnIds: Array.isArray(item.dependsOnIds) ? item.dependsOnIds : [],
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
    const projectId = activeFormProjectId;
    if (!projectId) return alert('Please select at least one project in Step 1.');
    if (!selectedProjectIds.includes(projectId)) return alert('Selected form project is outside Step 1 selection.');
    if (!form.title.trim()) return alert('Roadmap item title is required.');
    if (!form.startDate || !form.endDate) return alert('Start and end dates are required.');

    const start = parseDate(form.startDate);
    const end = parseDate(form.endDate);
    if (!start || !end) return alert('Invalid dates.');
    if (end.getTime() < start.getTime()) return alert('End date must be on or after start date.');

    const sanitizedDependsOnIds = Array.from(new Set(
      form.dependsOnIds.filter(depId =>
        depId !== editingItemId &&
        gantItems.some(item => item.projectId === projectId && item.id === depId)
      )
    ));

    const now = new Date().toISOString();
    const existing = editingItemId ? gantItems.find(item => item.id === editingItemId) : null;
    const nextItem: PMGanttItem = {
      id: editingItemId || generateId(),
      projectId,
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
      isMajorDelivery: form.isMajorDelivery,
      dependsOnIds: sanitizedDependsOnIds,
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
    setAiInfo('');
    setIsAIAutofilling(true);
    try {
      const extracted = await extractPMGanttItemFromText(aiInput, llmConfig);
      setForm(prev => ({
        ...prev,
        projectId: prev.projectId || selectedProjectIds[0] || '',
        title: extracted.title || prev.title,
        description: extracted.description || prev.description,
        owner: extracted.owner || prev.owner,
        startDate: extracted.startDate || prev.startDate,
        endDate: extracted.endDate || prev.endDate,
        ...(() => {
          const nextStatus = extracted.status || prev.status;
          const nextPriority = extracted.priority || prev.priority;
          const nextIsMilestone = extracted.isMilestone != null ? extracted.isMilestone : prev.isMilestone;
          return {
            status: nextStatus,
            priority: nextPriority,
            isMilestone: nextIsMilestone,
            isMajorDelivery: nextIsMilestone
              ? (prev.isMajorDelivery || nextPriority === 'Critical')
              : false,
          };
        })(),
        progressPct: extracted.progressPct != null ? clampPct(extracted.progressPct) : prev.progressPct,
        notes: extracted.notes || prev.notes,
      }));
    } catch (e: any) {
      setAiError(e?.message || 'AI autofill failed.');
    } finally {
      setIsAIAutofilling(false);
    }
  };

  const handleAICreateMilestonesFromText = async () => {
    const projectId = activeFormProjectId;
    if (!projectId) {
      setAiError('Select a project before creating milestones from text.');
      return;
    }
    if (!aiInput.trim()) {
      setAiError('Paste text containing milestones first.');
      return;
    }

    setAiError('');
    setAiInfo('');
    setIsAIBulkCreating(true);
    try {
      const drafts = await extractPMGanttItemsFromText(aiInput, llmConfig);
      if (drafts.length === 0) {
        setAiError('AI did not find explicit milestone information. You can continue manually in the form.');
        return;
      }
      const created = saveDraftItems(
        projectId,
        drafts.map(draft => ({
          ...draft,
          isMilestone: draft.isMilestone == null ? true : draft.isMilestone,
        }))
      );

      if (created === 0) {
        setAiError('No new milestones were created (duplicates or empty extraction).');
        return;
      }
      setAiInfo(`${created} milestone${created > 1 ? 's' : ''} created from pasted text. You can edit them anytime below.`);
      if (!selectedProjectIds.includes(projectId)) {
        setSelectedProjectIds(prev => [...prev, projectId]);
      }
    } catch (e: any) {
      setAiError(e?.message || 'Bulk milestone creation failed.');
    } finally {
      setIsAIBulkCreating(false);
    }
  };

  const handleAIGenerateFromProjectData = async () => {
    const projectId = activeFormProjectId;
    if (!projectId) {
      setAiError('Select a project first.');
      return;
    }
    const project = projectById.get(projectId);
    if (!project) {
      setAiError('Selected project is not available.');
      return;
    }

    setAiError('');
    setAiInfo('');
    setIsAIGeneratingProjectMilestones(true);
    try {
      const existingForProject = gantItems.filter(item => item.projectId === projectId);
      const drafts = await generatePMGanttMilestonesFromProject(project, existingForProject, llmConfig);
      if (drafts.length === 0) {
        setAiError('Not enough explicit project information to generate milestones automatically. You can fill the form manually or paste text.');
        return;
      }

      const created = saveDraftItems(
        projectId,
        drafts.map(draft => ({
          ...draft,
          isMilestone: true,
        }))
      );
      if (created === 0) {
        setAiError('AI suggestions were duplicates of existing milestones.');
        return;
      }
      setAiInfo(`${created} milestone${created > 1 ? 's' : ''} generated from project data. Edit them below if needed.`);
      if (!selectedProjectIds.includes(projectId)) {
        setSelectedProjectIds(prev => [...prev, projectId]);
      }
    } catch (e: any) {
      setAiError(e?.message || 'AI milestone generation from project data failed.');
    } finally {
      setIsAIGeneratingProjectMilestones(false);
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
          <div className="relative mb-3">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={projectSearch}
              onChange={e => setProjectSearch(e.target.value)}
              placeholder="Search project or team..."
              className="w-full pl-10 pr-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/30"
            />
          </div>
          <div className="space-y-2 max-h-[360px] overflow-y-auto pr-1">
            {filteredProjects.map(project => {
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
            {allProjects.length > 0 && filteredProjects.length === 0 && (
              <p className="text-sm text-gray-400 dark:text-gray-500 py-8 text-center">No project matches your search.</p>
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
              {!canUseForm ? (
                <div className="w-full px-3 py-2 text-xs border border-dashed border-amber-300 dark:border-amber-700 rounded-lg bg-amber-50 dark:bg-amber-900/20 text-amber-700 dark:text-amber-300">
                  Select at least one project in Step 1 to enable this form.
                </div>
              ) : hasSingleSelectedProject ? (
                <div className="w-full px-3 py-2.5 text-sm border border-emerald-200 dark:border-emerald-700 rounded-lg bg-emerald-50 dark:bg-emerald-900/20 text-emerald-800 dark:text-emerald-300 font-semibold">
                  {selectedProjects[0].teamName} • {selectedProjects[0].name}
                  <span className="ml-2 text-[11px] font-bold uppercase tracking-wide text-emerald-600 dark:text-emerald-400">Auto</span>
                </div>
              ) : (
                <select
                  value={activeFormProjectId}
                  onChange={e => setForm(prev => ({ ...prev, projectId: e.target.value, dependsOnIds: [] }))}
                  className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                >
                  <option value="">Select project...</option>
                  {selectedProjects.map(project => (
                    <option key={project.id} value={project.id}>{project.teamName} • {project.name}</option>
                  ))}
                </select>
              )}
            </div>
            <div className="flex items-end">
              <button
                onClick={handleAIGenerateFromProjectData}
                disabled={isAIGeneratingProjectMilestones || !activeFormProjectId}
                className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 text-sm font-bold text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 disabled:opacity-50 transition-colors"
              >
                <Sparkles className={`w-4 h-4 ${isAIGeneratingProjectMilestones ? 'animate-spin' : ''}`} />
                AI Create Milestones (Project Data)
              </button>
            </div>
            <div className="md:col-span-2 -mt-1">
              <p className="text-[11px] text-gray-500 dark:text-gray-400">
                If project data is too limited, AI will not invent milestones. In that case, add milestones manually or paste text in the AI assistant below.
              </p>
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Owner</label>
              <input
                list="pm-gant-owner-list"
                value={form.owner}
                onChange={e => setForm(prev => ({ ...prev, owner: e.target.value }))}
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                placeholder="Owner name"
              />
              <datalist id="pm-gant-owner-list">
                {ownerSuggestions.map(owner => (
                  <option key={owner} value={owner} />
                ))}
              </datalist>
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
                  onChange={e => setForm(prev => ({
                    ...prev,
                    isMilestone: e.target.checked,
                    isMajorDelivery: e.target.checked ? prev.isMajorDelivery : false
                  }))}
                  className="w-4 h-4 rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                />
                Milestone item
              </label>
            </div>
            <div className="flex items-end">
              <label className={`inline-flex items-center gap-2 text-sm font-medium ${form.isMilestone ? 'text-gray-700 dark:text-gray-300' : 'text-gray-400 dark:text-gray-500'} cursor-pointer`}>
                <input
                  type="checkbox"
                  checked={form.isMajorDelivery}
                  disabled={!form.isMilestone}
                  onChange={e => setForm(prev => ({ ...prev, isMajorDelivery: e.target.checked }))}
                  className="w-4 h-4 rounded border-gray-300 text-amber-500 focus:ring-amber-500 disabled:opacity-50"
                />
                Major delivery (gold star)
              </label>
            </div>
            <div className="md:col-span-2">
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Dependencies</label>
              {dependencyOptions.length === 0 ? (
                <div className="px-3 py-2 text-xs text-gray-500 dark:text-gray-400 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg bg-gray-50 dark:bg-gray-800/40">
                  No dependency candidate yet for this project. Create at least one other roadmap item first.
                </div>
              ) : (
                <div className="max-h-28 overflow-y-auto border border-gray-200 dark:border-gray-700 rounded-lg p-2 bg-gray-50 dark:bg-gray-800/40 space-y-1">
                  {dependencyOptions.map(option => {
                    const checked = form.dependsOnIds.includes(option.id);
                    return (
                      <label key={option.id} className="flex items-center gap-2 text-xs text-gray-700 dark:text-gray-300 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleDependency(option.id)}
                          className="w-3.5 h-3.5 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                        />
                        <span className="truncate">{option.title}</span>
                        <span className="text-gray-400 dark:text-gray-500">({formatShortDate(option.startDate)} → {formatShortDate(option.endDate)})</span>
                      </label>
                    );
                  })}
                </div>
              )}
              <p className="mt-1 text-[11px] text-gray-500 dark:text-gray-400">
                Dependencies will be shown in the Gant export to visualize sequence and critical chaining.
              </p>
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
              disabled={!activeFormProjectId}
              className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-white bg-indigo-600 rounded-lg hover:bg-indigo-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
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
                <Wand2 className="w-4 h-4" /> AI Milestone Assistant
              </h4>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleAIAutofill}
                  disabled={isAIAutofilling || !aiInput.trim()}
                  className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:opacity-50 transition-colors"
                >
                  <Sparkles className={`w-3.5 h-3.5 ${isAIAutofilling ? 'animate-spin' : ''}`} />
                  Autofill Form
                </button>
                <button
                  onClick={handleAICreateMilestonesFromText}
                  disabled={isAIBulkCreating || !aiInput.trim()}
                  className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-white bg-teal-600 rounded-md hover:bg-teal-700 disabled:opacity-50 transition-colors"
                >
                  <Sparkles className={`w-3.5 h-3.5 ${isAIBulkCreating ? 'animate-spin' : ''}`} />
                  Create Milestones
                </button>
              </div>
            </div>
            <p className="text-xs text-indigo-700 dark:text-indigo-300 mb-2">
              Paste rough notes, an email excerpt, or a text containing one or multiple milestones. Use "Autofill Form" for one draft, or "Create Milestones" for bulk creation.
            </p>
            {aiError && (
              <p className="text-xs text-red-600 dark:text-red-400 mb-2">{aiError}</p>
            )}
            {aiInfo && (
              <p className="text-xs text-emerald-700 dark:text-emerald-300 mb-2">{aiInfo}</p>
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
              const dependencyTitles = (item.dependsOnIds || [])
                .map(depId => gantItemById.get(depId)?.title)
                .filter((title): title is string => Boolean(title));
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
                    {isMajorDeliveryItem(item) && <span className="px-2 py-0.5 rounded-full bg-amber-100 dark:bg-amber-900/20 text-amber-700 dark:text-amber-300 border border-amber-200 dark:border-amber-700 font-bold">Major Delivery ★</span>}
                    {dependencyTitles.length > 0 && <span className="px-2 py-0.5 rounded-full bg-blue-100 dark:bg-blue-900/20 text-blue-700 dark:text-blue-300 border border-blue-200 dark:border-blue-700 font-bold">Depends on {dependencyTitles.length}</span>}
                    <span className="text-gray-500 dark:text-gray-400">{formatShortDate(item.startDate)} → {formatShortDate(item.endDate)}</span>
                  </div>
                  {item.description && <p className="mt-2 text-xs text-gray-600 dark:text-gray-300">{item.description}</p>}
                  {dependencyTitles.length > 0 && (
                    <p className="mt-1 text-[11px] text-blue-700 dark:text-blue-300">
                      <span className="font-semibold">Dependencies:</span> {dependencyTitles.join(' • ')}
                    </p>
                  )}
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
