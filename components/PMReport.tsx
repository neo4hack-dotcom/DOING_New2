
import React, { useState, useMemo, useRef, useEffect } from 'react';
import {
  FileBarChart, Plus, Trash2, Save, ChevronDown, AlertTriangle,
  CheckCircle2, TrendingUp, Shield, Megaphone, Target, Calendar,
  DollarSign, Users, Sparkles, Download, Edit3, Eye,
  CircleDot, Info, ChevronUp, Activity, Copy, History, Tag,
  Bot, Paperclip, Loader2, X, Send, Mail
} from 'lucide-react';
import {
  Team, User, Project, Task, TaskStatus, TaskPriority, LLMConfig, PMReportData, PMReportCostSplit, PMReportConfidentiality,
  RAGStatus
} from '../types';
import { generateId } from '../services/storage';
import { extractPMReportFromText } from '../services/llmService';

// ─── Props ───
interface PMReportProps {
  teams: Team[];
  users: User[];
  currentUser: User;
  llmConfig: LLMConfig;
  pmReportData: PMReportData[];
  onSavePMReport: (data: PMReportData) => void;
  onDeletePMReport: (id: string) => void;
}

type PMView = 'overview' | 'data-entry' | 'report-preview';
type PMReportExtractionDraft = Awaited<ReturnType<typeof extractPMReportFromText>>;

// ─── RAG helpers ───
const RAGDot: React.FC<{ status: RAGStatus; size?: 'sm' | 'md' | 'lg' }> = ({ status, size = 'md' }) => {
  const s = { sm: 'w-3 h-3', md: 'w-4 h-4', lg: 'w-5 h-5' }[size];
  const c = { Green: 'bg-emerald-500 shadow-emerald-500/50', Amber: 'bg-amber-500 shadow-amber-500/50', Red: 'bg-red-500 shadow-red-500/50' }[status];
  return <div className={`${s} rounded-full ${c} shadow-md`} />;
};

const RAGSelector: React.FC<{ value: RAGStatus; onChange: (v: RAGStatus) => void; label?: string }> = ({ value, onChange, label }) => (
  <div className="flex items-center gap-2">
    {label && <span className="text-xs font-medium text-gray-500 dark:text-gray-400 min-w-[70px]">{label}</span>}
    {(['Green', 'Amber', 'Red'] as RAGStatus[]).map(s => (
      <button key={s} onClick={() => onChange(s)}
        className={`px-3 py-1 rounded-md text-xs font-bold border transition-all ${
          value === s
            ? s === 'Green' ? 'bg-emerald-500 text-white border-emerald-600 shadow-sm' :
              s === 'Amber' ? 'bg-amber-500 text-white border-amber-600 shadow-sm' :
              'bg-red-500 text-white border-red-600 shadow-sm'
            : 'bg-white dark:bg-gray-800 text-gray-500 border-gray-200 dark:border-gray-700 hover:border-gray-400'
        }`}>{s}</button>
    ))}
  </div>
);

const createEmptyReport = (projectId: string, userId: string, version: number = 1): PMReportData => ({
  id: generateId(), projectId, userId,
  createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(),
  version, versionLabel: `v${version} — ${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}`,
  confidentialityLevel: 'Confidential',
  overallStatus: 'Green', scopeStatus: 'Green', scheduleStatus: 'Green', budgetStatus: 'Green', resourceStatus: 'Green',
  executiveSummary: '', keyDecisions: '', nextSteps: '',
  incidents: [], updates: [], news: [], milestones: [], risks: [],
  budgetAllocated: 0, budgetSpent: 0, budgetForecast: 0,
  costDistribution: [],
  overallCompletionPct: 0,
});

const cloneReport = (src: PMReportData, newVersion: number): PMReportData => ({
  ...JSON.parse(JSON.stringify(src)),
  id: generateId(),
  version: newVersion,
  versionLabel: `v${newVersion} — ${new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })}`,
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
});

const formatMD = (value: number) => `${value.toLocaleString(undefined, { maximumFractionDigits: 1 })} MD`;

const CONFIDENTIALITY_LEVELS: PMReportConfidentiality[] = [
  'Public',
  'Internal',
  'Confidential',
  'Strictly Confidential',
];

const CONFIDENTIALITY_RANK: Record<PMReportConfidentiality, number> = {
  Public: 0,
  Internal: 1,
  Confidential: 2,
  'Strictly Confidential': 3,
};

const getReportConfidentiality = (report: PMReportData): PMReportConfidentiality => {
  return report.confidentialityLevel || 'Confidential';
};

const getIncidentRAGStatus = (report: PMReportData): RAGStatus => {
  if (report.incidents.length === 0) return 'Green';

  const hasCriticalOpen = report.incidents.some(
    inc => inc.severity === 'Critical' && inc.status !== 'Resolved'
  );
  if (hasCriticalOpen) return 'Red';

  const hasMajorOpen = report.incidents.some(
    inc => inc.severity === 'Major' && inc.status !== 'Resolved'
  );
  if (hasMajorOpen) return 'Red';

  const hasOpenIncident = report.incidents.some(inc => inc.status !== 'Resolved');
  if (hasOpenIncident) return 'Amber';

  return 'Green';
};

interface MergeExtractedOptions {
  preserveExisting?: boolean;
}

const isMissingText = (value?: string | null): boolean => !value || !value.trim();
const isMissingNumeric = (value?: number | null): boolean => value == null || !Number.isFinite(value) || value === 0;
const normalizeKeyPart = (value: string | undefined | null) => (value || '').trim().toLowerCase();

const mergeUniqueByKey = <T,>(
  existing: T[],
  incoming: T[],
  keyFn: (item: T) => string
): T[] => {
  const merged = [...existing];
  const seen = new Set(existing.map(item => keyFn(item)).filter(Boolean));

  incoming.forEach(item => {
    const key = keyFn(item);
    if (!key || !seen.has(key)) {
      merged.push(item);
      if (key) seen.add(key);
    }
  });

  return merged;
};

const taskPriorityRank = (priority: TaskPriority) => {
  if (priority === TaskPriority.URGENT) return 4;
  if (priority === TaskPriority.HIGH) return 3;
  if (priority === TaskPriority.MEDIUM) return 2;
  return 1;
};

const buildProjectTaskPrefillDraft = (
  project: Project,
  usersById: Record<string, User>
): PMReportExtractionDraft => {
  const today = new Date().toISOString().split('T')[0];
  const tasks = project.tasks || [];
  const totalTasks = tasks.length;
  const doneTasks = tasks.filter(task => task.status === TaskStatus.DONE);
  const blockedTasks = tasks.filter(task => task.status === TaskStatus.BLOCKED);
  const overdueTasks = tasks.filter(task =>
    task.status !== TaskStatus.DONE &&
    Boolean(task.eta) &&
    task.eta < today
  );
  const unassignedTasks = tasks.filter(task => !task.assigneeId);
  const completionPct = totalTasks > 0
    ? Math.round((doneTasks.length / totalTasks) * 100)
    : (project.status === 'Done' ? 100 : 0);

  const hasCriticalBlocked = blockedTasks.some(task =>
    task.priority === TaskPriority.URGENT || task.priority === TaskPriority.HIGH
  );
  const hasRedExternalDependency = (project.externalDependencies || []).some(dep => dep.status === 'Red');

  let overallStatus: PMReportExtractionDraft['overallStatus'] = 'Green';
  if (project.status === 'Paused' || blockedTasks.length > 0 || overdueTasks.length > 0) overallStatus = 'Amber';
  if (hasCriticalBlocked || hasRedExternalDependency || overdueTasks.length >= 3) overallStatus = 'Red';

  let scopeStatus: PMReportExtractionDraft['scopeStatus'] = 'Green';
  if (blockedTasks.length > 0 || hasRedExternalDependency) scopeStatus = 'Amber';
  if (hasCriticalBlocked || hasRedExternalDependency) scopeStatus = 'Red';

  let scheduleStatus: PMReportExtractionDraft['scheduleStatus'] = 'Green';
  if (project.status === 'Paused' || overdueTasks.length > 0) scheduleStatus = 'Amber';
  if (overdueTasks.length >= 3) scheduleStatus = 'Red';

  const hasAnyTaskCost = tasks.some(task => typeof task.cost === 'number' && Number.isFinite(task.cost));
  const totalTaskCost = tasks.reduce((sum, task) => sum + (typeof task.cost === 'number' ? task.cost : 0), 0);
  const doneTaskCost = doneTasks.reduce((sum, task) => sum + (typeof task.cost === 'number' ? task.cost : 0), 0);
  const budgetAllocated = typeof project.cost === 'number' && Number.isFinite(project.cost)
    ? project.cost
    : (hasAnyTaskCost ? totalTaskCost : null);
  const budgetSpent = hasAnyTaskCost ? doneTaskCost : null;
  const budgetForecast = hasAnyTaskCost
    ? Math.max(totalTaskCost, budgetAllocated || 0)
    : (budgetAllocated != null ? budgetAllocated : null);

  let budgetStatus: PMReportExtractionDraft['budgetStatus'] = '';
  if (budgetAllocated != null && budgetAllocated > 0 && budgetForecast != null) {
    if (budgetForecast > budgetAllocated * 1.1) budgetStatus = 'Red';
    else if (budgetForecast > budgetAllocated * 0.95) budgetStatus = 'Amber';
    else budgetStatus = 'Green';
  }

  let resourceStatus: PMReportExtractionDraft['resourceStatus'] = 'Green';
  if (unassignedTasks.length > 0) resourceStatus = 'Amber';
  if (unassignedTasks.length >= 3 || (totalTasks > 0 && (unassignedTasks.length / totalTasks) >= 0.5)) resourceStatus = 'Red';

  const summaryBits = [
    `"${project.name}" is currently ${project.status.toLowerCase()}.`,
    totalTasks > 0
      ? `${completionPct}% completion (${doneTasks.length}/${totalTasks} tasks done${blockedTasks.length > 0 ? `, ${blockedTasks.length} blocked` : ''}${overdueTasks.length > 0 ? `, ${overdueTasks.length} overdue` : ''}).`
      : 'No project tasks are available yet.',
    project.description ? project.description.trim() : ''
  ].filter(Boolean);

  const keyDecisionBits: string[] = [];
  if (typeof project.tsdCreated === 'boolean') {
    keyDecisionBits.push(`TSD created: ${project.tsdCreated ? 'Yes' : 'No'}.`);
  }
  if ((project.dependencies || []).length > 0) {
    keyDecisionBits.push(`Cross-project dependencies to monitor: ${(project.dependencies || []).length}.`);
  }
  const redDeps = (project.externalDependencies || []).filter(dep => dep.status === 'Red').length;
  if (redDeps > 0) {
    keyDecisionBits.push(`External dependencies with RED status: ${redDeps}.`);
  }

  const nextStepTasks = tasks
    .filter(task => task.status !== TaskStatus.DONE)
    .sort((a, b) => {
      const prioCmp = taskPriorityRank(b.priority) - taskPriorityRank(a.priority);
      if (prioCmp !== 0) return prioCmp;
      const aEta = a.eta || '9999-12-31';
      const bEta = b.eta || '9999-12-31';
      return aEta.localeCompare(bEta);
    })
    .slice(0, 5);

  const updates = tasks
    .filter(task => task.status !== TaskStatus.TODO)
    .sort((a, b) => {
      const aEta = a.eta || '9999-12-31';
      const bEta = b.eta || '9999-12-31';
      return aEta.localeCompare(bEta);
    })
    .slice(0, 8)
    .map(task => {
      const isBlocked = task.status === TaskStatus.BLOCKED;
      const isDone = task.status === TaskStatus.DONE;
      const isOverdue = task.status !== TaskStatus.DONE && task.eta && task.eta < today;
      return {
        date: task.eta || today,
        category: isBlocked || isOverdue ? 'Timeline' : isDone ? 'Scope' : 'Other',
        title: task.title || '',
        description: task.description || '',
        impact: isBlocked || isOverdue ? 'Amber' : 'Green',
      };
    });

  const news = doneTasks
    .slice(0, 6)
    .map(task => ({
      date: task.eta || today,
      title: task.title || '',
      description: task.description || '',
      type: 'Achievement' as const,
    }));

  const milestones = tasks
    .filter(task => Boolean(task.eta))
    .sort((a, b) => (a.eta || '').localeCompare(b.eta || ''))
    .slice(0, 10)
    .map(task => ({
      name: task.title || '',
      plannedDate: task.eta || '',
      revisedDate: '',
      status: task.status === TaskStatus.BLOCKED
        ? 'Red'
        : task.status === TaskStatus.ONGOING
          ? 'Amber'
          : 'Green',
      completionPct: task.status === TaskStatus.DONE
        ? 100
        : task.status === TaskStatus.ONGOING
          ? 50
          : task.status === TaskStatus.BLOCKED
            ? 25
            : 0
    }));

  const getTaskOwner = (task: Task) => {
    const user = usersById[task.assigneeId || ''];
    return user ? `${user.firstName} ${user.lastName}`.trim() : '';
  };

  const riskFromTasks = blockedTasks.slice(0, 8).map(task => ({
    description: task.title || 'Blocked task',
    likelihood: task.priority === TaskPriority.URGENT ? 'High' : 'Medium',
    impact: task.priority === TaskPriority.URGENT || task.priority === TaskPriority.HIGH ? 'High' : 'Medium',
    mitigation: task.description || 'Follow-up with owner and unblock dependencies.',
    owner: getTaskOwner(task)
  }));
  const riskFromExternalDeps = (project.externalDependencies || [])
    .filter(dep => dep.status !== 'Green')
    .slice(0, 5)
    .map(dep => ({
      description: `External dependency: ${dep.label}`,
      likelihood: dep.status === 'Red' ? 'High' as const : 'Medium' as const,
      impact: dep.status === 'Red' ? 'High' as const : 'Medium' as const,
      mitigation: 'Align owners and update dependency plan.',
      owner: '',
    }));

  return {
    overallStatus,
    scopeStatus,
    scheduleStatus,
    budgetStatus,
    resourceStatus,
    overallCompletionPct: completionPct,
    executiveSummary: summaryBits.join(' '),
    keyDecisions: keyDecisionBits.join(' '),
    nextSteps: nextStepTasks.map(task => {
      const etaSuffix = task.eta ? ` (ETA ${task.eta})` : '';
      return `${task.title}${etaSuffix}`;
    }).join('\n'),
    budgetAllocated,
    budgetSpent,
    budgetForecast,
    incidents: blockedTasks.slice(0, 8).map(task => ({
      date: task.eta || today,
      title: task.title || '',
      description: task.description || '',
      severity: task.priority === TaskPriority.URGENT ? 'Critical' : 'Major',
      status: 'Open',
    })),
    updates,
    news,
    milestones,
    risks: [...riskFromTasks, ...riskFromExternalDeps],
  };
};

const mergeExtractedDataIntoReport = (
  report: PMReportData,
  extracted: PMReportExtractionDraft,
  options: MergeExtractedOptions = {}
): PMReportData => {
  const { preserveExisting = false } = options;
  const isRAG = (value: string): value is RAGStatus => value === 'Green' || value === 'Amber' || value === 'Red';
  const clampPct = (value: number) => Math.max(0, Math.min(100, Math.round(value)));
  const next: PMReportData = { ...report };
  const looksUntouched =
    isMissingText(report.executiveSummary) &&
    isMissingText(report.keyDecisions) &&
    isMissingText(report.nextSteps) &&
    report.incidents.length === 0 &&
    report.updates.length === 0 &&
    report.news.length === 0 &&
    report.milestones.length === 0 &&
    report.risks.length === 0 &&
    (report.overallCompletionPct || 0) === 0 &&
    (report.budgetAllocated || 0) === 0 &&
    (report.budgetSpent || 0) === 0 &&
    (report.budgetForecast || 0) === 0;

  const canUpdateRAG = () => !preserveExisting || looksUntouched;
  const shouldSetText = (value: string) => !preserveExisting || isMissingText(value);
  const shouldSetNumber = (value?: number | null) => !preserveExisting || isMissingNumeric(value);

  if (canUpdateRAG() && isRAG(extracted.overallStatus)) next.overallStatus = extracted.overallStatus;
  if (canUpdateRAG() && isRAG(extracted.scopeStatus)) next.scopeStatus = extracted.scopeStatus;
  if (canUpdateRAG() && isRAG(extracted.scheduleStatus)) next.scheduleStatus = extracted.scheduleStatus;
  if (canUpdateRAG() && isRAG(extracted.budgetStatus)) next.budgetStatus = extracted.budgetStatus;
  if (canUpdateRAG() && isRAG(extracted.resourceStatus)) next.resourceStatus = extracted.resourceStatus;

  if (extracted.overallCompletionPct != null && shouldSetNumber(next.overallCompletionPct)) {
    next.overallCompletionPct = clampPct(extracted.overallCompletionPct);
  }

  if (extracted.executiveSummary.trim() && shouldSetText(next.executiveSummary)) next.executiveSummary = extracted.executiveSummary.trim();
  if (extracted.keyDecisions.trim() && shouldSetText(next.keyDecisions)) next.keyDecisions = extracted.keyDecisions.trim();
  if (extracted.nextSteps.trim() && shouldSetText(next.nextSteps)) next.nextSteps = extracted.nextSteps.trim();

  if (extracted.budgetAllocated != null && shouldSetNumber(next.budgetAllocated)) next.budgetAllocated = extracted.budgetAllocated;
  if (extracted.budgetSpent != null && shouldSetNumber(next.budgetSpent)) next.budgetSpent = extracted.budgetSpent;
  if (extracted.budgetForecast != null && shouldSetNumber(next.budgetForecast)) next.budgetForecast = extracted.budgetForecast;

  const extractedIncidents = extracted.incidents.map(inc => ({
    id: generateId(),
    date: inc.date || '',
    title: inc.title || '',
    description: inc.description || '',
    severity: inc.severity || 'Minor',
    status: inc.status || 'Open',
  }));
  if (extractedIncidents.length > 0) {
    if (!preserveExisting || next.incidents.length === 0) {
      next.incidents = extractedIncidents;
    } else {
      next.incidents = mergeUniqueByKey(next.incidents, extractedIncidents, inc =>
        [normalizeKeyPart(inc.date), normalizeKeyPart(inc.title), inc.severity, inc.status].join('|')
      );
    }
  }

  const extractedUpdates = extracted.updates.map(upd => ({
    id: generateId(),
    date: upd.date || '',
    category: upd.category || 'Other',
    title: upd.title || '',
    description: upd.description || '',
    impact: isRAG(upd.impact) ? upd.impact : 'Green',
  }));
  if (extractedUpdates.length > 0) {
    if (!preserveExisting || next.updates.length === 0) {
      next.updates = extractedUpdates;
    } else {
      next.updates = mergeUniqueByKey(next.updates, extractedUpdates, upd =>
        [normalizeKeyPart(upd.date), normalizeKeyPart(upd.title), upd.category, upd.impact].join('|')
      );
    }
  }

  const extractedNews = extracted.news.map(item => ({
    id: generateId(),
    date: item.date || '',
    title: item.title || '',
    description: item.description || '',
    type: item.type || 'Info',
  }));
  if (extractedNews.length > 0) {
    if (!preserveExisting || next.news.length === 0) {
      next.news = extractedNews;
    } else {
      next.news = mergeUniqueByKey(next.news, extractedNews, item =>
        [normalizeKeyPart(item.date), normalizeKeyPart(item.title), item.type].join('|')
      );
    }
  }

  const extractedMilestones = extracted.milestones.map(m => ({
    id: generateId(),
    name: m.name || '',
    plannedDate: m.plannedDate || '',
    revisedDate: m.revisedDate || undefined,
    status: isRAG(m.status) ? m.status : 'Green',
    completionPct: m.completionPct != null ? clampPct(m.completionPct) : 0,
  }));
  if (extractedMilestones.length > 0) {
    if (!preserveExisting || next.milestones.length === 0) {
      next.milestones = extractedMilestones;
    } else {
      next.milestones = mergeUniqueByKey(next.milestones, extractedMilestones, m =>
        [normalizeKeyPart(m.name), normalizeKeyPart(m.plannedDate), normalizeKeyPart(m.revisedDate), m.status].join('|')
      );
    }
  }

  const extractedRisks = extracted.risks.map(r => ({
    id: generateId(),
    description: r.description || '',
    likelihood: r.likelihood || 'Medium',
    impact: r.impact || 'Medium',
    mitigation: r.mitigation || '',
    owner: r.owner || '',
  }));
  if (extractedRisks.length > 0) {
    if (!preserveExisting || next.risks.length === 0) {
      next.risks = extractedRisks;
    } else {
      next.risks = mergeUniqueByKey(next.risks, extractedRisks, r =>
        [normalizeKeyPart(r.description), normalizeKeyPart(r.owner), r.likelihood, r.impact].join('|')
      );
    }
  }

  return next;
};

const mergePrefillFromProjectIntoReport = (
  report: PMReportData,
  project: Project,
  usersById: Record<string, User>
): PMReportData => {
  const extracted = buildProjectTaskPrefillDraft(project, usersById);
  return mergeExtractedDataIntoReport(report, extracted, { preserveExisting: true });
};

// ════════════════════════════════════════
//  FALLBACK HTML GENERATOR (consulting-deck style)
// ════════════════════════════════════════
const buildConsultingDeckHTML = (
  data: { project: Project & { teamName: string }; report: PMReportData }[],
  teamNameById: Record<string, string> = {}
) => {
  const rc = (s: RAGStatus) => s === 'Green' ? '#10b981' : s === 'Amber' ? '#f59e0b' : '#ef4444';
  const rcBg = (s: RAGStatus) => s === 'Green' ? '#ecfdf5' : s === 'Amber' ? '#fffbeb' : '#fef2f2';
  const now = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  const highestConfidentiality: PMReportConfidentiality = data.length
    ? data.reduce<PMReportConfidentiality>((highest, item) => {
        const current = getReportConfidentiality(item.report);
        return CONFIDENTIALITY_RANK[current] > CONFIDENTIALITY_RANK[highest] ? current : highest;
      }, 'Public')
    : 'Confidential';

  const css = `
    <style>
      .deck{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:#1e293b;max-width:1140px;margin:0 auto}
      .deck *{box-sizing:border-box}
      .deck-header{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:4px solid #00915A;padding-bottom:14px;margin-bottom:28px}
      .deck-header h1{margin:0;font-size:26px;font-weight:800;background:linear-gradient(135deg,#00915A,#00A86B);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
      .deck-header .sub{font-size:12px;color:#64748b;margin-top:4px}
      .deck-header .conf{font-size:10px;color:#94a3b8;text-transform:uppercase;letter-spacing:1.5px;font-weight:700}
      .project-block{margin-bottom:24px;page-break-inside:avoid;break-inside:avoid-page;display:flex;flex-direction:column}
      .prj-banner{background:linear-gradient(135deg,#00915A 0%,#00A86B 50%,#007A4C 100%);color:#fff;padding:18px 24px;border-radius:14px 14px 0 0;display:flex;justify-content:space-between;align-items:center}
      .prj-banner h2{margin:0;font-size:20px;font-weight:700;letter-spacing:-0.3px}
      .prj-banner .meta{font-size:11px;opacity:.8;margin-top:3px}
      .prj-banner .pct{font-size:28px;font-weight:800;letter-spacing:-1px}
      .prj-banner .pct-label{font-size:10px;opacity:.7;text-transform:uppercase;letter-spacing:1px}
      .prj-body{border:1px solid #e2e8f0;border-top:none;border-radius:0 0 14px 14px;padding:24px;background:#fff;display:flex;flex-direction:column;min-height:620px}
      .prj-main{display:flex;flex-direction:column}
      .prj-body.compact{padding:18px}
      .report-section{margin-bottom:18px;page-break-inside:avoid;break-inside:avoid-page}
      .report-section.allow-split{page-break-inside:auto;break-inside:auto}
      .report-section:last-child{margin-bottom:0}
      .cost-bottom{margin-top:auto;padding-top:14px;border-top:2px solid #e2e8f0}
      .cost-bottom .section-title{margin-bottom:12px}
      .cost-split-title{margin-top:14px;margin-bottom:8px;font-size:11px;font-weight:700;color:#334155;text-transform:uppercase;letter-spacing:.45px}
      .rag-row{display:flex;gap:10px;margin-bottom:22px;flex-wrap:wrap}
      .rag-card{flex:1;min-width:100px;text-align:center;padding:14px 8px;border-radius:10px;border:1px solid #e2e8f0;background:#fafbfc}
      .rag-circle{width:24px;height:24px;border-radius:50%;margin:0 auto 8px;box-shadow:0 0 12px rgba(0,0,0,.15)}
      .rag-label{font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.7px}
      .progress-wrap{margin-bottom:0}
      .progress-bar-outer{height:12px;background:#e2e8f0;border-radius:6px;overflow:hidden}
      .progress-bar-inner{height:100%;border-radius:6px;background:linear-gradient(90deg,#00915A,#00A86B);transition:width .3s}
      .progress-label{display:flex;justify-content:space-between;font-size:12px;font-weight:600;color:#374151;margin-bottom:6px}
      .section-title{font-size:14px;font-weight:700;color:#1e293b;margin:0 0 10px;padding-bottom:6px;border-bottom:2px solid #e2e8f0;display:flex;align-items:center;gap:8px}
      .section-title .dot{width:8px;height:8px;border-radius:50%;background:#00915A}
      .card-summary{padding:14px 18px;border-radius:10px;margin-bottom:0;font-size:12px;line-height:1.6;color:#334155}
      .card-blue{background:#e9f8f2;border-left:4px solid #00915A}
      .card-green{background:#f0fdf4;border-left:4px solid #10b981}
      .card-amber{background:#fffbeb;border-left:4px solid #f59e0b}
      .metric-row{display:flex;gap:14px;margin-bottom:0;flex-wrap:wrap}
      .metric-box{flex:1;min-width:120px;background:#f8fafc;border-radius:10px;padding:14px;border:1px solid #e2e8f0;text-align:center}
      .metric-label{font-size:9px;color:#64748b;text-transform:uppercase;font-weight:700;letter-spacing:.5px}
      .metric-value{font-size:22px;font-weight:800;margin-top:4px}
      table.deck-table{width:100%;border-collapse:separate;border-spacing:0;font-size:12px;margin-bottom:0;border-radius:8px;overflow:hidden;border:1px solid #e2e8f0}
      table.deck-table thead{display:table-header-group}
      table.deck-table th{background:#f1f5f9;padding:10px 12px;text-align:left;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#475569}
      table.deck-table td{padding:10px 12px;border-top:1px solid #f1f5f9}
      table.deck-table tr{page-break-inside:avoid;break-inside:avoid-page}
      table.deck-table tr:hover td{background:#fafbfc}
      .severity-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;color:#fff}
      .sev-critical{background:#ef4444}.sev-major{background:#f59e0b}.sev-minor{background:#64748b}
      .status-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}
      .likelihood-badge{display:inline-block;padding:2px 6px;border-radius:4px;font-size:10px;font-weight:600}
      .update-item{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid #f1f5f9}
      .update-item{page-break-inside:avoid;break-inside:avoid-page}
      .update-dot{width:8px;height:8px;border-radius:50%;margin-top:5px;flex-shrink:0}
      .cat-tag{font-size:10px;font-weight:700;padding:1px 7px;border-radius:4px;background:#e7f7f0;color:#007A4C}
      .news-type{font-size:10px;font-weight:700;padding:2px 8px;border-radius:4px}

      .prj-body.compact .rag-card{padding:10px 6px}
      .prj-body.compact .metric-box{padding:10px}
      .prj-body.compact .metric-value{font-size:18px}
      .prj-body.compact .section-title{font-size:13px;margin-bottom:8px}
      .prj-body.compact .card-summary{padding:10px 12px;font-size:11px}
      .prj-body.compact table.deck-table th{padding:7px 8px;font-size:10px}
      .prj-body.compact table.deck-table td{padding:7px 8px;font-size:11px}
      .prj-body.compact .update-item{padding:6px 0}
      .prj-body.compact .cost-bottom{padding-top:10px}

      @media print {
        .deck{max-width:none}
        .deck-header{margin-bottom:16px}
        .project-block{margin-bottom:16px;page-break-inside:avoid;break-inside:avoid-page}
        .prj-banner{padding:12px 16px}
        .prj-banner .pct{font-size:24px}
        .prj-body{padding:16px;min-height:540px}
        .report-section{margin-bottom:12px;page-break-inside:avoid;break-inside:avoid-page}
        .report-section.allow-split{page-break-inside:auto;break-inside:auto}
        .cost-bottom{margin-top:auto}
        .prj-body.compact{padding:14px;font-size:11px}
      }
    </style>`;

  const projectBlocks = data.map(({ project, report }) => {
    const confidentiality = getReportConfidentiality(report);
    const budgetAllocated = report.budgetAllocated || 0;
    const budgetSpent = report.budgetSpent || 0;
    const budgetForecast = report.budgetForecast || 0;
    const budgetPct = budgetAllocated > 0 ? Math.round((budgetSpent / (budgetAllocated || 1)) * 100) : 0;
    const incidentStatus = getIncidentRAGStatus(report);
    const ragIndicators: Array<{ label: string; status: RAGStatus }> = [
      { label: 'Overall', status: report.overallStatus },
      { label: 'Scope', status: report.scopeStatus },
      { label: 'Schedule', status: report.scheduleStatus },
      { label: 'Budget', status: report.budgetStatus },
      { label: 'Resource', status: report.resourceStatus },
      { label: 'Incident', status: incidentStatus },
    ];

    const costDistribution = (report.costDistribution || []).filter(split =>
      Boolean((split.teamName || '').trim() || (split.teamId || '').trim())
    );
    const hasCostMetrics =
      budgetAllocated > 0 ||
      budgetSpent > 0 ||
      budgetForecast > 0 ||
      costDistribution.length > 0;

    const densityScore = report.milestones.length + report.incidents.length + report.risks.length + report.updates.length + report.news.length;
    const isCompact = densityScore > 16 || report.executiveSummary.length > 420 || report.keyDecisions.length > 320 || report.nextSteps.length > 320;

    return `
    <div class="project-block">
      <div class="prj-banner">
        <div>
          <h2>${project.name}</h2>
          <div class="meta">${project.teamName} &bull; ${project.status} &bull; Deadline: ${project.deadline} &bull; ${confidentiality} &bull; v${report.version}</div>
        </div>
        <div style="text-align:right">
          <div class="pct">${report.overallCompletionPct}%</div>
          <div class="pct-label">Complete</div>
        </div>
      </div>
      <div class="prj-body ${isCompact ? 'compact' : ''}">
        <div class="prj-main">
          <section class="report-section">
            <div class="rag-row">
              ${ragIndicators.map(({ label, status }) => {
                return `<div class="rag-card" style="background:${rcBg(status)}">
                  <div class="rag-circle" style="background:${rc(status)}"></div>
                  <div class="rag-label">${label}</div>
                </div>`;
              }).join('')}
            </div>
          </section>

          <section class="report-section">
            <div class="card-summary card-blue"><strong style="color:#006B46">Executive Summary</strong><br/>${report.executiveSummary || 'No executive summary provided.'}</div>
          </section>

          ${report.incidents.length > 0 ? `
          <section class="report-section allow-split">
            <div class="section-title"><div class="dot" style="background:#ef4444"></div>Incidents</div>
            <table class="deck-table">
              <thead>
                <tr><th style="width:80px">Severity</th><th>Incident</th><th style="width:200px">Description</th><th style="width:90px">Status</th><th style="width:90px">Date</th></tr>
              </thead>
              <tbody>
                ${report.incidents.map(inc => `<tr>
                  <td><span class="severity-badge sev-${inc.severity.toLowerCase()}">${inc.severity}</span></td>
                  <td style="font-weight:600">${inc.title}</td>
                  <td style="font-size:11px;color:#64748b">${inc.description}</td>
                  <td><span class="status-badge" style="background:${inc.status === 'Resolved' ? '#d1fae5' : inc.status === 'Investigating' ? '#fef3c7' : '#fee2e2'};color:${inc.status === 'Resolved' ? '#065f46' : inc.status === 'Investigating' ? '#92400e' : '#991b1b'}">${inc.status}</span></td>
                  <td style="font-size:11px;color:#64748b">${inc.date}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </section>` : ''}

          <section class="report-section">
            <div class="progress-wrap">
              <div class="progress-label"><span>Overall Progress</span><span style="color:#00915A">${report.overallCompletionPct}%</span></div>
              <div class="progress-bar-outer"><div class="progress-bar-inner" style="width:${report.overallCompletionPct}%"></div></div>
            </div>
          </section>

          ${report.milestones.length > 0 ? `
          <section class="report-section allow-split">
            <div class="section-title"><div class="dot"></div>Milestone Tracker</div>
            <table class="deck-table">
              <thead>
                <tr><th>Milestone</th><th style="text-align:center">Planned</th><th style="text-align:center">Revised</th><th style="text-align:center">Status</th><th style="text-align:center;width:160px">Progress</th></tr>
              </thead>
              <tbody>
                ${report.milestones.map(m => `<tr>
                  <td style="font-weight:600">${m.name}</td>
                  <td style="text-align:center;font-size:11px">${m.plannedDate}</td>
                  <td style="text-align:center;font-size:11px;${m.revisedDate && m.revisedDate > m.plannedDate ? 'color:#ef4444;font-weight:700' : ''}">${m.revisedDate || '—'}</td>
                  <td style="text-align:center"><div style="width:14px;height:14px;border-radius:50%;background:${rc(m.status)};margin:0 auto;box-shadow:0 0 6px ${rc(m.status)}40"></div></td>
                  <td><div style="display:flex;align-items:center;gap:8px"><div style="flex:1;height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden"><div style="height:100%;width:${m.completionPct}%;background:linear-gradient(90deg,#00915A,#00A86B);border-radius:4px"></div></div><span style="font-size:11px;font-weight:700;color:#00915A;min-width:32px;text-align:right">${m.completionPct}%</span></div></td>
                </tr>`).join('')}
              </tbody>
            </table>
          </section>` : ''}

          ${report.risks.length > 0 ? `
          <section class="report-section allow-split">
            <div class="section-title"><div class="dot" style="background:#f59e0b"></div>Risk Register</div>
            <table class="deck-table">
              <thead>
                <tr><th>Risk</th><th style="width:85px;text-align:center">Likelihood</th><th style="width:85px;text-align:center">Impact</th><th>Mitigation</th><th style="width:90px">Owner</th></tr>
              </thead>
              <tbody>
                ${report.risks.map(r => `<tr>
                  <td style="font-weight:600">${r.description}</td>
                  <td style="text-align:center"><span class="likelihood-badge" style="background:${r.likelihood === 'High' ? '#fee2e2' : r.likelihood === 'Medium' ? '#fef3c7' : '#d1fae5'};color:${r.likelihood === 'High' ? '#991b1b' : r.likelihood === 'Medium' ? '#92400e' : '#065f46'}">${r.likelihood}</span></td>
                  <td style="text-align:center"><span class="likelihood-badge" style="background:${r.impact === 'High' ? '#fee2e2' : r.impact === 'Medium' ? '#fef3c7' : '#d1fae5'};color:${r.impact === 'High' ? '#991b1b' : r.impact === 'Medium' ? '#92400e' : '#065f46'}">${r.impact}</span></td>
                  <td style="font-size:11px;color:#475569">${r.mitigation}</td>
                  <td style="font-size:11px;font-weight:600">${r.owner}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </section>` : ''}

          ${report.updates.length > 0 ? `
          <section class="report-section allow-split">
            <div class="section-title"><div class="dot" style="background:#3b82f6"></div>Key Updates</div>
            <div>
              ${report.updates.map(u => `<div class="update-item">
                <div class="update-dot" style="background:${rc(u.impact)}"></div>
                <div><span class="cat-tag">${u.category}</span> <span style="font-size:12px;font-weight:600;margin-left:4px">${u.title}</span><div style="font-size:11px;color:#64748b;margin-top:3px">${u.description}</div></div>
              </div>`).join('')}
            </div>
          </section>` : ''}

          ${report.news.length > 0 ? `
          <section class="report-section allow-split">
            <div class="section-title"><div class="dot" style="background:#10b981"></div>News & Achievements</div>
            <div>
              ${report.news.map(n => `<div class="update-item">
                <span class="news-type" style="background:${n.type === 'Achievement' ? '#d1fae5' : n.type === 'Change' ? '#fef3c7' : '#e0e7ff'};color:${n.type === 'Achievement' ? '#065f46' : n.type === 'Change' ? '#92400e' : '#3730a3'}">${n.type}</span>
                <div><div style="font-size:12px;font-weight:600">${n.title}</div><div style="font-size:11px;color:#64748b;margin-top:2px">${n.description}</div></div>
              </div>`).join('')}
            </div>
          </section>` : ''}

          ${report.keyDecisions ? `<section class="report-section"><div class="card-summary card-amber"><strong style="color:#92400e">Key Decisions</strong><br/>${report.keyDecisions}</div></section>` : ''}
          ${report.nextSteps ? `<section class="report-section"><div class="card-summary card-green"><strong style="color:#065f46">Next Steps</strong><br/>${report.nextSteps}</div></section>` : ''}
        </div>

        ${hasCostMetrics ? `
        <section class="report-section cost-bottom">
          <div class="section-title"><div class="dot"></div>Cost Overview (MD)</div>
          <div class="metric-row">
            <div class="metric-box"><div class="metric-label">Allocated</div><div class="metric-value" style="color:#1e293b">${formatMD(budgetAllocated)}</div></div>
            <div class="metric-box"><div class="metric-label">Spent</div><div class="metric-value" style="color:#f59e0b">${formatMD(budgetSpent)}</div></div>
            <div class="metric-box"><div class="metric-label">Forecast</div><div class="metric-value" style="color:${budgetForecast > budgetAllocated ? '#ef4444' : '#10b981'}">${formatMD(budgetForecast)}</div></div>
            <div class="metric-box"><div class="metric-label">Utilization</div>
              <div class="metric-value" style="color:${budgetPct > 90 ? '#ef4444' : budgetPct > 70 ? '#f59e0b' : '#10b981'}">${budgetPct}%</div>
            </div>
          </div>

          ${costDistribution.length > 0 ? `
            <div class="cost-split-title">Cost Distribution by Team (manual)</div>
            <table class="deck-table">
              <thead>
                <tr><th>Team</th><th style="text-align:right">Allocated (MD)</th><th style="text-align:right">Spent (MD)</th><th style="text-align:right">Forecast (MD)</th></tr>
              </thead>
              <tbody>
                ${costDistribution.map(split => {
                  const teamName = (split.teamName || '').trim() || teamNameById[split.teamId || ''] || split.teamId || 'N/A';
                  return `<tr>
                    <td style="font-weight:600">${teamName}</td>
                    <td style="text-align:right">${formatMD(split.allocatedMD || 0)}</td>
                    <td style="text-align:right">${formatMD(split.spentMD || 0)}</td>
                    <td style="text-align:right">${formatMD(split.forecastMD || 0)}</td>
                  </tr>`;
                }).join('')}
              </tbody>
            </table>
          ` : ''}
        </section>` : ''}
      </div>
    </div>`;
  }).join('');

  return `${css}<div class="deck">
    <div class="deck-header">
      <div><h1>DOINg - Project Status Report</h1><div class="sub">Generated ${now} &bull; ${data.length} project${data.length > 1 ? 's' : ''}</div></div>
      <div style="text-align:right"><div class="conf">Confidentiality: ${highestConfidentiality}</div></div>
    </div>
    ${projectBlocks}
  </div>`;
};

interface AIPMReportModalProps {
  llmConfig: LLMConfig;
  onClose: () => void;
  onApply: (extracted: PMReportExtractionDraft) => void;
}

interface AIPMReportChatMessage {
  id: string;
  role: 'assistant' | 'user';
  content: string;
}

const AIPMReportModal: React.FC<AIPMReportModalProps> = ({ llmConfig, onClose, onApply }) => {
  const [messages, setMessages] = useState<AIPMReportChatMessage[]>([
    {
      id: generateId(),
      role: 'assistant',
      content: 'Paste a paragraph (email, presentation excerpt, notes) or attach a text document. I will extract PM Report fields without inventing missing information.'
    }
  ]);
  const [inputText, setInputText] = useState('');
  const [attachedFile, setAttachedFile] = useState<{ name: string; content: string } | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [extractError, setExtractError] = useState<string | null>(null);
  const [extracted, setExtracted] = useState<PMReportExtractionDraft | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const endRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    endRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isExtracting]);

  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  }, [onClose]);

  const pushMessage = (role: 'assistant' | 'user', content: string) => {
    setMessages(prev => [...prev, { id: generateId(), role, content }]);
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (evt) => {
        setAttachedFile({ name: file.name, content: String(evt.target?.result || '') });
      };
      reader.readAsText(file);
    }
    e.target.value = '';
  };

  const handleExtract = async () => {
    const messageText = inputText.trim();
    const combined = [
      messageText,
      attachedFile ? `\n\n[Attached file: ${attachedFile.name}]\n${attachedFile.content}` : ''
    ].join('').trim();

    if (!combined) {
      setExtractError('Please write a message or attach a document before extraction.');
      return;
    }

    pushMessage(
      'user',
      [
        messageText || '(No free-text prompt)',
        attachedFile ? `Attachment: ${attachedFile.name}` : ''
      ].filter(Boolean).join('\n')
    );

    setInputText('');
    setAttachedFile(null);
    setExtractError(null);
    setIsExtracting(true);

    try {
      const result = await extractPMReportFromText(combined, llmConfig);
      setExtracted(result);

      const hasData =
        Boolean(result.executiveSummary.trim()) ||
        Boolean(result.keyDecisions.trim()) ||
        Boolean(result.nextSteps.trim()) ||
        result.overallCompletionPct != null ||
        result.budgetAllocated != null ||
        result.budgetSpent != null ||
        result.budgetForecast != null ||
        result.incidents.length > 0 ||
        result.updates.length > 0 ||
        result.news.length > 0 ||
        result.milestones.length > 0 ||
        result.risks.length > 0;

      pushMessage(
        'assistant',
        hasData
          ? 'Extraction completed. Review the extracted PM report mask on the right, then apply it to the standard PM Report editor.'
          : 'Extraction completed, but not enough explicit information was found. Fields are left blank as requested.'
      );
    } catch (e: any) {
      const err = e?.message || 'Extraction failed';
      setExtractError(err);
      pushMessage('assistant', `Extraction failed: ${err}`);
    } finally {
      setIsExtracting(false);
    }
  };

  const ragItems: Array<{ label: string; value: PMReportExtractionDraft['overallStatus'] }> = extracted ? [
    { label: 'Overall', value: extracted.overallStatus },
    { label: 'Scope', value: extracted.scopeStatus },
    { label: 'Schedule', value: extracted.scheduleStatus },
    { label: 'Budget', value: extracted.budgetStatus },
    { label: 'Resource', value: extracted.resourceStatus },
  ] : [];

  const renderRAGValue = (value: string) => {
    if (value === 'Green' || value === 'Amber' || value === 'Red') {
      return (
        <div className="flex items-center gap-2">
          <RAGDot status={value} size="sm" />
          <span className="text-xs font-semibold text-gray-700 dark:text-gray-200">{value}</span>
        </div>
      );
    }
    return <span className="text-xs text-gray-400">Not provided</span>;
  };

  const asLine = (label: string, value: string | number | null | undefined) => (
    <div className="flex items-start justify-between gap-3">
      <span className="text-xs font-semibold text-gray-500 dark:text-gray-400">{label}</span>
      <span className="text-xs text-right text-gray-700 dark:text-gray-200">{value == null || value === '' ? 'Not provided' : value}</span>
    </div>
  );

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-sm p-4 flex items-center justify-center" onClick={onClose}>
      <div className="bg-white dark:bg-gray-900 rounded-2xl border border-gray-200 dark:border-gray-800 shadow-2xl w-full max-w-7xl h-[90vh] overflow-hidden flex flex-col lg:flex-row" onClick={e => e.stopPropagation()}>
        <div className="w-full lg:w-[45%] flex-1 border-b lg:border-b-0 lg:border-r border-gray-200 dark:border-gray-800 flex flex-col min-h-0">
          <div className="px-5 py-4 border-b border-gray-200 dark:border-gray-800 flex items-center justify-between shrink-0">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-lg bg-indigo-100 dark:bg-indigo-900/30 flex items-center justify-center">
                <Bot className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
              </div>
              <div>
                <h3 className="font-bold text-gray-900 dark:text-white">AI PM Report</h3>
                <p className="text-xs text-gray-500 dark:text-gray-400">Chat + extraction to PM report template</p>
              </div>
            </div>
            <button onClick={onClose} className="p-2 rounded-lg hover:bg-gray-100 dark:hover:bg-gray-800 text-gray-500">
              <X className="w-4 h-4" />
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-4 space-y-3 bg-gray-50/60 dark:bg-gray-950/40">
            {messages.map(msg => (
              <div
                key={msg.id}
                className={`max-w-[90%] rounded-xl px-3 py-2 text-sm leading-relaxed ${
                  msg.role === 'assistant'
                    ? 'bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-700 dark:text-gray-200'
                    : 'ml-auto bg-indigo-600 text-white'
                }`}
              >
                {msg.content}
              </div>
            ))}
            {isExtracting && (
              <div className="max-w-[90%] rounded-xl px-3 py-2 text-sm bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-gray-600 dark:text-gray-300 flex items-center gap-2">
                <Loader2 className="w-4 h-4 animate-spin text-indigo-500" />
                Extracting PM report fields...
              </div>
            )}
            <div ref={endRef} />
          </div>

          <div className="p-4 border-t border-gray-200 dark:border-gray-800 space-y-3 shrink-0">
            {extractError && (
              <div className="text-xs text-red-600 dark:text-red-400 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg px-3 py-2">
                {extractError}
              </div>
            )}

            <textarea
              value={inputText}
              onChange={e => setInputText(e.target.value)}
              rows={4}
              className="w-full p-3 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 outline-none focus:ring-2 focus:ring-indigo-500 resize-none"
              placeholder="Describe the subject to analyze (email, presentation, project notes...)"
            />

            <input
              ref={fileInputRef}
              type="file"
              className="hidden"
              onChange={handleFileSelect}
              accept=".txt,.md,.csv,.json,.xml,.html,.log"
            />

            <div className="flex items-center justify-between gap-3">
              {attachedFile ? (
                <div className="flex items-center gap-2 min-w-0">
                  <span className="text-xs text-gray-500 dark:text-gray-400 truncate">{attachedFile.name}</span>
                  <button onClick={() => setAttachedFile(null)} className="text-red-500 hover:text-red-600">
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="inline-flex items-center gap-2 px-3 py-2 text-xs font-semibold border border-dashed border-gray-300 dark:border-gray-700 rounded-lg text-gray-600 dark:text-gray-300 hover:border-indigo-400 hover:text-indigo-600 transition-colors"
                >
                  <Paperclip className="w-3.5 h-3.5" />
                  Attach document
                </button>
              )}

              <button
                onClick={handleExtract}
                disabled={isExtracting || (!inputText.trim() && !attachedFile)}
                className="inline-flex items-center gap-2 px-4 py-2 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 disabled:opacity-50 text-white rounded-lg transition-colors"
              >
                <Send className="w-3.5 h-3.5" />
                Extract
              </button>
            </div>

            {extracted && (
              <button
                onClick={() => onApply(extracted)}
                className="w-full px-4 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors"
              >
                Apply Extracted Mask to Editor
              </button>
            )}
          </div>
        </div>

        <div className="w-full lg:w-[55%] flex-1 flex flex-col min-h-0">
          <div className="px-5 py-4 border-b border-gray-200 dark:border-gray-800 flex items-center justify-between shrink-0">
            <div>
              <h4 className="font-bold text-gray-900 dark:text-white">Extracted PM Mask</h4>
              <p className="text-xs text-gray-500 dark:text-gray-400">Review before injecting into the standard PM report editor</p>
            </div>
            <button
              onClick={() => extracted && onApply(extracted)}
              disabled={!extracted}
              className="px-4 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 rounded-lg transition-colors"
            >
              Apply to Editor
            </button>
          </div>

          <div className="flex-1 overflow-y-auto p-5 space-y-4">
            {!extracted && (
              <div className="h-full flex items-center justify-center text-center text-sm text-gray-400 dark:text-gray-500">
                Run an extraction from the chat panel to preview the AI-generated PM report mask.
              </div>
            )}

            {extracted && (
              <>
                <div className="border border-gray-200 dark:border-gray-800 rounded-xl p-4">
                  <h5 className="text-xs font-bold uppercase tracking-wide text-gray-500 dark:text-gray-400 mb-3">RAG</h5>
                  <div className="grid grid-cols-2 gap-3">
                    {ragItems.map(item => (
                      <div key={item.label} className="flex items-center justify-between border border-gray-100 dark:border-gray-800 rounded-lg px-3 py-2">
                        <span className="text-xs font-semibold text-gray-500 dark:text-gray-400">{item.label}</span>
                        {renderRAGValue(item.value)}
                      </div>
                    ))}
                  </div>
                  <div className="mt-3">{asLine('Completion', extracted.overallCompletionPct != null ? `${extracted.overallCompletionPct}%` : null)}</div>
                </div>

                <div className="border border-gray-200 dark:border-gray-800 rounded-xl p-4 space-y-2">
                  <h5 className="text-xs font-bold uppercase tracking-wide text-gray-500 dark:text-gray-400">Summary</h5>
                  {asLine('Executive Summary', extracted.executiveSummary || null)}
                  {asLine('Key Decisions', extracted.keyDecisions || null)}
                  {asLine('Next Steps', extracted.nextSteps || null)}
                </div>

                <div className="border border-gray-200 dark:border-gray-800 rounded-xl p-4 space-y-2">
                  <h5 className="text-xs font-bold uppercase tracking-wide text-gray-500 dark:text-gray-400">Cost (MD)</h5>
                  {asLine('Allocated', extracted.budgetAllocated)}
                  {asLine('Spent', extracted.budgetSpent)}
                  {asLine('Forecast', extracted.budgetForecast)}
                </div>

                <div className="border border-gray-200 dark:border-gray-800 rounded-xl p-4 space-y-2">
                  <h5 className="text-xs font-bold uppercase tracking-wide text-gray-500 dark:text-gray-400">Extracted Items</h5>
                  {asLine('Incidents', extracted.incidents.length)}
                  {asLine('Updates', extracted.updates.length)}
                  {asLine('News', extracted.news.length)}
                  {asLine('Milestones', extracted.milestones.length)}
                  {asLine('Risks', extracted.risks.length)}
                </div>

                <div className="text-xs text-gray-500 dark:text-gray-400 bg-indigo-50 dark:bg-indigo-900/20 border border-indigo-100 dark:border-indigo-800/60 rounded-lg p-3">
                  After applying, you can amend all fields in the standard PM Report edit form before saving.
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

// ════════════════════════════════════════
//  MAIN COMPONENT
// ════════════════════════════════════════
const PMReport: React.FC<PMReportProps> = ({
  teams, users, currentUser, llmConfig, pmReportData, onSavePMReport, onDeletePMReport
}) => {
  const [view, setView] = useState<PMView>('overview');
  const [selectedProjectIds, setSelectedProjectIds] = useState<string[]>([]);
  const [editingReport, setEditingReport] = useState<PMReportData | null>(null);
  const [generatedHTML, setGeneratedHTML] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [expandedSections, setExpandedSections] = useState<Record<string, boolean>>({
    rag: true, summary: true, incidents: true, updates: true, news: true, milestones: true, risks: true, budget: true
  });
  const [versionPanelProject, setVersionPanelProject] = useState<string | null>(null);
  const [showAIPMModal, setShowAIPMModal] = useState(false);
  const [prefillNotice, setPrefillNotice] = useState<string | null>(null);

  const allProjects = useMemo(() => {
    const projects: (Project & { teamName: string })[] = [];
    teams.forEach(t => t.projects.forEach(p => {
      if (!p.isArchived) projects.push({ ...p, teamName: t.name });
    }));
    return projects;
  }, [teams]);

  const teamNameById = useMemo(() => {
    const map: Record<string, string> = {};
    teams.forEach(team => {
      map[team.id] = team.name;
    });
    return map;
  }, [teams]);

  const usersById = useMemo(() => {
    const map: Record<string, User> = {};
    users.forEach(user => {
      map[user.id] = user;
    });
    return map;
  }, [users]);

  const getProjectVersions = (projectId: string): PMReportData[] => {
    return pmReportData
      .filter(r => r.projectId === projectId)
      .sort((a, b) => (b.version || 1) - (a.version || 1));
  };

  const getLatestVersion = (projectId: string): PMReportData | undefined => {
    const versions = getProjectVersions(projectId);
    return versions[0];
  };

  const toggleSection = (key: string) => setExpandedSections(prev => ({ ...prev, [key]: !prev[key] }));
  const toggleProject = (id: string) => setSelectedProjectIds(prev =>
    prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
  );

  // ─── Start editing a specific version ───
  const handleEditVersion = (report: PMReportData) => {
    setEditingReport({ ...report });
    setVersionPanelProject(null);
    setShowAIPMModal(false);
    setPrefillNotice(null);
    setView('data-entry');
  };

  // ─── Create a new version from the latest one (or blank) ───
  const handleNewVersion = (projectId: string) => {
    const latest = getLatestVersion(projectId);
    const newVersion = latest ? (latest.version || 1) + 1 : 1;
    const newReport = latest
      ? cloneReport(latest, newVersion)
      : createEmptyReport(projectId, currentUser.id, newVersion);
    setEditingReport(newReport);
    setVersionPanelProject(null);
    setShowAIPMModal(false);
    setPrefillNotice(null);
    setView('data-entry');
  };

  // ─── Quick edit: go to latest version for a project ───
  const handleEditReport = (projectId: string) => {
    const latest = getLatestVersion(projectId);
    if (latest) {
      handleEditVersion(latest);
    } else {
      handleNewVersion(projectId);
    }
  };

  const handleSave = () => {
    if (!editingReport) return;
    const toSave = { ...editingReport, updatedAt: new Date().toISOString() };
    if (!toSave.version) toSave.version = 1;
    if (!toSave.versionLabel) toSave.versionLabel = `v${toSave.version}`;
    onSavePMReport(toSave);
    setView('overview');
    setShowAIPMModal(false);
    setPrefillNotice(null);
    setEditingReport(null);
  };

  const handlePrefillFromProjectData = () => {
    if (!editingReport) return;
    const project = allProjects.find(p => p.id === editingReport.projectId);
    if (!project) return;

    setEditingReport(current => {
      if (!current) return current;
      return mergePrefillFromProjectIntoReport(current, project, usersById);
    });
    setPrefillNotice('Form pre-filled from existing project data and visible tasks. AI PM Report can now complete remaining missing fields.');
  };

  // ─── Report generation ───
  const handleGenerateReport = async () => {
    if (selectedProjectIds.length === 0) return;
    setIsGenerating(true);
    try {
      const reportsForGeneration = selectedProjectIds.map(pid => {
        const report = getLatestVersion(pid);
        const project = allProjects.find(p => p.id === pid);
        return { project, report };
      }).filter(x => x.report && x.project) as { project: Project & { teamName: string }; report: PMReportData }[];
      reportsForGeneration.sort((a, b) => {
        const teamCmp = a.project.teamName.localeCompare(b.project.teamName);
        if (teamCmp !== 0) return teamCmp;
        return a.project.name.localeCompare(b.project.name);
      });

      if (reportsForGeneration.length === 0) {
        setGeneratedHTML('<p style="color:#ef4444;font-family:sans-serif;padding:20px;">No report data found for selected projects. Please fill in data first.</p>');
        setView('report-preview');
        setIsGenerating(false);
        return;
      }

      // Deterministic export format to avoid style drift between generations.
      const html = buildConsultingDeckHTML(reportsForGeneration, teamNameById);
      setGeneratedHTML(html);
      setView('report-preview');
    } catch (err) {
      console.error('Report generation failed:', err);
      setGeneratedHTML('<p style="color:#ef4444;font-family:sans-serif;padding:20px;">Report generation failed. Please check your LLM configuration.</p>');
      setView('report-preview');
    }
    setIsGenerating(false);
  };

  // ─── PDF Export ───
  const buildExportDocumentHTML = (withPrintControls: boolean) => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>PM Status Report</title>
<style>@page{size:landscape;margin:9mm}@media print{body{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;font-size:11.5px;line-height:1.35}.no-print{display:none!important}.deck{max-width:none}.project-block,.report-section,.rag-card,.metric-box,.update-item{break-inside:avoid-page;page-break-inside:avoid}table,tr,th,td{break-inside:avoid-page;page-break-inside:avoid}thead{display:table-header-group}}*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:#1e293b;background:#fff;padding:20px;font-size:12px}</style>
</head><body>
${withPrintControls ? `
<div class="no-print" style="text-align:center;margin-bottom:24px;padding:16px;background:#f8fafc;border-radius:12px;border:1px solid #e2e8f0">
  <button onclick="window.print()" style="background:linear-gradient(135deg,#00915A,#00A86B);color:#fff;border:none;padding:14px 40px;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;box-shadow:0 4px 12px rgba(0,145,90,.3)">Print / Save as PDF</button>
  <button onclick="window.close()" style="margin-left:14px;background:#fff;color:#374151;border:1px solid #d1d5db;padding:14px 28px;border-radius:10px;font-size:15px;cursor:pointer">Cancel</button>
  <p style="margin-top:10px;font-size:12px;color:#64748b">Use your browser's print dialog to save as PDF in landscape orientation</p>
</div>` : ''}
${generatedHTML}</body></html>`;

  const handleExportPDF = () => {
    const w = window.open('', '_blank');
    if (!w) return;
    w.document.write(buildExportDocumentHTML(true));
    w.document.close();
  };

  const handleExportOutlookEmail = () => {
    if (!generatedHTML.trim()) return;
    const subjectDate = new Date().toLocaleDateString('en-GB');
    const fileDate = new Date().toISOString().split('T')[0];
    const htmlBody = buildExportDocumentHTML(false);
    const emlContent = [
      'X-Unsent: 1',
      'To: ',
      `Subject: PM Status Report - ${subjectDate}`,
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
    link.download = `pm-status-report-${fileDate}.eml`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  // ════════════════════════════════════════
  //  DATA ENTRY FORM
  // ════════════════════════════════════════
  const renderDataEntry = () => {
    if (!editingReport) return null;
    const project = allProjects.find(p => p.id === editingReport.projectId);

    const SectionHeader: React.FC<{ id: string; icon: React.ReactNode; title: string; count?: number }> = ({ id, icon, title, count }) => (
      <button onClick={() => toggleSection(id)} className="w-full flex items-center justify-between py-3 px-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg mb-3 hover:bg-gray-100 dark:hover:bg-gray-800 transition-colors">
        <div className="flex items-center gap-2">
          {icon}
          <span className="text-sm font-bold text-gray-800 dark:text-gray-200">{title}</span>
          {typeof count === 'number' && <span className="text-xs bg-indigo-100 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-2 py-0.5 rounded-full font-bold">{count}</span>}
        </div>
        {expandedSections[id] ? <ChevronUp className="w-4 h-4 text-gray-400" /> : <ChevronDown className="w-4 h-4 text-gray-400" />}
      </button>
    );

    const ic = "w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100 focus:ring-2 focus:ring-indigo-500 focus:border-transparent outline-none";
    const tc = `${ic} resize-none`;
    const costDistribution = editingReport.costDistribution || [];

    const updateCostSplit = (splitId: string, field: keyof PMReportCostSplit, value: string | number) => {
      const nextSplits = costDistribution.map(split =>
        split.id === splitId ? { ...split, [field]: value } : split
      );
      setEditingReport({ ...editingReport, costDistribution: nextSplits });
    };

    const addCostSplit = () => {
      setEditingReport({
        ...editingReport,
        costDistribution: [
          ...costDistribution,
          { id: generateId(), teamName: '', allocatedMD: 0, spentMD: 0, forecastMD: 0 }
        ]
      });
    };

    const removeCostSplit = (splitId: string) => {
      setEditingReport({
        ...editingReport,
        costDistribution: costDistribution.filter(split => split.id !== splitId)
      });
    };

    const splitTotals = costDistribution.reduce(
      (acc, split) => ({
        allocated: acc.allocated + (split.allocatedMD || 0),
        spent: acc.spent + (split.spentMD || 0),
        forecast: acc.forecast + (split.forecastMD || 0),
      }),
      { allocated: 0, spent: 0, forecast: 0 }
    );

    const hasSplitDelta =
      Math.abs(splitTotals.allocated - (editingReport.budgetAllocated || 0)) > 0.1 ||
      Math.abs(splitTotals.spent - (editingReport.budgetSpent || 0)) > 0.1 ||
      Math.abs(splitTotals.forecast - (editingReport.budgetForecast || 0)) > 0.1;

    return (
      <div className="max-w-4xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h2 className="text-xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
              <Edit3 className="w-5 h-5 text-indigo-500" />
              {project?.name || 'Unknown'}
            </h2>
            <div className="flex items-center gap-3 mt-1">
              <span className="text-sm text-gray-500">Version {editingReport.version || 1}</span>
              <span className="text-xs bg-indigo-100 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-2 py-0.5 rounded-full font-medium">{editingReport.versionLabel}</span>
            </div>
          </div>
          <div className="flex gap-2">
            <button
              onClick={handlePrefillFromProjectData}
              disabled={!project}
              className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-700 rounded-lg hover:bg-emerald-100 dark:hover:bg-emerald-900/40 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <Sparkles className="w-4 h-4" /> Prefill Project Data
            </button>
            <button onClick={() => setShowAIPMModal(true)} className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-lg hover:bg-indigo-100 dark:hover:bg-indigo-900/40 transition-colors">
              <Bot className="w-4 h-4" /> AI PM Report
            </button>
            <button onClick={() => { setView('overview'); setEditingReport(null); setShowAIPMModal(false); setPrefillNotice(null); }} className="px-4 py-2 text-sm font-medium text-gray-600 bg-gray-100 dark:bg-gray-800 dark:text-gray-400 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-700 transition-colors">Cancel</button>
            <button onClick={handleSave} className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-white bg-indigo-600 rounded-lg hover:bg-indigo-700 transition-colors shadow-sm">
              <Save className="w-4 h-4" /> Save
            </button>
          </div>
        </div>

        {prefillNotice && (
          <div className="mb-4 rounded-lg border border-emerald-200 dark:border-emerald-800 bg-emerald-50 dark:bg-emerald-900/20 px-4 py-2.5 text-xs text-emerald-800 dark:text-emerald-200">
            {prefillNotice}
          </div>
        )}

        {/* VERSION LABEL */}
        <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4 shadow-sm">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Version Label</label>
              <input className={ic} value={editingReport.versionLabel} onChange={e => setEditingReport({ ...editingReport, versionLabel: e.target.value })} placeholder="e.g. v3 — Sprint 12 Closing" />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Confidentiality Level</label>
              <select
                className={ic}
                value={getReportConfidentiality(editingReport)}
                onChange={e => setEditingReport({ ...editingReport, confidentialityLevel: e.target.value as PMReportConfidentiality })}
              >
                {CONFIDENTIALITY_LEVELS.map(level => (
                  <option key={level} value={level}>{level}</option>
                ))}
              </select>
            </div>
          </div>
        </div>

        {/* RAG */}
        <SectionHeader id="rag" icon={<CircleDot className="w-4 h-4 text-indigo-500" />} title="RAG Status" />
        {expandedSections.rag && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              <RAGSelector label="Overall" value={editingReport.overallStatus} onChange={v => setEditingReport({ ...editingReport, overallStatus: v })} />
              <RAGSelector label="Scope" value={editingReport.scopeStatus} onChange={v => setEditingReport({ ...editingReport, scopeStatus: v })} />
              <RAGSelector label="Schedule" value={editingReport.scheduleStatus} onChange={v => setEditingReport({ ...editingReport, scheduleStatus: v })} />
              <RAGSelector label="Budget" value={editingReport.budgetStatus} onChange={v => setEditingReport({ ...editingReport, budgetStatus: v })} />
              <RAGSelector label="Resource" value={editingReport.resourceStatus} onChange={v => setEditingReport({ ...editingReport, resourceStatus: v })} />
              <div className="flex items-center gap-2">
                <span className="text-xs font-medium text-gray-500 dark:text-gray-400 min-w-[70px]">Incident</span>
                <RAGDot status={getIncidentRAGStatus(editingReport)} size="md" />
                <span className="text-xs font-semibold text-gray-600 dark:text-gray-300">{getIncidentRAGStatus(editingReport)}</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-medium text-gray-500 dark:text-gray-400 min-w-[70px]">Complete</span>
                <input type="range" min={0} max={100} value={editingReport.overallCompletionPct} onChange={e => setEditingReport({ ...editingReport, overallCompletionPct: parseInt(e.target.value) })} className="flex-1 accent-indigo-600" />
                <span className="text-sm font-bold text-indigo-600 w-10 text-right">{editingReport.overallCompletionPct}%</span>
              </div>
            </div>
          </div>
        )}

        {/* SUMMARY */}
        <SectionHeader id="summary" icon={<FileBarChart className="w-4 h-4 text-indigo-500" />} title="Summary & Decisions" />
        {expandedSections.summary && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm space-y-4">
            <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Executive Summary</label>
              <textarea rows={3} className={tc} value={editingReport.executiveSummary} onChange={e => setEditingReport({ ...editingReport, executiveSummary: e.target.value })} placeholder="Brief overview of project status for senior management..." /></div>
            <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Key Decisions</label>
              <textarea rows={2} className={tc} value={editingReport.keyDecisions} onChange={e => setEditingReport({ ...editingReport, keyDecisions: e.target.value })} placeholder="Key decisions made or pending..." /></div>
            <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Next Steps</label>
              <textarea rows={2} className={tc} value={editingReport.nextSteps} onChange={e => setEditingReport({ ...editingReport, nextSteps: e.target.value })} placeholder="Planned actions for the coming period..." /></div>
          </div>
        )}

        {/* INCIDENTS */}
        <SectionHeader id="incidents" icon={<AlertTriangle className="w-4 h-4 text-red-500" />} title="Incidents" count={editingReport.incidents.length} />
        {expandedSections.incidents && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            {editingReport.incidents.map((inc, idx) => (
              <div key={inc.id} className="mb-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-100 dark:border-gray-700 relative">
                <button onClick={() => setEditingReport({ ...editingReport, incidents: editingReport.incidents.filter(i => i.id !== inc.id) })} className="absolute top-2 right-2 p-1 text-red-400 hover:text-red-600 rounded"><Trash2 className="w-3.5 h-3.5" /></button>
                <div className="grid grid-cols-3 gap-3 mb-3">
                  <input className={ic} placeholder="Title" value={inc.title} onChange={e => { const a = [...editingReport.incidents]; a[idx] = { ...a[idx], title: e.target.value }; setEditingReport({ ...editingReport, incidents: a }); }} />
                  <input type="date" className={ic} value={inc.date} onChange={e => { const a = [...editingReport.incidents]; a[idx] = { ...a[idx], date: e.target.value }; setEditingReport({ ...editingReport, incidents: a }); }} />
                  <div className="flex gap-2">
                    <select className={ic} value={inc.severity} onChange={e => { const a = [...editingReport.incidents]; a[idx] = { ...a[idx], severity: e.target.value as any }; setEditingReport({ ...editingReport, incidents: a }); }}><option>Critical</option><option>Major</option><option>Minor</option></select>
                    <select className={ic} value={inc.status} onChange={e => { const a = [...editingReport.incidents]; a[idx] = { ...a[idx], status: e.target.value as any }; setEditingReport({ ...editingReport, incidents: a }); }}><option>Open</option><option>Investigating</option><option>Resolved</option></select>
                  </div>
                </div>
                <textarea rows={2} className={tc} placeholder="Description..." value={inc.description} onChange={e => { const a = [...editingReport.incidents]; a[idx] = { ...a[idx], description: e.target.value }; setEditingReport({ ...editingReport, incidents: a }); }} />
              </div>
            ))}
            <button onClick={() => setEditingReport({ ...editingReport, incidents: [...editingReport.incidents, { id: generateId(), date: new Date().toISOString().split('T')[0], title: '', description: '', severity: 'Minor', status: 'Open' }] })} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 dark:text-indigo-400 border border-dashed border-indigo-300 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors w-full justify-center"><Plus className="w-4 h-4" /> Add Incident</button>
          </div>
        )}

        {/* UPDATES */}
        <SectionHeader id="updates" icon={<TrendingUp className="w-4 h-4 text-blue-500" />} title="Updates" count={editingReport.updates.length} />
        {expandedSections.updates && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            {editingReport.updates.map((upd, idx) => (
              <div key={upd.id} className="mb-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-100 dark:border-gray-700 relative">
                <button onClick={() => setEditingReport({ ...editingReport, updates: editingReport.updates.filter(u => u.id !== upd.id) })} className="absolute top-2 right-2 p-1 text-red-400 hover:text-red-600 rounded"><Trash2 className="w-3.5 h-3.5" /></button>
                <div className="grid grid-cols-4 gap-3 mb-3">
                  <input className={ic} placeholder="Title" value={upd.title} onChange={e => { const a = [...editingReport.updates]; a[idx] = { ...a[idx], title: e.target.value }; setEditingReport({ ...editingReport, updates: a }); }} />
                  <input type="date" className={ic} value={upd.date} onChange={e => { const a = [...editingReport.updates]; a[idx] = { ...a[idx], date: e.target.value }; setEditingReport({ ...editingReport, updates: a }); }} />
                  <select className={ic} value={upd.category} onChange={e => { const a = [...editingReport.updates]; a[idx] = { ...a[idx], category: e.target.value as any }; setEditingReport({ ...editingReport, updates: a }); }}>{['Scope', 'Timeline', 'Budget', 'Resource', 'Technical', 'Risk', 'Other'].map(c => <option key={c}>{c}</option>)}</select>
                  <RAGSelector value={upd.impact} onChange={v => { const a = [...editingReport.updates]; a[idx] = { ...a[idx], impact: v }; setEditingReport({ ...editingReport, updates: a }); }} />
                </div>
                <textarea rows={2} className={tc} placeholder="Description..." value={upd.description} onChange={e => { const a = [...editingReport.updates]; a[idx] = { ...a[idx], description: e.target.value }; setEditingReport({ ...editingReport, updates: a }); }} />
              </div>
            ))}
            <button onClick={() => setEditingReport({ ...editingReport, updates: [...editingReport.updates, { id: generateId(), date: new Date().toISOString().split('T')[0], category: 'Other', title: '', description: '', impact: 'Green' }] })} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 dark:text-indigo-400 border border-dashed border-indigo-300 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors w-full justify-center"><Plus className="w-4 h-4" /> Add Update</button>
          </div>
        )}

        {/* NEWS */}
        <SectionHeader id="news" icon={<Megaphone className="w-4 h-4 text-emerald-500" />} title="News & Achievements" count={editingReport.news.length} />
        {expandedSections.news && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            {editingReport.news.map((n, idx) => (
              <div key={n.id} className="mb-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-100 dark:border-gray-700 relative">
                <button onClick={() => setEditingReport({ ...editingReport, news: editingReport.news.filter(x => x.id !== n.id) })} className="absolute top-2 right-2 p-1 text-red-400 hover:text-red-600 rounded"><Trash2 className="w-3.5 h-3.5" /></button>
                <div className="grid grid-cols-3 gap-3 mb-3">
                  <input className={ic} placeholder="Title" value={n.title} onChange={e => { const a = [...editingReport.news]; a[idx] = { ...a[idx], title: e.target.value }; setEditingReport({ ...editingReport, news: a }); }} />
                  <input type="date" className={ic} value={n.date} onChange={e => { const a = [...editingReport.news]; a[idx] = { ...a[idx], date: e.target.value }; setEditingReport({ ...editingReport, news: a }); }} />
                  <select className={ic} value={n.type} onChange={e => { const a = [...editingReport.news]; a[idx] = { ...a[idx], type: e.target.value as any }; setEditingReport({ ...editingReport, news: a }); }}><option>Achievement</option><option>Announcement</option><option>Change</option><option>Info</option></select>
                </div>
                <textarea rows={2} className={tc} placeholder="Description..." value={n.description} onChange={e => { const a = [...editingReport.news]; a[idx] = { ...a[idx], description: e.target.value }; setEditingReport({ ...editingReport, news: a }); }} />
              </div>
            ))}
            <button onClick={() => setEditingReport({ ...editingReport, news: [...editingReport.news, { id: generateId(), date: new Date().toISOString().split('T')[0], title: '', description: '', type: 'Info' }] })} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 dark:text-indigo-400 border border-dashed border-indigo-300 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors w-full justify-center"><Plus className="w-4 h-4" /> Add News</button>
          </div>
        )}

        {/* MILESTONES */}
        <SectionHeader id="milestones" icon={<Target className="w-4 h-4 text-purple-500" />} title="Milestones / Planning" count={editingReport.milestones.length} />
        {expandedSections.milestones && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            {editingReport.milestones.map((m, idx) => (
              <div key={m.id} className="mb-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-100 dark:border-gray-700 relative">
                <button onClick={() => setEditingReport({ ...editingReport, milestones: editingReport.milestones.filter(x => x.id !== m.id) })} className="absolute top-2 right-2 p-1 text-red-400 hover:text-red-600 rounded"><Trash2 className="w-3.5 h-3.5" /></button>
                <div className="grid grid-cols-5 gap-3 mb-3">
                  <input className={`${ic} col-span-2`} placeholder="Milestone name" value={m.name} onChange={e => { const a = [...editingReport.milestones]; a[idx] = { ...a[idx], name: e.target.value }; setEditingReport({ ...editingReport, milestones: a }); }} />
                  <input type="date" className={ic} value={m.plannedDate} title="Planned" onChange={e => { const a = [...editingReport.milestones]; a[idx] = { ...a[idx], plannedDate: e.target.value }; setEditingReport({ ...editingReport, milestones: a }); }} />
                  <input type="date" className={ic} value={m.revisedDate || ''} title="Revised" onChange={e => { const a = [...editingReport.milestones]; a[idx] = { ...a[idx], revisedDate: e.target.value || undefined }; setEditingReport({ ...editingReport, milestones: a }); }} />
                  <RAGSelector value={m.status} onChange={v => { const a = [...editingReport.milestones]; a[idx] = { ...a[idx], status: v }; setEditingReport({ ...editingReport, milestones: a }); }} />
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-xs font-medium text-gray-500 min-w-[60px]">Progress</span>
                  <input type="range" min={0} max={100} value={m.completionPct} onChange={e => { const a = [...editingReport.milestones]; a[idx] = { ...a[idx], completionPct: parseInt(e.target.value) }; setEditingReport({ ...editingReport, milestones: a }); }} className="flex-1 accent-indigo-600" />
                  <span className="text-sm font-bold text-indigo-600 w-10 text-right">{m.completionPct}%</span>
                </div>
              </div>
            ))}
            <button onClick={() => setEditingReport({ ...editingReport, milestones: [...editingReport.milestones, { id: generateId(), name: '', plannedDate: '', status: 'Green', completionPct: 0 }] })} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 dark:text-indigo-400 border border-dashed border-indigo-300 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors w-full justify-center"><Plus className="w-4 h-4" /> Add Milestone</button>
          </div>
        )}

        {/* RISKS */}
        <SectionHeader id="risks" icon={<Shield className="w-4 h-4 text-orange-500" />} title="Risk Register" count={editingReport.risks.length} />
        {expandedSections.risks && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            {editingReport.risks.map((r, idx) => (
              <div key={r.id} className="mb-4 p-4 bg-gray-50 dark:bg-gray-800/50 rounded-lg border border-gray-100 dark:border-gray-700 relative">
                <button onClick={() => setEditingReport({ ...editingReport, risks: editingReport.risks.filter(x => x.id !== r.id) })} className="absolute top-2 right-2 p-1 text-red-400 hover:text-red-600 rounded"><Trash2 className="w-3.5 h-3.5" /></button>
                <div className="grid grid-cols-4 gap-3 mb-3">
                  <input className={`${ic} col-span-2`} placeholder="Risk description" value={r.description} onChange={e => { const a = [...editingReport.risks]; a[idx] = { ...a[idx], description: e.target.value }; setEditingReport({ ...editingReport, risks: a }); }} />
                  <select className={ic} value={r.likelihood} onChange={e => { const a = [...editingReport.risks]; a[idx] = { ...a[idx], likelihood: e.target.value as any }; setEditingReport({ ...editingReport, risks: a }); }}><option value="Low">Likelihood: Low</option><option value="Medium">Likelihood: Medium</option><option value="High">Likelihood: High</option></select>
                  <select className={ic} value={r.impact} onChange={e => { const a = [...editingReport.risks]; a[idx] = { ...a[idx], impact: e.target.value as any }; setEditingReport({ ...editingReport, risks: a }); }}><option value="Low">Impact: Low</option><option value="Medium">Impact: Medium</option><option value="High">Impact: High</option></select>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <textarea rows={2} className={tc} placeholder="Mitigation plan..." value={r.mitigation} onChange={e => { const a = [...editingReport.risks]; a[idx] = { ...a[idx], mitigation: e.target.value }; setEditingReport({ ...editingReport, risks: a }); }} />
                  <input className={ic} placeholder="Risk owner" value={r.owner} onChange={e => { const a = [...editingReport.risks]; a[idx] = { ...a[idx], owner: e.target.value }; setEditingReport({ ...editingReport, risks: a }); }} />
                </div>
              </div>
            ))}
            <button onClick={() => setEditingReport({ ...editingReport, risks: [...editingReport.risks, { id: generateId(), description: '', likelihood: 'Medium', impact: 'Medium', mitigation: '', owner: '' }] })} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-indigo-600 dark:text-indigo-400 border border-dashed border-indigo-300 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors w-full justify-center"><Plus className="w-4 h-4" /> Add Risk</button>
          </div>
        )}

        {/* BUDGET */}
        <SectionHeader id="budget" icon={<DollarSign className="w-4 h-4 text-green-500" />} title="Cost (MD)" />
        {expandedSections.budget && (
          <div className="mb-6 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
            <div className="grid grid-cols-3 gap-4">
              <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Allocated (MD)</label><input type="number" className={ic} value={editingReport.budgetAllocated || ''} onChange={e => setEditingReport({ ...editingReport, budgetAllocated: parseFloat(e.target.value) || 0 })} /></div>
              <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Spent (MD)</label><input type="number" className={ic} value={editingReport.budgetSpent || ''} onChange={e => setEditingReport({ ...editingReport, budgetSpent: parseFloat(e.target.value) || 0 })} /></div>
              <div><label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Forecast (MD)</label><input type="number" className={ic} value={editingReport.budgetForecast || ''} onChange={e => setEditingReport({ ...editingReport, budgetForecast: parseFloat(e.target.value) || 0 })} /></div>
            </div>
            {(editingReport.budgetAllocated || 0) > 0 && (
              <div className="mt-4">
                <div className="flex justify-between text-xs font-medium text-gray-500 mb-1">
                  <span>Cost Utilization</span>
                  <span className="font-bold text-indigo-600">{Math.round(((editingReport.budgetSpent || 0) / (editingReport.budgetAllocated || 1)) * 100)}%</span>
                </div>
                <div className="h-3 bg-gray-200 dark:bg-gray-700 rounded-full overflow-hidden">
                  <div className={`h-full rounded-full transition-all ${((editingReport.budgetSpent || 0) / (editingReport.budgetAllocated || 1)) > 0.9 ? 'bg-red-500' : ((editingReport.budgetSpent || 0) / (editingReport.budgetAllocated || 1)) > 0.7 ? 'bg-amber-500' : 'bg-emerald-500'}`} style={{ width: `${Math.min(100, Math.round(((editingReport.budgetSpent || 0) / (editingReport.budgetAllocated || 1)) * 100))}%` }} />
                </div>
              </div>
            )}

            <div className="mt-6 pt-5 border-t border-gray-200 dark:border-gray-800">
              <div className="flex items-center justify-between mb-3">
                <div>
                  <h4 className="text-sm font-bold text-gray-800 dark:text-gray-100">Manual Team Cost Distribution</h4>
                  <p className="text-xs text-gray-500 dark:text-gray-400">Distribute MD manually with free team labels (not linked to DOINg teams).</p>
                </div>
                <button
                  onClick={addCostSplit}
                  className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-lg hover:bg-indigo-100 dark:hover:bg-indigo-900/40 transition-colors"
                >
                  <Plus className="w-3.5 h-3.5" /> Add Team Split
                </button>
              </div>

              {costDistribution.length === 0 ? (
                <div className="text-xs text-gray-500 dark:text-gray-400 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg p-3">
                  No team distribution yet. Add lines with your own team names to split allocated/spent/forecast MD.
                </div>
              ) : (
                <div className="space-y-3">
                  {costDistribution.map(split => (
                    <div key={split.id} className="p-3 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/50">
                      <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
                        <input
                          className={ic}
                          value={split.teamName || ''}
                          placeholder="Team name (free text)"
                          onChange={e => updateCostSplit(split.id, 'teamName', e.target.value)}
                        />
                        <input
                          type="number"
                          className={ic}
                          value={split.allocatedMD || ''}
                          placeholder="Allocated (MD)"
                          onChange={e => updateCostSplit(split.id, 'allocatedMD', parseFloat(e.target.value) || 0)}
                        />
                        <input
                          type="number"
                          className={ic}
                          value={split.spentMD || ''}
                          placeholder="Spent (MD)"
                          onChange={e => updateCostSplit(split.id, 'spentMD', parseFloat(e.target.value) || 0)}
                        />
                        <input
                          type="number"
                          className={ic}
                          value={split.forecastMD || ''}
                          placeholder="Forecast (MD)"
                          onChange={e => updateCostSplit(split.id, 'forecastMD', parseFloat(e.target.value) || 0)}
                        />
                        <button
                          onClick={() => removeCostSplit(split.id)}
                          className="inline-flex items-center justify-center gap-1 px-3 py-2 text-xs font-semibold text-red-600 dark:text-red-400 border border-red-200 dark:border-red-800 rounded-lg hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" /> Remove
                        </button>
                      </div>
                    </div>
                  ))}

                  <div className="rounded-lg border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-800 p-3">
                    <div className="grid grid-cols-3 gap-3 text-xs">
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Allocated:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.allocated)}</span></div>
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Spent:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.spent)}</span></div>
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Forecast:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.forecast)}</span></div>
                    </div>
                    {hasSplitDelta && (
                      <p className="mt-2 text-xs text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-md px-2 py-1.5">
                        Distribution totals differ from global cost fields (Allocated / Spent / Forecast). Adjust manually if you want exact alignment.
                      </p>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    );
  };

  // ════════════════════════════════════════
  //  REPORT PREVIEW
  // ════════════════════════════════════════
  const renderReportPreview = () => (
    <div className="max-w-5xl mx-auto">
      <div className="flex items-center justify-between mb-6">
        <div>
          <h2 className="text-xl font-bold text-gray-900 dark:text-white flex items-center gap-2"><Eye className="w-5 h-5 text-indigo-500" /> Report Preview</h2>
          <p className="text-sm text-gray-500 mt-1">Review before exporting to PDF</p>
        </div>
        <div className="flex gap-2">
          <button onClick={() => setView('overview')} className="px-4 py-2 text-sm font-medium text-gray-600 bg-gray-100 dark:bg-gray-800 dark:text-gray-400 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-700 transition-colors">Back</button>
          <button onClick={handleExportOutlookEmail} className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-900/30 border border-emerald-200 dark:border-emerald-700 rounded-lg hover:bg-emerald-100 dark:hover:bg-emerald-900/40 transition-colors">
            <Mail className="w-4 h-4" /> Export Email (.eml)
          </button>
          <button onClick={handleExportPDF} className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-white bg-gradient-to-r from-indigo-600 to-purple-600 rounded-lg hover:from-indigo-700 hover:to-purple-700 transition-all shadow-md"><Download className="w-4 h-4" /> Export PDF (Landscape)</button>
        </div>
      </div>
      <div className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-8 shadow-sm overflow-auto" style={{ maxHeight: '75vh' }}>
        <div dangerouslySetInnerHTML={{ __html: generatedHTML }} />
      </div>
    </div>
  );

  // ════════════════════════════════════════
  //  VERSION HISTORY PANEL
  // ════════════════════════════════════════
  const renderVersionPanel = (projectId: string) => {
    const versions = getProjectVersions(projectId);
    const project = allProjects.find(p => p.id === projectId);
    return (
      <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={() => setVersionPanelProject(null)}>
        <div className="bg-white dark:bg-gray-900 rounded-2xl shadow-2xl max-w-lg w-full max-h-[80vh] overflow-hidden" onClick={e => e.stopPropagation()}>
          <div className="px-6 py-4 border-b border-gray-200 dark:border-gray-800 flex items-center justify-between">
            <div>
              <h3 className="text-lg font-bold text-gray-900 dark:text-white flex items-center gap-2"><History className="w-5 h-5 text-indigo-500" /> Version History</h3>
              <p className="text-sm text-gray-500">{project?.name}</p>
            </div>
            <button onClick={() => setVersionPanelProject(null)} className="p-2 hover:bg-gray-100 dark:hover:bg-gray-800 rounded-lg text-gray-400"><ChevronDown className="w-5 h-5" /></button>
          </div>
          <div className="px-6 py-4 overflow-y-auto max-h-[55vh] space-y-3">
            {versions.length === 0 && <p className="text-sm text-gray-400 text-center py-6">No versions yet</p>}
            {versions.map((v, i) => (
              <div key={v.id} className={`p-4 rounded-xl border transition-all ${i === 0 ? 'border-indigo-300 dark:border-indigo-700 bg-indigo-50/50 dark:bg-indigo-900/10' : 'border-gray-200 dark:border-gray-800 bg-gray-50 dark:bg-gray-800/30'}`}>
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <Tag className="w-3.5 h-3.5 text-indigo-500" />
                    <span className="text-sm font-bold text-gray-900 dark:text-white">{v.versionLabel || `v${v.version || 1}`}</span>
                    {i === 0 && <span className="text-[10px] bg-indigo-600 text-white px-2 py-0.5 rounded-full font-bold">LATEST</span>}
                  </div>
                  <RAGDot status={v.overallStatus} size="sm" />
                </div>
                <div className="text-xs text-gray-500 mb-3 flex items-center gap-3">
                  <span>Created: {new Date(v.createdAt).toLocaleDateString()}</span>
                  <span>Updated: {new Date(v.updatedAt).toLocaleDateString()}</span>
                  <span>{v.overallCompletionPct}% complete</span>
                </div>
                <div className="flex gap-2">
                  <button onClick={() => handleEditVersion(v)} className="flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-indigo-600 dark:text-indigo-400 bg-white dark:bg-gray-800 border border-indigo-200 dark:border-indigo-700 rounded-lg hover:bg-indigo-50 dark:hover:bg-indigo-900/20 transition-colors"><Edit3 className="w-3 h-3" /> Edit</button>
                  {versions.length > 1 && <button onClick={() => { if (confirm('Delete this version?')) onDeletePMReport(v.id); }} className="flex items-center gap-1 px-3 py-1.5 text-xs font-semibold text-red-500 bg-white dark:bg-gray-800 border border-red-200 dark:border-red-800 rounded-lg hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors"><Trash2 className="w-3 h-3" /> Delete</button>}
                </div>
              </div>
            ))}
          </div>
          <div className="px-6 py-4 border-t border-gray-200 dark:border-gray-800">
            <button onClick={() => handleNewVersion(projectId)} className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold text-white bg-indigo-600 rounded-xl hover:bg-indigo-700 transition-colors w-full justify-center shadow-sm"><Copy className="w-4 h-4" /> New Version (clone latest)</button>
          </div>
        </div>
      </div>
    );
  };

  // ════════════════════════════════════════
  //  OVERVIEW
  // ════════════════════════════════════════
  const renderOverview = () => (
    <div className="max-w-5xl mx-auto">
      <div className="mb-8 flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-extrabold text-gray-900 dark:text-white flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center shadow-lg"><FileBarChart className="w-5 h-5 text-white" /></div>
            PM Report
          </h2>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-2">Generate professional project status reports for senior management</p>
        </div>
        <button onClick={handleGenerateReport} disabled={selectedProjectIds.length === 0 || isGenerating}
          className={`flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-bold transition-all shadow-md ${selectedProjectIds.length === 0 ? 'bg-gray-200 text-gray-400 cursor-not-allowed dark:bg-gray-800 dark:text-gray-600' : 'bg-gradient-to-r from-indigo-600 to-purple-600 text-white hover:from-indigo-700 hover:to-purple-700 hover:shadow-lg'}`}>
          {isGenerating ? <><Activity className="w-4 h-4 animate-spin" /> Generating...</> : <><Sparkles className="w-4 h-4" /> Generate Report ({selectedProjectIds.length})</>}
        </button>
      </div>

      <div className="mb-6 bg-gradient-to-r from-indigo-50 to-purple-50 dark:from-indigo-900/20 dark:to-purple-900/20 border border-indigo-100 dark:border-indigo-800/50 rounded-xl p-4 flex items-start gap-3">
        <Info className="w-5 h-5 text-indigo-500 flex-shrink-0 mt-0.5" />
        <div className="text-sm text-indigo-800 dark:text-indigo-300">
          <span className="font-bold">How it works:</span> 1) Select projects 2) Click "Edit Data" to fill status data (each save creates a version) 3) Use "History" to manage versions 4) Select projects and "Generate Report" 5) Export to landscape PDF
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {allProjects.map(project => {
          const latest = getLatestVersion(project.id);
          const versionCount = getProjectVersions(project.id).length;
          const isSelected = selectedProjectIds.includes(project.id);
          const tasksDone = project.tasks.filter(t => t.status === 'Done').length;

          return (
            <div key={project.id} className={`relative bg-white dark:bg-gray-900 rounded-xl border-2 p-5 transition-all cursor-pointer hover:shadow-md ${isSelected ? 'border-indigo-500 shadow-md ring-2 ring-indigo-500/20' : 'border-gray-200 dark:border-gray-800'}`} onClick={() => toggleProject(project.id)}>
              <div className={`absolute top-3 right-3 w-6 h-6 rounded-full border-2 flex items-center justify-center transition-all ${isSelected ? 'bg-indigo-600 border-indigo-600' : 'border-gray-300 dark:border-gray-600'}`}>
                {isSelected && <CheckCircle2 className="w-4 h-4 text-white" />}
              </div>

              <div className="flex items-start gap-3 mb-3 pr-8">
                <div className={`w-2 h-8 rounded-full flex-shrink-0 ${project.status === 'Active' ? 'bg-emerald-500' : project.status === 'Planning' ? 'bg-blue-500' : project.status === 'Paused' ? 'bg-amber-500' : 'bg-gray-400'}`} />
                <div>
                  <h3 className="text-sm font-bold text-gray-900 dark:text-white">{project.name}</h3>
                  <p className="text-xs text-gray-500 dark:text-gray-400">{project.teamName} &bull; {project.status}</p>
                </div>
              </div>

              <div className="flex items-center gap-4 mb-3 text-xs text-gray-500 dark:text-gray-400">
                <span className="flex items-center gap-1"><Calendar className="w-3 h-3" /> {project.deadline}</span>
                <span className="flex items-center gap-1"><Users className="w-3 h-3" /> {project.members.length}</span>
                <span className="flex items-center gap-1"><CheckCircle2 className="w-3 h-3" /> {tasksDone}/{project.tasks.length}</span>
              </div>

              {latest ? (
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <RAGDot status={latest.overallStatus} size="sm" />
                    <span className="text-xs font-medium text-gray-500 dark:text-gray-400">{latest.versionLabel || `v${latest.version || 1}`}</span>
                    <span className="text-xs text-gray-400">&bull; {new Date(latest.updatedAt).toLocaleDateString()}</span>
                  </div>
                  <div className="flex items-center gap-1">
                    <button onClick={e => { e.stopPropagation(); setVersionPanelProject(project.id); }}
                      className="flex items-center gap-1 text-xs font-semibold text-gray-500 hover:text-indigo-600 dark:hover:text-indigo-400 transition-colors px-2 py-1 rounded-md hover:bg-gray-100 dark:hover:bg-gray-800">
                      <History className="w-3 h-3" /> {versionCount}
                    </button>
                    <button onClick={e => { e.stopPropagation(); handleEditReport(project.id); }}
                      className="flex items-center gap-1 text-xs font-semibold text-indigo-600 dark:text-indigo-400 hover:text-indigo-700 transition-colors px-2 py-1 rounded-md hover:bg-indigo-50 dark:hover:bg-indigo-900/20">
                      <Edit3 className="w-3 h-3" /> Edit
                    </button>
                  </div>
                </div>
              ) : (
                <button onClick={e => { e.stopPropagation(); handleEditReport(project.id); }}
                  className="flex items-center gap-1 text-xs font-semibold text-indigo-600 dark:text-indigo-400 hover:text-indigo-700 transition-colors">
                  <Plus className="w-3 h-3" /> Add Report Data
                </button>
              )}
            </div>
          );
        })}
      </div>

      {allProjects.length === 0 && (
        <div className="text-center py-16 text-gray-400 dark:text-gray-600">
          <FileBarChart className="w-12 h-12 mx-auto mb-3 opacity-50" />
          <p className="font-medium">No projects available</p>
          <p className="text-sm mt-1">Projects will appear here once they are created</p>
        </div>
      )}
    </div>
  );

  return (
    <div>
      {view === 'overview' && renderOverview()}
      {view === 'data-entry' && renderDataEntry()}
      {view === 'report-preview' && renderReportPreview()}
      {versionPanelProject && renderVersionPanel(versionPanelProject)}
      {showAIPMModal && view === 'data-entry' && editingReport && (
        <AIPMReportModal
          llmConfig={llmConfig}
          onClose={() => setShowAIPMModal(false)}
          onApply={(extracted) => {
            setEditingReport(current => current ? mergeExtractedDataIntoReport(current, extracted, { preserveExisting: true }) : current);
            setPrefillNotice('AI PM Report applied in completion mode: existing fields were preserved and only missing information was added when available.');
            setShowAIPMModal(false);
          }}
        />
      )}
    </div>
  );
};

export default PMReport;
