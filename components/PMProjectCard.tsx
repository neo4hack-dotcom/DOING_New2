import React, { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  Bot,
  Calendar,
  CheckCircle2,
  ClipboardList,
  DollarSign,
  Download,
  Eye,
  FileBarChart,
  GanttChartSquare,
  Mail,
  Plus,
  Sparkles,
  Trash2
} from 'lucide-react';
import { LLMConfig, OneOffQuery, PMGanttItem, PMReportData, Project, Team, User } from '../types';
import { generatePMProjectCardDraft, PMProjectCardLLMInput, PMProjectCardLLMOutput } from '../services/llmService';
import { generateId } from '../services/storage';

interface PMProjectCardProps {
  teams: Team[];
  users: User[];
  currentUser: User;
  llmConfig: LLMConfig;
  oneOffQueries: OneOffQuery[];
  pmReportData: PMReportData[];
  gantItems: PMGanttItem[];
}

type PMProjectCardView = 'workspace' | 'preview';
export type PMProject = Project & { teamName: string };
type CardHealth = 'Green' | 'Amber' | 'Red' | '';
type MilestoneStatus = 'On Track' | 'Delayed' | 'Completed' | '';

interface ProjectCardMilestone {
  id: string;
  description: string;
  baselineDate: string;
  forecastOrActualDate: string;
  status: MilestoneStatus;
}

interface ProjectCardRIDItem {
  id: string;
  risk: string;
  issue: string;
  dependency: string;
  mitigation: string;
}

interface ProjectCardTeamBudgetSplit {
  id: string;
  teamName: string; // Free label entered manually by PM, not necessarily a DOINg team
  totalBudgetMD: number | null;
  actualSpendMD: number | null;
  varianceMD: number | null;
  etcMD: number | null;
}

interface ProjectCardDraft {
  cardTitle: string;
  commonLink: string;
  executiveSponsor: string;
  projectManager: string;
  overallHealth: CardHealth;
  projectObjective: string;
  executiveSummary: string;
  keyAchievements: string[];
  milestones: ProjectCardMilestone[];
  totalBudgetMD: number | null;
  actualSpendMD: number | null;
  varianceMD: number | null;
  etcMD: number | null;
  teamBudgetSplits: ProjectCardTeamBudgetSplit[];
  ridItems: ProjectCardRIDItem[];
  keyDecisions: string[];
  approvalsNeeded: string[];
  resourceRequests: string[];
  escalations: string[];
}

const BRAND = '#00915A';
const BRAND_SECONDARY = '#00A86B';

const toISODate = (date: Date): string => date.toISOString().split('T')[0];
const todayISO = (): string => toISODate(new Date());

const parseDate = (value: string): Date | null => {
  if (!value) return null;
  const d = new Date(`${value}T00:00:00`);
  return Number.isNaN(d.getTime()) ? null : d;
};

const formatDate = (value: string): string => {
  const d = parseDate(value);
  return d ? d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
};

const escapeHTML = (value: string): string =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const round2 = (value: number): number => Math.round(value * 100) / 100;

const formatMD = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) return '—';
  return `${new Intl.NumberFormat('en-GB', { maximumFractionDigits: 2 }).format(value)} MD`;
};

const splitLines = (value: string): string[] =>
  value
    .split(/\r?\n/)
    .map(line => line.replace(/^[-*•]\s*/, '').trim())
    .filter(Boolean);

const uniqueStrings = (values: string[]): string[] => {
  const seen = new Set<string>();
  const out: string[] = [];
  values.forEach(raw => {
    const value = raw.trim();
    const key = value.toLowerCase();
    if (!value || seen.has(key)) return;
    seen.add(key);
    out.push(value);
  });
  return out;
};

const healthRank: Record<Exclude<CardHealth, ''>, number> = {
  Green: 1,
  Amber: 2,
  Red: 3,
};

const milestoneStatusBadgeClass = (status: MilestoneStatus): string => {
  if (status === 'Completed') return 'bg-emerald-100 text-emerald-700 border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-300 dark:border-emerald-800';
  if (status === 'Delayed') return 'bg-red-100 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800';
  if (status === 'On Track') return 'bg-sky-100 text-sky-700 border-sky-200 dark:bg-sky-900/20 dark:text-sky-300 dark:border-sky-800';
  return 'bg-slate-100 text-slate-700 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700';
};

const healthBadgeClass = (health: CardHealth): string => {
  if (health === 'Green') return 'bg-emerald-100 text-emerald-700 border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-300 dark:border-emerald-800';
  if (health === 'Amber') return 'bg-amber-100 text-amber-700 border-amber-200 dark:bg-amber-900/20 dark:text-amber-300 dark:border-amber-800';
  if (health === 'Red') return 'bg-red-100 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800';
  return 'bg-slate-100 text-slate-700 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700';
};

const createEmptyDraft = (): ProjectCardDraft => ({
  cardTitle: '',
  commonLink: '',
  executiveSponsor: '',
  projectManager: '',
  overallHealth: '',
  projectObjective: '',
  executiveSummary: '',
  keyAchievements: [],
  milestones: [],
  totalBudgetMD: null,
  actualSpendMD: null,
  varianceMD: null,
  etcMD: null,
  teamBudgetSplits: [],
  ridItems: [],
  keyDecisions: [],
  approvalsNeeded: [],
  resourceRequests: [],
  escalations: [],
});

const computeHealthFromStatus = (projects: PMProject[], reports: PMReportData[]): Exclude<CardHealth, ''> => {
  const reportHealth = reports
    .map(report => report.overallStatus)
    .filter((value): value is Exclude<CardHealth, ''> => value === 'Green' || value === 'Amber' || value === 'Red');

  if (reportHealth.length > 0) {
    return reportHealth.reduce((worst, current) => healthRank[current] > healthRank[worst] ? current : worst, 'Green');
  }

  const hasBlockedTask = projects.some(project => project.tasks.some(task => task.status === 'Blocked'));
  const hasRedDependency = projects.some(project => (project.externalDependencies || []).some(dep => dep.status === 'Red'));
  if (hasBlockedTask || hasRedDependency) return 'Red';

  const hasCaution = projects.some(project => project.status === 'Paused' || project.status === 'Planning');
  if (hasCaution) return 'Amber';

  return 'Green';
};

const normalizeFinancials = (draft: ProjectCardDraft): ProjectCardDraft => {
  const normalized: ProjectCardDraft = { ...draft };

  if (normalized.totalBudgetMD != null) normalized.totalBudgetMD = round2(normalized.totalBudgetMD);
  if (normalized.actualSpendMD != null) normalized.actualSpendMD = round2(normalized.actualSpendMD);

  if (normalized.varianceMD == null && normalized.totalBudgetMD != null && normalized.actualSpendMD != null) {
    normalized.varianceMD = round2(normalized.totalBudgetMD - normalized.actualSpendMD);
  } else if (normalized.varianceMD != null) {
    normalized.varianceMD = round2(normalized.varianceMD);
  }

  if (normalized.etcMD == null && normalized.totalBudgetMD != null && normalized.actualSpendMD != null) {
    normalized.etcMD = round2(Math.max(0, normalized.totalBudgetMD - normalized.actualSpendMD));
  } else if (normalized.etcMD != null) {
    normalized.etcMD = round2(normalized.etcMD);
  }

  normalized.teamBudgetSplits = (normalized.teamBudgetSplits || [])
    .map(split => {
      const next: ProjectCardTeamBudgetSplit = {
        id: split.id,
        teamName: (split.teamName || '').trim(),
        totalBudgetMD: split.totalBudgetMD == null ? null : round2(split.totalBudgetMD),
        actualSpendMD: split.actualSpendMD == null ? null : round2(split.actualSpendMD),
        varianceMD: split.varianceMD == null ? null : round2(split.varianceMD),
        etcMD: split.etcMD == null ? null : round2(split.etcMD),
      };

      if (next.varianceMD == null && next.totalBudgetMD != null && next.actualSpendMD != null) {
        next.varianceMD = round2(next.totalBudgetMD - next.actualSpendMD);
      }
      if (next.etcMD == null && next.totalBudgetMD != null && next.actualSpendMD != null) {
        next.etcMD = round2(Math.max(0, next.totalBudgetMD - next.actualSpendMD));
      }

      return next;
    })
    .filter(split => split.teamName.length > 0);

  return normalized;
};

const buildDeterministicDraft = (
  projects: PMProject[],
  users: User[],
  reports: PMReportData[],
  roadmapItems: PMGanttItem[],
  oneOffQueries: OneOffQuery[],
  currentUser: User,
  currentCommonLink: string
): ProjectCardDraft => {
  if (projects.length === 0) return createEmptyDraft();

  const projectCount = projects.length;
  const totalTasks = projects.reduce((acc, project) => acc + project.tasks.length, 0);
  const doneTasks = projects.reduce((acc, project) => acc + project.tasks.filter(task => task.status === 'Done').length, 0);
  const inProgressTasks = projects.reduce((acc, project) => acc + project.tasks.filter(task => task.status === 'In Progress').length, 0);
  const blockedTasks = projects.reduce((acc, project) => acc + project.tasks.filter(task => task.status === 'Blocked').length, 0);

  const managerNames = uniqueStrings(
    projects.map(project => {
      const manager = users.find(user => user.id === project.managerId);
      return manager ? `${manager.firstName} ${manager.lastName}`.trim() : '';
    })
  );
  const sponsors = uniqueStrings(projects.map(project => project.owner || ''));

  const health = computeHealthFromStatus(projects, reports);

  const reportNewsAchievements = reports
    .flatMap(report => report.news || [])
    .filter(news => news.type === 'Achievement')
    .map(news => news.title || news.description)
    .filter(Boolean);

  const doneTaskAchievements = projects
    .flatMap(project => project.tasks)
    .filter(task => task.status === 'Done')
    .map(task => task.title)
    .filter(Boolean);

  const roadmapAchievements = roadmapItems
    .filter(item => item.status === 'Done' || item.progressPct >= 100)
    .map(item => item.title)
    .filter(Boolean);

  const keyAchievements = uniqueStrings([
    ...doneTaskAchievements,
    ...reportNewsAchievements,
    ...roadmapAchievements,
  ]).slice(0, 10);

  const reportMilestones: ProjectCardMilestone[] = reports
    .flatMap(report => report.milestones || [])
    .map(milestone => {
      const status: MilestoneStatus = milestone.completionPct >= 100
        ? 'Completed'
        : milestone.status === 'Red'
          ? 'Delayed'
          : 'On Track';

      return {
        id: generateId(),
        description: milestone.name,
        baselineDate: milestone.plannedDate || '',
        forecastOrActualDate: milestone.revisedDate || milestone.plannedDate || '',
        status,
      };
    })
    .filter(item => item.description.trim().length > 0);

  const roadmapMilestones: ProjectCardMilestone[] = roadmapItems
    .filter(item => item.isMilestone)
    .map(item => ({
      id: generateId(),
      description: item.title,
      baselineDate: item.startDate || '',
      forecastOrActualDate: item.endDate || '',
      status: item.status === 'Done' || item.progressPct >= 100
        ? 'Completed'
        : item.status === 'Blocked'
          ? 'Delayed'
          : 'On Track',
    }))
    .filter(item => item.description.trim().length > 0);

  const projectDeadlineMilestones: ProjectCardMilestone[] = projects.map(project => ({
    id: generateId(),
    description: `${project.name} deadline`,
    baselineDate: project.deadline || '',
    forecastOrActualDate: project.deadline || '',
    status: project.status === 'Done' ? 'Completed' : (project.status === 'Paused' ? 'Delayed' : 'On Track'),
  }));

  const milestoneMap = new Map<string, ProjectCardMilestone>();
  [...reportMilestones, ...roadmapMilestones, ...projectDeadlineMilestones].forEach(milestone => {
    const key = milestone.description.trim().toLowerCase();
    if (!key || milestoneMap.has(key)) return;
    milestoneMap.set(key, milestone);
  });

  const milestones = Array.from(milestoneMap.values())
    .sort((a, b) => {
      const ad = parseDate(a.forecastOrActualDate || a.baselineDate)?.getTime() || Number.MAX_SAFE_INTEGER;
      const bd = parseDate(b.forecastOrActualDate || b.baselineDate)?.getTime() || Number.MAX_SAFE_INTEGER;
      return ad - bd;
    })
    .slice(0, 12);

  const risksFromReport = reports.flatMap(report => (report.risks || []).map(risk => ({
    id: generateId(),
    risk: risk.description,
    issue: '',
    dependency: '',
    mitigation: risk.mitigation,
  })));

  const incidentsFromReport = reports
    .flatMap(report => report.incidents || [])
    .filter(incident => incident.status !== 'Resolved')
    .map(incident => ({
      id: generateId(),
      risk: '',
      issue: `${incident.title}${incident.description ? `: ${incident.description}` : ''}`,
      dependency: '',
      mitigation: incident.status ? `Current status: ${incident.status}` : '',
    }));

  const blockedTaskItems = projects
    .flatMap(project => project.tasks)
    .filter(task => task.status === 'Blocked')
    .map(task => ({
      id: generateId(),
      risk: '',
      issue: `Blocked task: ${task.title}`,
      dependency: '',
      mitigation: '',
    }));

  const dependencyItems = projects
    .flatMap(project => project.externalDependencies || [])
    .filter(dep => dep.status !== 'Green')
    .map(dep => ({
      id: generateId(),
      risk: '',
      issue: '',
      dependency: `${dep.label} (${dep.status})`,
      mitigation: '',
    }));

  const oneOffRiskItems = oneOffQueries
    .filter(query => query.status === 'pending' || query.status === 'in_progress')
    .map(query => ({
      id: generateId(),
      risk: '',
      issue: `One-off query pending: ${query.title}`,
      dependency: '',
      mitigation: query.etaRequested ? `Requested ETA: ${query.etaRequested}` : '',
    }));

  const ridDedup = new Map<string, ProjectCardRIDItem>();
  [...risksFromReport, ...incidentsFromReport, ...blockedTaskItems, ...dependencyItems, ...oneOffRiskItems].forEach(item => {
    const key = `${item.risk}|${item.issue}|${item.dependency}|${item.mitigation}`.trim().toLowerCase();
    if (!key || ridDedup.has(key)) return;
    ridDedup.set(key, item);
  });

  const ridItems = Array.from(ridDedup.values()).slice(0, 12);

  const reportDecisionLines = reports.flatMap(report => splitLines(report.keyDecisions || ''));
  const keyDecisions = uniqueStrings(reportDecisionLines).slice(0, 10);

  const approvalsNeeded = uniqueStrings(
    keyDecisions.filter(line => /approve|approval|sign[- ]?off|validation/i.test(line))
  ).slice(0, 10);

  const resourceRequests = uniqueStrings(
    [
      ...keyDecisions.filter(line => /resource|fte|staff|capacity|headcount/i.test(line)),
      ...oneOffQueries
        .filter(query => query.status !== 'done' && query.status !== 'cancelled' && query.estimatedCostMD != null && query.estimatedCostMD >= 1)
        .map(query => `Resource attention required for one-off query: ${query.title}`)
    ]
  ).slice(0, 10);

  const escalations = uniqueStrings(
    reports
      .flatMap(report => report.incidents || [])
      .filter(incident => incident.severity === 'Critical' && incident.status !== 'Resolved')
      .map(incident => `Escalation: ${incident.title}`)
      .concat(blockedTaskItems.map(item => item.issue))
  ).slice(0, 10);

  const projectDescriptions = projects.map(project => project.description).filter(Boolean);
  const objective = projectDescriptions.length > 0
    ? projectDescriptions[0]
    : '';

  const summaryParts = [
    `${projectCount} project${projectCount > 1 ? 's' : ''} selected`,
    `${totalTasks} task${totalTasks > 1 ? 's' : ''} tracked`,
    `${doneTasks} done`,
    `${inProgressTasks} in progress`,
    `${blockedTasks} blocked`,
  ];

  const budgetFromReports = reports.reduce((acc, report) => acc + (report.budgetAllocated || 0), 0);
  const budgetFromProjects = projects.reduce((acc, project) => acc + (project.cost || 0), 0);
  const totalBudgetMD = budgetFromReports > 0 ? budgetFromReports : (budgetFromProjects > 0 ? budgetFromProjects : null);

  const actualSpendFromReports = reports.reduce((acc, report) => acc + (report.budgetSpent || 0), 0);
  const actualSpendFromOneOff = oneOffQueries.reduce((acc, query) => acc + (query.finalCostMD || 0), 0);
  const actualSpendMD = actualSpendFromReports > 0
    ? actualSpendFromReports
    : (actualSpendFromOneOff > 0 ? actualSpendFromOneOff : null);

  const forecastFromReports = reports.reduce((acc, report) => acc + (report.budgetForecast || 0), 0);

  const varianceMD = totalBudgetMD != null && actualSpendMD != null
    ? totalBudgetMD - actualSpendMD
    : null;

  const etcMD = forecastFromReports > 0 && actualSpendMD != null
    ? Math.max(0, forecastFromReports - actualSpendMD)
    : (totalBudgetMD != null && actualSpendMD != null ? Math.max(0, totalBudgetMD - actualSpendMD) : null);

  const defaultTitle = projectCount === 1
    ? `${projects[0].name} - Project Card`
    : `DOINg - SteerCo Project Card (${projectCount} Projects)`;

  const splitMap = new Map<string, { teamName: string; total: number; spent: number; forecast: number }>();
  reports.forEach(report => {
    (report.costDistribution || []).forEach(split => {
      const teamName = (split.teamName || split.teamId || '').trim();
      if (!teamName) return;
      const entry = splitMap.get(teamName.toLowerCase()) || { teamName, total: 0, spent: 0, forecast: 0 };
      entry.total += split.allocatedMD || 0;
      entry.spent += split.spentMD || 0;
      entry.forecast += split.forecastMD || 0;
      splitMap.set(teamName.toLowerCase(), entry);
    });
  });

  const teamBudgetSplits: ProjectCardTeamBudgetSplit[] = Array.from(splitMap.values())
    .sort((a, b) => a.teamName.localeCompare(b.teamName))
    .map(entry => {
      const variance = entry.total - entry.spent;
      const etc = entry.forecast > 0
        ? Math.max(0, entry.forecast - entry.spent)
        : Math.max(0, entry.total - entry.spent);

      return {
        id: generateId(),
        teamName: entry.teamName,
        totalBudgetMD: round2(entry.total),
        actualSpendMD: round2(entry.spent),
        varianceMD: round2(variance),
        etcMD: round2(etc),
      };
    });

  const draft: ProjectCardDraft = {
    cardTitle: defaultTitle,
    commonLink: currentCommonLink || '',
    executiveSponsor: sponsors.join(' / '),
    projectManager: managerNames.length > 0 ? managerNames.join(' / ') : `${currentUser.firstName} ${currentUser.lastName}`.trim(),
    overallHealth: health,
    projectObjective: objective,
    executiveSummary: `${summaryParts.join(', ')}.`,
    keyAchievements,
    milestones,
    totalBudgetMD: totalBudgetMD == null ? null : round2(totalBudgetMD),
    actualSpendMD: actualSpendMD == null ? null : round2(actualSpendMD),
    varianceMD: varianceMD == null ? null : round2(varianceMD),
    etcMD: etcMD == null ? null : round2(etcMD),
    teamBudgetSplits,
    ridItems,
    keyDecisions,
    approvalsNeeded,
    resourceRequests,
    escalations,
  };

  return normalizeFinancials(draft);
};

const mergeDrafts = (base: ProjectCardDraft, llm: PMProjectCardLLMOutput): ProjectCardDraft => {
  const merged: ProjectCardDraft = {
    ...base,
    cardTitle: llm.cardTitle || base.cardTitle,
    commonLink: llm.commonLink || base.commonLink,
    executiveSponsor: llm.executiveSponsor || base.executiveSponsor,
    projectManager: llm.projectManager || base.projectManager,
    overallHealth: llm.overallHealth || base.overallHealth,
    projectObjective: llm.projectObjective || base.projectObjective,
    executiveSummary: llm.executiveSummary || base.executiveSummary,
    keyAchievements: llm.keyAchievements.length > 0 ? llm.keyAchievements : base.keyAchievements,
    milestones: llm.milestones.length > 0
      ? llm.milestones.map(item => ({
          id: generateId(),
          description: item.description,
          baselineDate: item.baselineDate,
          forecastOrActualDate: item.forecastOrActualDate,
          status: item.status,
        }))
      : base.milestones,
    totalBudgetMD: llm.financials.totalBudgetMD != null ? llm.financials.totalBudgetMD : base.totalBudgetMD,
    actualSpendMD: llm.financials.actualSpendMD != null ? llm.financials.actualSpendMD : base.actualSpendMD,
    varianceMD: llm.financials.varianceMD != null ? llm.financials.varianceMD : base.varianceMD,
    etcMD: llm.financials.etcMD != null ? llm.financials.etcMD : base.etcMD,
    ridItems: llm.ridItems.length > 0
      ? llm.ridItems.map(item => ({ id: generateId(), ...item }))
      : base.ridItems,
    keyDecisions: llm.keyDecisions.length > 0 ? llm.keyDecisions : base.keyDecisions,
    approvalsNeeded: llm.approvalsNeeded.length > 0 ? llm.approvalsNeeded : base.approvalsNeeded,
    resourceRequests: llm.resourceRequests.length > 0 ? llm.resourceRequests : base.resourceRequests,
    escalations: llm.escalations.length > 0 ? llm.escalations : base.escalations,
  };

  return normalizeFinancials(merged);
};

const buildProjectCardHTML = (
  draftInput: ProjectCardDraft,
  projects: PMProject[],
  linkedOneOffCount: number
): string => {
  const draft = normalizeFinancials(draftInput);
  const generatedAt = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  const projectNames = projects.map(project => project.name);

  const achievements = draft.keyAchievements.length > 0 ? draft.keyAchievements : ['No explicit achievements captured yet.'];
  const milestones = draft.milestones.length > 0
    ? draft.milestones
    : [{ id: generateId(), description: 'No milestone explicitly available', baselineDate: '', forecastOrActualDate: '', status: '' as MilestoneStatus }];
  const ridItems = draft.ridItems.length > 0
    ? draft.ridItems
    : [{ id: generateId(), risk: '', issue: 'No explicit risk/issue/dependency captured.', dependency: '', mitigation: '' }];
  const teamBudgetSplits = draft.teamBudgetSplits.length > 0 ? draft.teamBudgetSplits : [];

  const denseMode = achievements.length + milestones.length + ridItems.length + teamBudgetSplits.length + draft.keyDecisions.length + draft.approvalsNeeded.length + draft.resourceRequests.length + draft.escalations.length > 20;

  const healthClass = draft.overallHealth === 'Green'
    ? 'health-green'
    : draft.overallHealth === 'Amber'
      ? 'health-amber'
      : draft.overallHealth === 'Red'
        ? 'health-red'
        : 'health-neutral';

  const css = `
  <style>
    .page{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;color:#0f172a;max-width:1120px;margin:0 auto}
    .page *{box-sizing:border-box}
    .header{background:linear-gradient(135deg,${BRAND} 0%,${BRAND_SECONDARY} 60%,#007A4C 100%);color:#fff;border-radius:16px;padding:18px 20px 16px 20px;box-shadow:0 8px 20px rgba(0,145,90,.25);margin-bottom:14px;page-break-inside:avoid;break-inside:avoid-page}
    .title{font-size:26px;font-weight:800;line-height:1.15;margin:0}
    .sub{font-size:12px;opacity:.9;margin-top:6px}
    .chips{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px}
    .chip{font-size:10px;font-weight:700;border:1px solid rgba(255,255,255,.35);background:rgba(255,255,255,.12);padding:3px 8px;border-radius:999px}
    .meta-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-top:14px}
    .meta{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);border-radius:10px;padding:9px 10px}
    .meta .label{font-size:10px;opacity:.85;text-transform:uppercase;letter-spacing:.6px;font-weight:700}
    .meta .value{font-size:13px;font-weight:700;margin-top:2px;line-height:1.3}
    .health-green{background:#dcfce7;color:#166534;border-color:#86efac}
    .health-amber{background:#fef3c7;color:#92400e;border-color:#fcd34d}
    .health-red{background:#fee2e2;color:#991b1b;border-color:#fca5a5}
    .health-neutral{background:#e2e8f0;color:#334155;border-color:#cbd5e1}
    .section{border:1px solid #dbe5ef;border-radius:12px;background:#fff;padding:12px 14px;margin-bottom:12px;page-break-inside:avoid;break-inside:avoid-page}
    .section h2{margin:0 0 8px 0;font-size:14px;font-weight:800;color:#0f172a}
    .objective{font-size:12px;line-height:1.5;color:#334155;background:#f8fafc;border-left:4px solid ${BRAND};padding:10px;border-radius:8px}
    .summary{font-size:12px;line-height:1.55;color:#334155;margin-top:8px}
    .list{margin:8px 0 0 0;padding-left:18px}
    .list li{font-size:11px;line-height:1.5;margin-bottom:4px}
    .grid-2{display:grid;grid-template-columns:1.15fr .85fr;gap:12px}
    .table{width:100%;border-collapse:collapse;margin-top:6px}
    .table th{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;background:#f8fafc;border-bottom:1px solid #e2e8f0;padding:6px;text-align:left}
    .table td{font-size:11px;color:#1e293b;border-bottom:1px solid #f1f5f9;padding:7px 6px;vertical-align:top}
    .badge{display:inline-flex;padding:2px 8px;border-radius:999px;border:1px solid #cbd5e1;font-size:10px;font-weight:700}
    .on-track{background:#dcfce7;color:#166534;border-color:#86efac}
    .delayed{background:#fee2e2;color:#991b1b;border-color:#fca5a5}
    .completed{background:#dbeafe;color:#1d4ed8;border-color:#93c5fd}
    .neutral{background:#e2e8f0;color:#334155;border-color:#cbd5e1}
    .kpi-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;margin-top:8px}
    .kpi{border:1px solid #dbe5ef;border-radius:10px;padding:8px;background:#f8fafc}
    .kpi .label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;font-weight:700}
    .kpi .value{font-size:14px;color:#0f172a;font-weight:800;margin-top:3px}
    .split-title{margin-top:12px;font-size:11px;font-weight:800;color:#0f172a}
    .split-note{margin-top:4px;font-size:10px;color:#64748b}
    .ask-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}
    .ask{border:1px solid #dbe5ef;border-radius:10px;padding:8px;background:#fff}
    .ask h3{margin:0 0 6px 0;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:#334155}
    .ask ul{margin:0;padding-left:16px}
    .ask li{font-size:11px;line-height:1.45;margin-bottom:3px}
    .footer{font-size:10px;color:#64748b;margin-top:10px;text-align:right}
    .dense .title{font-size:23px}
    .dense .section{padding:10px 12px;margin-bottom:10px}
    .dense .table th{font-size:9px;padding:5px}
    .dense .table td{font-size:10px;padding:5px}
    .dense .list li,.dense .ask li{font-size:10px}
    @media print {
      .page{max-width:none}
      .header,.section,.ask{break-inside:avoid-page;page-break-inside:avoid}
    }
  </style>`;

  const milestoneRows = milestones.map(milestone => {
    const statusClass = milestone.status === 'Completed'
      ? 'completed'
      : milestone.status === 'Delayed'
        ? 'delayed'
        : milestone.status === 'On Track'
          ? 'on-track'
          : 'neutral';

    return `
      <tr>
        <td>${escapeHTML(milestone.description)}</td>
        <td>${escapeHTML(formatDate(milestone.baselineDate))}</td>
        <td>${escapeHTML(formatDate(milestone.forecastOrActualDate))}</td>
        <td><span class="badge ${statusClass}">${escapeHTML(milestone.status || 'N/A')}</span></td>
      </tr>`;
  }).join('');

  const ridRows = ridItems.map(item => `
    <tr>
      <td>${escapeHTML(item.risk || '—')}</td>
      <td>${escapeHTML(item.issue || '—')}</td>
      <td>${escapeHTML(item.dependency || '—')}</td>
      <td>${escapeHTML(item.mitigation || '—')}</td>
    </tr>
  `).join('');

  const teamBudgetRows = teamBudgetSplits.map(split => `
    <tr>
      <td>${escapeHTML(split.teamName)}</td>
      <td>${escapeHTML(formatMD(split.totalBudgetMD))}</td>
      <td>${escapeHTML(formatMD(split.actualSpendMD))}</td>
      <td>${escapeHTML(formatMD(split.varianceMD))}</td>
      <td>${escapeHTML(formatMD(split.etcMD))}</td>
    </tr>
  `).join('');

  const listOrFallback = (items: string[], fallback: string) => {
    if (items.length === 0) return `<li>${escapeHTML(fallback)}</li>`;
    return items.map(item => `<li>${escapeHTML(item)}</li>`).join('');
  };

  return `${css}
  <div class="page ${denseMode ? 'dense' : ''}">
    <header class="header">
      <h1 class="title">${escapeHTML(draft.cardTitle || 'DOINg - Project Card')}</h1>
      <div class="sub">Generated ${escapeHTML(generatedAt)} • ${projects.length} selected project${projects.length > 1 ? 's' : ''}${linkedOneOffCount > 0 ? ` • ${linkedOneOffCount} linked one-off quer${linkedOneOffCount > 1 ? 'ies' : 'y'} included` : ''}</div>
      <div class="chips">${projectNames.map(name => `<span class="chip">${escapeHTML(name)}</span>`).join('')}</div>
      <div class="meta-grid">
        <div class="meta"><div class="label">Executive Sponsor</div><div class="value">${escapeHTML(draft.executiveSponsor || '—')}</div></div>
        <div class="meta"><div class="label">Project Manager</div><div class="value">${escapeHTML(draft.projectManager || '—')}</div></div>
        <div class="meta"><div class="label">Common Link</div><div class="value">${escapeHTML(draft.commonLink || '—')}</div></div>
        <div class="meta ${healthClass}"><div class="label">Overall Health</div><div class="value">${escapeHTML(draft.overallHealth || 'N/A')}</div></div>
      </div>
    </header>

    <section class="section">
      <h2>Executive Summary</h2>
      <div class="objective"><strong>Project Objective:</strong> ${escapeHTML(draft.projectObjective || 'Not explicitly provided in source data.')}</div>
      <div class="summary">${escapeHTML(draft.executiveSummary || 'No executive summary available yet.')}</div>
      <ul class="list">${listOrFallback(achievements, 'No explicit achievement identified yet.')}</ul>
    </section>

    <div class="grid-2">
      <section class="section">
        <h2>Schedule & Milestones</h2>
        <table class="table">
          <thead>
            <tr>
              <th>Milestone</th>
              <th>Baseline Date</th>
              <th>Forecast / Actual</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>${milestoneRows}</tbody>
        </table>
      </section>

      <section class="section">
        <h2>Financials (MD)</h2>
        <div class="kpi-grid">
          <div class="kpi"><div class="label">Total Budget</div><div class="value">${escapeHTML(formatMD(draft.totalBudgetMD))}</div></div>
          <div class="kpi"><div class="label">Actual Spend</div><div class="value">${escapeHTML(formatMD(draft.actualSpendMD))}</div></div>
          <div class="kpi"><div class="label">Variance</div><div class="value">${escapeHTML(formatMD(draft.varianceMD))}</div></div>
          <div class="kpi"><div class="label">ETC</div><div class="value">${escapeHTML(formatMD(draft.etcMD))}</div></div>
        </div>
        ${teamBudgetSplits.length > 0 ? `
          <div class="split-title">Manual Team Cost Distribution</div>
          <div class="split-note">Team labels are free text and may differ from DOINg team entities.</div>
          <table class="table">
            <thead>
              <tr>
                <th>Team Name</th>
                <th>Budget (MD)</th>
                <th>Spent (MD)</th>
                <th>Variance (MD)</th>
                <th>ETC (MD)</th>
              </tr>
            </thead>
            <tbody>${teamBudgetRows}</tbody>
          </table>
        ` : ''}
      </section>
    </div>

    <section class="section">
      <h2>Risks, Issues & Dependencies (RID)</h2>
      <table class="table">
        <thead>
          <tr>
            <th>Risk</th>
            <th>Issue</th>
            <th>Dependency</th>
            <th>Mitigation Plan</th>
          </tr>
        </thead>
        <tbody>${ridRows}</tbody>
      </table>
    </section>

    <section class="section">
      <h2>Key Decisions & Asks</h2>
      <div class="ask-grid">
        <div class="ask">
          <h3>Key Decisions</h3>
          <ul>${listOrFallback(draft.keyDecisions, 'No decision explicitly captured.')}</ul>
        </div>
        <div class="ask">
          <h3>Approvals Needed</h3>
          <ul>${listOrFallback(draft.approvalsNeeded, 'No specific approval explicitly identified.')}</ul>
        </div>
        <div class="ask">
          <h3>Resource Requests</h3>
          <ul>${listOrFallback(draft.resourceRequests, 'No explicit resource request identified.')}</ul>
        </div>
        <div class="ask">
          <h3>Escalations</h3>
          <ul>${listOrFallback(draft.escalations, 'No escalation currently identified.')}</ul>
        </div>
      </div>
    </section>

    <div class="footer">DOINg • Steering Committee Project Card</div>
  </div>`;
};

export const buildPMProjectCardHTMLFromSelection = (
  projects: PMProject[],
  users: User[],
  reports: PMReportData[],
  roadmapItems: PMGanttItem[],
  oneOffQueries: OneOffQuery[],
  currentUser: User,
  options?: {
    commonLink?: string;
    includeOneOffQueries?: boolean;
  }
): string => {
  const includeOneOff = options?.includeOneOffQueries !== false;
  const scopedQueries = includeOneOff ? oneOffQueries : [];
  const draft = buildDeterministicDraft(
    projects,
    users,
    reports,
    roadmapItems,
    scopedQueries,
    currentUser,
    options?.commonLink || ''
  );
  return buildProjectCardHTML(
    draft,
    projects,
    scopedQueries.length
  );
};

const PMProjectCard: React.FC<PMProjectCardProps> = ({
  teams,
  users,
  currentUser,
  llmConfig,
  oneOffQueries,
  pmReportData,
  gantItems,
}) => {
  const [view, setView] = useState<PMProjectCardView>('workspace');
  const [selectedProjectIds, setSelectedProjectIds] = useState<string[]>([]);
  const [includeOneOffQueries, setIncludeOneOffQueries] = useState(true);
  const [manualContext, setManualContext] = useState('');
  const [draft, setDraft] = useState<ProjectCardDraft>(createEmptyDraft());
  const [generatedHTML, setGeneratedHTML] = useState('');
  const [isAIMapping, setIsAIMapping] = useState(false);
  const [isGeneratingDocument, setIsGeneratingDocument] = useState(false);
  const [aiError, setAiError] = useState('');
  const [aiInfo, setAiInfo] = useState('');

  const inputClass = 'w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100';
  const labelClass = 'block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1';

  const allProjects = useMemo<PMProject[]>(() => {
    const list: PMProject[] = [];
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
    const map = new Map<string, PMProject>();
    allProjects.forEach(project => map.set(project.id, project));
    return map;
  }, [allProjects]);

  const visibleTeamIds = useMemo(() => new Set(teams.map(team => team.id)), [teams]);
  const visibleProjectIds = useMemo(() => new Set(allProjects.map(project => project.id)), [allProjects]);

  const scopedOneOffQueries = useMemo(() => {
    return (oneOffQueries || []).filter(query => {
      if (!visibleTeamIds.has(query.teamId)) return false;
      if (!query.projectId) return true;
      return visibleProjectIds.has(query.projectId);
    });
  }, [oneOffQueries, visibleTeamIds, visibleProjectIds]);

  const latestReportByProject = useMemo(() => {
    const map = new Map<string, PMReportData>();
    (pmReportData || []).forEach(report => {
      if (!visibleProjectIds.has(report.projectId)) return;
      const current = map.get(report.projectId);
      if (!current) {
        map.set(report.projectId, report);
        return;
      }
      const currentUpdated = new Date(current.updatedAt).getTime();
      const candidateUpdated = new Date(report.updatedAt).getTime();
      if (report.version > current.version || (report.version === current.version && candidateUpdated > currentUpdated)) {
        map.set(report.projectId, report);
      }
    });
    return map;
  }, [pmReportData, visibleProjectIds]);

  useEffect(() => {
    setSelectedProjectIds(prev => prev.filter(id => projectById.has(id)));
  }, [projectById]);

  useEffect(() => {
    if (selectedProjectIds.length === 0 && allProjects.length > 0) {
      setSelectedProjectIds([allProjects[0].id]);
    }
  }, [allProjects, selectedProjectIds.length]);

  const selectedProjects = useMemo(() => {
    return allProjects.filter(project => selectedProjectIds.includes(project.id));
  }, [allProjects, selectedProjectIds]);

  const selectedReports = useMemo(() => {
    return selectedProjects
      .map(project => latestReportByProject.get(project.id))
      .filter((report): report is PMReportData => Boolean(report));
  }, [selectedProjects, latestReportByProject]);

  const selectedRoadmapItems = useMemo(() => {
    return (gantItems || []).filter(item => selectedProjectIds.includes(item.projectId));
  }, [gantItems, selectedProjectIds]);

  const linkedSelectedOneOffQueries = useMemo(() => {
    return scopedOneOffQueries.filter(query => query.projectId && selectedProjectIds.includes(query.projectId));
  }, [scopedOneOffQueries, selectedProjectIds]);

  const sourceBundle = useMemo<PMProjectCardLLMInput>(() => {
    return {
      projects: selectedProjects.map(project => ({
        id: project.id,
        name: project.name,
        teamName: project.teamName,
        status: project.status,
        deadline: project.deadline,
        description: project.description || '',
        owner: project.owner || '',
        managerName: (() => {
          const manager = users.find(user => user.id === project.managerId);
          return manager ? `${manager.firstName} ${manager.lastName}`.trim() : '';
        })(),
        costMD: project.cost ?? null,
        tasks: project.tasks.map(task => ({
          title: task.title,
          status: task.status,
          priority: task.priority,
          eta: task.eta,
          assigneeName: (() => {
            const assignee = users.find(user => user.id === task.assigneeId);
            return assignee ? `${assignee.firstName} ${assignee.lastName}`.trim() : '';
          })(),
          description: task.description,
        })),
        externalDependencies: (project.externalDependencies || []).map(dep => ({
          label: dep.label,
          status: dep.status,
        })),
      })),
      pmReports: selectedReports.map(report => ({
        projectId: report.projectId,
        version: report.version,
        overallStatus: report.overallStatus,
        executiveSummary: report.executiveSummary,
        keyDecisions: report.keyDecisions,
        nextSteps: report.nextSteps,
        budgetAllocated: report.budgetAllocated ?? null,
        budgetSpent: report.budgetSpent ?? null,
        budgetForecast: report.budgetForecast ?? null,
        incidents: (report.incidents || []).map(incident => ({
          title: incident.title,
          severity: incident.severity,
          status: incident.status,
          description: incident.description,
        })),
        risks: (report.risks || []).map(risk => ({
          description: risk.description,
          likelihood: risk.likelihood,
          impact: risk.impact,
          mitigation: risk.mitigation,
          owner: risk.owner,
        })),
        milestones: (report.milestones || []).map(milestone => ({
          name: milestone.name,
          plannedDate: milestone.plannedDate,
          revisedDate: milestone.revisedDate || '',
          status: milestone.status,
          completionPct: milestone.completionPct,
        })),
      })),
      roadmapItems: selectedRoadmapItems.map(item => ({
        projectId: item.projectId,
        title: item.title,
        owner: item.owner,
        startDate: item.startDate,
        endDate: item.endDate,
        status: item.status,
        progressPct: item.progressPct,
        priority: item.priority,
        isMilestone: item.isMilestone === true,
      })),
      oneOffQueries: (includeOneOffQueries ? linkedSelectedOneOffQueries : []).map(query => ({
        id: query.id,
        projectId: query.projectId || null,
        title: query.title,
        requester: query.requester,
        sponsor: query.sponsor,
        status: query.status,
        etaRequested: query.etaRequested || '',
        description: query.description,
        estimatedCostMD: query.estimatedCostMD ?? null,
        finalCostMD: query.finalCostMD ?? null,
      })),
      commonLinkHint: draft.commonLink,
      manualContext,
    };
  }, [
    selectedProjects,
    selectedReports,
    selectedRoadmapItems,
    includeOneOffQueries,
    linkedSelectedOneOffQueries,
    users,
    draft.commonLink,
    manualContext,
  ]);

  useEffect(() => {
    if (selectedProjects.length === 0) return;
    if (draft.cardTitle) return;
    const nextDraft = buildDeterministicDraft(
      selectedProjects,
      users,
      selectedReports,
      selectedRoadmapItems,
      includeOneOffQueries ? linkedSelectedOneOffQueries : [],
      currentUser,
      draft.commonLink
    );
    setDraft(nextDraft);
  }, [
    selectedProjects,
    users,
    selectedReports,
    selectedRoadmapItems,
    includeOneOffQueries,
    linkedSelectedOneOffQueries,
    currentUser,
    draft.cardTitle,
    draft.commonLink,
  ]);

  const toggleProject = (projectId: string) => {
    setSelectedProjectIds(prev =>
      prev.includes(projectId) ? prev.filter(id => id !== projectId) : [...prev, projectId]
    );
  };

  const handleSelectAll = () => setSelectedProjectIds(allProjects.map(project => project.id));
  const handleClearSelection = () => setSelectedProjectIds([]);

  const handleRebuildFromData = () => {
    if (selectedProjects.length === 0) {
      alert('Select at least one project.');
      return;
    }

    const nextDraft = buildDeterministicDraft(
      selectedProjects,
      users,
      selectedReports,
      selectedRoadmapItems,
      includeOneOffQueries ? linkedSelectedOneOffQueries : [],
      currentUser,
      draft.commonLink
    );
    setDraft(nextDraft);
    setAiError('');
    setAiInfo('Draft rebuilt from existing DOINg data. You can now amend fields manually or run AI mapping.');
  };

  const handleAIMapping = async () => {
    if (selectedProjects.length === 0) {
      setAiError('Select at least one project before running AI mapping.');
      return;
    }

    setIsAIMapping(true);
    setAiError('');
    setAiInfo('');

    const baseDraft = buildDeterministicDraft(
      selectedProjects,
      users,
      selectedReports,
      selectedRoadmapItems,
      includeOneOffQueries ? linkedSelectedOneOffQueries : [],
      currentUser,
      draft.commonLink
    );

    try {
      const aiDraft = await generatePMProjectCardDraft(sourceBundle, llmConfig);
      const merged = mergeDrafts(baseDraft, aiDraft);
      setDraft(prev => ({
        ...merged,
        commonLink: merged.commonLink || prev.commonLink,
      }));
      setAiInfo('AI mapping completed. Review and edit all fields before generating the final one-pager.');
    } catch (e: any) {
      setDraft(baseDraft);
      setAiError(e?.message || 'AI mapping failed. Deterministic draft has been kept.');
    } finally {
      setIsAIMapping(false);
    }
  };

  const handleGenerateOnePager = () => {
    if (selectedProjects.length === 0) {
      alert('Select at least one project.');
      return;
    }

    setIsGeneratingDocument(true);
    try {
      const html = buildProjectCardHTML(
        normalizeFinancials(draft),
        selectedProjects,
        includeOneOffQueries ? linkedSelectedOneOffQueries.length : 0
      );
      setGeneratedHTML(html);
      setView('preview');
    } finally {
      setIsGeneratingDocument(false);
    }
  };

  const buildExportDocumentHTML = (withPrintControls: boolean) => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>DOINg - Project Card</title>
<style>
@page{size:A4 portrait;margin:9mm}
@media print{
  body{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}
  .no-print{display:none!important}
  .page,.section,.header{break-inside:avoid-page;page-break-inside:avoid}
}
</style>
</head><body style="margin:0;padding:18px;background:#fff;font-family:'Segoe UI',system-ui,-apple-system,sans-serif">
${withPrintControls ? `
<div class="no-print" style="text-align:center;margin-bottom:18px;padding:14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px">
  <button onclick="window.print()" style="background:linear-gradient(135deg,${BRAND},${BRAND_SECONDARY});color:#fff;border:none;padding:12px 32px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer">Print / Save as PDF</button>
  <button onclick="window.close()" style="margin-left:10px;background:#fff;color:#334155;border:1px solid #cbd5e1;padding:12px 22px;border-radius:10px;font-size:14px;cursor:pointer">Cancel</button>
</div>` : ''}
${generatedHTML}
</body></html>`;

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
    const emlContent = [
      'X-Unsent: 1',
      'To: ',
      `Subject: DOINg - Project Card - ${subjectDate}`,
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
      'Content-Transfer-Encoding: 8bit',
      '',
      buildExportDocumentHTML(false)
    ].join('\r\n');

    const blob = new Blob([emlContent], { type: 'message/rfc822;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `doing-project-card-${fileDate}.eml`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const updateNumberField = (field: 'totalBudgetMD' | 'actualSpendMD' | 'varianceMD' | 'etcMD', value: string) => {
    const cleaned = value.trim();
    if (!cleaned) {
      setDraft(prev => ({ ...prev, [field]: null }));
      return;
    }
    const parsed = parseFloat(cleaned.replace(',', '.'));
    setDraft(prev => ({ ...prev, [field]: Number.isFinite(parsed) ? parsed : null }));
  };

  const updateListField = (
    field: 'keyDecisions' | 'approvalsNeeded' | 'resourceRequests' | 'escalations',
    value: string
  ) => {
    setDraft(prev => ({ ...prev, [field]: splitLines(value) }));
  };

  const updateTeamBudgetSplitName = (splitId: string, teamName: string) => {
    setDraft(prev => ({
      ...prev,
      teamBudgetSplits: prev.teamBudgetSplits.map(split =>
        split.id === splitId ? { ...split, teamName } : split
      )
    }));
  };

  const updateTeamBudgetSplitNumber = (
    splitId: string,
    field: 'totalBudgetMD' | 'actualSpendMD' | 'varianceMD' | 'etcMD',
    value: string
  ) => {
    const cleaned = value.trim();
    const parsed = cleaned ? parseFloat(cleaned.replace(',', '.')) : null;
    setDraft(prev => ({
      ...prev,
      teamBudgetSplits: prev.teamBudgetSplits.map(split =>
        split.id === splitId
          ? { ...split, [field]: parsed != null && Number.isFinite(parsed) ? parsed : null }
          : split
      )
    }));
  };

  const addTeamBudgetSplit = () => {
    setDraft(prev => ({
      ...prev,
      teamBudgetSplits: [
        ...prev.teamBudgetSplits,
        {
          id: generateId(),
          teamName: '',
          totalBudgetMD: null,
          actualSpendMD: null,
          varianceMD: null,
          etcMD: null,
        }
      ]
    }));
  };

  const removeTeamBudgetSplit = (splitId: string) => {
    setDraft(prev => ({
      ...prev,
      teamBudgetSplits: prev.teamBudgetSplits.filter(split => split.id !== splitId)
    }));
  };

  const splitTotals = useMemo(() => {
    return draft.teamBudgetSplits.reduce(
      (acc, split) => ({
        total: acc.total + (split.totalBudgetMD || 0),
        spent: acc.spent + (split.actualSpendMD || 0),
        variance: acc.variance + (split.varianceMD || 0),
        etc: acc.etc + (split.etcMD || 0),
      }),
      { total: 0, spent: 0, variance: 0, etc: 0 }
    );
  }, [draft.teamBudgetSplits]);

  const hasSplitDelta =
    draft.teamBudgetSplits.length > 0 && (
      Math.abs(splitTotals.total - (draft.totalBudgetMD || 0)) > 0.1 ||
      Math.abs(splitTotals.spent - (draft.actualSpendMD || 0)) > 0.1 ||
      Math.abs(splitTotals.variance - (draft.varianceMD || 0)) > 0.1 ||
      Math.abs(splitTotals.etc - (draft.etcMD || 0)) > 0.1
    );

  if (view === 'preview') {
    return (
      <div className="max-w-7xl mx-auto">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h2 className="text-xl font-bold text-gray-900 dark:text-white flex items-center gap-2">
              <Eye className="w-5 h-5 text-indigo-500" /> PM - Project card Preview
            </h2>
            <p className="text-sm text-gray-500 dark:text-gray-400 mt-1">
              Final one-pager ready for Steering Committee export
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
              className="flex items-center gap-2 px-5 py-2 text-sm font-bold text-white bg-gradient-to-r from-emerald-600 to-teal-600 rounded-lg hover:from-emerald-700 hover:to-teal-700 transition-all shadow-md"
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
            <span className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-600 flex items-center justify-center text-white shadow-lg">
              <FileBarChart className="w-5 h-5" />
            </span>
            PM - Project card
          </h2>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-2">
            Select linked projects, map existing data with local AI, amend the card manually, then generate a board-ready one-pager.
          </p>
        </div>
        <button
          onClick={handleGenerateOnePager}
          disabled={isGeneratingDocument || selectedProjectIds.length === 0}
          className={`flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-bold transition-all shadow-md ${
            selectedProjectIds.length === 0
              ? 'bg-gray-200 text-gray-400 cursor-not-allowed dark:bg-gray-800 dark:text-gray-600'
              : 'bg-gradient-to-r from-emerald-600 to-teal-600 text-white hover:from-emerald-700 hover:to-teal-700'
          }`}
        >
          {isGeneratingDocument ? (
            <>
              <Sparkles className="w-4 h-4 animate-spin" /> Generating...
            </>
          ) : (
            <>
              <ClipboardList className="w-4 h-4" /> Generate Final Project Card
            </>
          )}
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <section className="lg:col-span-4 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-gray-900 dark:text-white">1. Select Projects</h3>
            <div className="flex items-center gap-2">
              <button onClick={handleSelectAll} className="text-xs font-semibold text-indigo-600 hover:underline">All</button>
              <button onClick={handleClearSelection} className="text-xs font-semibold text-gray-500 hover:underline">Clear</button>
            </div>
          </div>

          <div className="space-y-2 max-h-[270px] overflow-y-auto pr-1">
            {allProjects.map(project => {
              const selected = selectedProjectIds.includes(project.id);
              const openBlockers = project.tasks.filter(task => task.status === 'Blocked').length;
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
                    <span>{openBlockers} blocker{openBlockers > 1 ? 's' : ''}</span>
                  </div>
                </button>
              );
            })}
            {allProjects.length === 0 && (
              <p className="text-sm text-gray-400 dark:text-gray-500 py-8 text-center">No accessible project.</p>
            )}
          </div>

          <div className="p-3 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40 space-y-3">
            <label className="inline-flex items-center gap-2 text-sm font-medium text-gray-700 dark:text-gray-300 cursor-pointer">
              <input
                type="checkbox"
                checked={includeOneOffQueries}
                onChange={e => setIncludeOneOffQueries(e.target.checked)}
                className="w-4 h-4 rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
              />
              Include linked one-off queries
            </label>
            <p className="text-xs text-gray-500 dark:text-gray-400">
              {linkedSelectedOneOffQueries.length} linked one-off quer{linkedSelectedOneOffQueries.length > 1 ? 'ies' : 'y'} found for selected projects.
            </p>

            <div>
              <label className={labelClass}>Common Link Between Projects</label>
              <input
                className={inputClass}
                placeholder="Example: Shared ERP rollout and data migration"
                value={draft.commonLink}
                onChange={e => setDraft(prev => ({ ...prev, commonLink: e.target.value }))}
              />
            </div>

            <div>
              <label className={labelClass}>Manual Context (optional)</label>
              <textarea
                className={`${inputClass} resize-none`}
                rows={4}
                placeholder="Paste any board context, constraints, or latest steering guidance..."
                value={manualContext}
                onChange={e => setManualContext(e.target.value)}
              />
            </div>

            <div className="grid grid-cols-2 gap-2">
              <button
                onClick={handleRebuildFromData}
                disabled={selectedProjectIds.length === 0}
                className="inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-lg hover:bg-indigo-100 dark:hover:bg-indigo-900/40 disabled:opacity-50"
              >
                <GanttChartSquare className="w-3.5 h-3.5" /> Rebuild Draft
              </button>
              <button
                onClick={handleAIMapping}
                disabled={isAIMapping || selectedProjectIds.length === 0}
                className="inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-bold text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 disabled:opacity-50"
              >
                <Bot className={`w-3.5 h-3.5 ${isAIMapping ? 'animate-spin' : ''}`} /> AI Mapping
              </button>
            </div>
          </div>

          {(aiError || aiInfo) && (
            <div className={`p-3 rounded-lg text-xs border ${aiError ? 'bg-red-50 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800' : 'bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-300 dark:border-emerald-800'}`}>
              {aiError || aiInfo}
            </div>
          )}
        </section>

        <section className="lg:col-span-8 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm space-y-6">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-gray-900 dark:text-white">2. Edit Project Card Mask</h3>
            <span className="text-xs text-gray-500 dark:text-gray-400">All fields remain editable before final generation</span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className={labelClass}>Card Title</label>
              <input className={inputClass} value={draft.cardTitle} onChange={e => setDraft(prev => ({ ...prev, cardTitle: e.target.value }))} />
            </div>
            <div>
              <label className={labelClass}>Overall Health (RAG)</label>
              <select
                className={inputClass}
                value={draft.overallHealth}
                onChange={e => setDraft(prev => ({ ...prev, overallHealth: e.target.value as CardHealth }))}
              >
                <option value="">Not set</option>
                <option value="Green">Green</option>
                <option value="Amber">Amber</option>
                <option value="Red">Red</option>
              </select>
              <div className="mt-2">
                <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full border text-[11px] font-bold ${healthBadgeClass(draft.overallHealth)}`}>
                  {draft.overallHealth || 'N/A'}
                </span>
              </div>
            </div>
            <div>
              <label className={labelClass}>Executive Sponsor</label>
              <input className={inputClass} value={draft.executiveSponsor} onChange={e => setDraft(prev => ({ ...prev, executiveSponsor: e.target.value }))} />
            </div>
            <div>
              <label className={labelClass}>Project Manager</label>
              <input className={inputClass} value={draft.projectManager} onChange={e => setDraft(prev => ({ ...prev, projectManager: e.target.value }))} />
            </div>
            <div className="md:col-span-2">
              <label className={labelClass}>Project Objective</label>
              <textarea className={`${inputClass} resize-none`} rows={2} value={draft.projectObjective} onChange={e => setDraft(prev => ({ ...prev, projectObjective: e.target.value }))} />
            </div>
            <div className="md:col-span-2">
              <label className={labelClass}>Executive Summary</label>
              <textarea className={`${inputClass} resize-none`} rows={3} value={draft.executiveSummary} onChange={e => setDraft(prev => ({ ...prev, executiveSummary: e.target.value }))} />
            </div>
          </div>

          <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <h4 className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2"><CheckCircle2 className="w-4 h-4 text-emerald-500" /> Key Achievements</h4>
              <button
                onClick={() => setDraft(prev => ({ ...prev, keyAchievements: [...prev.keyAchievements, ''] }))}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-md hover:bg-indigo-100 dark:hover:bg-indigo-900/40"
              >
                <Plus className="w-3.5 h-3.5" /> Add
              </button>
            </div>
            <div className="space-y-2">
              {draft.keyAchievements.map((item, idx) => (
                <div key={`achievement-${idx}`} className="flex items-center gap-2">
                  <input
                    className={inputClass}
                    value={item}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      keyAchievements: prev.keyAchievements.map((entry, i) => i === idx ? e.target.value : entry)
                    }))}
                  />
                  <button
                    onClick={() => setDraft(prev => ({ ...prev, keyAchievements: prev.keyAchievements.filter((_, i) => i !== idx) }))}
                    className="inline-flex items-center justify-center w-8 h-8 rounded-md border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
              {draft.keyAchievements.length === 0 && <p className="text-xs text-gray-500 dark:text-gray-400">No achievement line yet.</p>}
            </div>
          </div>

          <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <h4 className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2"><Calendar className="w-4 h-4 text-sky-500" /> Schedule & Milestones</h4>
              <button
                onClick={() => setDraft(prev => ({
                  ...prev,
                  milestones: [...prev.milestones, { id: generateId(), description: '', baselineDate: '', forecastOrActualDate: '', status: '' }]
                }))}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-md hover:bg-indigo-100 dark:hover:bg-indigo-900/40"
              >
                <Plus className="w-3.5 h-3.5" /> Add
              </button>
            </div>
            <div className="space-y-2">
              {draft.milestones.map((milestone, idx) => (
                <div key={milestone.id} className="grid grid-cols-1 md:grid-cols-12 gap-2 items-center">
                  <input
                    className={`${inputClass} md:col-span-4`}
                    placeholder="Milestone description"
                    value={milestone.description}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      milestones: prev.milestones.map((entry, i) => i === idx ? { ...entry, description: e.target.value } : entry)
                    }))}
                  />
                  <input
                    type="date"
                    className={`${inputClass} md:col-span-2`}
                    value={milestone.baselineDate}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      milestones: prev.milestones.map((entry, i) => i === idx ? { ...entry, baselineDate: e.target.value } : entry)
                    }))}
                  />
                  <input
                    type="date"
                    className={`${inputClass} md:col-span-2`}
                    value={milestone.forecastOrActualDate}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      milestones: prev.milestones.map((entry, i) => i === idx ? { ...entry, forecastOrActualDate: e.target.value } : entry)
                    }))}
                  />
                  <select
                    className={`${inputClass} md:col-span-3`}
                    value={milestone.status}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      milestones: prev.milestones.map((entry, i) => i === idx ? { ...entry, status: e.target.value as MilestoneStatus } : entry)
                    }))}
                  >
                    <option value="">Not set</option>
                    <option value="On Track">On Track</option>
                    <option value="Delayed">Delayed</option>
                    <option value="Completed">Completed</option>
                  </select>
                  <button
                    onClick={() => setDraft(prev => ({ ...prev, milestones: prev.milestones.filter((_, i) => i !== idx) }))}
                    className="inline-flex items-center justify-center w-8 h-8 rounded-md border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 md:col-span-1"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                  <div className="md:col-span-12">
                    <span className={`inline-flex items-center px-2 py-0.5 rounded-full border text-[11px] font-bold ${milestoneStatusBadgeClass(milestone.status)}`}>
                      {milestone.status || 'N/A'}
                    </span>
                  </div>
                </div>
              ))}
              {draft.milestones.length === 0 && <p className="text-xs text-gray-500 dark:text-gray-400">No milestone line yet.</p>}
            </div>
          </div>

          <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
            <h4 className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2 mb-3"><DollarSign className="w-4 h-4 text-emerald-500" /> Financials (MD)</h4>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
              <div>
                <label className={labelClass}>Total Budget</label>
                <input type="number" className={inputClass} value={draft.totalBudgetMD ?? ''} onChange={e => updateNumberField('totalBudgetMD', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>Actual Spend</label>
                <input type="number" className={inputClass} value={draft.actualSpendMD ?? ''} onChange={e => updateNumberField('actualSpendMD', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>Variance</label>
                <input type="number" className={inputClass} value={draft.varianceMD ?? ''} onChange={e => updateNumberField('varianceMD', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>ETC</label>
                <input type="number" className={inputClass} value={draft.etcMD ?? ''} onChange={e => updateNumberField('etcMD', e.target.value)} />
              </div>
            </div>

            <div className="mt-5 pt-4 border-t border-gray-200 dark:border-gray-700">
              <div className="flex items-center justify-between mb-2">
                <div>
                  <h5 className="text-xs font-bold text-gray-900 dark:text-white">Manual Team Cost Distribution</h5>
                  <p className="text-[11px] text-gray-500 dark:text-gray-400">Team names are free text and do not need to match DOINg teams.</p>
                </div>
                <button
                  onClick={addTeamBudgetSplit}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-md hover:bg-indigo-100 dark:hover:bg-indigo-900/40"
                >
                  <Plus className="w-3.5 h-3.5" /> Add Team Split
                </button>
              </div>

              {draft.teamBudgetSplits.length === 0 ? (
                <div className="px-3 py-2 text-xs text-gray-500 dark:text-gray-400 border border-dashed border-gray-300 dark:border-gray-700 rounded-lg bg-gray-50 dark:bg-gray-800/40">
                  No team split yet. Add rows to distribute budget/spend manually across teams.
                </div>
              ) : (
                <div className="space-y-2">
                  {draft.teamBudgetSplits.map(split => (
                    <div key={split.id} className="grid grid-cols-1 md:grid-cols-12 gap-2 items-center p-2 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40">
                      <input
                        className={`${inputClass} md:col-span-3`}
                        placeholder="Team name (free text)"
                        value={split.teamName}
                        onChange={e => updateTeamBudgetSplitName(split.id, e.target.value)}
                      />
                      <input
                        type="number"
                        className={`${inputClass} md:col-span-2`}
                        placeholder="Budget MD"
                        value={split.totalBudgetMD ?? ''}
                        onChange={e => updateTeamBudgetSplitNumber(split.id, 'totalBudgetMD', e.target.value)}
                      />
                      <input
                        type="number"
                        className={`${inputClass} md:col-span-2`}
                        placeholder="Spent MD"
                        value={split.actualSpendMD ?? ''}
                        onChange={e => updateTeamBudgetSplitNumber(split.id, 'actualSpendMD', e.target.value)}
                      />
                      <input
                        type="number"
                        className={`${inputClass} md:col-span-2`}
                        placeholder="Variance MD"
                        value={split.varianceMD ?? ''}
                        onChange={e => updateTeamBudgetSplitNumber(split.id, 'varianceMD', e.target.value)}
                      />
                      <input
                        type="number"
                        className={`${inputClass} md:col-span-2`}
                        placeholder="ETC MD"
                        value={split.etcMD ?? ''}
                        onChange={e => updateTeamBudgetSplitNumber(split.id, 'etcMD', e.target.value)}
                      />
                      <button
                        onClick={() => removeTeamBudgetSplit(split.id)}
                        className="inline-flex items-center justify-center w-8 h-8 rounded-md border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 md:col-span-1"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}

                  <div className="rounded-lg border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-800 p-3">
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-xs">
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Budget:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.total)}</span></div>
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Spent:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.spent)}</span></div>
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split Variance:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.variance)}</span></div>
                      <div><span className="font-semibold text-gray-500 dark:text-gray-400">Split ETC:</span> <span className="font-bold text-gray-800 dark:text-gray-100">{formatMD(splitTotals.etc)}</span></div>
                    </div>
                    {hasSplitDelta && (
                      <p className="mt-2 text-xs text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-md px-2 py-1.5">
                        Split totals differ from global Financials fields (Budget / Spent / Variance / ETC). Adjust manually if you want full alignment.
                      </p>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>

          <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
            <div className="flex items-center justify-between mb-3">
              <h4 className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2"><AlertTriangle className="w-4 h-4 text-amber-500" /> Risks, Issues & Dependencies (RID)</h4>
              <button
                onClick={() => setDraft(prev => ({
                  ...prev,
                  ridItems: [...prev.ridItems, { id: generateId(), risk: '', issue: '', dependency: '', mitigation: '' }]
                }))}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-bold text-indigo-700 dark:text-indigo-300 bg-indigo-50 dark:bg-indigo-900/30 border border-indigo-200 dark:border-indigo-700 rounded-md hover:bg-indigo-100 dark:hover:bg-indigo-900/40"
              >
                <Plus className="w-3.5 h-3.5" /> Add
              </button>
            </div>
            <div className="space-y-2">
              {draft.ridItems.map((item, idx) => (
                <div key={item.id} className="grid grid-cols-1 md:grid-cols-12 gap-2 items-center">
                  <input
                    className={`${inputClass} md:col-span-3`}
                    placeholder="Risk"
                    value={item.risk}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      ridItems: prev.ridItems.map((entry, i) => i === idx ? { ...entry, risk: e.target.value } : entry)
                    }))}
                  />
                  <input
                    className={`${inputClass} md:col-span-3`}
                    placeholder="Issue"
                    value={item.issue}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      ridItems: prev.ridItems.map((entry, i) => i === idx ? { ...entry, issue: e.target.value } : entry)
                    }))}
                  />
                  <input
                    className={`${inputClass} md:col-span-2`}
                    placeholder="Dependency"
                    value={item.dependency}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      ridItems: prev.ridItems.map((entry, i) => i === idx ? { ...entry, dependency: e.target.value } : entry)
                    }))}
                  />
                  <input
                    className={`${inputClass} md:col-span-3`}
                    placeholder="Mitigation"
                    value={item.mitigation}
                    onChange={e => setDraft(prev => ({
                      ...prev,
                      ridItems: prev.ridItems.map((entry, i) => i === idx ? { ...entry, mitigation: e.target.value } : entry)
                    }))}
                  />
                  <button
                    onClick={() => setDraft(prev => ({ ...prev, ridItems: prev.ridItems.filter((_, i) => i !== idx) }))}
                    className="inline-flex items-center justify-center w-8 h-8 rounded-md border border-red-200 dark:border-red-800 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-900/20 md:col-span-1"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
              {draft.ridItems.length === 0 && <p className="text-xs text-gray-500 dark:text-gray-400">No RID line yet.</p>}
            </div>
          </div>

          <div className="border border-gray-200 dark:border-gray-700 rounded-lg p-4">
            <h4 className="text-sm font-bold text-gray-900 dark:text-white mb-3">Key Decisions & Asks</h4>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              <div>
                <label className={labelClass}>Key Decisions (1 line = 1 item)</label>
                <textarea className={`${inputClass} resize-none`} rows={4} value={draft.keyDecisions.join('\n')} onChange={e => updateListField('keyDecisions', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>Approvals Needed (1 line = 1 item)</label>
                <textarea className={`${inputClass} resize-none`} rows={4} value={draft.approvalsNeeded.join('\n')} onChange={e => updateListField('approvalsNeeded', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>Resource Requests (1 line = 1 item)</label>
                <textarea className={`${inputClass} resize-none`} rows={4} value={draft.resourceRequests.join('\n')} onChange={e => updateListField('resourceRequests', e.target.value)} />
              </div>
              <div>
                <label className={labelClass}>Escalations (1 line = 1 item)</label>
                <textarea className={`${inputClass} resize-none`} rows={4} value={draft.escalations.join('\n')} onChange={e => updateListField('escalations', e.target.value)} />
              </div>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
};

export default PMProjectCard;
