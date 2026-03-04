import React, { useMemo, useState } from 'react';
import {
  AlertTriangle,
  CalendarClock,
  CheckCircle2,
  Clock3,
  Mail,
  Search,
  Send,
  Trash2
} from 'lucide-react';
import {
  OneOffQuery,
  PMEmailJob,
  PMEmailReportType,
  PMGanttItem,
  PMReportData,
  SMTPConfig,
  Team,
  User
} from '../types';
import { generateId } from '../services/storage';
import { buildPMStatusReportHTMLFromSelection } from './PMReport';
import { buildPMGantHTMLFromSelection, PMGantProject } from './PMGant';
import { buildPMProjectCardHTMLFromSelection, PMProject } from './PMProjectCard';

interface PMMailCenterProps {
  teams: Team[];
  users: User[];
  currentUser: User;
  smtpConfig: SMTPConfig;
  pmReportData: PMReportData[];
  pmGantData: PMGanttItem[];
  oneOffQueries: OneOffQuery[];
  pmEmailJobs: PMEmailJob[];
  onSavePMEmailJob: (job: PMEmailJob) => void;
  onDeletePMEmailJob: (id: string) => void;
}

type SendMode = 'one-shot' | 'scheduled';
type ProjectOption = PMProject & { teamId: string };

const REPORT_LABELS: Record<PMEmailReportType, string> = {
  'pm-status-report': 'PM status report',
  'pm-project-card': 'PM - Project card',
  'pm-gant': 'PM Gant',
};

const statusBadgeClass = (status: PMEmailJob['status']) => {
  if (status === 'sent') return 'bg-emerald-100 text-emerald-700 border-emerald-200 dark:bg-emerald-900/20 dark:text-emerald-300 dark:border-emerald-800';
  if (status === 'failed') return 'bg-red-100 text-red-700 border-red-200 dark:bg-red-900/20 dark:text-red-300 dark:border-red-800';
  if (status === 'cancelled') return 'bg-slate-100 text-slate-600 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700';
  return 'bg-amber-100 text-amber-700 border-amber-200 dark:bg-amber-900/20 dark:text-amber-300 dark:border-amber-800';
};

const parseEmails = (value: string): string[] =>
  value
    .split(/[;,]/)
    .map(v => v.trim())
    .filter(Boolean);

const wrapEmailHTML = (title: string, reportHTML: string): string => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title></head><body style="margin:0;padding:18px;background:#ffffff;font-family:'Segoe UI',system-ui,-apple-system,sans-serif">${reportHTML}</body></html>`;

const PMMailCenter: React.FC<PMMailCenterProps> = ({
  teams,
  users,
  currentUser,
  smtpConfig,
  pmReportData,
  pmGantData,
  oneOffQueries,
  pmEmailJobs,
  onSavePMEmailJob,
  onDeletePMEmailJob,
}) => {
  const [sendMode, setSendMode] = useState<SendMode>('one-shot');
  const [selectedTeamId, setSelectedTeamId] = useState('');
  const [projectSearch, setProjectSearch] = useState('');
  const [selectedProjectIds, setSelectedProjectIds] = useState<string[]>([]);
  const [includeOneOffQueries, setIncludeOneOffQueries] = useState(true);
  const [selectedReportTypes, setSelectedReportTypes] = useState<Record<PMEmailReportType, boolean>>({
    'pm-status-report': true,
    'pm-project-card': false,
    'pm-gant': false,
  });
  const [toInput, setToInput] = useState('');
  const [ccInput, setCcInput] = useState('');
  const [bccInput, setBccInput] = useState('');
  const [scheduledAt, setScheduledAt] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [info, setInfo] = useState('');

  const allProjects = useMemo<ProjectOption[]>(() => {
    const projects: ProjectOption[] = [];
    teams.forEach(team => {
      (team.projects || []).forEach(project => {
        if (!project.isArchived) {
          projects.push({ ...project, teamName: team.name, teamId: team.id });
        }
      });
    });
    return projects;
  }, [teams]);

  const teamOptions = useMemo(() => teams.map(team => ({ id: team.id, name: team.name })), [teams]);

  const visibleProjects = useMemo(() => {
    const query = projectSearch.trim().toLowerCase();
    return allProjects.filter(project => {
      if (selectedTeamId && project.teamId !== selectedTeamId) return false;
      if (!query) return true;
      return (
        project.name.toLowerCase().includes(query) ||
        project.teamName.toLowerCase().includes(query) ||
        String(project.status || '').toLowerCase().includes(query)
      );
    });
  }, [allProjects, projectSearch, selectedTeamId]);

  const selectedProjects = useMemo(() => {
    const set = new Set(selectedProjectIds);
    return allProjects.filter(project => set.has(project.id));
  }, [allProjects, selectedProjectIds]);

  const sortedJobs = useMemo(() => {
    return [...(pmEmailJobs || [])].sort((a, b) => {
      const ad = new Date(a.createdAt).getTime();
      const bd = new Date(b.createdAt).getTime();
      return bd - ad;
    });
  }, [pmEmailJobs]);

  const smtpReady = Boolean(
    smtpConfig?.host?.trim() &&
    smtpConfig?.user?.trim() &&
    smtpConfig?.password &&
    Number(smtpConfig?.port || 0) > 0
  );

  const selectedReportList = useMemo(() => {
    return (Object.keys(selectedReportTypes) as PMEmailReportType[]).filter(key => selectedReportTypes[key]);
  }, [selectedReportTypes]);

  const toggleProject = (projectId: string) => {
    setSelectedProjectIds(prev =>
      prev.includes(projectId) ? prev.filter(id => id !== projectId) : [...prev, projectId]
    );
  };

  const selectAllVisibleProjects = () => {
    const visibleIds = new Set(visibleProjects.map(project => project.id));
    setSelectedProjectIds(prev => Array.from(new Set([...prev, ...visibleIds])));
  };

  const clearProjectSelection = () => setSelectedProjectIds([]);

  const handleTeamChange = (teamId: string) => {
    setSelectedTeamId(teamId);
    setSelectedProjectIds(prev => prev.filter(projectId => {
      const project = allProjects.find(item => item.id === projectId);
      if (!project) return false;
      return !teamId || project.teamId === teamId;
    }));
  };

  const buildReportPayloads = (): Array<{
    reportType: PMEmailReportType;
    reportTitle: string;
    subject: string;
    html: string;
  }> => {
    const selectedProjectSet = new Set(selectedProjectIds);
    const projectNames = selectedProjects.map(project => project.name);
    const scopeSuffix = projectNames.length === 1
      ? projectNames[0]
      : `${projectNames.length} projects`;
    const subjectDate = new Date().toLocaleDateString('en-GB');

    return selectedReportList.map(reportType => {
      if (reportType === 'pm-status-report') {
        const htmlFragment = buildPMStatusReportHTMLFromSelection(
          selectedProjects,
          pmReportData || [],
          selectedProjectIds
        );
        const reportTitle = REPORT_LABELS[reportType];
        return {
          reportType,
          reportTitle,
          subject: `${reportTitle} - ${scopeSuffix} - ${subjectDate}`,
          html: wrapEmailHTML(reportTitle, htmlFragment),
        };
      }

      if (reportType === 'pm-gant') {
        const htmlFragment = buildPMGantHTMLFromSelection(
          selectedProjects as PMGantProject[],
          (pmGantData || []).filter(item => selectedProjectSet.has(item.projectId))
        );
        const reportTitle = REPORT_LABELS[reportType];
        return {
          reportType,
          reportTitle,
          subject: `${reportTitle} - ${scopeSuffix} - ${subjectDate}`,
          html: wrapEmailHTML(reportTitle, htmlFragment),
        };
      }

      const linkedOneOff = (oneOffQueries || []).filter(query =>
        query.projectId && selectedProjectSet.has(query.projectId)
      );
      const htmlFragment = buildPMProjectCardHTMLFromSelection(
        selectedProjects as PMProject[],
        users,
        (pmReportData || []).filter(report => selectedProjectSet.has(report.projectId)),
        (pmGantData || []).filter(item => selectedProjectSet.has(item.projectId)),
        includeOneOffQueries ? linkedOneOff : [],
        currentUser,
        { includeOneOffQueries, commonLink: '' }
      );
      const reportTitle = REPORT_LABELS[reportType];
      return {
        reportType,
        reportTitle,
        subject: `${reportTitle} - ${scopeSuffix} - ${subjectDate}`,
        html: wrapEmailHTML(reportTitle, htmlFragment),
      };
    });
  };

  const validateBeforeSubmit = () => {
    if (!smtpReady) return 'SMTP config is incomplete. Please configure it first in Settings.';
    if (selectedProjects.length === 0) return 'Select at least one project.';
    if (selectedReportList.length === 0) return 'Select at least one PM report type.';
    const to = parseEmails(toInput);
    const cc = parseEmails(ccInput);
    const bcc = parseEmails(bccInput);
    if (to.length === 0 && cc.length === 0 && bcc.length === 0) return 'Add at least one recipient (To, Cc, or Bcc).';
    if (sendMode === 'scheduled' && !scheduledAt) return 'Select a schedule date/time.';
    return '';
  };

  const handleSubmit = async () => {
    setError('');
    setInfo('');
    const validationError = validateBeforeSubmit();
    if (validationError) {
      setError(validationError);
      return;
    }

    const recipients = {
      to: parseEmails(toInput),
      cc: parseEmails(ccInput),
      bcc: parseEmails(bccInput),
    };
    const payloads = buildReportPayloads();
    const scheduleAtISO = sendMode === 'scheduled'
      ? new Date(scheduledAt).toISOString()
      : new Date().toISOString();

    setIsSubmitting(true);
    try {
      if (sendMode === 'one-shot') {
        let sentCount = 0;
        for (const payload of payloads) {
          try {
            const response = await fetch('/api/mail/send', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                subject: payload.subject,
                html: payload.html,
                to: recipients.to,
                cc: recipients.cc,
                bcc: recipients.bcc,
              }),
            });
            if (!response.ok) {
              const body = await response.json().catch(() => ({}));
              throw new Error(body?.error || `SMTP send failed (${response.status})`);
            }

            const now = new Date().toISOString();
            onSavePMEmailJob({
              id: generateId(),
              reportType: payload.reportType,
              reportTitle: payload.reportTitle,
              subject: payload.subject,
              htmlBody: payload.html,
              teamId: selectedTeamId || null,
              teamName: selectedTeamId ? (teams.find(team => team.id === selectedTeamId)?.name || '') : '',
              projectIds: selectedProjects.map(project => project.id),
              projectNames: selectedProjects.map(project => project.name),
              recipients,
              createdByUserId: currentUser.id,
              createdAt: now,
              scheduleAt: now,
              status: 'sent',
              sentAt: now,
              lastTriedAt: now,
            });
            sentCount += 1;
          } catch (sendError: any) {
            const now = new Date().toISOString();
            onSavePMEmailJob({
              id: generateId(),
              reportType: payload.reportType,
              reportTitle: payload.reportTitle,
              subject: payload.subject,
              htmlBody: payload.html,
              teamId: selectedTeamId || null,
              teamName: selectedTeamId ? (teams.find(team => team.id === selectedTeamId)?.name || '') : '',
              projectIds: selectedProjects.map(project => project.id),
              projectNames: selectedProjects.map(project => project.name),
              recipients,
              createdByUserId: currentUser.id,
              createdAt: now,
              scheduleAt: now,
              status: 'failed',
              lastTriedAt: now,
              lastError: sendError?.message || 'Unknown SMTP error',
            });
          }
        }
        setInfo(`${sentCount}/${payloads.length} email(s) sent immediately.`);
      } else {
        payloads.forEach(payload => {
          onSavePMEmailJob({
            id: generateId(),
            reportType: payload.reportType,
            reportTitle: payload.reportTitle,
            subject: payload.subject,
            htmlBody: payload.html,
            teamId: selectedTeamId || null,
            teamName: selectedTeamId ? (teams.find(team => team.id === selectedTeamId)?.name || '') : '',
            projectIds: selectedProjects.map(project => project.id),
            projectNames: selectedProjects.map(project => project.name),
            recipients,
            createdByUserId: currentUser.id,
            createdAt: new Date().toISOString(),
            scheduleAt: scheduleAtISO,
            status: 'pending',
          });
        });
        setInfo(`${payloads.length} email job(s) scheduled for ${new Date(scheduleAtISO).toLocaleString()}.`);
      }
    } catch (e: any) {
      setError(e?.message || 'Email action failed.');
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-extrabold text-gray-900 dark:text-white flex items-center gap-3">
            <span className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-600 flex items-center justify-center text-white shadow-lg">
              <Mail className="w-5 h-5" />
            </span>
            PM Mail Center (Admin)
          </h2>
          <p className="text-sm text-gray-500 dark:text-gray-400 mt-2">
            Send PM reports now or schedule delivery using configured SMTP, with team/project scope and recipients (To/Cc/Bcc).
          </p>
        </div>
      </div>

      {!smtpReady && (
        <div className="rounded-xl border border-amber-200 bg-amber-50 text-amber-800 px-4 py-3 text-sm flex items-start gap-2 dark:bg-amber-900/20 dark:border-amber-700 dark:text-amber-300">
          <AlertTriangle className="w-4 h-4 mt-0.5 flex-shrink-0" />
          SMTP configuration is incomplete. Configure host, port, user and password in Settings before sending.
        </div>
      )}

      {error && (
        <div className="rounded-xl border border-red-200 bg-red-50 text-red-700 px-4 py-3 text-sm dark:bg-red-900/20 dark:border-red-700 dark:text-red-300">
          {error}
        </div>
      )}
      {info && (
        <div className="rounded-xl border border-emerald-200 bg-emerald-50 text-emerald-700 px-4 py-3 text-sm flex items-center gap-2 dark:bg-emerald-900/20 dark:border-emerald-700 dark:text-emerald-300">
          <CheckCircle2 className="w-4 h-4" />
          {info}
        </div>
      )}

      <div className="grid grid-cols-1 xl:grid-cols-12 gap-6">
        <section className="xl:col-span-5 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-bold text-gray-900 dark:text-white">Scope Selection</h3>
            <div className="text-xs text-gray-500 dark:text-gray-400">{selectedProjects.length} selected</div>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Team</label>
            <select
              value={selectedTeamId}
              onChange={e => handleTeamChange(e.target.value)}
              className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
            >
              <option value="">All teams</option>
              {teamOptions.map(team => (
                <option key={team.id} value={team.id}>{team.name}</option>
              ))}
            </select>
          </div>

          <div className="relative">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={projectSearch}
              onChange={e => setProjectSearch(e.target.value)}
              placeholder="Search project..."
              className="w-full pl-10 pr-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
            />
          </div>

          <div className="flex items-center gap-3 text-xs">
            <button onClick={selectAllVisibleProjects} className="font-semibold text-indigo-600 hover:underline">Select visible</button>
            <button onClick={clearProjectSelection} className="font-semibold text-gray-500 hover:underline">Clear</button>
          </div>

          <div className="space-y-2 max-h-[320px] overflow-y-auto pr-1">
            {visibleProjects.map(project => {
              const checked = selectedProjectIds.includes(project.id);
              return (
                <button
                  key={project.id}
                  onClick={() => toggleProject(project.id)}
                  className={`w-full text-left px-3 py-2.5 rounded-lg border transition-colors ${
                    checked
                      ? 'border-emerald-400 bg-emerald-50 dark:bg-emerald-900/20'
                      : 'border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40 hover:bg-gray-100 dark:hover:bg-gray-800'
                  }`}
                >
                  <div className="flex items-start justify-between gap-3">
                    <div>
                      <p className="text-sm font-semibold text-gray-900 dark:text-white">{project.name}</p>
                      <p className="text-xs text-gray-500 dark:text-gray-400">{project.teamName}</p>
                    </div>
                    {checked ? <CheckCircle2 className="w-4 h-4 text-emerald-500 mt-0.5" /> : <span className="w-4 h-4 rounded-full border border-gray-300 dark:border-gray-600 mt-0.5" />}
                  </div>
                </button>
              );
            })}
            {visibleProjects.length === 0 && (
              <p className="text-sm text-gray-400 dark:text-gray-500 py-8 text-center">No project matches current filters.</p>
            )}
          </div>
        </section>

        <section className="xl:col-span-7 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm space-y-4">
          <h3 className="text-sm font-bold text-gray-900 dark:text-white">Email Configuration</h3>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <button
              onClick={() => setSendMode('one-shot')}
              className={`p-3 rounded-lg border text-left ${sendMode === 'one-shot' ? 'border-emerald-400 bg-emerald-50 dark:bg-emerald-900/20' : 'border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40'}`}
            >
              <p className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2"><Send className="w-4 h-4" /> One-shot send</p>
              <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Generate and send immediately.</p>
            </button>
            <button
              onClick={() => setSendMode('scheduled')}
              className={`p-3 rounded-lg border text-left ${sendMode === 'scheduled' ? 'border-emerald-400 bg-emerald-50 dark:bg-emerald-900/20' : 'border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40'}`}
            >
              <p className="text-sm font-bold text-gray-900 dark:text-white flex items-center gap-2"><CalendarClock className="w-4 h-4" /> Scheduled send</p>
              <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">Generate snapshot now, deliver at selected time.</p>
            </button>
          </div>

          <div>
            <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">PM reports to send</label>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
              {(Object.keys(REPORT_LABELS) as PMEmailReportType[]).map(reportType => (
                <label key={reportType} className="inline-flex items-center gap-2 px-3 py-2 rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40 cursor-pointer text-sm text-gray-700 dark:text-gray-300">
                  <input
                    type="checkbox"
                    checked={selectedReportTypes[reportType]}
                    onChange={e => setSelectedReportTypes(prev => ({ ...prev, [reportType]: e.target.checked }))}
                    className="w-4 h-4 rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
                  />
                  {REPORT_LABELS[reportType]}
                </label>
              ))}
            </div>
          </div>

          <label className="inline-flex items-center gap-2 text-sm text-gray-700 dark:text-gray-300">
            <input
              type="checkbox"
              checked={includeOneOffQueries}
              onChange={e => setIncludeOneOffQueries(e.target.checked)}
              className="w-4 h-4 rounded border-gray-300 text-emerald-600 focus:ring-emerald-500"
            />
            Include linked one-off queries in PM - Project card generation
          </label>

          <div className="grid grid-cols-1 gap-3">
            <div>
              <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">To</label>
              <input
                value={toInput}
                onChange={e => setToInput(e.target.value)}
                placeholder="name@company.com; second@company.com"
                className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
              />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Cc</label>
                <input
                  value={ccInput}
                  onChange={e => setCcInput(e.target.value)}
                  placeholder="optional"
                  className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Bcc</label>
                <input
                  value={bccInput}
                  onChange={e => setBccInput(e.target.value)}
                  placeholder="optional"
                  className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                />
              </div>
            </div>
            {sendMode === 'scheduled' && (
              <div>
                <label className="block text-xs font-semibold text-gray-500 dark:text-gray-400 mb-1">Schedule at</label>
                <input
                  type="datetime-local"
                  value={scheduledAt}
                  onChange={e => setScheduledAt(e.target.value)}
                  className="w-full px-3 py-2 text-sm border border-gray-200 dark:border-gray-700 rounded-lg bg-white dark:bg-gray-800 text-gray-900 dark:text-gray-100"
                />
              </div>
            )}
          </div>

          <div className="flex justify-end pt-2">
            <button
              onClick={handleSubmit}
              disabled={isSubmitting}
              className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-bold text-white bg-gradient-to-r from-emerald-600 to-teal-600 rounded-lg hover:from-emerald-700 hover:to-teal-700 disabled:opacity-50"
            >
              {isSubmitting ? <Clock3 className="w-4 h-4 animate-spin" /> : sendMode === 'one-shot' ? <Send className="w-4 h-4" /> : <CalendarClock className="w-4 h-4" />}
              {sendMode === 'one-shot' ? 'Send now' : 'Schedule emails'}
            </button>
          </div>
        </section>
      </div>

      <section className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-gray-900 dark:text-white mb-4">Scheduled / Sent Jobs</h3>
        <div className="space-y-3 max-h-[420px] overflow-y-auto pr-1">
          {sortedJobs.map(job => (
            <div key={job.id} className="rounded-lg border border-gray-200 dark:border-gray-700 bg-gray-50 dark:bg-gray-800/40 p-3">
              <div className="flex items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className={`text-[11px] px-2 py-0.5 rounded-full border font-semibold ${statusBadgeClass(job.status)}`}>{job.status.toUpperCase()}</span>
                    <span className="text-xs text-gray-500 dark:text-gray-400">{REPORT_LABELS[job.reportType]}</span>
                  </div>
                  <p className="text-sm font-semibold text-gray-900 dark:text-white mt-1">{job.subject}</p>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    Team: {job.teamName || 'N/A'} • Projects: {job.projectNames.join(', ') || 'N/A'}
                  </p>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    To: {(job.recipients.to || []).join(', ') || '—'}{job.recipients.cc?.length ? ` • Cc: ${job.recipients.cc.join(', ')}` : ''}{job.recipients.bcc?.length ? ` • Bcc: ${job.recipients.bcc.join(', ')}` : ''}
                  </p>
                  <p className="text-[11px] text-gray-400 mt-1">
                    Created: {new Date(job.createdAt).toLocaleString()} • Scheduled: {new Date(job.scheduleAt).toLocaleString()}
                    {job.sentAt ? ` • Sent: ${new Date(job.sentAt).toLocaleString()}` : ''}
                  </p>
                  {job.lastError && <p className="text-[11px] text-red-600 dark:text-red-300 mt-1">Error: {job.lastError}</p>}
                </div>
                {(job.status === 'pending' || job.status === 'failed' || job.status === 'cancelled') && (
                  <button
                    onClick={() => {
                      if (!window.confirm('Delete this email job?')) return;
                      onDeletePMEmailJob(job.id);
                    }}
                    className="inline-flex items-center gap-1 px-2 py-1 text-xs font-semibold text-red-600 hover:text-red-700"
                    title="Delete job"
                  >
                    <Trash2 className="w-3 h-3" />
                    Delete
                  </button>
                )}
              </div>
            </div>
          ))}
          {sortedJobs.length === 0 && (
            <p className="text-sm text-gray-400 dark:text-gray-500 py-8 text-center">No PM email jobs yet.</p>
          )}
        </div>
      </section>
    </div>
  );
};

export default PMMailCenter;
