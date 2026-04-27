import React, { startTransition, useEffect, useRef, useState } from 'react';
import {
  AlertCircle,
  Calendar,
  ChevronDown,
  ChevronUp,
  CheckCircle2,
  Clock,
  Download,
  FileSpreadsheet,
  Info,
  Plus,
  RefreshCw,
  Shield,
  Trash2,
  UploadCloud,
  Users,
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import type { Workbook } from 'exceljs';
import {
  applyTimesheetWorkbook,
  buildPlanFileName,
  DEFAULT_STAFF_CONFIG,
  downloadWorkbook,
  generateTimesheetWorkbook,
  getMonthSheets,
  parseMonthSheetName,
  PLAN_TYPE_OPTIONS,
  type PlanType,
  readWorkbookFromFile,
} from './lib/workbookLogic';
import {
  applyTimesheetToGoogleSheet,
  loadGoogleSpreadsheetSource,
  listRecentSpreadsheets,
  type DriveFile,
} from './lib/googleSheets';

type SourceMode = 'sheets' | 'excel';
type BusyState = 'sourceUpload' | 'sourceConnect' | 'generate' | 'update' | null;
type SourceInfo =
  | {
      mode: 'excel';
      label: string;
    }
  | {
      mode: 'sheets';
      label: string;
      spreadsheetId: string;
    };

type RuntimeEnv = ImportMeta & {
  env?: Record<string, string | undefined>;
};

const STORAGE_KEYS = {
  sourceMode: 'viesu-saraksts-source-mode',
  staffRows: 'viesu-saraksts-staff-rows',
};

const PLAN_TYPE_UI_COPY: Record<PlanType, { description: string }> = {
  viesnica: {
    description: 'Reģistratūra, numuri un brokastu maiņa',
  },
  restor: {
    description: 'Restorāna un pasākumu apkalpošana',
  },
  virtuve: {
    description: 'Virtuves un ēdināšanas sagatavošana',
  },
};

type StaffEditorRow = {
  id: string;
  name: string;
  team: PlanType;
};

const STAFF_TEAM_OPTIONS: Array<{ value: PlanType; label: string }> = [
  { value: 'viesnica', label: 'Viesnīca' },
  { value: 'restor', label: 'Restorāns' },
  { value: 'virtuve', label: 'Virtuve' },
];

let nextStaffRowId = 1;

function createStaffRow(team: PlanType = 'viesnica', name = ''): StaffEditorRow {
  return {
    id: `staff-row-${Date.now()}-${nextStaffRowId++}`,
    name,
    team,
  };
}

function isPlanType(value: string): value is PlanType {
  return value === 'viesnica' || value === 'restor' || value === 'virtuve';
}

function inferTeamFromLegacyConfig(dept: string, role: string): PlanType {
  const normalizedDept = dept.trim().toLowerCase();
  const normalizedRole = role.trim().toLowerCase();

  if (normalizedDept === 'restorāns' || normalizedDept === 'restorans') return 'restor';
  if (normalizedDept === 'virtuve' && !normalizedRole.includes('brokast')) return 'virtuve';
  return 'viesnica';
}

function parseStaffConfigRows(text: string): StaffEditorRow[] {
  const rows = text
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [name = '', dept = '', role = ''] = line.split('|').map((part) => part.trim());
      return createStaffRow(inferTeamFromLegacyConfig(dept, role), name);
    });

  return rows.length ? rows : [createStaffRow()];
}

function serializeStaffConfigRows(rows: StaffEditorRow[]): string {
  return rows
    .map((row) => [row.name.trim(), row.team])
    .filter(([name]) => Boolean(name))
    .map(([name, team]) => `${name} | ${team} |`)
    .join('\n');
}

function getStoredStaffRows(): StaffEditorRow[] {
  if (typeof window === 'undefined') return parseStaffConfigRows(DEFAULT_STAFF_CONFIG);

  const raw = window.localStorage.getItem(STORAGE_KEYS.staffRows);
  if (!raw) return parseStaffConfigRows(DEFAULT_STAFF_CONFIG);

  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return parseStaffConfigRows(DEFAULT_STAFF_CONFIG);
    }

    const rows = parsed
      .map((item) => {
        if (!item || typeof item !== 'object') return null;
        const record = item as Partial<StaffEditorRow>;
        const team = typeof record.team === 'string' && isPlanType(record.team) ? record.team : 'viesnica';
        const name = typeof record.name === 'string' ? record.name : '';
        const id = typeof record.id === 'string' && record.id.trim() ? record.id : createStaffRow(team, name).id;
        return { id, name, team } satisfies StaffEditorRow;
      })
      .filter((row): row is StaffEditorRow => row !== null);

    if (!rows.length) {
      return [createStaffRow()];
    }

    nextStaffRowId =
      rows.reduce((maxId, row) => {
        const suffix = Number(row.id.split('-').at(-1));
        return Number.isFinite(suffix) ? Math.max(maxId, suffix + 1) : maxId;
      }, nextStaffRowId);

    return rows;
  } catch {
    return parseStaffConfigRows(DEFAULT_STAFF_CONFIG);
  }
}

function getStoredStaffRowsSnapshot(): string {
  return JSON.stringify(getStoredStaffRows());
}

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

function getStoredValue(key: string, fallback: string): string {
  if (typeof window === 'undefined') return fallback;
  return window.localStorage.getItem(key) ?? fallback;
}

function setStoredValue(key: string, value: string) {
  if (typeof window === 'undefined') return;
  window.localStorage.setItem(key, value);
}

function InfoHint({ text }: { text: string }) {
  return (
    <span className="relative inline-flex group">
      <button
        type="button"
        className="inline-flex h-7 w-7 items-center justify-center rounded-full text-on-surface-variant transition-colors hover:text-primary"
        aria-label="Papildu informācija"
      >
        <Info className="w-4 h-4" />
      </button>
      <span className="pointer-events-none absolute left-full top-1/2 z-20 ml-3 hidden w-72 -translate-y-1/2 rounded-[1rem] border border-outline-variant/20 bg-surface-container-lowest px-4 py-3 text-sm font-medium leading-6 text-on-surface shadow-xl group-hover:block group-focus-within:block">
        {text}
      </span>
    </span>
  );
}

function GoogleConnectLoader() {
  return (
    <div className="google-loader-shell" aria-live="polite" aria-label="Notiek savienojums ar Google Sheets">
      <figure className="google-loader">
        <div className="dot white" />
        <div className="dot" />
        <div className="dot" />
        <div className="dot" />
        <div className="dot" />
      </figure>
      <div className="google-loader-label">Savienojas ar Google Sheets...</div>
    </div>
  );
}

export default function App() {
  const runtimeEnv = ((import.meta as RuntimeEnv).env ?? {}) as Record<string, string | undefined>;
  const initialStaffRowsSnapshotRef = useRef<string>(getStoredStaffRowsSnapshot());

  const [sourceMode, setSourceMode] = useState<SourceMode>(
    () => (getStoredValue(STORAGE_KEYS.sourceMode, 'sheets') as SourceMode) || 'sheets',
  );
  const [recentSheets, setRecentSheets] = useState<DriveFile[] | null>(null);
  const [sourceInfo, setSourceInfo] = useState<SourceInfo | null>(null);
  const [availableMonths, setAvailableMonths] = useState<string[]>([]);
  const [selectedMonth, setSelectedMonth] = useState<string>('');
  const [selectedPlanType, setSelectedPlanType] = useState<PlanType>('viesnica');
  const [isSheetPickerOpen, setIsSheetPickerOpen] = useState(false);
  const [completedTimesheetFile, setCompletedTimesheetFile] = useState<File | null>(null);
  const [hasGeneratedPlan, setHasGeneratedPlan] = useState(false);
  const [staffRows, setStaffRows] = useState<StaffEditorRow[]>(() => JSON.parse(initialStaffRowsSnapshotRef.current));
  const [savedStaffRowsSnapshot, setSavedStaffRowsSnapshot] = useState<string>(() => initialStaffRowsSnapshotRef.current);
  const [isStaffConfigOpen, setIsStaffConfigOpen] = useState(false);
  const [isOnline, setIsOnline] = useState(() => (typeof navigator === 'undefined' ? true : navigator.onLine));
  const [isBusy, setIsBusy] = useState<BusyState>(null);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(
    null,
  );

  const originalWorkbookRef = useRef<Workbook | null>(null);
  const statusTimeoutRef = useRef<number | null>(null);

  useEffect(() => {
    setStoredValue(STORAGE_KEYS.sourceMode, sourceMode);
  }, [sourceMode]);

  useEffect(() => {
    return () => {
      if (statusTimeoutRef.current) {
        window.clearTimeout(statusTimeoutRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (typeof window === 'undefined') return;

    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
    };
  }, []);

  const showStatus = (type: 'success' | 'error' | 'info', text: string) => {
    if (statusTimeoutRef.current) {
      window.clearTimeout(statusTimeoutRef.current);
    }

    setStatusMessage({ type, text });
    statusTimeoutRef.current = window.setTimeout(() => setStatusMessage(null), 5000);
  };

  const applyLoadedWorkbook = (workbook: Workbook, source: SourceInfo) => {
    const monthSheets = getMonthSheets(workbook);
    originalWorkbookRef.current = workbook;
    setSourceInfo(source);
    setHasGeneratedPlan(false);
    setCompletedTimesheetFile(null);

    startTransition(() => {
      setAvailableMonths(monthSheets);
      setSelectedMonth(monthSheets[0] ?? '');
    });

    if (!monthSheets.length) {
      showStatus('error', 'Avota dokumentā netika atrastas mēneša lapas, piemēram, apr26 vai mai26');
      return;
    }

    showStatus('success', `Avots ielādēts: ${source.label}. Atrastas ${monthSheets.length} mēneša lapas`);
  };

  useEffect(() => {
    if (!sourceInfo) return;
    setHasGeneratedPlan(false);
    setCompletedTimesheetFile(null);
  }, [selectedMonth, selectedPlanType]);

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsBusy('sourceUpload');
    try {
      const workbook = await readWorkbookFromFile(file);
      applyLoadedWorkbook(workbook, {
        mode: 'excel',
        label: file.name,
      });
    } catch (error) {
      console.error(error);
      showStatus('error', 'Kļūda apstrādājot viesu saraksta Excel failu');
    } finally {
      setIsBusy(null);
    }
  };

  const fetchSheets = async () => {
    const clientId = (runtimeEnv.VITE_GOOGLE_CLIENT_ID || '').trim();
    if (!clientId) {
      showStatus('error', 'Lūdzu, pievienojiet VITE_GOOGLE_CLIENT_ID AI Studio iestatījumos (Settings -> Secrets)');
      return;
    }

    setIsBusy('sourceConnect');
    try {
      const sheets = await listRecentSpreadsheets(clientId);
      setRecentSheets(sheets);
      setIsSheetPickerOpen(true);
      showStatus('success', 'Atrasti nesenie Google Sheets faili');
    } catch (error) {
      console.error(error);
      showStatus('error', error instanceof Error ? error.message : 'Neizdevās pieslēgties Google Drive');
    } finally {
      setIsBusy(null);
    }
  };

  const connectGoogleSheet = async (spreadsheetId: string) => {
    const clientId = (runtimeEnv.VITE_GOOGLE_CLIENT_ID || '').trim();
    if (!clientId) {
      showStatus('error', 'Lūdzu, pievienojiet VITE_GOOGLE_CLIENT_ID AI Studio iestatījumos (Settings -> Secrets)');
      return;
    }

    setIsBusy('sourceConnect');
    try {
      const { workbook, title } = await loadGoogleSpreadsheetSource({
        clientId,
        spreadsheetId,
      });

      applyLoadedWorkbook(workbook, {
        mode: 'sheets',
        label: title,
        spreadsheetId,
      });
      setIsSheetPickerOpen(false);
    } catch (error) {
      console.error(error);
      showStatus('error', error instanceof Error ? error.message : 'Neizdevās pieslēgt Google Sheet');
    } finally {
      setIsBusy(null);
    }
  };

  const handleCompletedTimesheetUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setCompletedTimesheetFile(file);
    showStatus('success', `Fails ${file.name} augšupielādēts`);
  };

  const updateStaffRow = (id: string, field: keyof Omit<StaffEditorRow, 'id'>, value: string) => {
    setStaffRows((rows) => rows.map((row) => (row.id === id ? { ...row, [field]: value } : row)));
  };

  const addStaffRow = () => {
    setStaffRows((rows) => [createStaffRow(selectedPlanType), ...rows]);
  };

  const deleteStaffRow = (id: string) => {
    setStaffRows((rows) => {
      if (rows.length === 1) {
        return [createStaffRow()];
      }

      return rows.filter((row) => row.id !== id);
    });
  };

  const saveStaffRows = () => {
    const snapshot = JSON.stringify(staffRows);
    setStoredValue(STORAGE_KEYS.staffRows, snapshot);
    setSavedStaffRowsSnapshot(snapshot);
    showStatus('success', 'Personāla saraksts saglabāts');
  };

  const generateTimesheet = async () => {
    const staffConfig = serializeStaffConfigRows(staffRows);

    if (!originalWorkbookRef.current || !selectedMonth || !sourceInfo) {
      showStatus('error', 'Vispirms ielādējiet viesu saraksta avotu');
      return;
    }

    setIsBusy('generate');
    try {
      const { workbook, fileName, dayCount } = await generateTimesheetWorkbook(
        originalWorkbookRef.current,
        selectedMonth,
        staffConfig,
        selectedPlanType,
      );

      await downloadWorkbook(workbook, fileName);
      setHasGeneratedPlan(true);
      setCompletedTimesheetFile(null);
      showStatus('success', `Darba plāns sagatavots: ${fileName} (${dayCount} dienu bloki)`);
    } catch (error) {
      console.error(error);
      showStatus('error', error instanceof Error ? error.message : 'Kļūda ģenerējot darba plānu');
    } finally {
      setIsBusy(null);
    }
  };

  const updateViesuSaraksts = async () => {
    const staffConfig = serializeStaffConfigRows(staffRows);

    if (!originalWorkbookRef.current || !completedTimesheetFile || !sourceInfo) {
      showStatus('error', 'Vispirms ielādējiet avotu un aizpildīto darba plāna failu');
      return;
    }

    setIsBusy('update');
    try {
      const timesheetWorkbook = await readWorkbookFromFile(completedTimesheetFile);

      if (sourceInfo.mode === 'sheets') {
        const clientId = (runtimeEnv.VITE_GOOGLE_CLIENT_ID || '').trim();
        const result = await applyTimesheetToGoogleSheet({
          clientId,
          spreadsheetId: sourceInfo.spreadsheetId,
          originalWorkbook: originalWorkbookRef.current,
          timesheetWorkbook,
          staffConfigText: staffConfig,
        });

        showStatus('success', `Google Sheet atjaunots (${result.appliedDays} dienas)`);
      } else {
        const { workbook, fileName, appliedDays } = await applyTimesheetWorkbook(
          originalWorkbookRef.current,
          sourceInfo.label,
          timesheetWorkbook,
          staffConfig,
        );

        await downloadWorkbook(workbook, fileName);
        showStatus('success', `Viesu saraksts atjaunots: ${fileName} (${appliedDays} dienas)`);
      }
    } catch (error) {
      console.error(error);
      showStatus('error', error instanceof Error ? error.message : 'Kļūda atjaunojot viesu sarakstu');
    } finally {
      setIsBusy(null);
    }
  };

  const staffConfig = serializeStaffConfigRows(staffRows);
  const staffCount = staffRows.filter((row) => row.name.trim()).length;
  const hasUnsavedStaffChanges = JSON.stringify(staffRows) !== savedStaffRowsSnapshot;

  const sourceLoaded = sourceInfo !== null;
  const selectedPlanOption =
    PLAN_TYPE_OPTIONS.find((option) => option.value === selectedPlanType) ?? PLAN_TYPE_OPTIONS[0];
  const selectedMonthKey = selectedMonth ? parseMonthSheetName(selectedMonth) : null;
  const updateActionLabel =
    sourceMode === 'excel' ? '4. Lejupielādēt atjaunoto sarakstu' : '4. Atjaunot Google Sheet';
  const updateBusyLabel =
    sourceMode === 'excel' ? 'Atjauno viesu sarakstu...' : 'Atjauno Google Sheet...';
  const isGoogleConnecting = sourceMode === 'sheets' && isBusy === 'sourceConnect';

  return (
    <div className="app-shell min-h-screen text-on-surface font-body">
      <div className="app-page-background" aria-hidden="true" />
      <header className="w-full top-0 px-8 py-6 bg-surface-container-low transition-colors duration-300 sticky z-50">
        <div className="flex justify-between items-center max-w-[1440px] mx-auto w-full">
          <div
            className={cn(
              'inline-flex items-center gap-3 rounded-full border px-4 py-2 text-sm font-bold',
              isOnline
                ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                : 'border-red-200 bg-red-50 text-red-700',
            )}
          >
            <span
              className={cn(
                'h-2.5 w-2.5 rounded-full',
                isOnline ? 'bg-emerald-500 shadow-[0_0_0_4px_rgba(16,185,129,0.18)]' : 'bg-red-500 shadow-[0_0_0_4px_rgba(239,68,68,0.18)]',
              )}
            />
            {isOnline ? 'Darbojas' : 'Nav savienojuma'}
          </div>
        </div>
      </header>

      <main className="max-w-[1380px] mx-auto px-5 py-10 lg:px-8 lg:py-14 xl:px-10">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-10 xl:gap-12 items-start">
          <div className="relative bg-surface-container-lowest rounded-lg p-8 md:p-9 xl:p-10 shadow-[0_12px_32px_-4px_rgba(172,45,94,0.08)] flex flex-col h-full border border-outline-variant/15 min-w-0">
              {isGoogleConnecting && (
                <div className="absolute inset-0 z-20 rounded-lg bg-surface/78 backdrop-blur-[2px] flex items-center justify-center">
                  <GoogleConnectLoader />
                </div>
              )}
              <div className="flex items-start gap-4 mb-6">
                <div className="w-14 h-14 rounded-[1.4rem] bg-secondary-container flex items-center justify-center text-primary shrink-0">
                  <Calendar className="w-7 h-7" />
                </div>
                <div className="min-w-0 pt-1">
                  <div className="flex items-center gap-2">
                    <h2 className="editorial-header text-[clamp(1.65rem,2.1vw,2.05rem)] font-bold text-on-surface leading-[1.05]">
                      1. Ielādēt viesu sarakstu
                    </h2>
                    <InfoHint text="Izvēlieties Google Sheets dokumentu vai augšupielādējiet Excel avotu, no kura ģenerēt darba plānu." />
                  </div>
                </div>
              </div>

              <div className="grid w-full grid-cols-2 rounded-[1.6rem] bg-surface-container p-1.5 mb-7 gap-1.5">
                {(['sheets', 'excel'] as SourceMode[]).map((mode) => (
                  <button
                    key={mode}
                    type="button"
                    onClick={() => setSourceMode(mode)}
                    className={cn(
                      'ui-button min-h-[3.25rem] rounded-[1.15rem] px-4 py-3 text-sm font-bold text-center',
                      sourceMode === mode
                        ? 'ui-button-primary shadow-[0_10px_24px_-12px_rgba(172,45,94,0.85)]'
                        : 'ui-button-soft',
                    )}
                  >
                    {mode === 'sheets' ? 'Google Sheets' : 'Excel fails'}
                  </button>
                ))}
              </div>

              <div className="space-y-7 flex-grow min-w-0">
                {sourceMode === 'sheets' ? (
                  <>
                    {!recentSheets ? (
                      <>
                        <button
                          type="button"
                          onClick={fetchSheets}
                          disabled={isBusy !== null}
                          className="ui-button ui-button-primary font-bold rounded-full py-4 px-8 flex items-center justify-center gap-2 shadow-lg shadow-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
                        >
                          <Calendar className="w-5 h-5 shrink-0" />
                          <span>{isBusy === 'sourceConnect' ? 'Pieslēdzas...' : 'Pieslēgties Google Drive'}</span>
                        </button>
                      </>
                    ) : isSheetPickerOpen || sourceInfo?.mode !== 'sheets' ? (
                      <div className="space-y-4">
                        <div className="flex items-center justify-between">
                          <label className="text-sm font-semibold text-on-surface-variant ml-1">Izvēlieties failu</label>
                          <button
                            onClick={fetchSheets}
                            className="ui-button ui-button-soft rounded-full px-4 py-2 text-primary text-sm font-bold"
                          >
                            Atsvaidzināt
                          </button>
                        </div>
                        <div className="flex flex-col gap-2.5 max-h-[260px] overflow-y-auto pr-2 custom-scrollbar">
                          {recentSheets.length === 0 ? (
                            <div className="text-sm text-on-surface-variant p-4 text-center bg-surface-container-lowest rounded-xl border border-outline-variant/30">
                              Netika atrasts neviens Google Sheets fails.
                            </div>
                          ) : (
                            recentSheets.map((sheet) => (
                              <button
                                key={sheet.id}
                                onClick={() => connectGoogleSheet(sheet.id)}
                                disabled={isBusy !== null}
                                className="ui-button ui-button-soft flex items-center gap-3 p-4 rounded-[1.35rem] border border-outline-variant/30 text-left disabled:opacity-50 min-w-0"
                              >
                                <FileSpreadsheet className="w-5 h-5 text-primary shrink-0" />
                                <span className="font-medium text-on-surface truncate min-w-0">{sheet.name}</span>
                              </button>
                            ))
                          )}
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        <div className="flex items-center justify-between gap-4">
                          <label className="text-sm font-semibold text-on-surface-variant ml-1">Izvēlētais Google Sheet</label>
                          <button
                            type="button"
                            onClick={() => setIsSheetPickerOpen(true)}
                            className="ui-button ui-button-soft text-primary text-sm font-bold rounded-full px-4 py-2 shrink-0"
                          >
                            Mainīt failu
                          </button>
                        </div>
                        <div className="flex items-center gap-3 bg-surface-container-low p-4 rounded-[1.5rem] min-w-0">
                          <div className="w-8 h-8 rounded-full bg-primary flex items-center justify-center text-white shrink-0">
                            <CheckCircle2 className="w-5 h-5" />
                          </div>
                          <span className="text-sm font-medium text-on-surface truncate min-w-0" title={sourceInfo.label}>
                            Google Sheet: {sourceInfo.label}
                          </span>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <label className="group relative flex flex-col items-center justify-center border-2 border-dashed border-outline-variant/50 rounded-[1.8rem] p-10 bg-surface-container-low hover:bg-secondary-container transition-colors cursor-pointer text-center min-h-[240px]">
                    <input type="file" className="hidden" accept=".xlsx" onChange={handleExcelUpload} />
                    <UploadCloud className="w-12 h-12 text-primary mb-4" />
                    <p className="text-on-surface font-semibold mb-1">Ievelciet viesu saraksta Excel failu šeit</p>
                    <p className="text-on-surface-variant text-sm">Atbalstītais formāts: .xlsx</p>
                  </label>
                )}

                <div className="space-y-2">
                  <label className="text-sm font-semibold text-on-surface-variant ml-1">Izvēlēties mēnesi</label>
                  <div className="relative">
                    <select
                      className="w-full bg-surface-container-lowest border border-outline-variant/30 rounded-full px-6 py-3.5 appearance-none focus:ring-2 focus:ring-primary/20 focus:border-primary outline-none transition-all font-medium text-on-surface"
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(e.target.value)}
                      disabled={availableMonths.length === 0}
                    >
                      {availableMonths.length === 0 ? (
                        <option value="">Vispirms ielādējiet avotu</option>
                      ) : (
                        availableMonths.map((month) => (
                          <option key={month} value={month}>
                            {month}
                          </option>
                        ))
                      )}
                    </select>
                    <div className="absolute right-5 top-1/2 -translate-y-1/2 pointer-events-none text-outline">
                      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M6 9L12 15L18 9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                  </div>
                </div>

                <div className="space-y-2">
                  <label className="text-sm font-semibold text-on-surface-variant ml-1">Darba plāna tips</label>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    {PLAN_TYPE_OPTIONS.map((option) => (
                      <button
                        key={option.value}
                        type="button"
                        onClick={() => setSelectedPlanType(option.value)}
                        className={cn(
                          'ui-button rounded-[1.75rem] border px-4 py-4 text-left min-h-[112px] flex flex-col justify-between gap-3 min-w-0 overflow-hidden',
                          selectedPlanType === option.value
                            ? 'ui-button-primary border-primary text-white'
                            : 'ui-button-soft border-outline-variant/30',
                        )}
                      >
                        <div>
                          <div className="font-bold text-lg leading-tight">{option.label}</div>
                          <div className="text-xs mt-1.5 opacity-80 leading-5">
                            {PLAN_TYPE_UI_COPY[option.value].description}
                          </div>
                        </div>
                        <div
                          className="text-[11px] leading-4 opacity-70 text-wrap-anywhere"
                          title={selectedMonthKey ? buildPlanFileName(selectedMonthKey, option.value) : option.fileName}
                        >
                          {selectedMonthKey ? buildPlanFileName(selectedMonthKey, option.value) : option.fileName}
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <button
                onClick={generateTimesheet}
                disabled={!sourceLoaded || !selectedMonth || isBusy !== null}
                className="ui-button ui-button-primary text-white font-bold rounded-full py-4 px-8 mt-8 flex items-center justify-center gap-2 shadow-lg shadow-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Download className="w-5 h-5 shrink-0" />
                <span>{isBusy === 'generate' ? 'Ģenerē darba plānu...' : `2. Lejupielādēt ${selectedPlanOption.label} Excel`}</span>
              </button>
            </div>

            <div className="bg-surface-container-lowest rounded-lg p-8 md:p-9 xl:p-10 shadow-[0_12px_32px_-4px_rgba(172,45,94,0.08)] flex flex-col h-full border border-outline-variant/15 min-w-0">
              <div className="flex items-start gap-4 mb-6">
                <div className="w-14 h-14 rounded-[1.4rem] bg-secondary-container flex items-center justify-center text-primary shrink-0">
                  <RefreshCw className="w-7 h-7" />
                </div>
                <div className="min-w-0 pt-1">
                  <div className="flex items-center gap-2">
                    <h2 className="editorial-header text-[clamp(1.65rem,2.1vw,2.05rem)] font-bold text-on-surface leading-[1.05]">
                      3. Augšupielādēt aizpildīto darba plānu
                    </h2>
                    <InfoHint text="Ievietojiet aizpildīto Excel failu, lai sistēma ierakstītu maiņas atpakaļ sākotnējā avotā." />
                  </div>
                </div>
              </div>

              <div className="space-y-7 flex-grow min-w-0">
                <label className="group relative flex flex-col items-center justify-center border-2 border-dashed border-outline-variant/50 rounded-[1.8rem] p-10 bg-surface-container-low hover:bg-secondary-container transition-colors cursor-pointer text-center min-h-[240px]">
                  <input type="file" className="hidden" accept=".xlsx" onChange={handleCompletedTimesheetUpload} />
                  <UploadCloud className="w-12 h-12 text-primary mb-4" />
                  <p className="text-on-surface font-semibold mb-1">Ievelciet aizpildīto darba laiku failu</p>
                  <p className="text-on-surface-variant text-sm">Sistēma pārlasīs Excel un atjaunos sākotnējo avotu</p>
                </label>

                {completedTimesheetFile && (
                  <div className="flex items-center gap-3 bg-surface-container-low p-4 rounded-[1.5rem] min-w-0">
                    <div className="w-8 h-8 rounded-full bg-primary flex items-center justify-center text-white">
                      <CheckCircle2 className="w-5 h-5" />
                    </div>
                    <span className="text-sm font-medium text-on-surface truncate min-w-0" title={completedTimesheetFile.name}>
                      {completedTimesheetFile.name} augšupielādēts
                    </span>
                  </div>
                )}

                <div className="flex items-center gap-2 text-primary text-sm font-bold px-1">
                  <Info className="w-4 h-4" />
                  Apstrādes laiks: ~5-10 sekundes
                </div>
              </div>

              <button
                onClick={updateViesuSaraksts}
                disabled={!completedTimesheetFile || !sourceLoaded || isBusy !== null}
                className="ui-button ui-button-primary text-white font-bold rounded-full py-4 px-8 mt-8 flex items-center justify-center gap-2 shadow-lg shadow-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <FileSpreadsheet className="w-5 h-5 shrink-0" />
                <span>{isBusy === 'update' ? updateBusyLabel : updateActionLabel}</span>
              </button>
            </div>
        </div>

        <div className="mt-10 max-w-[1000px] mx-auto">
          <div className="bg-surface-container-lowest rounded-lg p-7 md:p-8 xl:p-9 shadow-[0_12px_32px_-4px_rgba(172,45,94,0.08)] border border-outline-variant/15">
            <div className="flex flex-col gap-5 md:flex-row md:items-start md:justify-between">
              <div className="flex items-start gap-3">
                <div className="w-12 h-12 rounded-full bg-secondary-container flex items-center justify-center text-primary shrink-0">
                  <Users className="w-6 h-6" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <h2 className="editorial-header text-2xl font-bold text-on-surface">Personāla saraksts plāniem</h2>
                    <InfoHint text="Šis ir galvenais darbinieku saraksts. Sistēma no tā izvēlas, kuri cilvēki jāiekļauj konkrētajā Excel plānā, un pēc tam izmanto to pašu kartējumu ierakstīšanai atpakaļ." />
                  </div>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setIsStaffConfigOpen((value) => !value)}
                className="ui-button ui-button-soft inline-flex items-center justify-center gap-2 self-start rounded-full border border-outline-variant/30 px-5 py-3 text-sm font-bold text-on-surface"
              >
                {isStaffConfigOpen ? (
                  <>
                    <ChevronUp className="w-4 h-4" />
                    Paslēpt sarakstu
                  </>
                ) : (
                  <>
                    <ChevronDown className="w-4 h-4" />
                    Rādīt un rediģēt
                  </>
                )}
              </button>
            </div>

            <div className="mt-6">
              <div className="rounded-[1.5rem] bg-surface-container px-5 py-4">
                <div className="text-xs font-bold uppercase tracking-[0.18em] text-on-surface-variant/70">
                  Ko tas ietekmē
                </div>
                <div className="mt-3 space-y-2 text-sm leading-6 text-on-surface-variant">
                  <p>
                    <span className="font-semibold text-on-surface">Ģenerēšanu:</span> kuri darbinieki parādās Excel failā.
                  </p>
                  <p>
                    <span className="font-semibold text-on-surface">Importu atpakaļ:</span> kuru aizpildītās rindas sistēma
                    nolasa no Excel.
                  </p>
                  <p>
                    <span className="font-semibold text-on-surface">Aktīvi darbinieki:</span> {staffCount}
                  </p>
                </div>
              </div>
            </div>

            {isStaffConfigOpen && (
              <div className="mt-6">
                <div className="mb-4 flex items-center justify-between gap-4">
                  <div className="text-sm text-on-surface-variant leading-6">
                    Rediģējiet katru darbinieku atsevišķi un nospiediet <span className="font-semibold text-on-surface">Saglabāt sarakstu</span>.
                  </div>
                  <div className="flex items-center gap-3 shrink-0">
                    {hasUnsavedStaffChanges && (
                      <span className="text-xs font-bold text-primary">Ir nesaglabātas izmaiņas</span>
                    )}
                    <button
                      type="button"
                      onClick={saveStaffRows}
                      className="ui-button ui-button-soft inline-flex items-center gap-2 rounded-full border border-outline-variant/30 px-4 py-2.5 text-sm font-bold text-on-surface"
                    >
                      <CheckCircle2 className="w-4 h-4" />
                      Saglabāt sarakstu
                    </button>
                    <button
                      type="button"
                      onClick={addStaffRow}
                      className="ui-button ui-button-primary inline-flex items-center gap-2 rounded-full px-4 py-2.5 text-sm font-bold text-white"
                    >
                      <Plus className="w-4 h-4" />
                      Pievienot
                    </button>
                  </div>
                </div>

                <div className="overflow-hidden rounded-[1.75rem] border border-outline-variant/30 bg-surface-container-low">
                  <div className="hidden md:grid grid-cols-[1.2fr_0.9fr_auto] gap-3 border-b border-outline-variant/20 px-5 py-4 text-xs font-bold uppercase tracking-[0.14em] text-on-surface-variant/80">
                    <div>Vārds</div>
                    <div>Komanda</div>
                    <div>Darbība</div>
                  </div>

                  <div className="max-h-[420px] overflow-y-auto">
                    {staffRows.map((row) => (
                      <div
                        key={row.id}
                        className="grid gap-3 border-b border-outline-variant/15 px-4 py-4 md:grid-cols-[1.2fr_0.9fr_auto] md:px-5 last:border-b-0"
                      >
                        <label className="flex flex-col gap-1.5">
                          <span className="text-xs font-semibold text-on-surface-variant md:hidden">Vārds</span>
                          <input
                            value={row.name}
                            onChange={(e) => updateStaffRow(row.id, 'name', e.target.value)}
                            className="rounded-[1rem] border border-outline-variant/25 bg-surface-container-lowest px-4 py-3 text-sm font-medium text-on-surface outline-none focus:border-primary focus:ring-2 focus:ring-primary/15"
                            placeholder="Piemēram, Anita"
                          />
                        </label>
                        <label className="flex flex-col gap-1.5">
                          <span className="text-xs font-semibold text-on-surface-variant md:hidden">Komanda</span>
                          <select
                            value={row.team}
                            onChange={(e) => updateStaffRow(row.id, 'team', e.target.value)}
                            className="rounded-[1rem] border border-outline-variant/25 bg-surface-container-lowest px-4 py-3 text-sm font-medium text-on-surface outline-none focus:border-primary focus:ring-2 focus:ring-primary/15"
                          >
                            {STAFF_TEAM_OPTIONS.map((option) => (
                              <option key={option.value} value={option.value}>
                                {option.label}
                              </option>
                            ))}
                          </select>
                        </label>
                        <div className="flex items-end md:items-center">
                          <button
                            type="button"
                            onClick={() => deleteStaffRow(row.id)}
                            className="ui-button ui-button-danger inline-flex items-center justify-center gap-2 rounded-full border border-outline-variant/30 px-4 py-3 text-sm font-semibold"
                          >
                            <Trash2 className="w-4 h-4" />
                            Dzēst
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="mt-20 flex flex-col md:flex-row items-center justify-between gap-6 border-t border-outline-variant/10 pt-10">
          <div className="flex items-center gap-6">
            <div className="text-on-surface-variant text-sm font-medium flex items-center gap-2">
              <Shield className="w-5 h-5" />
              Datu apstrāde lokāli pārlūkā
            </div>
            <div className="text-on-surface-variant text-sm font-medium flex items-center gap-2">
              <Clock className="w-5 h-5" />
              Pēdējā darbība: tikko
            </div>
          </div>
          <div className="text-on-surface-variant text-xs opacity-50 font-medium">Viesnīcas Suite © 2026 | Google Sheets edition</div>
        </div>
      </main>

      {statusMessage && (
        <div className="fixed bottom-8 right-8 z-[100]">
          <div
            className={cn(
              'bg-surface-container-lowest shadow-2xl rounded-lg p-4 pr-12 border-l-4 flex items-center gap-4 animate-in slide-in-from-right duration-500',
              statusMessage.type === 'error' ? 'border-error' : 'border-primary',
            )}
          >
            <div
              className={cn(
                'w-10 h-10 rounded-full flex items-center justify-center text-white',
                statusMessage.type === 'error' ? 'bg-error' : 'bg-primary',
              )}
            >
              {statusMessage.type === 'error' ? <AlertCircle className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
            </div>
            <div>
              <p className="font-bold text-on-surface text-sm">{statusMessage.type === 'error' ? 'Kļūda' : 'Gatavs darbam'}</p>
              <p className="text-xs text-on-surface-variant">{statusMessage.text}</p>
            </div>
            <button
              onClick={() => setStatusMessage(null)}
              className="absolute top-2 right-2 text-on-surface-variant hover:text-primary transition-colors"
            >
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="18" y1="6" x2="6" y2="18" />
                <line x1="6" y1="6" x2="18" y2="18" />
              </svg>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
