import type { Cell, Workbook, Worksheet } from 'exceljs';
import {
  buildLatvianMonthLabel,
  getDaysInMonth,
  isWeekendDay,
  parseMonthKeyFromSheetName,
} from './calendar';
import { parseActivity, type EventCategory } from './eventParsing';

export const DEFAULT_STAFF = [
  'Ruta | Vadība | Projektu vadītāja',
  'Anita | Reģistratūra | Administratore',
  'Agate | Reģistratūra | Administratore',
  'Dina | Reģistratūra | Administratore',
  'Inese | Reģistratūra | Administratore',
  'Evija | Reģistratūra | Administratore',
  'Liāna | Reģistratūra | Administratore',
  'Una | Reģistratūra | Administratore',
  'Ivita | Numuri | Numuriņu uzkopšana',
  'Sarmīte | Numuri | Numuriņu uzkopšana',
  'Ināra | Numuri | Numuriņu uzkopšana',
  'Solvita | Numuri | Numuriņu uzkopšana',
  'Santa | Virtuve | Brokastu maiņa',
  'Airita | Virtuve | Brokastu maiņa',
  'Inga | Restorāns | Maiņa',
  'Anda | Restorāns | Maiņa',
  'Ligita | Restorāns | Maiņa',
  'Sanita | Restorāns | Viesmīle',
  'Liliāna | Restorāns | Viesmīle',
  'Aleksis | Restorāns | Bārmenis',
  'Amanda | Restorāns | Maiņa',
  'Kristīne | Restorāns | Maiņa',
  'Alise | Restorāns | Viesmīle',
  'Iveta | Restorāns | Viesmīle',
  'Līva | Restorāns | Viesmīle',
  'Gustavs | Restorāns | Maiņa',
  'Daiga | Virtuve | Pavāre',
  'Dace | Virtuve | Pavāre',
  'Edīte | Virtuve | Palīgs',
  'Elīna | Virtuve | Pavāre',
  'Ance | Virtuve | Maiņa',
  'Anete | Virtuve | Maiņa',
  'Jānis | Virtuve | Maiņa',
  'Marguss | Virtuve | Maiņa',
  'Marija | Virtuve | Pavāre',
];

export const DEFAULT_STAFF_CONFIG = DEFAULT_STAFF.join('\n');
export const DEFAULT_STAFF_NOTE =
  'Sākuma saraksts balstīts uz vēsturiskajiem ruta, anita, virtuve un darbin ierakstiem. Ja vajag, to var rediģēt.';

type StaffMember = {
  id: string;
  name: string;
  dept: string;
  role: string;
};

export type PlanType = 'viesnica' | 'restor' | 'virtuve';

export const PLAN_TYPE_OPTIONS: Array<{ value: PlanType; label: string; fileName: string }> = [
  { value: 'viesnica', label: 'Viesnīca', fileName: 'YYYY_MM_viesnica.xlsx' },
  { value: 'restor', label: 'Restorāns', fileName: 'YYYY_MM_restorans.xlsx' },
  { value: 'virtuve', label: 'Virtuve', fileName: 'YYYY_MM_virtuve.xlsx' },
];

type MonthBlock = {
  headerRowIndex: number;
  dayRowIndex: number;
  dayNumber: number;
  endRow: number;
  staffRowIndex: number | null;
  reservedText: string;
  reservationRows: number[];
};

type MetaRow = {
  staffId: string;
  staffName: string;
  dept: string;
  role: string;
  startRow: number;
  endRow: number;
  hoursRow: number;
  noteRow: number;
};

type ParsedMeta = {
  sheetName: string;
  monthKey: string;
  planType: PlanType;
  dayColumnOffset: number;
  rowMap: MetaRow[];
};

type WritebackOperation = {
  dayNumber: number;
  targetRowIndex: number;
  insertRow: boolean;
  targetColumn: number;
  summary: string;
};

export type GoogleSheetWritebackPlan = {
  sheetName: string;
  operations: WritebackOperation[];
  appliedDays: number;
};

type ExcelModule = Awaited<typeof import('exceljs')>;
type SummaryRowDefinition = {
  label: string;
  buildValue: (sheet: Worksheet, block: MonthBlock) => string | number;
};
type PlanTemplate = {
  planType: PlanType;
  sheetTitlePrefix: string;
  summaryRows: SummaryRowDefinition[];
  staffStartRow: number;
  dayColumnOffset: number;
  helperRow?: {
    rowNumber: number;
    buildValue: (sheet: Worksheet, block: MonthBlock) => string;
    highlightDay: (sheet: Worksheet, block: MonthBlock) => boolean;
  };
  footerLegendRows?: string[];
  hideStaffMetaLabels?: boolean;
  staffNameTransform?: (name: string) => string;
};

async function getExcelJsModule(): Promise<ExcelModule> {
  return import('exceljs');
}

async function createWorkbook(): Promise<Workbook> {
  const exceljs: any = await getExcelJsModule();
  const WorkbookCtor = exceljs.Workbook ?? exceljs.default?.Workbook;

  if (!WorkbookCtor) {
    throw new Error('Neizdevās ielādēt Excel apstrādes bibliotēku.');
  }

  return new WorkbookCtor();
}

function cellText(cell: Cell | null | undefined): string {
  const value = cell?.value;

  if (value == null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim();
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') return value.text.trim();
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('').trim();
    }
    if ('result' in value && value.result != null) return String(value.result).trim();
    if ('formula' in value) return String(value.result ?? `=${value.formula}`).trim();
    if ('hyperlink' in value) return String(value.text || value.hyperlink).trim();
  }

  return String(value).trim();
}

function normalizeText(value: string): string {
  return String(value || '').trim().toLowerCase();
}

export function parseMonthSheetName(sheetName: string): string | null {
  return parseMonthKeyFromSheetName(sheetName);
}

export function getMonthSheets(workbook: Workbook): string[] {
  return workbook.worksheets
    .map((sheet) => sheet.name)
    .filter((name) => parseMonthSheetName(name))
    .sort((left, right) => left.localeCompare(right));
}

export async function readWorkbookFromFile(file: File): Promise<Workbook> {
  const workbook = await createWorkbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  return workbook;
}

async function cloneWorkbook(workbook: Workbook): Promise<Workbook> {
  const clone = await createWorkbook();
  const buffer = await workbook.xlsx.writeBuffer();
  await clone.xlsx.load(buffer);
  return clone;
}

export function parseStaffConfig(text: string): StaffMember[] {
  return text
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, index) => {
      const [name, dept = 'Citi', role = 'Darbinieks'] = line.split('|').map((part) => part.trim());

      return {
        id: `staff_${index + 1}`,
        name,
        dept,
        role,
      };
    });
}

function filterStaffForPlan(staff: StaffMember[], planType: PlanType): StaffMember[] {
  const filtered = staff.filter((member) => {
    const dept = normalizeText(member.dept);
    const role = normalizeText(member.role);

    switch (planType) {
      case 'restor':
        return dept === 'restor' || dept === 'restorāns' || dept === 'restorans';
      case 'virtuve':
        return dept === 'virtuve';
      case 'viesnica':
      default:
        return (
          dept === 'viesnica' ||
          dept === 'viesnīca' ||
          dept === 'vadība' ||
          dept === 'vadiba' ||
          dept === 'reģistratūra' ||
          dept === 'registratūra' ||
          dept === 'registratura' ||
          dept === 'numuri' ||
          (dept === 'virtuve' && role.includes('brokast'))
        );
    }
  });

  return filtered.length ? filtered : staff;
}

function parseMonthBlocks(sheet: Worksheet): MonthBlock[] {
  const blocks: MonthBlock[] = [];

  for (let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const firstCell = normalizeText(cellText(sheet.getRow(rowIndex).getCell(1)));
    if (firstCell !== 'dat') continue;

    let dayRowIndex = rowIndex + 1;
    let rawDay = cellText(sheet.getRow(dayRowIndex).getCell(1));
    for (let candidate = rowIndex + 1; candidate <= Math.min(rowIndex + 3, sheet.rowCount); candidate += 1) {
      const candidateDay = cellText(sheet.getRow(candidate).getCell(1));
      if (/\d+/.test(candidateDay)) {
        dayRowIndex = candidate;
        rawDay = candidateDay;
        break;
      }
    }
    const dayNumber = Number(String(rawDay).match(/\d+/)?.[0]);
    if (!dayNumber) continue;

    let nextHeader = sheet.rowCount + 1;
    for (let pointer = dayRowIndex + 1; pointer <= sheet.rowCount; pointer += 1) {
      if (normalizeText(cellText(sheet.getRow(pointer).getCell(1))) === 'dat') {
        nextHeader = pointer;
        break;
      }
    }

    const endRow = nextHeader - 1;
    let staffRowIndex: number | null = null;
    const reservationRows: number[] = [];
    const dataStartRow = Math.min(dayRowIndex + 3, endRow + 1);

    let reservedText = '';
    for (let candidate = rowIndex + 1; candidate < dataStartRow; candidate += 1) {
      const candidateText = cellText(sheet.getRow(candidate).getCell(2));
      if (/^rezervē/i.test(candidateText.trim())) {
        reservedText = candidateText;
        break;
      }
    }
    if (!reservedText) {
      reservedText = cellText(sheet.getRow(dayRowIndex).getCell(2));
    }

    for (let pointer = rowIndex + 1; pointer <= endRow; pointer += 1) {
      const row = sheet.getRow(pointer);
      const first = normalizeText(cellText(row.getCell(1)));
      const values = [1, 2, 3, 4, 5, 6, 7, 8].map((column) => cellText(row.getCell(column)));
      const hasData = values.some(Boolean);

      if (first === 'darbin') {
        staffRowIndex = pointer;
      } else if (pointer >= dataStartRow && hasData) {
        reservationRows.push(pointer);
      }
    }

    blocks.push({
      headerRowIndex: rowIndex,
      dayRowIndex,
      dayNumber,
      endRow,
      staffRowIndex,
      reservedText,
      reservationRows,
    });

    rowIndex = endRow;
  }

  return blocks;
}

function buildDayEventSummary(sheet: Worksheet, block: MonthBlock): string {
  const items: string[] = [];

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const room = cellText(row.getCell(3));
    const guest = cellText(row.getCell(4));
    const parsed = parseActivity(activity, guest);

    if (
      !activity ||
      ![
        'restaurant',
        'lunch',
        'excursion',
        'horse',
        'ceremony',
        'seminar',
        'unknown',
      ].includes(parsed.category)
    ) {
      continue;
    }

    items.push(`${activity}${guest ? ` · ${guest}` : ''}`);
  }

  return items.join(' | ');
}

function parseRoomTokens(value: string): string[] {
  const cleaned = String(value || '').trim();
  if (!cleaned) return [];

  const matches = cleaned.match(/\??\d+\s*[a-z]{0,3}/gi) ?? [];
  const tokens: string[] = [];
  const seen = new Set<string>();

  for (const match of matches) {
    const token = match.replace(/\?/g, '').replace(/\s+/g, '').toLowerCase();
    if (!token || seen.has(token)) continue;
    seen.add(token);
    tokens.push(token);
  }

  return tokens;
}

function parseBlockedRoomSet(reservedText: string): Set<string> {
  const normalized = reservedText.replace(/^rezervē\s*/i, '').trim();
  return new Set(parseRoomTokens(normalized));
}

function buildBroadRoomNote(activity: string, roomText: string): string | null {
  const normalized = normalizeText(`${activity} ${roomText}`);
  if (!normalized) return null;

  if (normalized.includes('aizņemtas visas pārejās istabas') || normalized.includes('visas ist pil')) {
    return 'visa pils un dm';
  }

  if (!normalized.includes('visa pils')) {
    return null;
  }

  const excludedRooms = parseRoomTokens(roomText || activity);
  if (normalized.includes('izņemot') && excludedRooms.length) {
    return `visa pils (izņemot ${excludedRooms.join(', ')})`;
  }

  return 'visa pils';
}

function getHotelRoomMovement(activity: string, guest: string): 'arrival' | 'stayover' | null {
  const normalized = normalizeText(activity);

  if (normalized.includes('turp')) {
    return 'stayover';
  }

  if (normalized.includes('iebr')) {
    return 'arrival';
  }

  const parsed = parseActivity(activity, guest);
  if (parsed.category === 'continueStay') return 'stayover';
  if (parsed.category === 'checkin') return 'arrival';

  return null;
}

function collectHotelRoomSignals(sheet: Worksheet, block: MonthBlock) {
  const arrivalRooms: string[] = [];
  const stayoverRooms: string[] = [];
  const broadNotes: string[] = [];
  const arrivalSeen = new Set<string>();
  const stayoverSeen = new Set<string>();
  const noteSeen = new Set<string>();

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const roomText = cellText(row.getCell(3));
    const guest = cellText(row.getCell(4));
    const broadNote = buildBroadRoomNote(activity, roomText);
    const movement = getHotelRoomMovement(activity, guest);

    if (broadNote && !noteSeen.has(broadNote)) {
      noteSeen.add(broadNote);
      broadNotes.push(broadNote);
    }

    const roomTokens = parseRoomTokens(roomText);
    if (!roomTokens.length) continue;

    if (movement === 'stayover') {
      for (const room of roomTokens) {
        if (stayoverSeen.has(room)) continue;
        stayoverSeen.add(room);
        stayoverRooms.push(room);
      }
      continue;
    }

    if (movement === 'arrival') {
      for (const room of roomTokens) {
        if (arrivalSeen.has(room)) continue;
        arrivalSeen.add(room);
        arrivalRooms.push(room);
      }
    }
  }

  return {
    arrivalRooms,
    stayoverRooms,
    broadNotes,
    roomCount: new Set([...arrivalRooms, ...stayoverRooms]).size,
  };
}

function buildHotelSignals(sheet: Worksheet, block: MonthBlock) {
  const specialLabels: string[] = [];
  const seenLabels = new Set<string>();
  let eventGuests = 0;
  let spaCount = 0;
  let checkinCount = 0;
  let breakfastCount = 0;
  const roomSignals = collectHotelRoomSignals(sheet, block);

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const room = cellText(row.getCell(3));
    const guest = cellText(row.getCell(4));
    const parsed = parseActivity(activity, guest);
    const normalizedActivity = normalizeText(activity);
    const guestCount = parsed.guestCount || extractGuestCount(guest) || extractGuestCount(activity);

    if (parsed.category === 'spa') {
      spaCount += 1;
    }

    if (parsed.category === 'checkin') {
      checkinCount += 1;
    }

    if (parsed.category === 'breakfast') {
      breakfastCount += guestCount || 1;
    }

    if (['restaurant', 'dinner', 'lunch', 'ceremony', 'seminar'].includes(parsed.category)) {
      const countsAsSharedEvent =
        !room ||
        guestCount >= 6 ||
        ['ceremony', 'seminar'].includes(parsed.category) ||
        /telts|banket|pasāk|pasak|laul|krist/.test(normalizedActivity);

      if (!countsAsSharedEvent) {
        continue;
      }

      eventGuests += guestCount;

      if (/telts|banket|pasāk|pasak|laul|krist|semin/.test(normalizedActivity)) {
        const label = activity.trim();
        if (label && !seenLabels.has(label)) {
          specialLabels.push(label);
          seenLabels.add(label);
        }
      }
    }
  }

  return {
    roomCount: roomSignals.roomCount,
    eventGuests,
    spaCount,
    checkinCount,
    breakfastCount,
    specialLabels,
  };
}

function buildHotelEventSummary(sheet: Worksheet, block: MonthBlock): string {
  const signals = buildHotelSignals(sheet, block);
  const items: string[] = [];

  if (signals.eventGuests > 0) {
    items.push(`${signals.eventGuests}p`);
  }

  signals.specialLabels.forEach((label) => {
    if (!items.includes(label)) {
      items.push(label);
    }
  });

  return items.join(' | ');
}

function buildDaySpaSummary(sheet: Worksheet, block: MonthBlock): string {
  const items: string[] = [];

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const guest = cellText(row.getCell(4));
    const parsed = parseActivity(activity, guest);

    if (parsed.category === 'spa') {
      items.push(activity);
    }
  }

  return items.join(' | ');
}

function buildDayRoomSummary(sheet: Worksheet, block: MonthBlock): string {
  const roomSignals = collectHotelRoomSignals(sheet, block);
  const items: string[] = [];

  if (roomSignals.stayoverRooms.length) {
    items.push(`(${roomSignals.stayoverRooms.join(' ')})`);
  }

  if (roomSignals.arrivalRooms.length) {
    items.push(roomSignals.arrivalRooms.join(' '));
  }

  for (const note of roomSignals.broadNotes) {
    items.push(items.length ? `+ ${note}` : note);
  }

  return items.join(' ');
}

function buildHotelHelperMarker(sheet: Worksheet, block: MonthBlock): string {
  const signals = buildHotelSignals(sheet, block);
  return signals.roomCount >= 5 || signals.eventGuests >= 10 || signals.spaCount > 0 ? 'x' : '';
}

function shouldHighlightHotelDay(sheet: Worksheet, block: MonthBlock): boolean {
  const signals = buildHotelSignals(sheet, block);
  return (
    signals.roomCount >= 4 ||
    signals.eventGuests >= 8 ||
    signals.spaCount > 0 ||
    signals.checkinCount >= 2 ||
    signals.breakfastCount >= 6
  );
}

function buildGenericEventSummary(sheet: Worksheet, block: MonthBlock, matcher?: (activity: string) => boolean): string {
  const items: string[] = [];

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const guest = cellText(row.getCell(4));
    const parsed = parseActivity(activity, guest);

    if (!activity) continue;
    if (matcher && !matcher(parsed.normalized)) continue;

    items.push(`${activity}${guest ? ` · ${guest}` : ''}`);
  }

  return items.join(' | ');
}

function buildUniqueGuestRows(
  sheet: Worksheet,
  block: MonthBlock,
  includeRow: (
    parsed: ReturnType<typeof parseActivity>,
    guest: string,
    room: string,
    bk: string,
    activity: string,
  ) => boolean,
): Array<{ activity: string; guest: string; room: string; bk: string }> {
  const seen = new Set<string>();
  const rows: Array<{ activity: string; guest: string; room: string; bk: string }> = [];

  for (const rowIndex of block.reservationRows) {
    const row = sheet.getRow(rowIndex);
    const activity = cellText(row.getCell(2));
    const room = cellText(row.getCell(3));
    const guest = cellText(row.getCell(4));
    const bk = cellText(row.getCell(5));
    const parsed = parseActivity(activity, guest);

    if (!includeRow(parsed, guest, room, normalizeText(bk), activity)) continue;

    const dedupeKey = `${room}|${guest}`.trim();
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);
    rows.push({ activity, guest, room, bk });
  }

  return rows;
}

function extractGuestCount(text: string): number {
  const matches = text.match(/\d+/g);
  if (!matches) return 0;
  return matches.map(Number).reduce((sum, value) => sum + value, 0);
}

function buildRestaurantGuestCount(sheet: Worksheet, block: MonthBlock): number {
  const rows = buildUniqueGuestRows(sheet, block, (parsed, guest, room) => {
    if (!guest && !parsed.display) return false;
    return ['restaurant', 'dinner', 'lunch', 'ceremony', 'seminar'].includes(parsed.category) && (!room || !!guest);
  });

  let total = 0;
  for (const row of rows) {
    total += extractGuestCount(row.guest) || extractGuestCount(row.activity);
  }

  return total;
}

function buildBreakfastGuestCount(sheet: Worksheet, block: MonthBlock): number {
  const rows = buildUniqueGuestRows(sheet, block, (parsed, guest, room, bk, activity) => {
    if (!room && !guest) return false;
    const normalizedActivity = normalizeText(activity);
    return parsed.category === 'breakfast' || normalizedActivity.includes('brok') || bk.includes('b');
  });

  return rows.reduce((sum, row) => sum + extractGuestCount(row.guest), 0);
}

function buildHotelGuestCount(sheet: Worksheet, block: MonthBlock): number {
  const rows = buildUniqueGuestRows(sheet, block, (parsed, guest, room) => {
    if (!room || !guest) return false;
    return ['checkin', 'continueStay', 'breakfast', 'dinner'].includes(parsed.category) || parsed.category === 'unknown';
  });

  return rows.reduce((sum, row) => sum + extractGuestCount(row.guest), 0);
}

function getPlanTemplate(planType: PlanType): PlanTemplate {
  switch (planType) {
    case 'restor':
      return {
        planType,
        sheetTitlePrefix: 'Restorāns',
        summaryRows: [
          { label: 'restor pers sk', buildValue: buildRestaurantGuestCount },
          { label: 'vies n sk', buildValue: buildHotelGuestCount },
          { label: 'brok pers sk', buildValue: buildBreakfastGuestCount },
        ],
        staffStartRow: 8,
        dayColumnOffset: 1,
        hideStaffMetaLabels: true,
      };
    case 'virtuve':
      return {
        planType,
        sheetTitlePrefix: 'Virtuve',
        summaryRows: [
          { label: 'Brokastis', buildValue: buildBreakfastGuestCount },
          { label: 'Restorāns', buildValue: buildRestaurantGuestCount },
          {
            label: 'pasākums',
            buildValue: (sheet, block) =>
              buildGenericEventSummary(
                sheet,
                block,
                (activity) => /rest|vak|pusd|banket|laul|krist|pasāk|ēdin/.test(activity),
              ),
          },
        ],
        staffStartRow: 7,
        dayColumnOffset: 1,
        hideStaffMetaLabels: true,
      };
    case 'viesnica':
    default:
      return {
        planType: 'viesnica',
        sheetTitlePrefix: 'Viesnīca',
        summaryRows: [
          { label: 'Pasākums', buildValue: buildHotelEventSummary },
          { label: 'pirts beidzas', buildValue: buildDaySpaSummary },
          { label: 'istabas', buildValue: buildDayRoomSummary },
        ],
        staffStartRow: 7,
        dayColumnOffset: 1,
        helperRow: {
          rowNumber: 6,
          buildValue: buildHotelHelperMarker,
          highlightDay: shouldHighlightHotelDay,
        },
        footerLegendRows: [
          'ist - istabu kopšana',
          'p - pirtis kopšana',
          'br - brokastu stundas reģistratūras grafikā',
          'x - brīvs',
          '(16) - istaba uz otru nakti',
        ],
        hideStaffMetaLabels: true,
        staffNameTransform: (name) => name.toUpperCase(),
      };
  }
}

function applyWeekendFill(cell: Cell) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF92D050' },
  };
}

function applyMutedFill(cell: Cell) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF3F3F3' },
  };
}

function applyBorder(cell: Cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
  };
}

function getDayColumnIndex(dayNumber: number, dayColumnOffset: number): number {
  return dayNumber + dayColumnOffset;
}

function getTotalColumnIndex(dayCount: number, dayColumnOffset: number): number {
  return getDayColumnIndex(dayCount, dayColumnOffset) + 1;
}

export function buildPlanFileName(monthKey: string, planType: PlanType): string {
  const [year, month] = monthKey.split('-');
  const typePart = planType === 'restor' ? 'restorans' : planType;
  return `${year}_${month}_${typePart}.xlsx`;
}

function getPlanWritebackLabel(planType: PlanType): string {
  switch (planType) {
    case 'restor':
      return 'Restorāns';
    case 'virtuve':
      return 'Virtuve';
    case 'viesnica':
    default:
      return 'Viesnīca';
  }
}

function getPlanWritebackColumn(planType: PlanType): number {
  switch (planType) {
    case 'restor':
      return 4;
    case 'virtuve':
      return 6;
    case 'viesnica':
    default:
      return 2;
  }
}

function applyClassicPlanStyles(
  rotaSheet: Worksheet,
  sourceSheet: Worksheet,
  blocks: MonthBlock[],
  staff: StaffMember[],
  template: PlanTemplate,
  monthKey: string,
  dateRowNumber: number,
  footerStartRow: number,
) {
  const dayCount = getDaysInMonth(monthKey);
  const totalColumnIndex = getTotalColumnIndex(dayCount, template.dayColumnOffset);
  rotaSheet.views = [
    {
      state: 'normal',
      showGridLines: true,
      zoomScale: 47,
      zoomScaleNormal: 100,
    },
  ];

  rotaSheet.getCell('A1').value = buildLatvianMonthLabel(monthKey);
  rotaSheet.getCell('A1').font = { bold: true, size: 16, name: 'Calibri' };

  rotaSheet.getColumn(1).width = 13.9;
  for (let dayNumber = 1; dayNumber <= dayCount; dayNumber += 1) {
    rotaSheet.getColumn(getDayColumnIndex(dayNumber, template.dayColumnOffset)).width = 7.6;
  }
  rotaSheet.getColumn(totalColumnIndex).width = 7.6;

  const topHeights = new Map<number, number>([
    [1, 22.5],
    [2, 66.75],
    [3, 81.75],
    [4, 108.75],
    [5, 29.25],
    [6, 19.5],
  ]);
  topHeights.forEach((height, rowNumber) => {
    rotaSheet.getRow(rowNumber).height = height;
  });

  template.summaryRows.forEach((_, index) => {
    const cell = rotaSheet.getCell(`A${index + 2}`);
    cell.font = { bold: true, size: 16, name: 'Calibri' };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true, textRotation: 90 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
    };
  });

  const dateHeader = rotaSheet.getCell(`A${dateRowNumber}`);
  dateHeader.font = { bold: true, size: 16, name: 'Calibri' };
  dateHeader.alignment = { horizontal: 'center', vertical: 'middle' };
  dateHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E2F3' } };
  dateHeader.border = {
    top: { style: 'thick', color: { argb: 'FF000000' } },
    left: { style: 'thick', color: { argb: 'FF000000' } },
    right: { style: 'medium', color: { argb: 'FF000000' } },
    bottom: { style: 'thick', color: { argb: 'FF000000' } },
  };

  const totalHeader = rotaSheet.getRow(dateRowNumber).getCell(totalColumnIndex);
  totalHeader.value = 'KOPĀ';
  totalHeader.font = { bold: true, size: 16, name: 'Calibri' };
  totalHeader.alignment = { horizontal: 'center', vertical: 'middle' };
  totalHeader.border = {
    top: { style: 'thick', color: { argb: 'FF000000' } },
    left: { style: 'hair', color: { argb: 'FF000000' } },
    right: { style: 'medium', color: { argb: 'FF000000' } },
    bottom: { style: 'thick', color: { argb: 'FF000000' } },
  };

  if (template.helperRow) {
    const helperCell = rotaSheet.getCell(`A${template.helperRow.rowNumber}`);
    helperCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' }, name: 'Calibri' };
    helperCell.alignment = { horizontal: 'center', vertical: 'middle' };
    helperCell.border = {
      top: { style: 'thick', color: { argb: 'FF000000' } },
      left: { style: 'thick', color: { argb: 'FF000000' } },
      right: { style: 'hair', color: { argb: 'FF000000' } },
      bottom: { style: 'hair', color: { argb: 'FF000000' } },
    };
  }

  staff.forEach((_, index) => {
    const startRow = template.staffStartRow + index * 4;
    const endRow = startRow + 1;
    const hoursRow = startRow + 2;
    const noteRow = startRow + 3;

    rotaSheet.getRow(startRow).height = index === 0 ? 23.25 : 22.5;
    rotaSheet.getRow(endRow).height = 23.25;
    rotaSheet.getRow(hoursRow).height = 19.5;
    rotaSheet.getRow(noteRow).height = 54;

    const nameCell = rotaSheet.getCell(`A${startRow}`);
    nameCell.font = { bold: true, size: 16, name: 'Calibri' };
    nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
    nameCell.border = {
      left: { style: index === 0 ? 'thick' : 'hair', color: { argb: 'FF000000' } },
    };
    if (index > 0) {
      nameCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F3F3' } };
    }

    const noteCell = rotaSheet.getCell(`A${noteRow}`);
    noteCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F3F3' } };
    noteCell.border = {
      top: { style: 'medium', color: { argb: 'FF000000' } },
      left: { style: 'hair', color: { argb: 'FF000000' } },
      right: { style: 'hair', color: { argb: 'FF000000' } },
      bottom: { style: 'hair', color: { argb: 'FF000000' } },
    };
    noteCell.alignment = { wrapText: true };

    for (let dayNumber = 1; dayNumber <= dayCount; dayNumber += 1) {
      const columnIndex = getDayColumnIndex(dayNumber, template.dayColumnOffset);
      const shouldHighlight = isWeekendDay(monthKey, dayNumber);

      [startRow, endRow, hoursRow, noteRow].forEach((rowNumber) => {
        const cell = rotaSheet.getRow(rowNumber).getCell(columnIndex);
        cell.font = {
          size: rowNumber === noteRow || rowNumber === hoursRow ? 16 : 16,
          bold: rowNumber === noteRow,
          name: rowNumber === noteRow ? 'Calibri' : 'Calibri',
        };
        cell.alignment = {
          horizontal: 'center',
          vertical: rowNumber === noteRow ? 'middle' : undefined,
          wrapText: rowNumber === noteRow,
        };
        cell.border = {
          top: { style: rowNumber === startRow ? 'thick' : rowNumber === noteRow ? 'medium' : 'hair', color: { argb: 'FF000000' } },
          left: { style: 'hair', color: { argb: 'FF000000' } },
          right: { style: 'hair', color: { argb: 'FF000000' } },
          bottom: { style: rowNumber === noteRow ? 'hair' : 'hair', color: { argb: 'FF000000' } },
        };
        if (shouldHighlight) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
        }
      });
    }

    const totalCell = rotaSheet.getRow(hoursRow).getCell(totalColumnIndex);
    totalCell.font = { size: 18, name: 'Arial' };
    totalCell.alignment = { horizontal: 'center', vertical: 'middle' };
    totalCell.border = {
      left: { style: 'hair', color: { argb: 'FF000000' } },
      right: { style: 'medium', color: { argb: 'FF000000' } },
      top: { style: 'hair', color: { argb: 'FF000000' } },
      bottom: { style: 'hair', color: { argb: 'FF000000' } },
    };
  });

  rotaSheet.getRow(footerStartRow).height = 29.25;
  const footerDateCell = rotaSheet.getCell(`A${footerStartRow + 1}`);
  footerDateCell.value = buildLatvianMonthLabel(monthKey, true);
  footerDateCell.font = { size: 16, name: 'Calibri' };

  const legendColumn = 4;
  (template.footerLegendRows ?? []).forEach((legend, index) => {
    const rowNumber = footerStartRow + index + 1;
    if (index > 0) {
      rotaSheet.getCell(`A${rowNumber}`).value = '';
    }
    const cell = rotaSheet.getRow(rowNumber).getCell(legendColumn);
    cell.value = legend;
    cell.font = {
      size: index === template.footerLegendRows!.length - 1 ? 14 : 16,
      name: index === template.footerLegendRows!.length - 1 ? 'Arial' : 'Calibri',
    };
  });
}

function createMetaSheet(workbook: Workbook, meta: ParsedMeta) {
  const metaSheet = workbook.addWorksheet('__meta', { state: 'veryHidden' });
  metaSheet.addRow(['template', 'viesu_logic_v1']);
  metaSheet.addRow(['sheetName', meta.sheetName]);
  metaSheet.addRow(['monthKey', meta.monthKey]);
  metaSheet.addRow(['planType', meta.planType]);
  metaSheet.addRow(['dayColumnOffset', meta.dayColumnOffset]);
  metaSheet.addRow(['staffId', 'staffName', 'dept', 'role', 'startRow', 'endRow', 'hoursRow', 'noteRow']);

  meta.rowMap.forEach((row) => {
    metaSheet.addRow([
      row.staffId,
      row.staffName,
      row.dept,
      row.role,
      row.startRow,
      row.endRow,
      row.hoursRow,
      row.noteRow,
    ]);
  });
}

function readMeta(timesheetWorkbook: Workbook): ParsedMeta | null {
  const metaSheet = timesheetWorkbook.getWorksheet('__meta');
  if (!metaSheet) return null;

  const template = cellText(metaSheet.getRow(1).getCell(2));
  if (template !== 'viesu_logic_v1') return null;

  const sheetName = cellText(metaSheet.getRow(2).getCell(2));
  const monthKey = cellText(metaSheet.getRow(3).getCell(2));
  const row4Key = cellText(metaSheet.getRow(4).getCell(1));
  const rawPlanType = row4Key === 'planType' ? cellText(metaSheet.getRow(4).getCell(2)) : '';
  const planType: PlanType =
    rawPlanType === 'restor' || rawPlanType === 'virtuve' || rawPlanType === 'viesnica' ? rawPlanType : 'viesnica';
  const dayColumnOffset =
    Number(cellText(metaSheet.getRow(row4Key === 'planType' ? 5 : 4).getCell(2))) || 3;
  const rowMapStart = row4Key === 'planType' ? 7 : 6;
  const rowMap: MetaRow[] = [];

  for (let rowIndex = rowMapStart; rowIndex <= metaSheet.rowCount; rowIndex += 1) {
    const row = metaSheet.getRow(rowIndex);
    const staffId = cellText(row.getCell(1));
    if (!staffId) continue;

    rowMap.push({
      staffId,
      staffName: cellText(row.getCell(2)),
      dept: cellText(row.getCell(3)),
      role: cellText(row.getCell(4)),
      startRow: Number(cellText(row.getCell(5))),
      endRow: Number(cellText(row.getCell(6))),
      hoursRow: Number(cellText(row.getCell(7))),
      noteRow: Number(cellText(row.getCell(8))),
    });
  }

  return { sheetName, monthKey, planType, dayColumnOffset, rowMap };
}

function buildDayPlans(timesheetWorkbook: Workbook, staffList: StaffMember[]) {
  const meta = readMeta(timesheetWorkbook);
  if (!meta) {
    throw new Error('Šis nav sistēmas ģenerēts darba plāna fails.');
  }

  const rotaSheet = timesheetWorkbook.getWorksheet('ruta');
  if (!rotaSheet) {
    throw new Error('Timesheet failā nav atrasta lapa "ruta".');
  }

  const grouped = new Map<string, string[]>();
  const dayCount = getDaysInMonth(meta.monthKey);

  meta.rowMap.forEach((rowMap) => {
    const staff = staffList.find((item) => item.id === rowMap.staffId) || rowMap;

    for (let day = 1; day <= dayCount; day += 1) {
      const columnIndex = getDayColumnIndex(day, meta.dayColumnOffset);
      const start = cellText(rotaSheet.getRow(rowMap.startRow).getCell(columnIndex));
      const end = cellText(rotaSheet.getRow(rowMap.endRow).getCell(columnIndex));
      const note = cellText(rotaSheet.getRow(rowMap.noteRow).getCell(columnIndex));
      const normalizedNote = normalizeText(note);

      if (!start && !end) continue;

      const staffLabel = 'name' in staff ? staff.name : staff.staffName;
      const label = `${staffLabel} ${start && end ? `${start}-${end}` : start || end || ''}${
        note && normalizedNote !== 'x' ? ` (${note})` : ''
      }`.trim();
      const key = String(day);

      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key)!.push(label);
    }
  });

  return { meta, grouped };
}

function writeStaffSummary(row: ReturnType<Worksheet['getRow']>, targetColumn: number, summary: string) {
  row.getCell(1).value = 'darbin';
  row.getCell(targetColumn).value = summary;
}

function buildWritebackOperations(
  originalWorkbook: Workbook,
  timesheetWorkbook: Workbook,
  staffConfigText: string,
): { sheetName: string; operations: WritebackOperation[] } {
  const staffList = parseStaffConfig(staffConfigText);
  const { meta, grouped } = buildDayPlans(timesheetWorkbook, staffList);
  const sourceSheet = originalWorkbook.getWorksheet(meta.sheetName);

  if (!sourceSheet) {
    throw new Error(`Oriģinālajā failā nav atrasta lapa "${meta.sheetName}".`);
  }

  const blocks = parseMonthBlocks(sourceSheet);
  const operations: WritebackOperation[] = [];

  blocks.forEach((block) => {
    const entries = grouped.get(String(block.dayNumber));
    if (!entries?.length) return;

    const targetColumn = getPlanWritebackColumn(meta.planType);
    const summary = `${getPlanWritebackLabel(meta.planType)}: ${entries.join(', ')}`;
    const targetRowIndex = block.staffRowIndex ?? block.endRow + 1;
    operations.push({
      dayNumber: block.dayNumber,
      targetRowIndex,
      insertRow: block.staffRowIndex == null,
      targetColumn,
      summary,
    });
  });

  operations.sort((left, right) => right.targetRowIndex - left.targetRowIndex);

  return { sheetName: meta.sheetName, operations };
}

export async function generateTimesheetWorkbook(
  originalWorkbook: Workbook,
  sheetName: string,
  staffConfigText: string,
  planType: PlanType = 'viesnica',
): Promise<{ workbook: Workbook; fileName: string; dayCount: number }> {
  const monthKey = parseMonthSheetName(sheetName);
  if (!monthKey) {
    throw new Error('Neizdevās noteikt mēnesi no lapas nosaukuma.');
  }

  const sourceSheet = originalWorkbook.getWorksheet(sheetName);
  if (!sourceSheet) {
    throw new Error(`Oriģinālajā failā nav atrasta lapa "${sheetName}".`);
  }

  const blocks = parseMonthBlocks(sourceSheet);
  if (!blocks.length) {
    throw new Error('Izvēlētajā mēneša lapā netika atrasti dienu bloki.');
  }
  const blocksByDay = new Map(blocks.map((block) => [block.dayNumber, block]));
  const dayCount = getDaysInMonth(monthKey);

  const staff = filterStaffForPlan(parseStaffConfig(staffConfigText), planType);
  if (!staff.length) {
    throw new Error('Darbinieku saraksts ir tukšs.');
  }

  const workbook = await createWorkbook();
  const template = getPlanTemplate(planType);
  const rotaSheet = workbook.addWorksheet('ruta');
  const meta: ParsedMeta = { sheetName, monthKey, planType, dayColumnOffset: template.dayColumnOffset, rowMap: [] };
  const dateRowNumber = template.summaryRows.length + 2;
  rotaSheet.getCell('A1').value = buildLatvianMonthLabel(monthKey);
  rotaSheet.getCell('A1').font = { bold: true, size: 14 };

  template.summaryRows.forEach((rowDef, index) => {
    rotaSheet.getCell(`A${index + 2}`).value = rowDef.label;
  });
  rotaSheet.getCell(`A${dateRowNumber}`).value = 'DATUMS';
  if (template.helperRow) {
    applyBorder(rotaSheet.getCell(`A${template.helperRow.rowNumber}`));
  }

  const summaryRowNumbers = new Set(template.summaryRows.map((_, index) => index + 2));

  rotaSheet.getColumn(1).width = 16;
  rotaSheet.getColumn(2).width = 10;
  rotaSheet.getColumn(3).width = 14;

  for (let dayNumber = 1; dayNumber <= dayCount; dayNumber += 1) {
    const block = blocksByDay.get(dayNumber);
    const columnIndex = getDayColumnIndex(dayNumber, template.dayColumnOffset);
    rotaSheet.getColumn(columnIndex).width = 9;
    template.summaryRows.forEach((rowDef, index) => {
      rotaSheet.getRow(index + 2).getCell(columnIndex).value = block ? rowDef.buildValue(sourceSheet, block) : '';
    });
    rotaSheet.getRow(dateRowNumber).getCell(columnIndex).value = dayNumber;

    const rowNumbers = [...template.summaryRows.map((_, index) => index + 2), dateRowNumber];
    const shouldHighlight = isWeekendDay(monthKey, dayNumber);

    if (template.helperRow) {
      const helperCell = rotaSheet.getRow(template.helperRow.rowNumber).getCell(columnIndex);
      helperCell.value = block ? template.helperRow.buildValue(sourceSheet, block) : '';
      applyBorder(helperCell);
      helperCell.alignment = { vertical: 'middle', horizontal: 'center' };
      rowNumbers.push(template.helperRow.rowNumber);
    }

    rowNumbers.forEach((rowNumber) => {
      const cell = rotaSheet.getRow(rowNumber).getCell(columnIndex);
      applyBorder(cell);
      if (shouldHighlight) applyWeekendFill(cell);
      if (rowNumber !== dateRowNumber) {
        const isSummaryRow = summaryRowNumbers.has(rowNumber);
        cell.alignment = {
          vertical: isSummaryRow ? 'bottom' : rowNumber === template.helperRow?.rowNumber ? 'middle' : 'top',
          horizontal: isSummaryRow ? 'left' : 'center',
          wrapText: true,
          textRotation: isSummaryRow ? 90 : 0,
        };
      }
    });
  }

  const leftLabels = [...template.summaryRows.map((_, index) => `A${index + 2}`), `A${dateRowNumber}`];
  if (template.helperRow) {
    leftLabels.push(`A${template.helperRow.rowNumber}`);
  }

  leftLabels.forEach((address) => {
    const cell = rotaSheet.getCell(address);
    cell.font = { bold: true };
    applyBorder(cell);
    if (address !== `A${template.helperRow?.rowNumber}`) {
      applyMutedFill(cell);
    }
  });

  let currentRow = template.staffStartRow;
  const totalColumnIndex = getTotalColumnIndex(dayCount, template.dayColumnOffset);
  const firstDayColumnLetter = rotaSheet.getColumn(getDayColumnIndex(1, template.dayColumnOffset)).letter;
  const lastDayColumnLetter = rotaSheet.getColumn(getDayColumnIndex(dayCount, template.dayColumnOffset)).letter;
  staff.forEach((member) => {
    const startRow = currentRow;
    const endRow = currentRow + 1;
    const hoursRow = currentRow + 2;
    const noteRow = currentRow + 3;

    meta.rowMap.push({
      staffId: member.id,
      staffName: member.name,
      dept: member.dept,
      role: member.role,
      startRow,
      endRow,
      hoursRow,
      noteRow,
    });

    rotaSheet.getCell(`A${startRow}`).value = template.staffNameTransform?.(member.name) ?? member.name;
    rotaSheet.getCell(`A${startRow}`).font = { bold: true };

    if (!template.hideStaffMetaLabels) {
      rotaSheet.getCell(`A${noteRow}`).value = member.dept;
      rotaSheet.getCell(`B${noteRow}`).value = member.role;
    }

    [startRow, endRow, hoursRow, noteRow].forEach((rowNumber) => {
      ['A', 'B', 'C'].forEach((column) => {
        const cell = rotaSheet.getCell(`${column}${rowNumber}`);
        applyBorder(cell);
        if (column === 'A' || (!template.hideStaffMetaLabels && rowNumber === noteRow)) {
          applyMutedFill(cell);
        }
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
    });

    for (let dayNumber = 1; dayNumber <= dayCount; dayNumber += 1) {
      const columnIndex = getDayColumnIndex(dayNumber, template.dayColumnOffset);
      const shouldHighlight = isWeekendDay(monthKey, dayNumber);
      const columnLetter = rotaSheet.getColumn(columnIndex).letter;

      [startRow, endRow, hoursRow, noteRow].forEach((rowNumber) => {
        const cell = rotaSheet.getRow(rowNumber).getCell(columnIndex);
        applyBorder(cell);
        if (shouldHighlight) applyWeekendFill(cell);
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      rotaSheet.getRow(hoursRow).getCell(columnIndex).value = {
        formula: `IF(AND(ISNUMBER(${columnLetter}${startRow}),ISNUMBER(${columnLetter}${endRow})),(${columnLetter}${endRow}-${columnLetter}${startRow})*IF(AND(${columnLetter}${startRow}<1,${columnLetter}${endRow}<1),24,1),"")`,
      };
      rotaSheet.getRow(hoursRow).getCell(columnIndex).numFmt = '0.##';
    }

    rotaSheet.getRow(hoursRow).getCell(totalColumnIndex).value = {
      formula: `SUM(${firstDayColumnLetter}${hoursRow}:${lastDayColumnLetter}${hoursRow})`,
    };
    rotaSheet.getRow(hoursRow).getCell(totalColumnIndex).numFmt = '0.##';

    currentRow += 4;
  });

  if (template.footerLegendRows?.length) {
    rotaSheet.getCell(`A${currentRow}`).value = 'DATUMS';
    rotaSheet.getCell(`A${currentRow}`).font = { bold: true };
    applyBorder(rotaSheet.getCell(`A${currentRow}`));

    for (let dayNumber = 1; dayNumber <= dayCount; dayNumber += 1) {
      const columnIndex = getDayColumnIndex(dayNumber, template.dayColumnOffset);
      const cell = rotaSheet.getRow(currentRow).getCell(columnIndex);
      cell.value = dayNumber;
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      applyBorder(cell);
      if (isWeekendDay(monthKey, dayNumber)) {
        applyWeekendFill(cell);
      }
    }

    template.footerLegendRows.forEach((legend, index) => {
      const rowNumber = currentRow + index + 1;
      const cell = rotaSheet.getCell(`A${rowNumber}`);
      cell.value = legend;
      applyBorder(cell);
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      applyMutedFill(cell);
    });
  } else {
    rotaSheet.getCell(`A${currentRow + 1}`).value = 'Piezīme';
    rotaSheet.getCell(`B${currentRow + 1}`).value =
      'Aizpildi sākuma laiku, beigu laiku un, ja vajag, piezīmi ceturtajā rindā.';
    rotaSheet.mergeCells(currentRow + 1, 2, currentRow + 1, 8);
    applyBorder(rotaSheet.getCell(`A${currentRow + 1}`));
    applyBorder(rotaSheet.getCell(`B${currentRow + 1}`));
  }

  applyClassicPlanStyles(rotaSheet, sourceSheet, blocks, staff, template, monthKey, dateRowNumber, currentRow);

  createMetaSheet(workbook, meta);

  return {
    workbook,
    fileName: buildPlanFileName(monthKey, planType),
    dayCount,
  };
}

export async function applyTimesheetWorkbook(
  originalWorkbook: Workbook,
  originalFileName: string,
  timesheetWorkbook: Workbook,
  staffConfigText: string,
): Promise<{ workbook: Workbook; fileName: string; appliedDays: number }> {
  const workbookCopy = await cloneWorkbook(originalWorkbook);
  const { sheetName, operations } = buildWritebackOperations(workbookCopy, timesheetWorkbook, staffConfigText);
  const sourceSheet = workbookCopy.getWorksheet(sheetName);

  if (!sourceSheet) {
    throw new Error(`Oriģinālajā failā nav atrasta lapa "${sheetName}".`);
  }

  operations.forEach((operation) => {
    if (operation.insertRow) {
      sourceSheet.insertRow(operation.targetRowIndex, []);
    }

    const row = sourceSheet.getRow(operation.targetRowIndex);
    writeStaffSummary(row, operation.targetColumn, operation.summary);
  });

  const outputName = originalFileName.replace(/\.xlsx$/i, '') + '_updated.xlsx';

  return {
    workbook: workbookCopy,
    fileName: outputName,
    appliedDays: operations.length,
  };
}

export function buildGoogleSheetWritebackPlan(
  originalWorkbook: Workbook,
  timesheetWorkbook: Workbook,
  staffConfigText: string,
): GoogleSheetWritebackPlan {
  const { sheetName, operations } = buildWritebackOperations(originalWorkbook, timesheetWorkbook, staffConfigText);

  return {
    sheetName,
    operations,
    appliedDays: operations.length,
  };
}

export async function downloadWorkbook(workbook: Workbook, fileName: string): Promise<void> {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
