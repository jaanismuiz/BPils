const MONTH_MAP: Record<string, string> = {
  jan: '01',
  feb: '02',
  mar: '03',
  apr: '04',
  mai: '05',
  jun: '06',
  jul: '07',
  aug: '08',
  sep: '09',
  okt: '10',
  nov: '11',
  dec: '12',
};

const LATVIAN_MONTH_NAMES = [
  'Janvāris',
  'Februāris',
  'Marts',
  'Aprīlis',
  'Maijs',
  'Jūnijs',
  'Jūlijs',
  'Augusts',
  'Septembris',
  'Oktobris',
  'Novembris',
  'Decembris',
];

export function parseMonthKeyFromSheetName(sheetName: string): string | null {
  const match = String(sheetName)
    .trim()
    .toLowerCase()
    .match(/^([a-zāčēģīķļņšūž]{3})(\d{2})$/);

  if (!match) return null;

  const month = MONTH_MAP[match[1]];
  if (!month) return null;

  return `20${match[2]}-${month}`;
}

export function getDaysInMonth(monthKey: string): number {
  const [year, month] = monthKey.split('-').map(Number);
  return new Date(year, month, 0).getDate();
}

export function isWeekendDay(monthKey: string, dayNumber: number): boolean {
  const [year, month] = monthKey.split('-').map(Number);
  const weekday = new Date(year, month - 1, dayNumber).getDay();
  return weekday === 0 || weekday === 6;
}

export function buildLatvianMonthLabel(monthKey: string, includeYear = false): string {
  const [year, month] = monthKey.split('-').map(Number);
  const monthName = LATVIAN_MONTH_NAMES[month - 1] ?? monthKey;

  return includeYear ? `${monthName} ${year}` : monthName;
}
