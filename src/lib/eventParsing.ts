export type EventCategory =
  | 'breakfast'
  | 'checkin'
  | 'checkout'
  | 'continueStay'
  | 'dinner'
  | 'lunch'
  | 'restaurant'
  | 'spa'
  | 'excursion'
  | 'horse'
  | 'ceremony'
  | 'seminar'
  | 'reservationMeta'
  | 'admin'
  | 'unknown';

export type ParsedActivity = {
  normalized: string;
  display: string;
  category: EventCategory;
  guestCount: number;
  isOperationalEvent: boolean;
};

function normalizeActivity(raw: string): string {
  return raw
    .toLowerCase()
    .replace(/\u00d7/g, 'x')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function countGuests(text: string): number {
  const matches = text.match(/\d+/g);
  if (!matches) return 0;
  return matches.map(Number).reduce((sum, value) => sum + value, 0);
}

function isReservationMeta(normalized: string): boolean {
  return (
    normalized.startsWith('rezervē') ||
    /^(\d+[a-z]*)(,\d+[a-z\s]*)+$/.test(normalized.replace(/\s+/g, '')) ||
    normalized === 'laiks un darbība' ||
    normalized === 'pieejams' ||
    normalized === 'piejams' ||
    normalized.includes('nav pieej') ||
    normalized.includes('netirgojam') ||
    normalized.includes('pils slēgta') ||
    normalized.includes('numurs slēgts') ||
    /^(jan|feb|mar|apr|mai|jun|jul|aug|sep|okt|nov|dec)(\s+(jan|feb|mar|apr|mai|jun|jul|aug|sep|okt|nov|dec))*$/.test(
      normalized,
    ) ||
    normalized.startsWith('https://')
  );
}

function isAdminNoise(normalized: string): boolean {
  return (
    normalized === '[object object]' ||
    normalized === '180' ||
    normalized === '120' ||
    normalized === '95' ||
    normalized === '250' ||
    normalized === 'x pas' ||
    normalized === 'x sv. galds' ||
    normalized.includes('patst') ||
    normalized.includes('dzirnavas nav')
  );
}

export function parseActivity(activity: string, guestText = ''): ParsedActivity {
  const normalized = normalizeActivity(activity);
  const display = activity.trim();
  const guestCount = countGuests(guestText || activity);

  if (!normalized) {
    return {
      normalized,
      display,
      category: 'unknown',
      guestCount,
      isOperationalEvent: false,
    };
  }

  if (isReservationMeta(normalized)) {
    return { normalized, display, category: 'reservationMeta', guestCount, isOperationalEvent: false };
  }

  if (isAdminNoise(normalized)) {
    return { normalized, display, category: 'admin', guestCount, isOperationalEvent: false };
  }

  if (normalized.includes('turpina nakš') || normalized.includes('turpin nakš') || normalized.includes('turpina naš')) {
    return { normalized, display, category: 'continueStay', guestCount, isOperationalEvent: true };
  }

  if (normalized.includes('naksšņ')) {
    return { normalized, display, category: 'continueStay', guestCount, isOperationalEvent: true };
  }

  if (
    /\bbrok\b|\bbro\b/.test(normalized) ||
    normalized.includes('xbrok') ||
    normalized.includes('x rok') ||
    normalized.includes('9:00brok') ||
    normalized.includes('10:00brok') ||
    normalized.includes('11:00brok')
  ) {
    return { normalized, display, category: 'breakfast', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('iebr') ||
    normalized.includes('iebrauc') ||
    normalized.includes('xier') ||
    normalized.includes('x ieb') ||
    normalized.includes('x ier') ||
    normalized.includes('pārceļas no pils')
  ) {
    return { normalized, display, category: 'checkin', guestCount, isOperationalEvent: true };
  }

  if (normalized.includes('aizbrauc')) {
    return { normalized, display, category: 'checkout', guestCount, isOperationalEvent: true };
  }

  if (normalized.includes('vak') || normalized.includes('sv. galds') || normalized.includes('vest. uzkodas')) {
    return { normalized, display, category: 'dinner', guestCount, isOperationalEvent: true };
  }

  if (normalized.includes('pusd') || normalized.includes('pus.')) {
    return { normalized, display, category: 'lunch', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('rest') ||
    normalized.includes('res') ||
    normalized.includes('redt') ||
    normalized.includes('rets') ||
    normalized.includes('dzirnav') ||
    normalized.includes('banket') ||
    normalized.includes('kafijas pauze') ||
    normalized.includes('kaf p') ||
    normalized.includes('kaf.p') ||
    normalized.includes('kaf z')
  ) {
    return { normalized, display, category: 'restaurant', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('pirt') ||
    normalized.includes('pitis') ||
    normalized.includes('spa') ||
    normalized.includes('basein') ||
    normalized.includes('mas')
  ) {
    return { normalized, display, category: 'spa', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('izj') ||
    normalized.includes('kaman') ||
    normalized.includes('drošk') ||
    normalized.includes('pajūg') ||
    normalized.includes('zirg') ||
    normalized.includes('puskariet') ||
    normalized.includes('ponij')
  ) {
    return { normalized, display, category: 'horse', guestCount, isOperationalEvent: true };
  }

  if (normalized.includes('eksk') || normalized.includes('eks.') || normalized.includes('gidu') || normalized.includes('ekskurs')) {
    return { normalized, display, category: 'excursion', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('laul') ||
    normalized.includes('cer') ||
    normalized.includes('krist') ||
    normalized.includes('salūt') ||
    normalized.includes('kāz') ||
    normalized.includes('foto sesija') ||
    normalized.includes('ugunskurs') ||
    normalized.includes('vasarsvētki')
  ) {
    return { normalized, display, category: 'ceremony', guestCount, isOperationalEvent: true };
  }

  if (
    normalized.includes('sem') ||
    normalized.includes('pasāk') ||
    normalized.includes('pas.') ||
    /\bpas\b/.test(normalized) ||
    normalized.includes('pasut') ||
    normalized.includes('pasūt') ||
    normalized.includes('telts') ||
    normalized.includes('zāl') ||
    normalized.includes('balkon') ||
    normalized.includes('piknika vieta') ||
    normalized.includes('filmēt') ||
    normalized.includes('vizīte') ||
    normalized.includes('kopsapulce')
  ) {
    return { normalized, display, category: 'seminar', guestCount, isOperationalEvent: true };
  }

  return { normalized, display, category: 'unknown', guestCount, isOperationalEvent: false };
}
