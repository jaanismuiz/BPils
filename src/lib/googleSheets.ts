import type { Workbook } from 'exceljs';
import { buildGoogleSheetWritebackPlan, readWorkbookFromFile } from './workbookLogic';

const GOOGLE_IDENTITY_SCRIPT = 'https://accounts.google.com/gsi/client';
const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.readonly',
].join(' ');

type GoogleTokenResponse = {
  access_token: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

type GoogleTokenClient = {
  callback: ((response: GoogleTokenResponse) => void) | null;
  requestAccessToken: (options?: { prompt?: string }) => void;
};

type GoogleSheetsMetadata = {
  properties?: {
    title?: string;
  };
  sheets?: Array<{
    properties?: {
      sheetId?: number;
      title?: string;
      index?: number;
    };
  }>;
};

type TokenState = {
  accessToken: string;
  expiresAt: number;
};

declare global {
  interface Window {
    google?: {
      accounts?: {
        oauth2?: {
          initTokenClient: (config: {
            client_id: string;
            scope: string;
            callback: (response: GoogleTokenResponse) => void;
          }) => GoogleTokenClient;
        };
      };
    };
  }
}

let googleScriptPromise: Promise<void> | null = null;
let tokenState: TokenState | null = null;
let tokenClient: GoogleTokenClient | null = null;
let tokenClientId: string | null = null;

function quoteSheetName(sheetName: string): string {
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) {
    return sheetName;
  }

  return `'${sheetName.replace(/'/g, "''")}'`;
}

function columnNumberToA1(columnNumber: number): string {
  let current = columnNumber;
  let result = '';

  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }

  return result;
}

async function loadGoogleIdentityScript(): Promise<void> {
  if (googleScriptPromise) return googleScriptPromise;

  googleScriptPromise = new Promise<void>((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>(`script[src="${GOOGLE_IDENTITY_SCRIPT}"]`);
    if (existing) {
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error('Neizdevās ielādēt Google autorizācijas skriptu.')), {
        once: true,
      });
      if (window.google?.accounts?.oauth2) {
        resolve();
      }
      return;
    }

    const script = document.createElement('script');
    script.src = GOOGLE_IDENTITY_SCRIPT;
    script.async = true;
    script.defer = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('Neizdevās ielādēt Google autorizācijas skriptu.'));
    document.head.appendChild(script);
  });

  return googleScriptPromise;
}

function buildTokenClient(clientId: string): GoogleTokenClient {
  const initTokenClient = window.google?.accounts?.oauth2?.initTokenClient;
  if (!initTokenClient) {
    throw new Error('Google autorizācija nav pieejama šajā pārlūkā.');
  }

  tokenClientId = clientId;
  tokenClient = initTokenClient({
    client_id: clientId,
    scope: GOOGLE_SCOPES,
    callback: () => {
      // Callback is reassigned per request.
    },
  });

  return tokenClient;
}

async function ensureAccessToken(clientId: string): Promise<string> {
  await loadGoogleIdentityScript();

  if (tokenState && tokenState.expiresAt > Date.now() + 30_000 && tokenClientId === clientId) {
    return tokenState.accessToken;
  }

  const client = !tokenClient || tokenClientId !== clientId ? buildTokenClient(clientId) : tokenClient;

  return new Promise<string>((resolve, reject) => {
    client.callback = (response) => {
      if (response.error || !response.access_token) {
        reject(new Error(response.error_description || 'Google autorizācija neizdevās.'));
        return;
      }

      tokenState = {
        accessToken: response.access_token,
        expiresAt: Date.now() + (response.expires_in ?? 3600) * 1000,
      };

      resolve(response.access_token);
    };

    client.requestAccessToken({
      prompt: tokenState ? '' : 'consent',
    });
  });
}

async function fetchGoogleJson<T>(url: string, accessToken: string, init?: RequestInit): Promise<T> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Google API kļūda: ${response.status} ${text}`);
  }

  return response.json() as Promise<T>;
}

async function fetchGoogleBinary(url: string, accessToken: string): Promise<Blob> {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Google API kļūda: ${response.status} ${text}`);
  }

  return response.blob();
}

export function extractSpreadsheetId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;

  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch) return urlMatch[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

export async function loadGoogleSpreadsheetSource(args: {
  clientId: string;
  spreadsheetId: string;
}): Promise<{
  workbook: Workbook;
  title: string;
}> {
  const accessToken = await ensureAccessToken(args.clientId);
  const metadata = await fetchGoogleJson<GoogleSheetsMetadata>(
    `https://sheets.googleapis.com/v4/spreadsheets/${args.spreadsheetId}?fields=properties.title`,
    accessToken,
  );

  const blob = await fetchGoogleBinary(
    `https://www.googleapis.com/drive/v3/files/${args.spreadsheetId}/export?mimeType=${encodeURIComponent(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )}`,
    accessToken,
  );

  const file = new File([blob], `${metadata.properties?.title || 'viesu-saraksts'}.xlsx`, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const workbook = await readWorkbookFromFile(file);

  return {
    workbook,
    title: metadata.properties?.title || 'Google Sheet',
  };
}

export async function applyTimesheetToGoogleSheet(args: {
  clientId: string;
  spreadsheetId: string;
  originalWorkbook: Workbook;
  timesheetWorkbook: Workbook;
  staffConfigText: string;
}): Promise<{ appliedDays: number }> {
  const accessToken = await ensureAccessToken(args.clientId);
  const plan = buildGoogleSheetWritebackPlan(args.originalWorkbook, args.timesheetWorkbook, args.staffConfigText);
  const metadata = await fetchGoogleJson<GoogleSheetsMetadata>(
    `https://sheets.googleapis.com/v4/spreadsheets/${args.spreadsheetId}?fields=sheets.properties(sheetId,title,index)`,
    accessToken,
  );

  const sheetProperties = metadata.sheets?.find((sheet) => sheet.properties?.title === plan.sheetName)?.properties;
  if (sheetProperties?.sheetId == null) {
    throw new Error(`Google Sheet dokumentā nav atrasta lapa "${plan.sheetName}".`);
  }

  const insertRequests = plan.operations
    .filter((operation) => operation.insertRow)
    .map((operation) => ({
      insertDimension: {
        range: {
          sheetId: sheetProperties.sheetId!,
          dimension: 'ROWS',
          startIndex: operation.targetRowIndex - 1,
          endIndex: operation.targetRowIndex,
        },
        inheritFromBefore: true,
      },
    }));

  if (insertRequests.length > 0) {
    await fetchGoogleJson(
      `https://sheets.googleapis.com/v4/spreadsheets/${args.spreadsheetId}:batchUpdate`,
      accessToken,
      {
        method: 'POST',
        body: JSON.stringify({ requests: insertRequests }),
      },
    );
  }

  const data = plan.operations.flatMap((operation) => {
    const targetColumnA1 = columnNumberToA1(operation.targetColumn);
    return [
      {
        range: `${quoteSheetName(plan.sheetName)}!A${operation.targetRowIndex}`,
        values: [['darbin']],
      },
      {
        range: `${quoteSheetName(plan.sheetName)}!${targetColumnA1}${operation.targetRowIndex}`,
        values: [[operation.summary]],
      },
    ];
  });

  if (data.length > 0) {
    await fetchGoogleJson(
      `https://sheets.googleapis.com/v4/spreadsheets/${args.spreadsheetId}/values:batchUpdate`,
      accessToken,
      {
        method: 'POST',
        body: JSON.stringify({
          valueInputOption: 'RAW',
          data,
        }),
      },
    );
  }

  const wrapRequests = plan.operations.flatMap((operation) => [
    {
      repeatCell: {
        range: {
          sheetId: sheetProperties.sheetId!,
          startRowIndex: operation.targetRowIndex - 1,
          endRowIndex: operation.targetRowIndex,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
        cell: {
          userEnteredFormat: {
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredFormat.wrapStrategy',
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: sheetProperties.sheetId!,
          startRowIndex: operation.targetRowIndex - 1,
          endRowIndex: operation.targetRowIndex,
          startColumnIndex: operation.targetColumn - 1,
          endColumnIndex: operation.targetColumn,
        },
        cell: {
          userEnteredFormat: {
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredFormat.wrapStrategy',
      },
    },
  ]);

  if (wrapRequests.length > 0) {
    await fetchGoogleJson(
      `https://sheets.googleapis.com/v4/spreadsheets/${args.spreadsheetId}:batchUpdate`,
      accessToken,
      {
        method: 'POST',
        body: JSON.stringify({ requests: wrapRequests }),
      },
    );
  }

  return {
    appliedDays: plan.appliedDays,
  };
}

export type DriveFile = {
  id: string;
  name: string;
};

export async function listRecentSpreadsheets(clientId: string): Promise<DriveFile[]> {
  const accessToken = await ensureAccessToken(clientId);
  const query = encodeURIComponent("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
  const url = `https://www.googleapis.com/drive/v3/files?q=${query}&orderBy=modifiedTime desc&pageSize=15&fields=files(id,name)`;

  const response = await fetchGoogleJson<{ files: DriveFile[] }>(url, accessToken);
  return response.files || [];
}
