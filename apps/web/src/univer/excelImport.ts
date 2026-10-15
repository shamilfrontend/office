import LuckyExcel from '@mertdeveci55/univer-import-export';

type WorkbookSnapshot = Record<string, unknown>;

function baseName(filename: string) {
  return filename.replace(/\.[^.]+$/, '') || filename;
}

function isSheetSpreadsheet(file: File) {
  const lower = file.name.toLowerCase();
  return (
    lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv')
  );
}

function excelToUniver(file: File): Promise<WorkbookSnapshot> {
  return new Promise((resolve, reject) => {
    void LuckyExcel.transformExcelToUniver(
      file,
      (data) => {
        if (!data) {
          reject(new Error('Не удалось прочитать Excel'));
          return;
        }
        const snapshot = data as unknown as WorkbookSnapshot;
        if (!snapshot.name) snapshot.name = baseName(file.name);
        resolve(snapshot);
      },
      (err) => {
        reject(err instanceof Error ? err : new Error(String(err)));
      },
    );
  });
}

function csvToUniver(file: File): Promise<WorkbookSnapshot> {
  return new Promise((resolve, reject) => {
    LuckyExcel.transformCsvToUniver(
      file,
      (data) => {
        if (!data) {
          reject(new Error('Не удалось прочитать CSV'));
          return;
        }
        const snapshot = data as unknown as WorkbookSnapshot;
        if (!snapshot.name) snapshot.name = baseName(file.name);
        resolve(snapshot);
      },
      (err) => {
        reject(err instanceof Error ? err : new Error(String(err)));
      },
    );
  });
}

/** Convert spreadsheet File to Univer IWorkbookData and wrap as uploadable JSON. */
export async function spreadsheetFileToUpload(file: File): Promise<File> {
  if (!isSheetSpreadsheet(file)) {
    throw new Error('Ожидался файл .xlsx, .xls или .csv');
  }
  const lower = file.name.toLowerCase();
  const snapshot = lower.endsWith('.csv')
    ? await csvToUniver(file)
    : await excelToUniver(file);

  const name = String(snapshot.name || baseName(file.name));
  snapshot.name = name;

  return new File([JSON.stringify(snapshot)], `${name}.json`, {
    type: 'application/json',
  });
}

export function downloadWorkbookXlsx(
  snapshot: unknown,
  fileName: string,
): Promise<void> {
  const safeName = fileName.toLowerCase().endsWith('.xlsx')
    ? fileName
    : `${fileName}.xlsx`;

  return new Promise((resolve, reject) => {
    void LuckyExcel.transformUniverToExcel({
      snapshot,
      fileName: safeName,
      getBuffer: true,
      success: (buffer) => {
        try {
          if (!buffer) {
            reject(new Error('Пустой буфер Excel'));
            return;
          }
          const bytes =
            buffer instanceof ArrayBuffer
              ? new Uint8Array(buffer)
              : new Uint8Array(buffer as ArrayLike<number>);
          const blob = new Blob([bytes], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = safeName;
          a.click();
          URL.revokeObjectURL(url);
          resolve();
        } catch (err) {
          reject(err instanceof Error ? err : new Error(String(err)));
        }
      },
      error: (err) => {
        reject(err instanceof Error ? err : new Error(String(err)));
      },
    });
  });
}
