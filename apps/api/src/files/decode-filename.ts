/**
 * Multer/busboy often treat UTF-8 filename bytes as Latin-1.
 * Result: "Результаты.xlsx" → "Ð ÐµÐ·ÑƒÐ»ÑŒÑ‚Ð°Ñ‚Ñ‹.xlsx"
 */
export function decodeUploadFilename(name: string): string {
  if (!name) return name;
  // Already valid Cyrillic/CJK — leave as-is
  if (/[\u0400-\u04FF\u4E00-\u9FFF]/.test(name)) return name;
  // Typical UTF-8-as-Latin-1 markers
  if (!/[ÐÑÃÂ]/.test(name)) return name;
  try {
    const decoded = Buffer.from(name, 'latin1').toString('utf8');
    if (decoded.includes('\uFFFD')) return name;
    return decoded;
  } catch {
    return name;
  }
}
