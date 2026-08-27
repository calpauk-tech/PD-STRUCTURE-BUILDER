// Heuristic to guess 19xx vs 20xx for 2-digit years
export const guessCentury = (twoDigitYear: number, fieldName: string): number => {
    const CURRENT_TWO_DIGIT_YEAR = parseInt(new Date().getFullYear().toString().slice(-2), 10);
    const lowerName = fieldName.toLowerCase();
    
    // Birth Date Logic: Strictly pivot on the current year. 
    if (lowerName.includes('birth')) {
        return twoDigitYear > CURRENT_TWO_DIGIT_YEAR ? 1900 : 2000;
    }
    // Other Dates: Window 1951 - 2050
    return twoDigitYear > 50 ? 1900 : 2000;
};

// Core parser: Handles JS Dates, Excel Serials, and String Formats
export const parseDateWithFormat = (input: any, forceUS: boolean, yearCorrection?: number): string | null => {
    if (!input) return null;

    // 1. Handle actual JS Date objects
    if (input instanceof Date) {
        if (isNaN(input.getTime())) return null;
        const y = input.getFullYear();
        const m = String(input.getMonth() + 1).padStart(2, '0');
        const d = String(input.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    // 2. Handle Excel Serial Numbers (e.g. 43465)
    if (typeof input === 'number' || /^\d{5}$/.test(String(input))) {
        const serial = parseInt(input, 10);
        if (serial > 20000 && serial < 80000) { // Valid bounds for recent dates
            // 25569 is the difference in days between Jan 1 1900 and Jan 1 1970
            const dateObj = new Date(Math.round((serial - 25569) * 86400 * 1000));
            const y = dateObj.getUTCFullYear();
            const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
            const d = String(dateObj.getUTCDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
        }
    }

    const str = String(input).trim();
    if (!str) return null;

    // 3. Handle Year-First format (YYYY-MM-DD, YYYY.MM.DD, YYYY/MM/DD)
    const matchYearFirst = str.match(/^(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})$/);
    if (matchYearFirst) {
        const y = parseInt(matchYearFirst[1], 10);
        const p1 = parseInt(matchYearFirst[2], 10);
        const p2 = parseInt(matchYearFirst[3], 10);
        let m, d;
        if (p1 > 12) { d = p1; m = p2; } else { m = p1; d = p2; }
        if (m > 12 || d > 31 || m < 1 || d < 1) return null;
        return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }

    // 4. Handle Year-Last format (D/M/Y or M/D/Y)
    const match = str.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})$/);
    if (match) {
        let part1 = parseInt(match[1], 10);
        let part2 = parseInt(match[2], 10);
        let y = parseInt(match[3], 10);

        if (yearCorrection) {
            y = yearCorrection;
        } else if (y < 100) {
            y += (y > 50 ? 1900 : 2000); 
        }

        let d, m;
        if (forceUS) { m = part1; d = part2; } else { d = part1; m = part2; }
        if (m > 12 || d > 31 || m < 1 || d < 1) return null;
        return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }

    // 5. Fallback standard JS Date parse ("Tue Aug 15 2000")
    const fallbackDate = new Date(str.replace(/-/g, '/'));
    if (!isNaN(fallbackDate.getTime()) && /[A-Za-z]/.test(str)) {
        const y = fallbackDate.getFullYear();
        if (y > 1900 && y < 2100) {
            return `${y}-${String(fallbackDate.getMonth() + 1).padStart(2, '0')}-${String(fallbackDate.getDate()).padStart(2, '0')}`;
        }
    }
    return null;
};

// Auto-detect US vs EU format by scanning the dataset for impossible months (e.g. 13/05/2024 -> must be EU)
export const checkForUSDateFormat = (rows: any[], dateKeys: string[]): boolean => {
    let usMarkerFound = false;
    let euMarkerFound = false;
    for (const row of rows) {
        for (const key of dateKeys) {
            const val = row[key];
            if (val && typeof val === 'string') {
                const match = val.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})$/);
                if (match) {
                    const p1 = parseInt(match[1], 10);
                    const p2 = parseInt(match[2], 10);
                    if (p2 > 12) usMarkerFound = true; // Middle number > 12 means it must be DD/MM/YYYY (EU) -> therefore false for US format detection logic (wait, actually if middle is > 12 it's DD/MM/YYYY. If first is > 12 it's MM/DD/YYYY).
                    if (p1 > 12) euMarkerFound = true; 
                }
            }
        }
    }
    // If we find a first part > 12 (MM/DD/YYYY), it's likely US.
    return usMarkerFound && !euMarkerFound; 
};
