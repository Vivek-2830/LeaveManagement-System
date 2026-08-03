/** mm/dd/yyyy for display */
// export function formatDisplayDate(dateValue: string | Date): string {
//   const d = new Date(dateValue);
//   const month = d.getMonth() + 1;
//   const day = d.getDate();
//   return `${month < 10 ? '0' : ''}${month}/${day < 10 ? '0' : ''}${day}/${d.getFullYear()}`;
// }

export function formatDisplayDate(dateValue: string | Date): string {
  const d = typeof dateValue === 'string' ? new Date(dateValue) : dateValue;
  const month = d.getMonth() + 1;
  const day = d.getDate();
  return `${month < 10 ? '0' : ''}${month}/${day < 10 ? '0' : ''}${day}/${d.getFullYear()}`;
}

/** ISO date-only string (yyyy-MM-dd), matches the value of an <input type="date"> */
export function toIsoDate(date: Date): string {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return `${date.getFullYear()}-${month < 10 ? '0' : ''}${month}-${day < 10 ? '0' : ''}${day}`;
}

export function parseIsoDate(value: string): Date | undefined {
  if (!value) return undefined;
  const [y, m, d] = value.split('-').map((v) => parseInt(v, 10));
  if (!y || !m || !d) return undefined;
  return new Date(y, m - 1, d);
}

export function isSameCalendarDay(a: Date, b: Date): boolean {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

/**
 * Counts weekdays (Mon-Fri) between two dates, inclusive on both ends.
 */
export function calcBusinessDays(dDate1: Date, dDate2: Date): number {
  let iAdjust = 0;
  if (dDate2 < dDate1) return -1;
  let iWeekday1 = dDate1.getDay();
  let iWeekday2 = dDate2.getDay();
  iWeekday1 = iWeekday1 === 0 ? 7 : iWeekday1;
  iWeekday2 = iWeekday2 === 0 ? 7 : iWeekday2;
  if (iWeekday1 > 5 && iWeekday2 > 5) iAdjust = 1;
  iWeekday1 = iWeekday1 > 5 ? 5 : iWeekday1;
  iWeekday2 = iWeekday2 > 5 ? 5 : iWeekday2;

  const iWeeks = Math.floor((dDate2.getTime() - dDate1.getTime()) / 604800000);

  let iDateDiff: number;
  if (iWeekday1 <= iWeekday2) {
    iDateDiff = iWeeks * 5 + (iWeekday2 - iWeekday1);
  } else {
    iDateDiff = (iWeeks + 1) * 5 - (iWeekday1 - iWeekday2);
  }
  iDateDiff -= iAdjust;
  return iDateDiff + 1;
}
