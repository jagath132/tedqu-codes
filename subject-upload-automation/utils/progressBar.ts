export class ProgressBar {
  private static barLength: number = 30;

  public static update(current: number, total: number, itemName: string = '', status: string = ''): void {
    if (total === 0) return;
    const progress = Math.min(Math.max(current / total, 0), 1);
    const filledCount = Math.round(progress * this.barLength);
    const emptyCount = this.barLength - filledCount;

    const filledBar = '█'.repeat(filledCount);
    const emptyBar = '░'.repeat(emptyCount);
    const percentage = Math.floor(progress * 100);

    const safeName = itemName.length > 25 ? itemName.substring(0, 22) + '...' : itemName;
    const statusTag = status ? `[${status}]` : '';

    const text = `\r[${filledBar}${emptyBar}] ${current}/${total} (${percentage}%) ${statusTag} ${safeName}`.padEnd(80);
    process.stdout.write(text);

    if (current >= total) {
      process.stdout.write('\n');
    }
  }
}
