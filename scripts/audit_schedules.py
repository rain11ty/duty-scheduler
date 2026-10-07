"""Read local PDFs through the real parser without importing/saving application data."""
import argparse
from collections import Counter
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from logic import parse_pdf_schedule


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('folder', type=Path, help='课表文件夹（包含子目录）')
    parser.add_argument('--term', help='校验课表正文学期，如 2026-2027-1')
    args = parser.parse_args()
    if not args.folder.is_dir():
        parser.error('课表文件夹不存在')
    files = sorted(f for f in args.folder.rglob('*') if f.is_file() and f.suffix.lower() == '.pdf')
    if not files:
        parser.error('文件夹内没有 PDF')
    counts = Counter()
    for index, file in enumerate(files, 1):
        try:
            slots = parse_pdf_schedule(file, expected_term=args.term)
        except Exception as error:
            counts['failed'] += 1
            print(f'[{index}/{len(files)}] FAIL {file.relative_to(args.folder)}: {error}', flush=True)
            continue
        classes = sum(slot.get('type') == 'class' for slot in slots)
        practices = sum(slot.get('type') == 'practice' for slot in slots)
        counts['success'] += 1
        counts['class_slots'] += classes
        counts['practice_slots'] += practices
        print(f'[{index}/{len(files)}] OK class_slots={classes} practice_slots={practices}', flush=True)
    print(f"总计 {len(files)} 份：成功 {counts['success']}，失败 {counts['failed']}；"
          f"普通课程节次 {counts['class_slots']}，实践节次 {counts['practice_slots']}")
    return 1 if counts['failed'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
