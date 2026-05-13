from __future__ import annotations

import argparse
import shutil
import time
from pathlib import Path

from docx_editor import Document


def split_paragraphs(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines()]


def ref_only(paragraph_listing: str) -> str:
    return paragraph_listing.split("|", 1)[0].strip()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("original", type=Path)
    parser.add_argument("revised", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()

    work_dir = args.output.parent / f"_redline_work_{int(time.time())}"
    work_dir.mkdir(parents=True, exist_ok=False)
    work_input = work_dir / "original_source.docx"
    work_revised = work_dir / "revised_source.docx"
    shutil.copyfile(args.original, work_input)
    shutil.copyfile(args.revised, work_revised)

    revised_doc = Document.open(work_revised, author="AI Refiner", force_recreate=True)
    try:
        revised_paragraphs = split_paragraphs(revised_doc.get_visible_text())
    finally:
        revised_doc.close(cleanup=False)

    original_doc = Document.open(work_input, author="AI Refiner", force_recreate=True)
    try:
        original_paragraphs = split_paragraphs(original_doc.get_visible_text())
        paragraph_refs = [ref_only(item) for item in original_doc.list_paragraphs(max_chars=0)]

        count = min(len(original_paragraphs), len(revised_paragraphs), len(paragraph_refs))
        rewrites: list[tuple[str, str]] = []
        for idx in range(count):
            old_text = original_paragraphs[idx]
            new_text = revised_paragraphs[idx]
            if not old_text or old_text == new_text:
                continue
            rewrites.append((paragraph_refs[idx], new_text))
            if args.limit and len(rewrites) >= args.limit:
                break

        print(f"Original paragraphs: {len(original_paragraphs)}")
        print(f"Revised paragraphs: {len(revised_paragraphs)}")
        print(f"Paragraph refs: {len(paragraph_refs)}")
        print(f"Rewrites: {len(rewrites)}")

        if rewrites:
            original_doc.batch_rewrite(rewrites)
        original_doc.save(args.output)
    finally:
        original_doc.close(cleanup=False)

    print(f"Saved: {args.output}")
    print(f"Workspace kept for inspection: {work_dir}")


if __name__ == "__main__":
    main()
