#!/usr/bin/env python3
"""
Heading-Based DOCX Organizer

Rules implemented:
1. The directory hierarchy is Heading 1 > Heading 2 > Heading 3.
2. Every Heading 1 and Heading 2 creates a folder. A DOCX whose filename
   exactly matches that heading is placed in the same directory as the folder.
3. Matching is exact after normalization: the complete heading/file stem must
   match through the final character. Prefix/partial matches are rejected.
   Example: "Residential Roofing" != "Residential Roofing Repair".
4. Headings are the source of truth. Files never create categories.
5. Heading 3 is used as a destination level but does not independently create
   a required category outside its parent hierarchy unless a matching file is
   found; see CREATE_HEADING3_FOLDERS below.
6. Ambiguous duplicate headings are reported rather than guessed.

Usage:
    python organize_by_headings.py --source "C:/path/to/files" \
        --headings "C:/path/to/Organized.docx" \
        --destination "C:/path/to/organized"

Optional:
    --copy       Copy files instead of moving them.
    --dry-run    Show planned actions without changing files.
"""

from __future__ import annotations

import argparse
import re
import shutil
from collections import defaultdict
from pathlib import Path

from docx import Document


# Set True if you want Heading 3 folders to be created even when no matching
# DOCX exists. Set False to create Heading 3 folders only when a file belongs
# there. Heading 1 and Heading 2 folders are ALWAYS created.
CREATE_HEADING3_FOLDERS = False


DUPLICATE_SUFFIX_RE = re.compile(r"\s*\(\d+\)\s*$")


def clean_name(value: str) -> str:
    """Normalize a heading/folder name without changing meaningful text."""
    value = value.strip()
    if value.lower().endswith(".docx"):
        value = value[:-5]
    value = re.sub(r"\s+", " ", value)
    return value


def filename_match_name(value: str) -> str:
    """
    Normalize a filename for heading matching.

    Safe filename artifacts ignored:
      - .docx extension
      - surrounding/repeated whitespace
      - Windows duplicate suffixes such as (1), (2), (12)
      - a trailing underscore

    Meaningful words and punctuation are otherwise preserved.
    """
    value = clean_name(value)

    # Windows duplicate copies: "Name (1)", "Name (2)", etc.
    value = DUPLICATE_SUFFIX_RE.sub("", value)

    # Ignore only a trailing underscore. Internal underscores are meaningful.
    value = value.rstrip("_").rstrip()

    return value


def exact_key(value: str) -> str:
    """
    Strict comparison key.

    The complete remaining filename must equal the complete heading.
    No substring, prefix, fuzzy, or partial matching is performed.
    Comparison is case-insensitive.
    """
    return filename_match_name(value).casefold()


def read_heading_tree(docx_path: Path):
    """
    Read Heading 1/2/3 paragraphs and preserve their hierarchy.

    Heading 3 belongs to the most recent Heading 2.
    Heading 2 belongs to the most recent Heading 1.
    """
    doc = Document(docx_path)
    tree = []
    current_h1 = None
    current_h2 = None

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        style = paragraph.style.name if paragraph.style else ""

        if style.startswith("Heading 1"):
            current_h1 = text
            current_h2 = None
            tree.append({
                "h1": current_h1,
                "h2": None,
                "h3": None,
            })

        elif style.startswith("Heading 2"):
            if current_h1 is None:
                raise ValueError(
                    f'Heading 2 "{text}" appears before any Heading 1.'
                )
            current_h2 = text
            tree.append({
                "h1": current_h1,
                "h2": current_h2,
                "h3": None,
            })

        elif style.startswith("Heading 3"):
            if current_h1 is None:
                raise ValueError(
                    f'Heading 3 "{text}" appears before any Heading 1.'
                )

            # Normally the hierarchy is H1 > H2 > H3. However, if the
            # source sitemap contains an H3 directly under an H1, preserve
            # the source heading relationship instead of crashing.
            #
            # The headings remain the source of truth; the script never
            # invents an H2 category.
            tree.append({
                "h1": current_h1,
                "h2": current_h2,
                "h3": text,
            })

    return tree


def build_expected_destinations(tree):
    """
    Build a heading-derived map:
        exact heading name -> possible heading destinations.

    Each match records its heading level. Duplicate heading names are kept
    separate so the organizer can refuse to guess.
    """
    destinations = defaultdict(list)

    for item in tree:
        h1 = item["h1"]
        h2 = item["h2"]
        h3 = item["h3"]

        if h3 is not None:
            if h2 is None:
                # Preserve an H1 > H3 relationship if the sitemap skips H2.
                destinations[exact_key(h3)].append({
                    "level": 3,
                    "parts": (h1, h3),
                })
            else:
                destinations[exact_key(h3)].append({
                    "level": 3,
                    "parts": (h1, h2, h3),
                })
        elif h2 is not None:
            destinations[exact_key(h2)].append({
                "level": 2,
                "parts": (h1, h2),
            })
        else:
            destinations[exact_key(h1)].append({
                "level": 1,
                "parts": (h1,),
            })

    return destinations


def safe_destination(root: Path, parts: tuple[str, ...]) -> Path:
    """Create the requested heading hierarchy below the output root."""
    path = root
    for part in parts:
        path /= clean_name(part)
    return path


def unique_target(path: Path) -> Path:
    """
    Prevent overwriting an existing file.

    If file.docx exists, produce file (1).docx, file (2).docx, etc.
    """
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    counter = 1

    while True:
        candidate = path.with_name(f"{stem} ({counter}){suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def organize(source: Path, headings_docx: Path, destination: Path,
             copy_files: bool = False, dry_run: bool = False):
    tree = read_heading_tree(headings_docx)
    destinations = build_expected_destinations(tree)

    skipped_h2 = [
        item["h3"] for item in tree
        if item["h3"] is not None and item["h2"] is None
    ]
    if skipped_h2:
        print(
            "WARNING: The sitemap contains Heading 3 entries directly under "
            "Heading 1 (no Heading 2). The script will preserve that source "
            "hierarchy and will NOT invent a Heading 2."
        )
        for heading in skipped_h2:
            print(f"  H1 > H3: {heading}")

    # Create only the directories dictated by headings.
    required_dirs = set()

    for item in tree:
        h1 = item["h1"]
        h2 = item["h2"]
        h3 = item["h3"]

        # H1 is always a folder.
        required_dirs.add((h1,))

        # H2 is always a folder.
        if h2 is not None:
            required_dirs.add((h1, h2))

        # H3 behavior is configurable. If the source has H1 > H3,
        # preserve that exact source hierarchy rather than inventing H2.
        if h3 is not None and CREATE_HEADING3_FOLDERS:
            if h2 is None:
                required_dirs.add((h1, h3))
            else:
                required_dirs.add((h1, h2, h3))

    print("=== Creating heading-based directory hierarchy ===")
    for parts in sorted(required_dirs, key=lambda x: (len(x), x)):
        path = safe_destination(destination, parts)
        print(f"FOLDER: {path}")
        if not dry_run:
            path.mkdir(parents=True, exist_ok=True)

    files = sorted(
        p for p in source.rglob("*")
        if p.is_file()
        and p.suffix.lower() == ".docx"
        and p.resolve() != headings_docx.resolve()
    )

    matched = 0
    unmatched = []
    ambiguous = []

    for file_path in files:
        key = exact_key(file_path.name)
        candidates = destinations.get(key, [])

        if not candidates:
            unmatched.append(file_path)
            continue

        # Exact filename match, but duplicate heading names can point to
        # multiple places. Never guess.
        if len(candidates) > 1:
            ambiguous.append((file_path, candidates))
            continue

        match = candidates[0]
        parts = match["parts"]
        target_dir = safe_destination(destination, parts)

        # H3 files belong directly in their H2 folder when H3 folders are
        # disabled. For an H1 > H3 source relationship, the H1 is the parent.
        if match["level"] == 3 and not CREATE_HEADING3_FOLDERS:
            target_dir = safe_destination(destination, parts[:-1])

        target_dir.mkdir(parents=True, exist_ok=True) if not dry_run else None

        target = unique_target(target_dir / file_path.name)

        action = "COPY" if copy_files else "MOVE"
        print(f"{action}: {file_path} -> {target}")

        if not dry_run:
            if copy_files:
                shutil.copy2(file_path, target)
            else:
                shutil.move(str(file_path), str(target))

        matched += 1

    print("\n=== SUMMARY ===")
    print(f"Heading records read: {len(tree)}")
    print(f"DOCX files examined:  {len(files)}")
    print(f"Matched exactly:      {matched}")
    print(f"Unmatched:            {len(unmatched)}")
    print(f"Ambiguous:            {len(ambiguous)}")

    if not CREATE_HEADING3_FOLDERS:
        print("Heading 3 folders:    DISABLED (H3 files go into H2 folders)")
    else:
        print("Heading 3 folders:    ENABLED")

    if unmatched:
        print("\nUNMATCHED FILES (not moved):")
        for path in unmatched:
            print(f"  {path}")

    if ambiguous:
        print("\nAMBIGUOUS FILES (not moved):")
        for path, candidates in ambiguous:
            print(f"  {path}")
            for candidate in candidates:
                print("    -> " + " > ".join(candidate["parts"]))


def main():
    parser = argparse.ArgumentParser(
        description="Organize DOCX files strictly from Heading 1/2/3."
    )
    parser.add_argument("--source", required=True,
                        help="Folder containing DOCX files.")
    parser.add_argument("--headings", required=True,
                        help="DOCX containing the Heading 1/2/3 source hierarchy.")
    parser.add_argument("--destination", required=True,
                        help="Root folder for the generated hierarchy.")
    parser.add_argument("--copy", action="store_true",
                        help="Copy files instead of moving them.")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview folders/actions without changing files.")

    args = parser.parse_args()

    source = Path(args.source).expanduser().resolve()
    headings_docx = Path(args.headings).expanduser().resolve()
    destination = Path(args.destination).expanduser().resolve()

    if not source.is_dir():
        raise SystemExit(f"Source folder does not exist: {source}")
    if not headings_docx.is_file():
        raise SystemExit(f"Heading DOCX does not exist: {headings_docx}")

    organize(
        source=source,
        headings_docx=headings_docx,
        destination=destination,
        copy_files=args.copy,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
