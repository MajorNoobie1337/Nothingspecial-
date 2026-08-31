#!/usr/bin/env python3
"""Read a txt file of UUIDs and emit them in batches of 16.

Usage:
    python batch_uuids.py uuids.txt                  # "a","b","c" per line
    python batch_uuids.py uuids.txt --format list    # ["a","b","c"] per line
    python batch_uuids.py uuids.txt --format plain   # a,b,c per line
    python batch_uuids.py uuids.txt --format json    # one JSON list of lists
    python batch_uuids.py uuids.txt --size 50
    python batch_uuids.py uuids.txt --strict         # drop anything not a valid UUID

Import it instead if you want the batches in Python:
    from batch_uuids import read_uuids, batched
    for chunk in batched(read_uuids("uuids.txt"), 16):
        ...  # chunk is a list[str] of at most 16 UUIDs
"""

import argparse
import json
import sys
import uuid as uuid_mod


def read_uuids(path, strict=False, dedupe=True):
    """Return the UUIDs in *path* as a list of strings, order preserved."""
    values = []
    seen = set()
    with open(path) as fh:
        for lineno, line in enumerate(fh, 1):
            for token in line.replace(",", " ").split():
                if strict:
                    try:
                        uuid_mod.UUID(token)
                    except ValueError:
                        print(f"line {lineno}: skipping {token!r}", file=sys.stderr)
                        continue
                if dedupe and token in seen:
                    continue
                seen.add(token)
                values.append(token)
    return values


def batched(items, size=16):
    """Yield successive lists of at most *size* items."""
    if size < 1:
        raise ValueError("size must be >= 1")
    for start in range(0, len(items), size):
        yield items[start:start + size]


def main():
    ap = argparse.ArgumentParser(description="Batch a file of UUIDs.")
    ap.add_argument("file", help="txt file, one or more UUIDs per line")
    ap.add_argument("--size", type=int, default=16, help="batch size (default 16)")
    ap.add_argument("--sep", default=",", help="separator within a batch (default ,)")
    ap.add_argument(
        "--format",
        choices=["quoted", "list", "plain", "json"],
        default="quoted",
        help='quoted: "a","b"  |  list: ["a","b"]  |  plain: a,b  |  '
             "json: whole thing as one JSON document",
    )
    ap.add_argument("--strict", action="store_true", help="skip malformed UUIDs")
    ap.add_argument("--keep-duplicates", action="store_true")
    args = ap.parse_args()

    uuids = read_uuids(args.file, strict=args.strict,
                       dedupe=not args.keep_duplicates)
    batches = list(batched(uuids, args.size))

    if args.format == "json":
        print(json.dumps(batches, indent=2))
    else:
        for chunk in batches:
            if args.format == "plain":
                print(args.sep.join(chunk))
            else:
                body = args.sep.join(json.dumps(u) for u in chunk)
                print(f"[{body}]" if args.format == "list" else body)

    print(f"{len(uuids)} uuid(s) -> {len(batches)} batch(es) of {args.size}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
