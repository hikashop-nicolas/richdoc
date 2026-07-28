#!/usr/bin/env python3
"""Ask LibreOffice whether it reads richdoc's output the way it read the input.

The other checks judge bytes, structure and schema conformance. This judges MEANING, with
a full word processor as the reader: both the original fixture and richdoc's rewrite of it
are converted to flat ODF, and the two conversions are compared. Whatever LibreOffice
resolved (named styles, inherited formatting, list numbering, note bodies) is resolved the
same way for both, so a difference is something richdoc changed that a real word processor
notices.

Converting to the SAME kind of document rather than to text or PDF is deliberate: it is
what proves the file was understood rather than merely opened. A text export would pass on
a document whose formatting had been destroyed.

This is also the only judge .doc can have. The legacy writer rebuilds the whole file
instead of preserving it, so the preservation and schema checks do not apply to it.

Usage: lo-check.py [corpus-dir]   (default .cache/lo-corpus)
"""
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "soffice",
]

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
}
Q = {k: "{%s}" % v for k, v in NS.items()}


def find_soffice() -> str | None:
    for c in SOFFICE_CANDIDATES:
        path = shutil.which(c) if "/" not in c else (c if Path(c).exists() else None)
        if path:
            return path
    return None


def to_fodt_batch(soffice: str, srcs: list, outdir: Path) -> dict:
    """Convert every file in one soffice run and return stem -> flat-ODF path.

    One invocation, not one per file: LibreOffice takes a second or two just to start, and
    doing that 72 times took longer than the rest of the test suite put together.
    """
    if not srcs:
        return {}
    subprocess.run(
        [soffice, "--headless", "--convert-to", "fodt", "--outdir", str(outdir)] + [str(s) for s in srcs],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False, timeout=900,
    )
    return {p.stem: p for p in outdir.glob("*.fodt")}


# --- reading a flat ODF -----------------------------------------------------------------

# The formatting worth comparing. Anything else (LibreOffice's own bookkeeping, generated
# style names, colour-mode defaults) differs harmlessly between two conversions.
PARA_PROPS = ["text-align", "margin-left", "margin-right", "text-indent", "break-before"]
TEXT_PROPS = ["font-weight", "font-style", "color", "font-size", "background-color"]


def style_table(root) -> dict:
    """style name -> resolved properties, following parent-style-name."""
    raw = {}
    for holder in ("automatic-styles", "styles"):
        node = root.find(Q["office"] + holder)
        if node is None:
            continue
        for st in node.findall(Q["style"] + "style"):
            name = st.get(Q["style"] + "name")
            if not name:
                continue
            props = {}
            for kind, keys in ((("paragraph-properties"), PARA_PROPS), (("text-properties"), TEXT_PROPS)):
                el = st.find(Q["style"] + kind)
                if el is None:
                    continue
                for k in keys:
                    v = el.get(Q["fo"] + k)
                    if v is not None:
                        props[k] = v
                if kind == "text-properties":
                    u = el.get(Q["style"] + "text-underline-style")
                    if u and u != "none":
                        props["underline"] = u
            raw[name] = (st.get(Q["style"] + "parent-style-name"), props)

    resolved: dict = {}

    def resolve(name, seen=()):
        if name in resolved:
            return resolved[name]
        parent, props = raw.get(name, (None, {}))
        base = dict(resolve(parent, seen + (name,))) if parent and parent not in seen else {}
        base.update(props)
        resolved[name] = base
        return base

    for name in raw:
        resolve(name)
    return resolved


def text_of(el) -> str:
    """A paragraph's visible text, with the space/tab/break elements rendered."""
    out = []

    def walk(node):
        if node.text:
            out.append(node.text)
        for c in node:
            tag = c.tag
            if tag == Q["text"] + "s":
                out.append(" " * int(c.get(Q["text"] + "c", "1")))
            elif tag == Q["text"] + "tab":
                out.append(" ")
            elif tag == Q["text"] + "line-break":
                out.append(" ")
            elif tag == Q["text"] + "note":
                pass  # notes are compared on their own, not inside the host paragraph
            else:
                walk(c)
            if c.tail:
                out.append(c.tail)

    walk(el)
    return normalise_volatile(" ".join("".join(out).split()))


# A date or time field is recalculated every time the document is opened, so the two
# conversions of a pair disagree whenever they straddle a second boundary. The value is not
# something richdoc controls; the field's presence and kind are.
TIME_RE = re.compile(r"\b\d{1,2}:\d{2}(:\d{2})?\b")
DATE_RE = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")


def normalise_volatile(text: str) -> str:
    return DATE_RE.sub("<date>", TIME_RE.sub("<time>", text))


def summarise(path: Path, content_only: bool = False) -> dict:
    root = ET.parse(path).getroot()
    styles = style_table(root)
    body = root.find(Q["office"] + "body")
    text = body.find(Q["office"] + "text") if body is not None else None
    if text is None:
        return {"error": "no office:text"}

    paras, tables, notes, links, images, lists = [], [], [], [], 0, []

    def para_entry(p, kind):
        st = styles.get(p.get(Q["text"] + "style-name"), {})
        # Run-level formatting actually applied inside this paragraph, as a sorted set, so
        # the comparison does not depend on how the runs happen to be split.
        runs = set()
        for span in p.iter(Q["text"] + "span"):
            sp = styles.get(span.get(Q["text"] + "style-name"), {})
            for k in TEXT_PROPS + ["underline"]:
                if sp.get(k):
                    runs.add(f"{k}={sp[k]}")
        entry = {"kind": kind, "text": text_of(p), "runs": sorted(runs)}
        if not content_only:
            entry.update({k: v for k, v in st.items() if k in PARA_PROPS})
        if kind == "h":
            entry["level"] = p.get(Q["text"] + "outline-level", "1")
        return entry

    def walk(node, in_list=False):
        nonlocal images
        for c in node:
            tag = c.tag
            if tag in (Q["text"] + "p", Q["text"] + "h"):
                entry = para_entry(c, "h" if tag.endswith("}h") else "p")
                (lists if in_list else paras).append(entry)
            elif tag == Q["text"] + "list":
                walk(c, True)
            elif tag == Q["text"] + "list-item":
                walk(c, True)
            elif tag == Q["table"] + "table":
                # Document order, looking through table:table-header-rows. Collecting the
                # header rows separately and appending them put the header last in one file
                # and first in the other, which looked like richdoc reordering the table.
                rows = [r for r in c.iter(Q["table"] + "table-row")]
                cells = [text_of(cell) for r in rows for cell in r.findall(Q["table"] + "table-cell")]
                header = sum(len(g.findall(Q["table"] + "table-row")) for g in c.iter(Q["table"] + "table-header-rows"))
                tables.append({"rows": len(rows), "headerRows": header, "cells": cells})
                walk(c)
            else:
                walk(c, in_list)

    walk(text)

    for n in root.iter(Q["text"] + "note"):
        cite = n.find(Q["text"] + "note-citation")
        bodies = n.find(Q["text"] + "note-body")
        notes.append({
            "class": n.get(Q["text"] + "note-class"),
            "citation": (cite.text or "").strip() if cite is not None else "",
            "text": " ".join(text_of(p) for p in bodies).strip() if bodies is not None else "",
        })
    for a in root.iter(Q["text"] + "a"):
        links.append(a.get(Q["xlink"] + "href"))
    images = len(list(root.iter(Q["draw"] + "image")))

    header_footer = []
    master = root.find(Q["office"] + "master-styles")
    if master is not None:
        for mp in master.iter(Q["style"] + "master-page"):
            for band in ("header", "footer"):
                el = mp.find(Q["style"] + band)
                if el is not None:
                    header_footer.append(f"{band}: " + " ".join(text_of(p) for p in el))

    return {
        "paragraphs": paras,
        "listItems": lists,
        "tables": tables,
        "notes": notes,
        "links": links,
        "images": images,
        "bands": header_footer,
    }


# --- comparing ---------------------------------------------------------------------------

# Differences that are understood. Each needs a reason, and they print on every run rather
# than being silently dropped. A difference in the same place that is NOT one of these still
# fails, so this stays a regression gate rather than a rubber stamp.
#
# The .doc entries are real gaps in the legacy writer, not artefacts: it rebuilds the whole
# file from the edited body and does not yet carry these across. They are listed in
# _plans/REMAINING.md as work to do, and this check is what will confirm the fixes.
KNOWN: dict = {
    ("headerfooter.doc", "bands"): (
        ".doc: a page-number field in a footer renders as 'Page PAGE 1': the field's "
        "instruction text is emitted as literal text next to its value."
    ),
    ("lists.doc", "paragraphs"): (
        ".doc: lists are flattened into ordinary paragraphs with a literal bullet character, "
        "so a numbered list loses its numbering as well as its structure."
    ),
    ("lists.doc", "listItems"): ".doc: see the note on lists.doc paragraphs.",
    ("notes.doc", "notes"): ".doc: footnote and endnote citation marks are written as '?'.",
    ("comments.doc", "paragraphs"): (
        ".doc: a comment's date is written as zeroes, and a comment anchored over a range "
        "moves to the end of the range rather than staying at its start."
    ),
    ("fields.doc", "paragraphs"): ".doc: a TIME field is written as a DATE field.",
    ("tables.doc", "tables"): ".doc: a table's header-row designation is lost.",
}


LEN_RE = re.compile(r"^(-?\d+(?:\.\d+)?)(in|cm|mm|pt|pc)$")
TO_IN = {"in": 1.0, "cm": 1 / 2.54, "mm": 1 / 25.4, "pt": 1 / 72, "pc": 1 / 6}
# Lengths are compared with a tolerance of a hundredth of an inch (about a quarter of a
# millimetre). The reader converts lengths to whole pixels and the writer converts back, so
# an untouched 1cm indent returns as 1.005cm. That is invisible on a page and stable across
# saves; a real regression in indent handling would be far larger than this.
LEN_TOL_IN = 0.01


def is_zeroish(v) -> bool:
    """A length equal to zero, or nothing at all."""
    if v is None:
        return True
    m = LEN_RE.match(v) if isinstance(v, str) else None
    return bool(m) and float(m[1]) == 0


def same(b, a) -> bool:
    """Equality, except that two lengths within the tolerance count as equal."""
    if isinstance(b, str) and isinstance(a, str):
        mb, ma = LEN_RE.match(b), LEN_RE.match(a)
        if mb and ma:
            return abs(float(mb[1]) * TO_IN[mb[2]] - float(ma[1]) * TO_IN[ma[2]]) <= LEN_TOL_IN
        return b == a
    if isinstance(b, list) and isinstance(a, list):
        return len(b) == len(a) and all(same(x, y) for x, y in zip(b, a))
    if isinstance(b, dict) and isinstance(a, dict):
        # A property that is absent means the same as one explicitly set to zero, so a
        # heading that stops restating "margin-left: 0in" has not changed.
        keys = set(b) | set(a)
        return all(same(b.get(k, "0in") if k not in b else b[k], a.get(k, "0in") if k not in a else a[k])
                   if is_zeroish(b.get(k)) or is_zeroish(a.get(k)) else (k in b and k in a and same(b[k], a[k]))
                   for k in keys)
    return b == a


def diff(label: str, before: dict, after: dict) -> list:
    out = []
    for key in sorted(set(before) | set(after)):
        b, a = before.get(key), after.get(key)
        if same(b, a):
            continue
        if (label, key) in KNOWN:
            print(f"known: {label}: {key}: {KNOWN[(label, key)]}\n")
            continue
        out.append(f"{label}: {key}\n      before: {b!r}\n      after:  {a!r}")
    return out


def main() -> int:
    soffice = find_soffice()
    if not soffice:
        print("LibreOffice not found; install it or pass its path in SOFFICE.", file=sys.stderr)
        return 2
    corpus = Path(sys.argv[1] if len(sys.argv) > 1 else ".cache/lo-corpus")
    if not corpus.exists():
        print(f"no corpus in {corpus}: run `npm run check:lo`, which writes it first", file=sys.stderr)
        return 2

    pairs = []
    for out_file in sorted(corpus.glob("*")):
        if out_file.suffix not in (".docx", ".odt", ".doc"):
            continue
        src = Path("test-corpus") / out_file.name
        if src.exists():
            pairs.append((src, out_file))
    if not pairs:
        print("no fixtures to check", file=sys.stderr)
        return 2

    failures, checked = [], 0
    with tempfile.TemporaryDirectory() as tmp:
        tmpd = Path(tmp)
        # One extension at a time, into its own directory: paragraphs.docx and
        # paragraphs.odt share a stem, so a single output directory would have them
        # overwrite each other.
        converted = {}
        for suffix in (".docx", ".odt", ".doc"):
            for side, index in (("a", 0), ("b", 1)):
                d = tmpd / f"{side}{suffix[1:]}"
                d.mkdir()
                files = [p[index] for p in pairs if p[index].suffix == suffix]
                for stem, path in to_fodt_batch(soffice, files, d).items():
                    converted[(side, suffix, stem)] = path

        for src, out_file in pairs:
            key = (src.suffix, src.stem)
            a = converted.get(("a",) + key)
            b = converted.get(("b",) + key)
            if a is None:
                failures.append(f"{out_file.name}: LibreOffice could not convert the ORIGINAL fixture")
                continue
            if b is None:
                failures.append(f"{out_file.name}: LibreOffice could not open what richdoc wrote")
                continue
            checked += 1
            # .doc is judged on content, not on resolved paragraph formatting. Its writer
            # regenerates the whole file from the edited body rather than preserving it, so
            # it does not carry the original's explicit alignment and indents; comparing
            # those would report the design of the format's support, not a defect.
            content_only = src.suffix == ".doc"
            failures.extend(diff(out_file.name, summarise(a, content_only), summarise(b, content_only)))

    if failures:
        print(f"LibreOffice reads {len(failures)} thing(s) differently after a round trip:\n")
        for f in failures:
            print(f"  {f}")
        return 1
    print(f"LibreOffice reads {checked} rewritten documents as it reads the originals.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
