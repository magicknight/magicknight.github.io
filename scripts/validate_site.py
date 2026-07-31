#!/usr/bin/env python3
"""Validate the public ScanLine website and application deck."""

from __future__ import annotations

import re
import sys
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parents[1]
HTML_FILES = [ROOT / "index.html", ROOT / "index_zh.html"]
DECK = ROOT / "deck" / "ScanLine_AI_Gemini_Startup_Forum_2026.pdf"
PPTX = ROOT / "deck" / "ScanLine_AI_Gemini_Startup_Forum_2026.pptx"


class LocalReferenceParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.references: list[tuple[str, str]] = []

    def handle_starttag(self, tag, attrs):
        values = dict(attrs)
        for attr in ("href", "src", "action"):
            value = values.get(attr)
            if value:
                self.references.append((attr, value))


def local_target(value: str) -> Path | None:
    parsed = urlparse(value)
    if parsed.scheme or parsed.netloc or value.startswith(("#", "mailto:", "tel:")):
        return None
    clean = parsed.path
    if not clean:
        return None
    if clean.startswith("/"):
        clean = clean[1:]
    target = ROOT / clean
    if clean.endswith("/"):
        target = target / "index.html"
    return target


def require(condition: bool, message: str, errors: list[str]):
    if not condition:
        errors.append(message)


def validate():
    errors: list[str] = []
    for page in HTML_FILES:
        require(page.exists(), f"missing page: {page.relative_to(ROOT)}", errors)
        if not page.exists():
            continue
        text = page.read_text(encoding="utf-8")
        parser = LocalReferenceParser()
        parser.feed(text)
        for attr, reference in parser.references:
            target = local_target(reference)
            if target is not None:
                require(
                    target.exists(),
                    f"{page.name}: missing {attr} target {reference}",
                    errors,
                )
        require("<h1" in text, f"{page.name}: missing h1", errors)
        require(
            "ScanLine_AI_Gemini_Startup_Forum_2026.pdf" in text,
            f"{page.name}: missing current public deck link",
            errors,
        )

    english = (ROOT / "index.html").read_text(encoding="utf-8")
    chinese = (ROOT / "index_zh.html").read_text(encoding="utf-8")
    required_english = [
        "See every part. Decide before it leaves the line.",
        "OPEN BRIDGE",
        "One part. One defect. One line.",
        "Independent alternative",
        "Dr. Juan Liu",
    ]
    required_chinese = ["开放桥梁", "一个工件，一个缺陷，一条产线。", "刘娟 博士"]
    for phrase in required_english:
        require(phrase in english, f"index.html: missing required phrase: {phrase}", errors)
    for phrase in required_chinese:
        require(phrase in chinese, f"index_zh.html: missing required phrase: {phrase}", errors)

    stale_claims = [
        "2-5 Second Scans",
        "100% Online Inspection",
        "patent pending",
        "Milan office",
        "Europe HQ",
        "Request a Demo",
    ]
    combined = english.lower() + "\n" + chinese.lower()
    for phrase in stale_claims:
        require(
            phrase.lower() not in combined,
            f"stale or unsupported claim remains: {phrase}",
            errors,
        )

    require((ROOT / "CNAME").read_text(encoding="utf-8").strip() == "scanline.cn", "CNAME mismatch", errors)
    require(DECK.exists() and DECK.stat().st_size > 100_000, "missing or implausibly small PDF deck", errors)
    if DECK.exists():
        require(DECK.read_bytes()[:5] == b"%PDF-", "deck is not a PDF", errors)
    require(PPTX.exists() and PPTX.stat().st_size > 100_000, "missing or implausibly small PPTX", errors)
    if PPTX.exists():
        with zipfile.ZipFile(PPTX) as archive:
            slide_names = [
                name
                for name in archive.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            ]
        require(len(slide_names) == 13, f"expected 13 slides, found {len(slide_names)}", errors)

    if errors:
        print("VALIDATION FAILED")
        for error in errors:
            print(f"- {error}")
        return 1
    print("VALIDATION PASSED")
    print("- bilingual pages and local references")
    print("- truthful-state language and stale-claim guard")
    print("- CNAME")
    print("- public PDF and 13-slide editable PPTX")
    return 0


if __name__ == "__main__":
    sys.exit(validate())
