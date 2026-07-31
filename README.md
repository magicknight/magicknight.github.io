# ScanLine public website

This repository publishes the bilingual ScanLine website at
<https://scanline.cn/>.

- English: `index.html`
- Chinese: `index_zh.html`
- Shared styles and behavior: `assets/`
- Public pitch materials and reproducible builder: `deck/`

Build the English deck:

```bash
python3 deck/build_deck.py
libreoffice --headless --convert-to pdf --outdir deck \
  deck/ScanLine_AI_Gemini_Startup_Forum_2026.pptx
```

Validate the website and generated pitch files:

```bash
python3 scripts/validate_site.py
```

The site deliberately distinguishes established founder experience, work in
progress, and open product-validation milestones. Target performance figures
must not be rewritten as achieved results without a qualified benchmark.
