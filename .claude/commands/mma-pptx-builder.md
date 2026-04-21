---
name: mma-pptx-builder
description: Generates PowerPoint decks, presentations, and slides for MMA using the official template, correct slide masters, and approved shape patterns. Use when creating, modifying, or rebuilding .pptx files, decks, presentations, or slides for MMA. Also use when asked to "make a deck", "build a presentation", "create slides", "generate a PowerPoint", or any PPTX/PPT task. Covers template workflow, master selection per think tank, shape construction (accent bar cards, callouts, tables, flow diagrams), and python-pptx implementation patterns.
---

# MMA PowerPoint Builder

Generate slides that inherit MMA's real template structure, use the correct master layouts, preserve logo/header positioning, and produce visually strong shape-based slides.

> **Colors, fonts, tone, naming:** See **mma-brand-guidelines** skill. This skill covers PPTX-specific construction rules only.

---

## 1. Template — Finding It

Always start from the official MMA template. Never build from a blank Presentation().

**Canonical template (as of April 2026):** `MMA PPT Template v2.1 2026.potx`. Identical masters/layouts/colors to the Dec 2025 template, minus one duplicate (`1_Title and Content Black Bullets`) that was removed from the Core Gold master.

### Template discovery — try in this order

**Step 1 — Local project directory (Claude Code / Cowork only; skip on claude.ai).**
Search the current working directory and common project subfolders for any file matching `MMA PPT Template*.potx` or `MMA_Template*.pptx`. If the user already has a local copy in the project, use that.

```bash
# Bash tool
find . -maxdepth 4 -iname "MMA*Template*.potx" -o -iname "MMA_Template*.pptx" 2>/dev/null
```

**Step 2 — SharePoint via Microsoft connector / MCP.**
Available to anyone with a MMA account. Most staff have at least read access.

- File name: `MMA PPT Template v2.1 2026.potx`
- Path: `MMA Internal > Documents > General > Marketing + Media Alliance MMA Brand Kit 2025`
- Direct link (anyone with the link can view): https://mmaglobalcom.sharepoint.com/:p:/s/MMAGlobal/IQASzy8PH_6cTZVea1p9qoTTAUzphRo3ozdN8T_qyHznFW4?e=QploEn

Use `mcp__claude_ai_Microsoft_365__sharepoint_search` or `mcp__ms365__search-onedrive-files` with the file name as the query. For most users this is the most reliable path.

**Step 3 — Synced OneDrive app (Claude Code / Cowork only).**
If the user has the OneDrive app on macOS and syncs the MMA Internal site, the file is at:

```
/Users/<username>/Library/CloudStorage/OneDrive-MMAGlobal/MMA Internal - Documents/General/Marketing + Media Alliance MMA Brand Kit 2025/MMA PPT Template v2.1 2026.potx
```

### POTX → PPTX one-time conversion

python-pptx cannot open `.potx` directly — it throws `ValueError: not a PowerPoint file`. Patch the content type:

```python
import zipfile

def potx_to_pptx(src_potx, dst_pptx):
    with zipfile.ZipFile(src_potx, "r") as zin, \
         zipfile.ZipFile(dst_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"presentationml.template.main+xml",
                    b"presentationml.presentation.main+xml",
                )
            zout.writestr(item, data)

# Then open normally:
from pptx import Presentation
prs = Presentation("MMA_Template.pptx")
```

Once converted and saved locally as `.pptx`, new slides inherit the real masters, logo placement, slide numbers, and footer positions.

---

## 2. Slide Masters — pick the right one for the think tank

The template ships with **six masters**. Most LLMs reach for `prs.slide_layouts[i]` which only exposes **master 0**. To use the themed masters, access them via `prs.slide_masters[i].slide_layouts`.

| Idx | Master | Brand | Layouts | Use for |
|-----|--------|-------|---------|---------|
| 0 | **MMA Core Gold** | Gold `#FFA400` | 13 (rich) | Default MMA decks. Has Section Header, Appendix, 1/3 and 2/3 splits, Two Content, Content with Caption. |
| 1 | MMA MATT | Blue `#0047BB` | 6 | Measurement & Attribution Think Tank decks. |
| 2 | MMA MOSTT | Orange `#E25700` | 6 | Marketing Org Strategy Think Tank decks. |
| 3 | MMA DATT | Teal `#00AB84` | 6 | Data & CX Think Tank decks. |
| 4 | **MMA ALTT** | Peridot `#B5C900` | 6 | AI Leadership Think Tank decks. |
| 5 | Charts and Tables Clean | neutral | 6 | Data-heavy appendix slides. |

**Think tank masters (1–4) only have 6 layouts each:** `1_Title Slide`, `Title Slide`, `2_Title Slide`, `Title and Content`, `Title Only`, `1_Title Only`. They **do not have** Section Header, Appendix, or 1/3 / 2/3 splits — those only exist on Core Gold.

**Rule:** Use the think tank master for *every slide* in a think tank deck (including section dividers — see §7). Do not mix in Core Gold slides for dividers; the visual inconsistency is worse than repurposing a `2_Title Slide` as a divider.

### Helper — resolve a layout by master and name

```python
def layout_by_name(prs, master_idx, layout_name):
    master = prs.slide_masters[master_idx]
    for l in master.slide_layouts:
        if l.name == layout_name:
            return l
    raise KeyError(f"Layout '{layout_name}' not found on master {master_idx}")

# Core Gold usage
layout = layout_by_name(prs, 0, "Title and Content Black Bullets")
# ALTT usage
layout = layout_by_name(prs, 4, "Title and Content")
```

### Layout map — Core Gold master (master 0)

| Name | Use for |
|------|---------|
| `1_Title Slide` | First slide of a Core Gold deck. |
| `Title and Content Black Bullets` | **Default workhorse** for content slides. Preserves title + accent bar; no unused body placeholder. |
| `Title and Content Gold Bullets` | Bullet-list slides only. Has body placeholder. |
| `Section Header` | Section dividers (Core Gold only). |
| `Title and Content 2/3` / `1/3` | Two-column ratio layouts. |
| `Two Content` | True two-column with native placeholders. |
| `1_Blank` | Full-bleed imagery ONLY. |
| `Appendix` | Appendix divider. |

### Remove sample slides before building

```python
from pptx.oxml.ns import qn

def strip_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sld in list(sldIdLst):
        rId = sld.get(qn("r:id"))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sld)
```

---

## 3. Placeholder Hygiene — clean up what you don't use

Layouts come with default placeholders (title, subtitle, body, footer, slide number). **If you don't use a placeholder, delete it from the slide** — empty placeholders render "Click to add..." ghost text and break shape-driven layouts.

```python
def delete_unused_placeholders(slide, keep_idxs=(0, 12)):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep_idxs:
            sp = ph._element
            sp.getparent().remove(sp)
```

**When to delete:**
- Title Only / 1_Title Only layouts: keep placeholder 0 (title) and 12 (slide number). Delete all others.
- Title Slide: keep 0 (title), 1 (subtitle), 12. Delete the rest.
- Anything with a body placeholder you're not filling: delete it.

**When to keep:**
- If the layout's native subtitle placeholder is in the right position, use it instead of adding a fresh text box.

---

## 4. Typography (PPTX-Specific Sizes)

**14pt is the preferred minimum for all text on a slide. 12pt is the absolute floor.** Never go below 12pt.

| Element | Font | Size |
|---------|------|------|
| Slide title (content slides) | Söhne Halbfett | 36pt default, 32pt only if wrapping |
| Title slide title | Söhne Halbfett | 44–54pt |
| Section header title | Söhne Halbfett | 60pt |
| Subtitle | Söhne Leicht | 18–20pt, gray #666666 |
| Body / bullets | Söhne Leicht | 16–20pt |
| Card titles | Söhne Halbfett | 18–22pt |
| Card body | Söhne Leicht | 14–16pt |
| Kicker labels | Söhne Halbfett | 14pt, uppercase, accent color |
| Table text | Söhne Leicht | 14–16pt |
| Footnotes / source lines | Söhne Leicht | 14pt preferred, 12pt absolute min |

### Title + subtitle live in ONE text box

Content slide subtitles must be the second paragraph of the title placeholder, not a separate text box.

```python
def set_title_with_subtitle(slide, title_text, subtitle_text=None,
                            title_size=36, subtitle_size=18):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p1 = tf.paragraphs[0]
            r1 = p1.add_run()
            r1.text = title_text
            r1.font.name = "Söhne Halbfett"
            r1.font.size = Pt(title_size)
            r1.font.bold = True
            r1.font.color.rgb = BLACK
            if subtitle_text:
                p2 = tf.add_paragraph()
                p2.space_before = Pt(4)
                r2 = p2.add_run()
                r2.text = subtitle_text
                r2.font.name = "Söhne Leicht"
                r2.font.size = Pt(subtitle_size)
                r2.font.bold = False
                r2.font.color.rgb = MED_GRAY
            return
```

---

## 5. Title Slide — use the native layout, don't freelance

Use the existing placeholders on `Title Slide` / `1_Title Slide`. Don't add freshly positioned text boxes.

```python
def set_title_slide(slide, title, subtitle=None):
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text_frame.text = title
            for p in ph.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = "Söhne Halbfett"
                    r.font.bold = True
        elif idx == 1 and subtitle is not None:
            ph.text_frame.text = subtitle
            for p in ph.text_frame.paragraphs:
                for r in p.runs:
                    r.font.name = "Söhne Leicht"
```

**Never add edge-to-edge accent bars to the title slide.** The template already has its own colored diagonal slash.

---

## 6. Section Divider Slides

For section dividers inside a think tank deck, repurpose the master's `2_Title Slide` layout. Do not add full-height colored rectangles.

---

## 7. Logo Protection

The MMA logo sits at the bottom-right (~1.2" × 0.8"). Content that extends into that zone needs a white backing rect.

```python
def protect_logo(slide):
    sw = slide.part.package.presentation_part.presentation.slide_width
    sh = slide.part.package.presentation_part.presentation.slide_height
    box_w, box_h = Inches(1.4), Inches(0.95)
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        sw - box_w - Inches(0.1), sh - box_h - Inches(0.05),
        box_w, box_h)
    rect.fill.solid(); rect.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    rect.line.fill.background()
    return rect
```

---

## 8. Color Usage

```python
from pptx.dml.color import RGBColor

def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

GOLD      = hex_rgb("FFA400")
SAPPHIRE  = hex_rgb("0047BB")
TOPAZ     = hex_rgb("E25700")
EMERALD   = hex_rgb("00AB84")
PERIDOT   = hex_rgb("B5C900")

LT_GOLD     = hex_rgb("FFBB4D")
LT_SAPPHIRE = hex_rgb("4079D6")
LT_TOPAZ    = hex_rgb("E98B4D")
LT_EMERALD  = hex_rgb("66CDB7")
LT_PERIDOT  = hex_rgb("D6E466")

BLACK       = hex_rgb("000000")
WHITE       = hex_rgb("FFFFFF")
DARK_GRAY   = hex_rgb("333333")
MED_GRAY    = hex_rgb("666666")
LIGHT_BG    = hex_rgb("F5F5F5")
CREAM       = hex_rgb("FFF5E0")
ALT_ROW     = hex_rgb("EBEBEB")
BORDER_CLR  = hex_rgb("E0E0E0")
```

- 75% tints for shape fills, accent bars, card backgrounds, diagram nodes.
- Full-saturation for text accents, borders, emphasis.
- NEVER same-hue text on same-hue background.

---

## 9. Shape Construction Patterns

### Centering a row of cards

```python
def center_x(total_width, slide_width):
    return (slide_width - total_width) / 2
```

### Rounded card

```python
def add_rounded_card(slide, left, top, width, height, fill_color=LIGHT_BG):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    s.fill.solid(); s.fill.fore_color.rgb = fill_color
    s.line.fill.background()
    set_tight_corners(s)
    return s

def set_tight_corners(shape, adj=5000):
    sp = shape._element
    prstGeom = sp.find(".//" + qn("a:prstGeom"))
    if prstGeom is not None:
        avLst = prstGeom.find(qn("a:avLst"))
        if avLst is None:
            avLst = etree.SubElement(prstGeom, qn("a:avLst"))
        for gd in avLst.findall(qn("a:gd")):
            avLst.remove(gd)
        gd = etree.SubElement(avLst, qn("a:gd"))
        gd.set("name", "adj")
        gd.set("fmla", f"val {adj}")
```

### Text box helper

```python
def add_textbox(slide, left, top, width, height, text, font_name, size_pt,
                color=BLACK, bold=False, align=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    p = tf.paragraphs[0]
    if align: p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font_name
    r.font.size = Pt(size_pt)
    r.font.bold = bold
    r.font.color.rgb = color
    return tb
```

---

## 10. Accent Bar Cards (top bar preferred)

```python
def add_top_bar_card(slide, left, top, width, height,
                     fill_color=LIGHT_BG, bar_color=GOLD, bar_height=None):
    if bar_height is None: bar_height = Inches(0.08)
    card = add_rounded_card(slide, left, top, width, height, fill_color)
    card.line.color.rgb = BORDER_CLR; card.line.width = Pt(0.75)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, bar_height)
    bar.fill.solid(); bar.fill.fore_color.rgb = bar_color
    bar.line.fill.background()
    return card, bar
```

---

## 11. Tables

Always use native PowerPoint table objects. Never fake tables with shapes.

| Element | Style |
|---------|-------|
| Header row | Dark gray #333333 fill, white bold Halbfett |
| Odd data rows | White fill |
| Even data rows | #EBEBEB fill |
| Borders | #D0D0D0, all sides |
| Text | Söhne Leicht 14pt |

---

## 12. Critical DO NOT Rules

| Rule | Why |
|------|-----|
| Never use custGeom | Renders invisible shapes |
| Never rotate shapes 90/270° | Width/height swap breaks layouts |
| Never use flipV/flipH on text shapes | Text renders mirrored |
| Never fake tables from shapes | Use native tables |
| Never build from blank Presentation() | Always use MMA template |
| Never add edge-to-edge accent bars | Masters already have accent treatments |
| Never leave unused placeholders | Delete them |
| Never shrink titles below 32pt | 36pt default |
| Never mix Core Gold dividers into think tank decks | Use think tank master's 2_Title Slide |
| Never set text below 12pt | 14pt preferred, 12pt absolute floor |
| Never put subtitle in separate text box | Second paragraph of title placeholder |

---

## 13. Required Imports

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
```

---

## 14. Design Behavior Rules

- Most slides should be shape-driven, not plain bullets.
- One idea per slide. Lead with the takeaway.
- Don't let 3+ consecutive slides feel identically templated.
- Prefer clean layouts with white space.
- When in doubt, add less — the template carries its own visual language.
- If the user says a slide is good, preserve it. Iterate only on parts they flag.
