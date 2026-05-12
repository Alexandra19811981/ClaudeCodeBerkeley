import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
import os

# ── 1. DATA ──────────────────────────────────────────────────────────────────

# Segment ARR ($M) at three anchor points
years  = ["Q4 2022", "Q4 2024", "Q4 2025"]
ent    = [74,  138, 167]
mm     = [96,  169, 194]
smb    = [105,  68,  52]

# NRR by segment, quarterly Q1 2024 → Q4 2025
quarters = ["Q1\n2024","Q2\n2024","Q3\n2024","Q4\n2024",
            "Q1\n2025","Q2\n2025","Q3\n2025","Q4\n2025"]
nrr_ent = [127,127,126,126,125,125,125,125]
nrr_mm  = [108,107,106,105,104,103,103,102]
nrr_smb = [ 98, 96, 93, 91, 89, 88, 86, 84]

# ── 2. CHART ─────────────────────────────────────────────────────────────────

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6))
fig.patch.set_facecolor("#F9F9F9")

BLUE   = "#1A3A5C"
TEAL   = "#2A7F8C"
RED    = "#C0392B"
GREY   = "#7F8C8D"

# ── Panel A: Stacked bar — ARR by segment ────────────────────────────────────
x = np.arange(len(years))
w = 0.45

b1 = ax1.bar(x, ent, w, label="Enterprise", color=BLUE)
b2 = ax1.bar(x, mm,  w, bottom=ent, label="Mid-market", color=TEAL)
b3 = ax1.bar(x, smb, w, bottom=[e+m for e,m in zip(ent,mm)],
             label="SMB", color=RED)

ax1.set_facecolor("#F9F9F9")
ax1.set_xticks(x)
ax1.set_xticklabels(years, fontsize=11)
ax1.set_ylabel("ARR ($M)", fontsize=11)
ax1.set_title("A.  ARR Mix Shift by Segment", fontsize=13, fontweight="bold",
              pad=12, loc="left")
ax1.legend(loc="upper left", fontsize=10)
ax1.set_ylim(0, 480)
ax1.spines["top"].set_visible(False)
ax1.spines["right"].set_visible(False)

# Annotate totals
totals = [e+m+s for e,m,s in zip(ent,mm,smb)]
for xi, tot in zip(x, totals):
    ax1.text(xi, tot + 5, f"${tot}M", ha="center", fontsize=10,
             fontweight="bold", color="#333333")

# Annotate SMB % labels inside bars
for xi, (e, m, s) in enumerate(zip(ent, mm, smb)):
    pct = round(s / (e+m+s) * 100)
    ax1.text(xi, e + m + s/2, f"SMB\n{pct}%", ha="center", va="center",
             fontsize=9, color="white", fontweight="bold")

# ── Panel B: NRR by segment over time ────────────────────────────────────────
xq = np.arange(len(quarters))
ax2.plot(xq, nrr_ent, "o-", color=BLUE,  lw=2.2, ms=6, label="Enterprise")
ax2.plot(xq, nrr_mm,  "s-", color=TEAL,  lw=2.2, ms=6, label="Mid-market")
ax2.plot(xq, nrr_smb, "^-", color=RED,   lw=2.2, ms=6, label="SMB")
ax2.axhline(100, color=GREY, lw=1.2, ls="--")
ax2.text(7.05, 100.4, "100%\n(breakeven)", fontsize=8, color=GREY)

ax2.set_facecolor("#F9F9F9")
ax2.set_xticks(xq)
ax2.set_xticklabels(quarters, fontsize=10)
ax2.set_ylabel("Net Revenue Retention (%)", fontsize=11)
ax2.set_title("B.  NRR Compression — Mid-market Approaching Breakeven",
              fontsize=13, fontweight="bold", pad=12, loc="left")
ax2.legend(loc="upper right", fontsize=10)
ax2.set_ylim(78, 134)
ax2.spines["top"].set_visible(False)
ax2.spines["right"].set_visible(False)

# End-point labels
for vals, col in [(nrr_ent, BLUE), (nrr_mm, TEAL), (nrr_smb, RED)]:
    ax2.text(7.15, vals[-1], f"{vals[-1]}%", va="center",
             fontsize=9, color=col, fontweight="bold")

fig.suptitle("Meridian Technologies — Key Strategic Trends (2022–2025)",
             fontsize=14, fontweight="bold", y=1.01, color=BLUE)

plt.tight_layout()
chart_path = "/home/user/ClaudeCodeBerkeley/meridian_board_chart.png"
plt.savefig(chart_path, dpi=160, bbox_inches="tight",
            facecolor=fig.get_facecolor())
plt.close()
print(f"Chart saved: {chart_path}")


# ── 3. WORD DOCUMENT ─────────────────────────────────────────────────────────

doc = Document()

# Page margins
for section in doc.sections:
    section.top_margin    = Inches(0.9)
    section.bottom_margin = Inches(0.9)
    section.left_margin   = Inches(1.0)
    section.right_margin  = Inches(1.0)

def set_font(run, name="Calibri", size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1, color=(26,58,92)):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.color.rgb = RGBColor(*color)
    return p

def add_para(doc, text, size=11, bold=False, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    return p

def add_bullet(doc, text, bold_prefix=None, size=11):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        set_font(r1, size=size, bold=True)
        r2 = p.add_run(text)
        set_font(r2, size=size)
    else:
        r = p.add_run(text)
        set_font(r, size=size)
    return p

# ── Header ───────────────────────────────────────────────────────────────────
h = doc.add_heading("", level=0)
h.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = h.add_run("MERIDIAN TECHNOLOGIES")
run.font.name  = "Calibri"
run.font.size  = Pt(20)
run.font.bold  = True
run.font.color.rgb = RGBColor(26,58,92)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(2)
r = sub.add_run("Annual Board Strategic Review  |  CEO Opening Remarks")
set_font(r, size=12, color=(42,127,140))

meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
meta.paragraph_format.space_after = Pt(14)
r = meta.add_run("Catherine Park, President & CEO  |  May 2026  |  5 minutes")
set_font(r, size=10, color=(127,140,141))

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ── Headline ─────────────────────────────────────────────────────────────────
hl = doc.add_paragraph()
hl.alignment = WD_ALIGN_PARAGRAPH.CENTER
hl.paragraph_format.space_before = Pt(0)
hl.paragraph_format.space_after  = Pt(16)
r = hl.add_run(
    "Our enterprise franchise is the strongest it has ever been — "
    "and 2026 is the year we decide whether AI makes us a platform or leaves us a feature."
)
r.font.name  = "Calibri"
r.font.size  = Pt(13)
r.font.bold  = True
r.font.italic = True
r.font.color.rgb = RGBColor(26,58,92)

# ── Chart ────────────────────────────────────────────────────────────────────
add_heading(doc, "The Picture in One Chart", level=2)
doc.add_picture(chart_path, width=Inches(6.3))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

cap = doc.add_paragraph()
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap.paragraph_format.space_after = Pt(14)
r = cap.add_run(
    "Left: ARR by segment — Enterprise doubled, SMB collapsed 50%. "
    "Right: NRR compression — Mid-market trending toward breakeven."
)
set_font(r, size=9, color=(127,140,141))

# ── Issue 1 ──────────────────────────────────────────────────────────────────
add_heading(doc, "Issue 1 — AI Competitive Gap: We Are Behind and Time Is Not Neutral", level=2)

add_para(doc,
    "This is the most urgent issue. We shipped Copilot GA in September 2025 — roughly "
    "18 months after Asana, which bundled AI into its standard tier in October. Today we "
    "have 710 paying Copilot seats and $3.5M in AI ARR. That is a beginning, not a moat.",
    size=11, space_after=6)

add_bullet(doc, "710 Copilot paying seats at Q4 2025 vs. zero twelve months prior — "
           "real traction, but $3.5M ARR on a $413M base.", bold_prefix="Evidence: ")
add_bullet(doc, "Magic number declined from 1.20 (Q1 2024) to 0.92 (Q4 2025) as R&D "
           "rose to 25.4% of revenue — we are spending more to grow less.",
           bold_prefix="Evidence: ")
add_bullet(doc, "ClearAI Work raised $120M at $1B+ valuation. Atlassian announced "
           "agentic Jira. Asana full agent suite shipped November 2025.",
           bold_prefix="Evidence: ")
add_bullet(doc, "Engineering net hiring is ~25 heads/year at 15% attrition — the "
           "2026 plan requires 80. Helio retention cliff hits in 2026.",
           bold_prefix="Risk: ")

doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ── Issue 2 ──────────────────────────────────────────────────────────────────
add_heading(doc, "Issue 2 — Mid-Market NRR at 102%: One Bad Year From a Net ARR Drain", level=2)

add_para(doc,
    "Mid-market is 47% of our ARR and growing at 6%. NRR has compressed "
    "from 115% in 2022 to 102% today. If it crosses 100%, our largest segment "
    "becomes a headwind, not a tailwind — and enterprise cannot grow fast enough to compensate.",
    size=11, space_after=6)

add_bullet(doc, "NRR: 115% (2022) → 108% (Q1 2024) → 102% (Q4 2025). "
           "GRR has declined from 89% to 86% — customers are leaving, not just expanding less.",
           bold_prefix="Evidence: ")
add_bullet(doc, "Mid-market renewals now routinely demand a price hold, expanded seats "
           "at flat price, or Copilot bundled at no cost (Q3 2025 earnings call).",
           bold_prefix="Evidence: ")
add_bullet(doc, "Resource management module — the #1 mid-market customer ask — "
           "deferred twice; now targeting Q2 2026.",
           bold_prefix="Evidence: ")

doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ── Issue 3 ──────────────────────────────────────────────────────────────────
add_heading(doc, "Issue 3 — We Need a 3-Year Plan: The Board Authorized One for Q1 2026", level=2)

add_para(doc,
    "The board authorized a successor strategic plan in Q1 2026. The March 11 Investor Day "
    "will be the public version. Today I am asking the board to approve the internal version "
    "that drives our R&D, M&A, and capital allocation decisions this year.",
    size=11, space_after=6)

add_bullet(doc, "The 2023–2026 plan explicitly noted it was 'no longer the right plan for 2025' "
           "and authorized the new CEO to write a successor (Strategic Plan document).",
           bold_prefix="Context: ")
add_bullet(doc, "The central strategic question — agentic work platform vs. PM tool with AI features "
           "— changes R&D priorities, sales motion, pricing model, and positioning "
           "(Q4 2025 earnings call).",
           bold_prefix="Stakes: ")
add_bullet(doc, "Two additional AI-native acquisition targets in early diligence. "
           "Capital allocation in 2026 requires a resolved strategic direction.",
           bold_prefix="Dependency: ")
add_bullet(doc, "31% of engineers cite 'roadmap thrash' in the 2025 employee survey. "
           "The organization will not execute a plan it doesn't believe is durable.",
           bold_prefix="Internal risk: ")

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# ── Ask ──────────────────────────────────────────────────────────────────────
ask_p = doc.add_paragraph()
ask_p.paragraph_format.space_before = Pt(4)
ask_p.paragraph_format.space_after  = Pt(4)

# Horizontal rule approximation via border — just use a bold label
r = ask_p.add_run("MY ONE ASK OF THE BOARD TODAY")
r.font.name  = "Calibri"
r.font.size  = Pt(12)
r.font.bold  = True
r.font.color.rgb = RGBColor(192, 57, 43)

ask_body = doc.add_paragraph()
ask_body.paragraph_format.space_after = Pt(6)
r = ask_body.add_run(
    "Approve the 3-year strategic direction — specifically, the decision to position Meridian "
    "as an agentic work platform for regulated enterprises — so that every R&D, M&A, and "
    "go-to-market decision in 2026 flows from a single, board-sanctioned answer to the question: "
    "what is Meridian?"
)
set_font(r, size=12, bold=True, color=(26,58,92))

note = doc.add_paragraph()
note.paragraph_format.space_after = Pt(4)
r = note.add_run(
    "Supporting detail on the 3-year financial model, AI roadmap, and M&A pipeline "
    "follows in the board package. I will spend the balance of today's session on "
    "Q&A and the strategic direction decision."
)
set_font(r, size=10, color=(127,140,141))

# ── Footer ───────────────────────────────────────────────────────────────────
doc.add_paragraph()
footer_p = doc.add_paragraph()
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = footer_p.add_run(
    "CONFIDENTIAL — For Board Use Only  |  Meridian Technologies, Inc.  |  NASDAQ: MRDN"
)
set_font(r, size=9, color=(127,140,141))

# ── Save ─────────────────────────────────────────────────────────────────────
docx_path = "/home/user/ClaudeCodeBerkeley/board_opening_remarks.docx"
doc.save(docx_path)
print(f"Word document saved: {docx_path}")
