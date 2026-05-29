"""Generate all 7 English-only academic figures using matplotlib.
Rewritten to match original Korean thesis figure sizes, box styles, and readability.
Key: larger fonts, bordered rectangles, proper dimensions matching originals."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'output', 'figures')
os.makedirs(OUT_DIR, exist_ok=True)

# Common style - match original thesis formatting
FONT = 'Times New Roman'
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = [FONT, 'DejaVu Serif']
plt.rcParams['font.size'] = 12


def draw_box(ax, x, y, w, h, text, fontsize=12, bold=False, lw=1.8,
             fill='white', edgecolor='black', ha='center', va='center',
             rounded=False):
    """Draw a rectangle with centered text. Default: square corners like original."""
    style = "round,pad=0.02" if rounded else "square,pad=0"
    box = FancyBboxPatch((x - w/2, y - h/2), w, h, boxstyle=style,
                          facecolor=fill, edgecolor=edgecolor, linewidth=lw)
    ax.add_patch(box)
    weight = 'bold' if bold else 'normal'
    ax.text(x, y, text, ha=ha, va=va, fontsize=fontsize, fontweight=weight,
            fontfamily='serif', linespacing=1.3)


def draw_arrow(ax, x1, y1, x2, y2, style='->', lw=1.8, color='black'):
    """Draw an arrow between two points."""
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle=style, lw=lw, color=color))


# ============================================================
# Figure 1.1: Overall Logical Flow and Analytical Framework
# Original: 1202 x 846 px
# ============================================================
def fig_1_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(8.0, 5.6), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')

    steps = [
        ("Step 1: Research Problem",
         "AI utilization training judgment formation\n"
         "-> [Structural Disconnection] -> Strategic decision-making not generated"),
        ("Step 2: Theory & Method",
         "Dynamic Capabilities Theory (Seizing stage focus)\n"
         "/ Process Tracing & Pattern Matching analysis"),
        ("Step 3: Empirical Findings",
         "Judgment exists but decision absent /\n"
         "Self-reinforcing stagnation loop of non-decision identified"),
        ("Step 4: Theoretical Extension",
         "Reconceptualization of Seizing\n"
         "(from 'moment of decision' to 'processual infrastructure')"),
        ("Step 5: Recommendations",
         "Introduction of Conditional Trigger Model\n"
         "and time-lagged escalation"),
    ]

    y_positions = [6.1, 4.9, 3.7, 2.5, 1.3]
    box_w = 8.8
    header_h = 0.40
    body_h = 0.70

    for i, (title, body) in enumerate(steps):
        y = y_positions[i]
        header_y = y + body_h/2 + header_h/2

        # Header box
        header_box = FancyBboxPatch((5 - box_w/2, header_y - header_h/2), box_w, header_h,
                                     boxstyle="square,pad=0",
                                     facecolor='#1a1a1a', edgecolor='black', linewidth=1.5)
        ax.add_patch(header_box)
        # Triangle marker
        tri_x = 5 - box_w/2 + 0.35
        tri = plt.Polygon([[tri_x, header_y - 0.09], [tri_x, header_y + 0.09],
                            [tri_x + 0.2, header_y]], closed=True,
                           facecolor='white', edgecolor='white', linewidth=0.5)
        ax.add_patch(tri)
        ax.text(5, header_y, title, ha='center', va='center', fontsize=12,
                fontweight='bold', color='white', fontfamily='serif')

        # Body box
        body_box = FancyBboxPatch((5 - box_w/2, y - body_h/2), box_w, body_h,
                                   boxstyle="square,pad=0",
                                   facecolor='white', edgecolor='black', linewidth=1.2)
        ax.add_patch(body_box)
        ax.text(5, y, body, ha='center', va='center', fontsize=11, fontfamily='serif')

        # Arrow to next step
        if i < len(steps) - 1:
            next_y = y_positions[i+1]
            arrow_start = y - body_h/2 - 0.02
            arrow_end = next_y + body_h/2 + header_h + 0.02
            draw_arrow(ax, 5, arrow_start, 5, arrow_end, lw=1.8)

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_1_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_1_1.png")
    return path


# ============================================================
# Figure 4.1: Structural Disconnection of Judgment-Decision
#              Conversion and the Seizing Non-initiation Mechanism
# Original: 1358 x 1341 px -- MUST be wide, not narrow
# ============================================================
def fig_4_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(9.0, 9.0), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    box_w = 6.5
    box_h = 0.75

    # Sensing box (top)
    draw_box(ax, 5, 9.3, 5.0, 0.75, "Sensing", fontsize=18, bold=True, lw=2.0)
    draw_arrow(ax, 5, 8.92, 5, 8.55, lw=2.0)

    # "Seizing" label
    ax.text(5, 8.3, "Seizing", ha='center', va='center', fontsize=17,
            fontweight='bold', fontfamily='serif')

    # Judgment Formation
    draw_box(ax, 5, 7.65, box_w, box_h, "Judgment Formation",
             fontsize=16, bold=True, lw=1.8)
    draw_arrow(ax, 5, 7.27, 5, 6.9, lw=2.0)

    # Discontinuity of Decision Initiation
    draw_box(ax, 5, 6.5, box_w, box_h, "Discontinuity of Decision Initiation",
             fontsize=16, bold=True, lw=1.8)
    draw_arrow(ax, 5, 6.12, 5, 5.8, lw=2.0)

    # Decision Initiation Conditions Not Activated (large box)
    big_w, big_h = 7.0, 2.2
    big_cy = 5.0
    big_box = FancyBboxPatch((5 - big_w/2, big_cy - big_h/2), big_w, big_h,
                              boxstyle="square,pad=0",
                              facecolor='white', edgecolor='black', linewidth=1.8)
    ax.add_patch(big_box)
    ax.text(5, 5.7, "Decision Initiation Conditions Not Activated",
            ha='center', va='center', fontsize=16, fontweight='bold', fontfamily='serif')
    ax.text(5, 5.2, "- Absence of Trigger", ha='center', va='center',
            fontsize=15, fontfamily='serif')
    ax.text(5, 4.8, "- Failed Agenda-setting", ha='center', va='center',
            fontsize=15, fontfamily='serif')
    ax.text(5, 4.4, "- Uninitiated Decision Pathway", ha='center', va='center',
            fontsize=15, fontfamily='serif')

    draw_arrow(ax, 5, big_cy - big_h/2, 5, 3.55, lw=2.0)

    # Sustained Non-Event State
    draw_box(ax, 5, 3.15, box_w, box_h, "Sustained Non-Event State",
             fontsize=16, bold=True, lw=1.8)

    # Seizing bracket (left side)
    bracket_x = 1.0
    bracket_top = 8.05
    bracket_bot = 2.75
    ax.plot([bracket_x, bracket_x], [bracket_bot, bracket_top],
            color='black', lw=1.5)
    ax.plot([bracket_x, bracket_x + 0.25], [bracket_top, bracket_top],
            color='black', lw=1.5)
    ax.plot([bracket_x, bracket_x + 0.25], [bracket_bot, bracket_bot],
            color='black', lw=1.5)

    draw_arrow(ax, 5, 2.75, 5, 2.25, lw=2.0)

    # Transforming box (bottom)
    draw_box(ax, 5, 1.8, 5.5, 0.85, "Transforming\n(Not Initiated)",
             fontsize=17, bold=True, lw=2.0)

    fig.tight_layout(pad=0.5)
    path = os.path.join(OUT_DIR, 'fig_4_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_4_1.png")
    return path


# ============================================================
# Figure 4.2: Processual Bottleneck Mechanism at the Seizing Stage
# Original: 969 x 1702 px
# ============================================================
def fig_4_2():
    dpi = 150
    fig, ax = plt.subplots(figsize=(6.5, 11.3), dpi=dpi)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 14)
    ax.axis('off')

    # Top: Organisational Judgment on AI Utilisation Request
    draw_box(ax, 4, 13.2, 6.5, 0.95,
             "Organisational Judgment\non AI Utilisation Request",
             fontsize=15, bold=True, lw=1.8)
    draw_arrow(ax, 4, 12.72, 4, 12.35, lw=1.8)

    # Completion of Sensing
    draw_box(ax, 4, 12.0, 5.8, 0.7,
             "Completion of Sensing", fontsize=15, bold=True, lw=1.8)
    draw_arrow(ax, 4, 11.65, 4, 11.25, lw=1.8)

    # Seizing Stage Processual Bottleneck (large group box)
    group_box = FancyBboxPatch((0.5, 4.6), 7.0, 6.5,
                                boxstyle="square,pad=0",
                                facecolor='#f5f5f5', edgecolor='black', linewidth=2.0)
    ax.add_patch(group_box)
    ax.text(4, 10.8, "Seizing Stage Processual Bottleneck",
            ha='center', va='center', fontsize=16, fontweight='bold', fontfamily='serif')

    # 6 sub-items inside
    items = [
        "Absence of Formal Authorisation",
        "Ambiguity of Responsible Actor",
        "Undefined Budget and Schedule",
        "Non-Agenda Formation",
        "Procedural Non-Activation",
        "Non-Event State",
    ]
    sub_y_start = 10.05
    sub_gap = 0.95
    for i, item in enumerate(items):
        y = sub_y_start - i * sub_gap
        draw_box(ax, 4, y, 5.8, 0.65, item, fontsize=14, bold=False,
                 lw=1.2, fill='white')
        if i < len(items) - 1:
            draw_arrow(ax, 4, y - 0.32, 4, y - sub_gap + 0.32, lw=1.2)

    draw_arrow(ax, 4, 4.6, 4, 4.1, lw=1.8)

    # Formal Decision Non-Occurrence
    draw_box(ax, 4, 3.65, 6.5, 0.8,
             "Formal Decision Non-Occurrence",
             fontsize=15, bold=True, lw=1.8)
    draw_arrow(ax, 4, 3.25, 4, 2.75, lw=1.8)

    # Non-Initiation of Transforming
    draw_box(ax, 4, 2.3, 6.5, 0.85,
             "Non-Initiation of Transforming",
             fontsize=15, bold=True, lw=1.8)

    fig.tight_layout(pad=0.4)
    path = os.path.join(OUT_DIR, 'fig_4_2.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_4_2.png")
    return path


# ============================================================
# Figure 5.1: Self-Reinforcing Mechanism of the Non-decision State
# Original: 973 x 684 px  (circular diagram)
# ============================================================
def fig_5_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(6.5, 4.6), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')

    # Center label
    draw_box(ax, 5, 3.5, 3.5, 1.1,
             "Non-decision State\n(Seizing Non-activation\nState)",
             fontsize=11, bold=True, lw=2.0)

    # 5 surrounding boxes
    bw, bh = 2.8, 0.75
    boxes = [
        (5.0, 6.2, "No Timetable"),
        (8.5, 5.0, "Absence of\nAccountable Ownership"),
        (7.5, 1.5, "Procedural\nNon-activation"),
        (2.5, 1.5, "Invisibility\nof Stagnation"),
        (1.5, 5.0, "Absence of\nExplicit Rejection"),
    ]
    for (bx, by, label) in boxes:
        draw_box(ax, bx, by, bw, bh, label, fontsize=10, bold=True, lw=1.8)

    # Curved arrows clockwise
    arrow_pairs = [
        ((5 + bw/2 + 0.05, 6.2), (8.5, 5.0 + bh/2 + 0.05), 0.3),
        ((8.5, 5.0 - bh/2 - 0.05), (7.5 + bw/2*0.3, 1.5 + bh/2 + 0.05), 0.3),
        ((7.5 - bw/2 - 0.05, 1.5), (2.5 + bw/2 + 0.05, 1.5), -0.3),
        ((2.5 - bw/2*0.3, 1.5 + bh/2 + 0.05), (1.5, 5.0 - bh/2 - 0.05), 0.3),
        ((1.5, 5.0 + bh/2 + 0.05), (5 - bw/2 - 0.05, 6.2), 0.3),
    ]
    for (start, end, rad) in arrow_pairs:
        arrow = FancyArrowPatch(
            start, end,
            connectionstyle=f"arc3,rad={rad}",
            arrowstyle='->', lw=2.0, color='black',
            mutation_scale=18
        )
        ax.add_patch(arrow)

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_5_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_1.png")
    return path


# ============================================================
# Figure 5.2: Processual Mechanism of Judgment-Decision-Making
#              Conversion Stagnation at the Seizing Stage
# Original: 1024 x 410 px  (horizontal flow)
# ============================================================
def fig_5_2():
    dpi = 150
    fig, ax = plt.subplots(figsize=(6.8, 2.8), dpi=dpi)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Box 1: Organizational Judgment
    draw_box(ax, 1.3, 3.2, 2.2, 1.3,
             "Organizational\nJudgment", fontsize=10, bold=False, lw=1.8)
    draw_arrow(ax, 2.4, 3.2, 3.0, 3.2, style='->', lw=1.8)

    # Box 2: Legitimated Judgment
    draw_box(ax, 4.0, 3.2, 2.0, 1.3,
             "Legitimated\nJudgment", fontsize=10, bold=False, lw=1.8)
    draw_arrow(ax, 5.0, 3.2, 5.7, 3.2, style='->', lw=1.8)

    # Box 3: Process Initiation Failure Zone (large)
    zone_box = FancyBboxPatch((5.8, 2.0), 3.8, 2.4,
                               boxstyle="square,pad=0",
                               facecolor='white', edgecolor='black', linewidth=1.8)
    ax.add_patch(zone_box)
    ax.text(7.7, 4.05, "Process Initiation Failure Zone",
            ha='center', va='center', fontsize=10, fontweight='bold', fontfamily='serif')
    items = [
        "\u2022 No Agenda Setter",
        "\u2022 No Responsibility Allocation",
        "\u2022 No Procedural Trigger",
        "\u2022 No Timetable",
    ]
    for j, item in enumerate(items):
        ax.text(6.1, 3.55 - j * 0.38, item, ha='left', va='center',
                fontsize=9, fontfamily='serif')

    # Arrow down to Non-Decision State
    draw_arrow(ax, 7.7, 2.0, 7.7, 1.45, lw=1.8)
    draw_box(ax, 7.7, 0.95, 4.2, 0.7,
             "Non-Decision State (Stagnation)",
             fontsize=10, bold=True, lw=1.8)

    # Dashed arrow to Decision-Making (Not Activated)
    draw_arrow(ax, 9.6, 3.2, 10.5, 3.2, style='->', lw=1.2, color='gray')
    dashed_box = FancyBboxPatch((10.5, 2.4), 1.3, 1.6,
                                 boxstyle="square,pad=0",
                                 facecolor='white', edgecolor='gray',
                                 linewidth=1.2, linestyle='dashed')
    ax.add_patch(dashed_box)
    ax.text(11.15, 3.2, "Decision-\nMaking\n(Not\nActivated)",
            ha='center', va='center', fontsize=8, fontfamily='serif', color='gray')

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_5_2.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_2.png")
    return path


# ============================================================
# Figure 5.3: Judgment Life Cycle Framework and
#              Conditional Conversion Pathways
# Original: 1534 x 1322 px -- MUST be large
# ============================================================
def fig_5_3():
    dpi = 150
    fig, ax = plt.subplots(figsize=(10.2, 8.8), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    # Stage 1: Conception
    draw_box(ax, 5, 9.3, 7.0, 0.85,
             "[Stage 1: Conception]\nAI Training Request Approval",
             fontsize=15, bold=False, lw=2.0)
    draw_arrow(ax, 5, 8.87, 5, 8.45, lw=2.0)

    # Stage 2: Floating
    draw_box(ax, 5, 8.1, 7.0, 0.7,
             "[Stage 2: Floating]\nNo Agenda-setting",
             fontsize=15, bold=False, lw=2.0)
    draw_arrow(ax, 5, 7.75, 5, 7.3, lw=2.0)

    # Stage 3: Stagnation (large box with gray fill)
    stag_w, stag_h = 8.0, 3.0
    stag_cy = 5.65
    stag_box = FancyBboxPatch((5 - stag_w/2, stag_cy - stag_h/2), stag_w, stag_h,
                               boxstyle="square,pad=0",
                               facecolor='#f0f0f0', edgecolor='black', linewidth=2.0)
    ax.add_patch(stag_box)
    ax.text(5, 6.9, "[Stage 3: Stagnation]", ha='center', va='center',
            fontsize=17, fontweight='bold', fontfamily='serif')
    ax.text(5, 6.45, "Self-reinforcing Structure", ha='center', va='center',
            fontsize=14, fontfamily='serif')

    items = [
        "\u2022 No Timetable",
        "\u2022 No Ownership",
        "\u2022 Non-activation",
        "\u2022 No Visibility",
    ]
    for j, item in enumerate(items):
        ax.text(3.2, 5.9 - j * 0.4, item, ha='left', va='center',
                fontsize=14, fontfamily='serif')

    # Self-reinforcing loop: circular arrow on the right side of the list
    ax.text(7.0, 5.5, "Self-reinforcing\n     Loop", ha='center', va='center',
            fontsize=12, fontstyle='italic', fontfamily='serif')
    loop_arrow = FancyArrowPatch((7.7, 5.9), (7.7, 5.1),
                                  connectionstyle="arc3,rad=-1.2",
                                  arrowstyle='->', lw=2.0, color='black',
                                  mutation_scale=16)
    ax.add_patch(loop_arrow)

    draw_arrow(ax, 5, stag_cy - stag_h/2, 5, 3.75, lw=2.0)

    # Stage 4: Bifurcation Point
    draw_box(ax, 5, 3.35, 6.0, 0.7,
             "[Stage 4: Bifurcation Point]",
             fontsize=16, bold=True, lw=2.0)

    # Two branches with angled lines
    draw_arrow(ax, 3.2, 3.0, 2.0, 2.35, lw=2.0)
    draw_arrow(ax, 6.8, 3.0, 8.0, 2.35, lw=2.0)

    # Extinction (left)
    draw_box(ax, 2.0, 1.9, 3.0, 0.7, "Extinction",
             fontsize=16, bold=True, lw=2.0)

    # Revival (right)
    draw_box(ax, 8.0, 1.9, 3.0, 0.7, "Revival",
             fontsize=16, bold=True, lw=2.0)

    draw_arrow(ax, 8.0, 1.55, 8.0, 1.05, lw=2.0)

    # Formal Decision-Making
    draw_box(ax, 8.0, 0.6, 3.5, 0.7,
             "Formal Decision-Making",
             fontsize=15, bold=True, lw=2.0)

    fig.tight_layout(pad=0.4)
    path = os.path.join(OUT_DIR, 'fig_5_3.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_3.png")
    return path


# ============================================================
# Figure 6.1: The Judgment-Decision Non-Initiation Process
# Original: 912 x 566 px
# ============================================================
def fig_6_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(6.1, 3.8), dpi=dpi)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Box 1: Organizational Judgment Formed
    draw_box(ax, 4, 5.0, 5.5, 0.9,
             "Organizational Judgment\nFormed",
             fontsize=13, bold=False, lw=2.0)

    draw_arrow(ax, 4, 4.55, 4, 3.95, style='->', lw=2.5, color='#333333')

    # Box 2: Process Initiation Failure
    draw_box(ax, 4, 3.4, 5.5, 0.9,
             "Process Initiation Failure",
             fontsize=13, bold=False, lw=2.0)

    draw_arrow(ax, 4, 2.95, 4, 2.35, style='->', lw=2.5, color='#333333')

    # Box 3: Decision Not Activated (Non-Decision)
    draw_box(ax, 4, 1.8, 5.5, 0.9,
             "Decision Not Activated\n(Non-Decision)",
             fontsize=13, bold=False, lw=2.0)

    # Note at bottom
    ax.text(4, 0.65,
            "Note. Organizational judgment was formed, but the decision-making\n"
            "process was not initiated at the Seizing stage.",
            ha='center', va='center', fontsize=9.5, fontstyle='italic', fontfamily='serif')

    fig.tight_layout(pad=0.4)
    path = os.path.join(OUT_DIR, 'fig_6_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_6_1.png")
    return path


# ============================================================
# Main: Generate all figures
# ============================================================
if __name__ == '__main__':
    print("Generating English-only figures (v2 - matching original format)...")
    paths = []
    paths.append(fig_1_1())
    paths.append(fig_4_1())
    paths.append(fig_4_2())
    paths.append(fig_5_1())
    paths.append(fig_5_2())
    paths.append(fig_5_3())
    paths.append(fig_6_1())
    print(f"\nAll {len(paths)} figures generated in: {OUT_DIR}")
    for p in paths:
        sz = os.path.getsize(p)
        from PIL import Image
        img = Image.open(p)
        print(f"  {os.path.basename(p)}: {img.size[0]}x{img.size[1]}px, {sz:,} bytes")
        img.close()
