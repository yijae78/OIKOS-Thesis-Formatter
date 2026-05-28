"""Generate all 7 English-only academic figures using matplotlib.
Replaces bilingual (Korean+English) originals with English-only versions."""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np
import os

OUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'output', 'figures')
os.makedirs(OUT_DIR, exist_ok=True)

# Common style
FONT = 'Times New Roman'
plt.rcParams['font.family'] = 'serif'
plt.rcParams['font.serif'] = [FONT, 'DejaVu Serif']
plt.rcParams['font.size'] = 11


def draw_box(ax, x, y, w, h, text, fontsize=11, bold=False, lw=1.5,
             fill='white', edgecolor='black', ha='center', va='center',
             rounded=True, text_wrap=True):
    """Draw a rounded rectangle with centered text."""
    if rounded:
        box = FancyBboxPatch((x - w/2, y - h/2), w, h,
                             boxstyle="round,pad=0.02",
                             facecolor=fill, edgecolor=edgecolor, linewidth=lw)
    else:
        box = FancyBboxPatch((x - w/2, y - h/2), w, h,
                             boxstyle="square,pad=0",
                             facecolor=fill, edgecolor=edgecolor, linewidth=lw)
    ax.add_patch(box)
    weight = 'bold' if bold else 'normal'
    ax.text(x, y, text, ha=ha, va=va, fontsize=fontsize, fontweight=weight,
            fontfamily='serif', wrap=text_wrap)


def draw_arrow(ax, x1, y1, x2, y2, style='->', lw=1.5, color='black'):
    """Draw an arrow between two points."""
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle=style, lw=lw, color=color))


# ============================================================
# Figure 1.1: Overall Logical Flow and Analytical Framework
# 1202 x 846 px
# ============================================================
def fig_1_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(1202/dpi, 846/dpi), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')

    steps = [
        ("Step 1: Research Problem",
         "AI utilization training judgment formation\n"
         "\u2192 [Structural Disconnection] \u2192\n"
         "Strategic decision-making not generated"),
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
    box_w = 8.5
    header_h = 0.35
    body_h = 0.65

    for i, (title, body) in enumerate(steps):
        y = y_positions[i]
        # Header box (dark fill with triangle marker)
        header_y = y + body_h/2 + header_h/2
        box_header = FancyBboxPatch((5 - box_w/2, header_y - header_h/2), box_w, header_h,
                                     boxstyle="square,pad=0",
                                     facecolor='#2a2a2a', edgecolor='black', linewidth=1.5)
        ax.add_patch(box_header)
        # Draw filled triangle marker instead of unicode
        tri = plt.Polygon([[0.9, header_y - 0.08], [0.9, header_y + 0.08],
                            [1.1, header_y]], closed=True, facecolor='white',
                           edgecolor='white', linewidth=0.5)
        ax.add_patch(tri)
        ax.text(1.4, header_y, title, ha='left', va='center', fontsize=10,
                fontweight='bold', color='white', fontfamily='serif')

        # Body box
        body_box = FancyBboxPatch((5 - box_w/2, y - body_h/2), box_w, body_h,
                                   boxstyle="square,pad=0",
                                   facecolor='white', edgecolor='black', linewidth=1.0)
        ax.add_patch(body_box)
        ax.text(5, y, body, ha='center', va='center', fontsize=8.5, fontfamily='serif')

        # Arrow to next step
        if i < len(steps) - 1:
            next_y = y_positions[i+1]
            draw_arrow(ax, 5, y - body_h/2 - 0.02, 5,
                       next_y + body_h/2 + header_h + 0.02, lw=1.5)

    fig.tight_layout(pad=0.2)
    path = os.path.join(OUT_DIR, 'fig_1_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_1_1.png")
    return path


# ============================================================
# Figure 4.1: Structural Disconnection of Judgment-Decision
#              Conversion and the Seizing Non-initiation Mechanism
# 1358 x 1341 px
# ============================================================
def fig_4_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(1358/dpi, 1341/dpi), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    # Sensing box (top)
    draw_box(ax, 5, 9.3, 4.5, 0.6, "Sensing", fontsize=13, bold=True)

    draw_arrow(ax, 5, 9.0, 5, 8.6, lw=1.5)

    # Seizing label
    ax.text(5, 8.35, "Seizing", ha='center', va='center', fontsize=13,
            fontweight='bold', fontfamily='serif')

    # Seizing bracket area
    seizing_top = 8.1
    seizing_bot = 3.2

    # Judgment Formation
    draw_box(ax, 5, 7.6, 6.0, 0.6, "Judgment Formation", fontsize=11, bold=True)
    draw_arrow(ax, 5, 7.3, 5, 6.9, lw=1.5)

    # Discontinuity
    draw_box(ax, 5, 6.5, 6.0, 0.6,
             "Discontinuity of Decision Initiation", fontsize=11, bold=True)
    draw_arrow(ax, 5, 6.2, 5, 5.8, lw=1.5)

    # Decision Initiation Conditions Not Activated (large box)
    big_box = FancyBboxPatch((2.0, 4.3), 6.0, 1.3,
                              boxstyle="round,pad=0.03",
                              facecolor='white', edgecolor='black', linewidth=1.5)
    ax.add_patch(big_box)
    ax.text(5, 5.3, "Decision Initiation Conditions Not Activated",
            ha='center', va='center', fontsize=11, fontweight='bold', fontfamily='serif')
    ax.text(5, 4.95, "- Absence of Trigger", ha='center', va='center',
            fontsize=9.5, fontfamily='serif')
    ax.text(5, 4.7, "- Failed Agenda-setting", ha='center', va='center',
            fontsize=9.5, fontfamily='serif')
    ax.text(5, 4.45, "- Uninitiated Decision Pathway", ha='center', va='center',
            fontsize=9.5, fontfamily='serif')

    draw_arrow(ax, 5, 4.3, 5, 3.9, lw=1.5)

    # Sustained Non-Event State
    draw_box(ax, 5, 3.5, 6.0, 0.6,
             "Sustained Non-Event State", fontsize=11, bold=True)

    # Seizing bracket (left side)
    ax.plot([1.5, 1.5], [seizing_bot, seizing_top], color='black', lw=1.0)
    ax.plot([1.5, 1.7], [seizing_top, seizing_top], color='black', lw=1.0)
    ax.plot([1.5, 1.7], [seizing_bot, seizing_bot], color='black', lw=1.0)

    draw_arrow(ax, 5, 3.2, 5, 2.8, lw=1.5)

    # Transforming box (bottom, with note)
    trans_box = FancyBboxPatch((2.5, 1.8), 5.0, 0.8,
                                boxstyle="round,pad=0.03",
                                facecolor='white', edgecolor='black', linewidth=1.5)
    ax.add_patch(trans_box)
    ax.text(5, 2.3, "Transforming", ha='center', va='center',
            fontsize=13, fontweight='bold', fontfamily='serif')
    ax.text(5, 2.0, "(Not Initiated)", ha='center', va='center',
            fontsize=10, fontstyle='italic', fontfamily='serif')

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_4_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_4_1.png")
    return path


# ============================================================
# Figure 4.2: Processual Bottleneck Mechanism at the Seizing Stage
# 969 x 1702 px
# ============================================================
def fig_4_2():
    dpi = 150
    fig, ax = plt.subplots(figsize=(969/dpi, 1702/dpi), dpi=dpi)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 14)
    ax.axis('off')

    # Top: Organisational Judgment on AI Utilisation Request
    draw_box(ax, 4, 13.3, 6.5, 0.8,
             "Organisational Judgment\non AI Utilisation Request",
             fontsize=10, bold=True)
    draw_arrow(ax, 4, 12.9, 4, 12.5, lw=1.5)

    # Completion of Sensing
    draw_box(ax, 4, 12.1, 5.0, 0.6,
             "Completion of Sensing", fontsize=10, bold=True)
    draw_arrow(ax, 4, 11.8, 4, 11.3, lw=1.5)

    # Seizing Stage Processual Bottleneck (large group box)
    group_box = FancyBboxPatch((0.8, 4.8), 6.4, 6.3,
                                boxstyle="round,pad=0.05",
                                facecolor='#f8f8f8', edgecolor='black', linewidth=2.0)
    ax.add_patch(group_box)
    ax.text(4, 10.8, "Seizing Stage Processual Bottleneck",
            ha='center', va='center', fontsize=11, fontweight='bold', fontfamily='serif')

    # 6 sub-items inside
    items = [
        "Absence of Formal Authorisation",
        "Ambiguity of Responsible Actor",
        "Undefined Budget and Schedule",
        "Non-Agenda Formation",
        "Procedural Non-Activation",
        "Non-Event State",
    ]
    sub_y_start = 10.1
    sub_gap = 0.95
    for i, item in enumerate(items):
        y = sub_y_start - i * sub_gap
        draw_box(ax, 4, y, 5.5, 0.6, item, fontsize=9.5, bold=False,
                 lw=1.0, fill='white')
        if i < len(items) - 1:
            draw_arrow(ax, 4, y - 0.3, 4, y - sub_gap + 0.3, lw=1.0)

    draw_arrow(ax, 4, 4.8, 4, 4.3, lw=1.5)

    # Formal Decision Non-Occurrence
    draw_box(ax, 4, 3.8, 6.5, 0.7,
             "Formal Decision Non-Occurrence",
             fontsize=10, bold=True)
    draw_arrow(ax, 4, 3.45, 4, 2.95, lw=1.5)

    # Non-Initiation of Transforming
    draw_box(ax, 4, 2.4, 6.5, 0.8,
             "Non-Initiation of Transforming",
             fontsize=10, bold=True)

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_4_2.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_4_2.png")
    return path


# ============================================================
# Figure 5.1: Self-Reinforcing Mechanism of the Non-decision State
# 973 x 684 px  (circular diagram)
# ============================================================
def fig_5_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(973/dpi, 684/dpi), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis('off')

    # Center label
    draw_box(ax, 5, 3.5, 3.5, 1.0,
             "Non-decision State\n(Seizing Non-activation\nState)",
             fontsize=9, bold=True, lw=1.5)

    # 5 surrounding boxes: top-center, top-right, bottom-right, bottom-left, top-left
    bw, bh = 2.5, 0.7
    boxes = [
        (5.0, 6.2, "No Timetable"),
        (8.5, 5.0, "Absence of\nAccountable Ownership"),
        (7.5, 1.5, "Procedural\nNon-activation"),
        (2.5, 1.5, "Invisibility\nof Stagnation"),
        (1.5, 5.0, "Absence of\nExplicit Rejection"),
    ]
    for (bx, by, label) in boxes:
        draw_box(ax, bx, by, bw, bh, label, fontsize=8, bold=True, lw=1.5)

    # Draw curved arrows clockwise between boxes (outside the center)
    arrow_pairs = [
        # from box0(top) to box1(top-right): go right
        ((5 + bw/2 + 0.05, 6.2), (8.5, 5.0 + bh/2 + 0.05), 0.3),
        # from box1(top-right) to box2(bottom-right): go down
        ((8.5, 5.0 - bh/2 - 0.05), (7.5 + bw/2*0.3, 1.5 + bh/2 + 0.05), 0.3),
        # from box2(bottom-right) to box3(bottom-left): go left
        ((7.5 - bw/2 - 0.05, 1.5), (2.5 + bw/2 + 0.05, 1.5), -0.3),
        # from box3(bottom-left) to box4(top-left): go up
        ((2.5 - bw/2*0.3, 1.5 + bh/2 + 0.05), (1.5, 5.0 - bh/2 - 0.05), 0.3),
        # from box4(top-left) to box0(top): go right
        ((1.5, 5.0 + bh/2 + 0.05), (5 - bw/2 - 0.05, 6.2), 0.3),
    ]
    for (start, end, rad) in arrow_pairs:
        arrow = FancyArrowPatch(
            start, end,
            connectionstyle=f"arc3,rad={rad}",
            arrowstyle='->', lw=2.0, color='black',
            mutation_scale=15
        )
        ax.add_patch(arrow)

    fig.tight_layout(pad=0.2)
    path = os.path.join(OUT_DIR, 'fig_5_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_1.png")
    return path


# ============================================================
# Figure 5.2: Processual Mechanism of Judgment-Decision-Making
#              Conversion Stagnation at the Seizing Stage
# 1024 x 410 px  (horizontal flow)
# ============================================================
def fig_5_2():
    dpi = 150
    fig, ax = plt.subplots(figsize=(1024/dpi, 410/dpi), dpi=dpi)
    ax.set_xlim(0, 12)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Box 1: Organizational Judgment
    draw_box(ax, 1.3, 3.2, 2.2, 1.2,
             "Organizational\nJudgment", fontsize=9, bold=False, lw=1.5)

    draw_arrow(ax, 2.4, 3.2, 3.0, 3.2, style='->', lw=1.5)

    # Box 2: Legitimated Judgment
    draw_box(ax, 4.0, 3.2, 2.0, 1.2,
             "Legitimated\nJudgment", fontsize=9, bold=False, lw=1.5)

    draw_arrow(ax, 5.0, 3.2, 5.7, 3.2, style='->', lw=1.5)

    # Box 3: Process Initiation Failure Zone (large)
    zone_box = FancyBboxPatch((5.8, 2.1), 3.8, 2.2,
                               boxstyle="round,pad=0.05",
                               facecolor='white', edgecolor='black', linewidth=1.5)
    ax.add_patch(zone_box)
    ax.text(7.7, 4.0, "Process Initiation Failure Zone",
            ha='center', va='center', fontsize=8.5, fontweight='bold', fontfamily='serif')
    items = [
        "\u2022 No Agenda Setter",
        "\u2022 No Responsibility Allocation",
        "\u2022 No Procedural Trigger",
        "\u2022 No Timetable",
    ]
    for j, item in enumerate(items):
        ax.text(6.1, 3.55 - j * 0.35, item, ha='left', va='center',
                fontsize=7.5, fontfamily='serif')

    # Arrow down to Non-Decision State
    draw_arrow(ax, 7.7, 2.1, 7.7, 1.5, lw=1.5)
    draw_box(ax, 7.7, 1.0, 4.0, 0.7,
             "Non-Decision State (Stagnation)",
             fontsize=8.5, bold=True, lw=1.5)

    # Dashed arrow to Decision-Making (Not Activated)
    draw_arrow(ax, 9.6, 3.2, 10.5, 3.2, style='->', lw=1.0, color='gray')
    # Dashed box
    dashed_box = FancyBboxPatch((10.5, 2.5), 1.3, 1.4,
                                 boxstyle="round,pad=0.03",
                                 facecolor='white', edgecolor='gray',
                                 linewidth=1.0, linestyle='dashed')
    ax.add_patch(dashed_box)
    ax.text(11.15, 3.2, "Decision-\nMaking\n(Not\nActivated)",
            ha='center', va='center', fontsize=7, fontfamily='serif', color='gray')

    fig.tight_layout(pad=0.2)
    path = os.path.join(OUT_DIR, 'fig_5_2.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_2.png")
    return path


# ============================================================
# Figure 5.3: Judgment Life Cycle Framework and
#              Conditional Conversion Pathways
# 1534 x 1322 px
# ============================================================
def fig_5_3():
    dpi = 150
    fig, ax = plt.subplots(figsize=(1534/dpi, 1322/dpi), dpi=dpi)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)
    ax.axis('off')

    # Stage 1: Conception
    draw_box(ax, 5, 9.2, 5.5, 0.8,
             "[Stage 1: Conception]\nAI Training Request Approval",
             fontsize=10, bold=False, lw=1.5)
    draw_arrow(ax, 5, 8.8, 5, 8.3, lw=1.5)

    # Stage 2: Floating
    draw_box(ax, 5, 7.9, 5.5, 0.6,
             "[Stage 2: Floating]\nNo Agenda-setting",
             fontsize=10, bold=False, lw=1.5)
    draw_arrow(ax, 5, 7.6, 5, 7.1, lw=1.5)

    # Stage 3: Stagnation (large box with gray fill)
    stag_box = FancyBboxPatch((1.5, 4.5), 7.0, 2.5,
                               boxstyle="round,pad=0.05",
                               facecolor='#f0f0f0', edgecolor='black', linewidth=2.0)
    ax.add_patch(stag_box)
    ax.text(5, 6.7, "[Stage 3: Stagnation]", ha='center', va='center',
            fontsize=11, fontweight='bold', fontfamily='serif')
    ax.text(5, 6.3, "Self-reinforcing Structure", ha='center', va='center',
            fontsize=10, fontfamily='serif')
    items = [
        "\u2022 No Timetable",
        "\u2022 No Ownership",
        "\u2022 Non-activation",
        "\u2022 No Visibility",
    ]
    for j, item in enumerate(items):
        ax.text(3.5, 5.8 - j * 0.3, item, ha='left', va='center',
                fontsize=9.5, fontfamily='serif')

    # Self-reinforcing loop label (no unicode symbols)
    ax.text(5, 4.7, "~  Self-reinforcing Loop  ~",
            ha='center', va='center', fontsize=9.5, fontstyle='italic', fontfamily='serif')
    # Draw small curved arrows on each side
    loop_l = FancyArrowPatch((3.0, 4.85), (3.0, 4.55),
                              connectionstyle="arc3,rad=-0.8",
                              arrowstyle='->', lw=1.2, color='black', mutation_scale=10)
    ax.add_patch(loop_l)
    loop_r = FancyArrowPatch((7.0, 4.55), (7.0, 4.85),
                              connectionstyle="arc3,rad=-0.8",
                              arrowstyle='->', lw=1.2, color='black', mutation_scale=10)
    ax.add_patch(loop_r)

    draw_arrow(ax, 5, 4.5, 5, 3.9, lw=1.5)

    # Stage 4: Bifurcation Point
    draw_box(ax, 5, 3.5, 4.5, 0.6,
             "[Stage 4: Bifurcation Point]",
             fontsize=10, bold=True, lw=1.5)

    # Two branches
    draw_arrow(ax, 3.5, 3.2, 2.0, 2.5, lw=1.5)
    draw_arrow(ax, 6.5, 3.2, 8.0, 2.5, lw=1.5)

    # Extinction (left)
    ax.text(2.0, 2.2, "Extinction", ha='center', va='center',
            fontsize=11, fontweight='bold', fontfamily='serif')

    # Revival (right)
    ax.text(8.0, 2.2, "Revival", ha='center', va='center',
            fontsize=11, fontweight='bold', fontfamily='serif')

    draw_arrow(ax, 8.0, 1.95, 8.0, 1.5, lw=1.5)

    # Formal Decision-Making
    draw_box(ax, 8.0, 1.0, 4.0, 0.7,
             "Formal Decision-Making",
             fontsize=10, bold=True, lw=1.5)

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_5_3.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_5_3.png")
    return path


# ============================================================
# Figure 6.1: The Judgment-Decision Non-Initiation Process
# 912 x 566 px
# ============================================================
def fig_6_1():
    dpi = 150
    fig, ax = plt.subplots(figsize=(912/dpi, 566/dpi), dpi=dpi)
    ax.set_xlim(0, 8)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Box 1: Organizational Judgment Formed
    draw_box(ax, 4, 5.0, 5.5, 0.9,
             "Organizational Judgment\nFormed",
             fontsize=11, bold=False, lw=2.0)

    # Arrow
    draw_arrow(ax, 4, 4.55, 4, 3.95, style='->', lw=2.5, color='#333333')

    # Box 2: Process Initiation Failure
    draw_box(ax, 4, 3.4, 5.5, 0.9,
             "Process Initiation Failure",
             fontsize=11, bold=False, lw=2.0)

    # Arrow
    draw_arrow(ax, 4, 2.95, 4, 2.35, style='->', lw=2.5, color='#333333')

    # Box 3: Decision Not Activated (Non-Decision)
    draw_box(ax, 4, 1.8, 5.5, 0.9,
             "Decision Not Activated\n(Non-Decision)",
             fontsize=11, bold=False, lw=2.0)

    # Note at bottom
    ax.text(4, 0.7,
            "Note. Organizational judgment was formed, but the decision-making\n"
            "process was not initiated at the Seizing stage.",
            ha='center', va='center', fontsize=8, fontstyle='italic', fontfamily='serif')

    fig.tight_layout(pad=0.3)
    path = os.path.join(OUT_DIR, 'fig_6_1.png')
    fig.savefig(path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  Created: fig_6_1.png")
    return path


# ============================================================
# Main: Generate all figures
# ============================================================
if __name__ == '__main__':
    print("Generating English-only figures...")
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
        print(f"  {os.path.basename(p)}: {sz:,} bytes")
