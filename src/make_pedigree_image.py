import argparse
import csv
import math
import os
import re
import textwrap


def load_rows(csv_path):
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    by_pk = {}
    for row in rows:
        pk = (row.get("PrimaryKey") or "").strip()
        if not pk:
            continue
        by_pk[pk] = row
    return by_pk


def compute_max_depth(pk, by_pk):
    memo = {}
    visiting = set()

    def depth_for(node_pk):
        if node_pk in memo:
            return memo[node_pk]
        if node_pk in visiting:
            return 0
        visiting.add(node_pk)
        data = by_pk.get(node_pk)
        if data is None:
            memo[node_pk] = 0
            visiting.remove(node_pk)
            return 0
        sire = (data.get("Sire") or "").strip()
        dam = (data.get("Dam") or "").strip()
        depths = [0]
        if sire:
            depths.append(1 + depth_for(sire))
        if dam:
            depths.append(1 + depth_for(dam))
        max_depth = max(depths)
        memo[node_pk] = max_depth
        visiting.remove(node_pk)
        return max_depth

    return depth_for(pk)


def clamp_depth(pk, by_pk, requested_depth):
    max_depth = compute_max_depth(pk, by_pk)
    if requested_depth is None:
        return max_depth
    return max(0, min(max_depth, requested_depth))


def strip_country(name):
    return re.sub(r"\s*\([^)]*\)\s*$", "", name).strip()


def format_horse_label(
    pk, by_pk, include_year, strip_country_tag, single_line=False
):
    data = by_pk.get(pk)
    if data is None:
        return pk
    raw_name = (data.get("Horse Name") or "").strip() or pk
    if strip_country_tag:
        name = strip_country(raw_name)
        country = ""
    else:
        if "(" in raw_name:
            head, tail = raw_name.split("(", 1)
            name = head.strip()
            country = f"({tail}".strip()
        else:
            name = raw_name
            country = ""
    year = (data.get("Year") or "").strip()
    if include_year and year:
        second = f"{country} {year}".strip()
        if single_line:
            return f"{name} {second}".strip()
        return f"{name}\n{second}" if second else f"{name}\n{year}"
    if country:
        return f"{name} {country}".strip() if single_line else f"{name}\n{country}"
    return name


def collect_inbreeding(pk, by_pk, max_depth):
    occurrences = {}

    def walk(node_pk, gen, path, visiting, side):
        if gen > max_depth:
            return
        if node_pk in visiting:
            return
        visiting.add(node_pk)
        new_path = path + [node_pk]
        occurrences.setdefault(node_pk, []).append(
            {"gen": gen, "path": new_path, "side": side}
        )
        data = by_pk.get(node_pk)
        if data is None:
            visiting.remove(node_pk)
            return
        sire = (data.get("Sire") or "").strip()
        dam = (data.get("Dam") or "").strip()
        if gen < max_depth:
            if sire:
                walk(sire, gen + 1, new_path, visiting, side)
            if dam:
                walk(dam, gen + 1, new_path, visiting, side)
        visiting.remove(node_pk)

    data = by_pk.get(pk)
    if data:
        sire = (data.get("Sire") or "").strip()
        dam = (data.get("Dam") or "").strip()
        if sire:
            walk(sire, 1, [], set(), "Sire")
        if dam:
            walk(dam, 1, [], set(), "Dam")

    inbred = {}
    inbred_types = {}
    for ancestor, occs in occurrences.items():
        if len(occs) < 2:
            continue
        gens = [occ["gen"] for occ in occs]
        percentage = sum((0.5**g) for g in gens) * 100.0
        gens_sorted = sorted(gens, reverse=True)
        paths = [occ["path"] for occ in occs]
        sides = {occ["side"] for occ in occs}
        inbred[ancestor] = {"gens": gens_sorted, "percentage": percentage, "paths": paths}
        if sides == {"Sire"}:
            inbred_types[ancestor] = "sire"
        elif sides == {"Dam"}:
            inbred_types[ancestor] = "dam"
        else:
            inbred_types[ancestor] = "both"

    subsumed = set()
    candidates = list(inbred.keys())
    for ancestor in candidates:
        for other in candidates:
            if ancestor == other:
                continue
            all_contained = True
            for path in inbred[ancestor]["paths"]:
                if other not in path:
                    all_contained = False
                    break
                if path.index(other) > path.index(ancestor):
                    all_contained = False
                    break
            if all_contained:
                subsumed.add(ancestor)
                break

    return inbred, subsumed, inbred_types


def polar_to_xy(radius, angle_rad):
    return radius * math.cos(angle_rad), radius * math.sin(angle_rad)


def place_wedge_text(
    ax, text, r_inner, r_outer, angle_start, angle_end, base_size, gen
):
    if not text:
        return
    angle_mid = (angle_start + angle_end) / 2
    angle_deg = math.degrees(angle_mid) % 360
    angle_range = abs(angle_end - angle_start)
    radius_mid = (r_inner + r_outer) / 2

    if gen >= 5:
        rotation = angle_deg
    else:
        rotation = angle_deg - 90

    ring_thickness = max(r_outer - r_inner, 0.001)
    arc_length = max(angle_range * radius_mid, 0.001)
    char_width = max(len(text), 1) * 0.06
    scale_arc = arc_length / char_width
    scale_rad = ring_thickness / 0.08
    scale = min(1.0, scale_arc, scale_rad)
    font_size = max(4, base_size * scale)
    x, y = polar_to_xy(radius_mid, angle_mid)
    ha = "center"
    ax.text(
        x,
        y,
        text,
        ha=ha,
        va="center",
        rotation=rotation,
        rotation_mode="anchor",
        fontsize=font_size,
    )


def draw_chart(pk, by_pk, max_depth, out_path):
    try:
        import matplotlib.pyplot as plt
        from matplotlib.patches import Wedge
    except ImportError as exc:
        raise SystemExit(
            "matplotlib is required. Install with: pip install -r requirements.txt"
        ) from exc

    if max_depth <= 0:
        raise SystemExit("Generations must be at least 1 for the circular chart.")

    inbred, subsumed, inbred_types = collect_inbreeding(pk, by_pk, max_depth)
    inbred_pks = set(inbred.keys()) - subsumed

    fig_size = max(8, 2 + max_depth * 1.3)
    fig, ax = plt.subplots(figsize=(fig_size * 0.9, fig_size * 1.2))
    ax.set_aspect("equal")
    ax.axis("off")

    sire_fill = "#DDEBF7"
    dam_fill = "#FCE4EC"
    unknown_fill = "#F2F2F2"
    edge_default = "#FFFFFF"
    edge_inbred_dam = "#FF0000"
    edge_inbred_sire = "#1F4E9E"
    edge_inbred_both = "#800080"

    extra_per_gen = 0.012
    max_extra = extra_per_gen * max(0, max_depth - 7)

    def ring_outer(gen):
        base = gen / max_depth
        if gen >= 8:
            return base + extra_per_gen * (gen - 7)
        return base

    def ring_inner(gen):
        base = (gen - 1) / max_depth
        if gen >= 8:
            return base + extra_per_gen * (gen - 8)
        return base

    def draw_wedge(node_pk, gen, angle_start, angle_end):
        data = by_pk.get(node_pk)
        sex = (data.get("Sex") or "").strip().upper() if data else ""
        if sex == "H" or sex == "G" or sex == "C":
            face = sire_fill
        elif sex == "M" or sex == "F":
            face = dam_fill
        else:
            face = unknown_fill

        r_inner = ring_inner(gen)
        r_outer = ring_outer(gen)
        wedge = Wedge(
            center=(0.0, 0.0),
            r=r_outer,
            theta1=math.degrees(angle_start),
            theta2=math.degrees(angle_end),
            width=r_outer - r_inner,
            facecolor=face,
            edgecolor=edge_default,
            linewidth=0.5,
        )
        ax.add_patch(wedge)
        if node_pk in inbred_pks:
            inbred_type = inbred_types.get(node_pk)
            if inbred_type == "sire":
                wedge.set_edgecolor(edge_inbred_sire)
            elif inbred_type == "dam":
                wedge.set_edgecolor(edge_inbred_dam)
            else:
                wedge.set_edgecolor(edge_inbred_both)
            wedge.set_linewidth(1.2)

        if gen >= 8:
            label = format_horse_label(
                node_pk,
                by_pk,
                include_year=False,
                strip_country_tag=True,
                single_line=True,
            )
        else:
            label = format_horse_label(
                node_pk,
                by_pk,
                include_year=True,
                strip_country_tag=False,
                single_line=False,
            )

        font_size = max(0.3, 11 / (2 ** (gen - 1)))
        place_wedge_text(
            ax, label, r_inner, r_outer, angle_start, angle_end, font_size, gen
        )

        if gen >= max_depth:
            return

        if data is None:
            return

        sire = (data.get("Sire") or "").strip()
        dam = (data.get("Dam") or "").strip()
        if not sire and not dam:
            return

        mid = (angle_start + angle_end) / 2
        # Left half (positive angles): top = dam, bottom = sire.
        # Right half (negative to positive): top = sire, bottom = dam.
        if angle_start > 0 and angle_end > 0:
            sire_range = (mid, angle_end)
            dam_range = (angle_start, mid)
        else:
            sire_range = (mid, angle_end)
            dam_range = (angle_start, mid)
        if sire:
            draw_wedge(sire, gen + 1, sire_range[0], sire_range[1])
        if dam:
            draw_wedge(dam, gen + 1, dam_range[0], dam_range[1])

    data = by_pk.get(pk)
    if data is None:
        raise SystemExit(f"PrimaryKey not found in CSV: {pk}")

    sire = (data.get("Sire") or "").strip()
    dam = (data.get("Dam") or "").strip()

    sire_start, sire_end = math.radians(90), math.radians(270)
    dam_start, dam_end = math.radians(-90), math.radians(90)

    if sire:
        draw_wedge(sire, 1, sire_start, sire_end)
    if dam:
        draw_wedge(dam, 1, dam_start, dam_end)

    data = by_pk.get(pk) or {}
    raw_name = (data.get("Horse Name") or "").strip() or pk
    year = (data.get("Year") or "").strip()
    center_label = f"{raw_name} {year}".strip()
    title_text = ax.text(
        0,
        1.14,
        center_label,
        ha="center",
        va="bottom",
        fontsize=12,
        weight="bold",
    )

    if inbred_pks:
        summary_parts = []
        for ancestor_pk in sorted(
            inbred_pks, key=lambda key: inbred[key]["percentage"], reverse=True
        ):
            data = by_pk.get(ancestor_pk)
            if data:
                name = strip_country((data.get("Horse Name") or "").strip())
            else:
                name = ancestor_pk
            gens_text = " x ".join(
                str(g) for g in sorted(inbred[ancestor_pk]["gens"])
            )
            summary_parts.append(
                f"{name} {inbred[ancestor_pk]['percentage']:.2f}% {gens_text}"
            )
        summary = " / ".join(summary_parts)
    else:
        summary = "No inbreeding detected within selected generations."

    fig_width_in = fig.get_size_inches()[0]
    max_chars = max(60, int(fig_width_in * 10))
    parts = summary.split(" / ")
    lines = []
    current = ""
    for part in parts:
        candidate = part if not current else f"{current} / {part}"
        if len(candidate) > max_chars and current:
            lines.append(current)
            current = part
        else:
            current = candidate
    if current:
        lines.append(current)
    forced_lines = []
    for line in lines:
        if len(line) <= max_chars:
            forced_lines.append(line)
            continue
        forced_lines.extend(
            textwrap.wrap(
                line,
                width=max_chars,
                break_long_words=True,
                break_on_hyphens=False,
            )
        )
    summary_wrapped = "\n".join(forced_lines)
    summary_text = ax.text(0, -1.12, summary_wrapped, ha="center", va="top", fontsize=10)

    radius_limit = 1.05 + max_extra
    ax.set_xlim(-radius_limit, radius_limit)
    ax.set_ylim(-1.4 - max_extra, 1.35 + max_extra)

    fig.savefig(out_path, dpi=200, bbox_inches="tight")
    circle_path = os.path.splitext(out_path)[0] + "_circle.png"
    summary_text.set_visible(False)
    title_text.set_visible(False)
    radius_limit = 1.05 + max_extra
    fig.set_size_inches(fig_size, fig_size)
    ax.set_xlim(-radius_limit, radius_limit)
    ax.set_ylim(-radius_limit, radius_limit)
    fig.savefig(circle_path, dpi=200, bbox_inches="tight", pad_inches=0)
    plt.close(fig)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Create a circular pedigree image from bloodline.csv"
    )
    parser.add_argument("pk", nargs="?", help="PrimaryKey")
    parser.add_argument(
        "--csv",
        default="bloodline.csv",
        help="Path to input CSV (default: bloodline.csv)",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output image path (default: <PK>.png in current directory)",
    )
    parser.add_argument(
        "--gen",
        type=int,
        default=None,
        help="Max generations (default: prompt, fallback to 5)",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    pk = (args.pk or "").strip()
    if not pk:
        pk = input("Enter PrimaryKey: ").strip()
    if not pk:
        raise SystemExit("PrimaryKey is required.")

    csv_path = args.csv
    if not os.path.exists(csv_path):
        raise SystemExit(f"CSV not found: {csv_path}")

    by_pk = load_rows(csv_path)

    gen = args.gen
    if gen is None:
        gen_input = input("Enter generations (default 9): ").strip()
        if gen_input:
            try:
                gen = int(gen_input)
            except ValueError as exc:
                raise SystemExit("Generations must be an integer.") from exc
        else:
            gen = 9
    if gen < 1:
        raise SystemExit("Generations must be >= 1.")

    max_depth = clamp_depth(pk, by_pk, gen)

    out_path = args.out or f"{pk}.png"
    draw_chart(pk, by_pk, max_depth, out_path)
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
