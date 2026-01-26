import argparse
import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional


@dataclass
class Horse:
    key: str
    sire: Optional[str]
    dam: Optional[str]
    sex: str
    color: str
    year: str
    details: str
    url: str
    name: str
    is_placeholder: bool = False


@dataclass
class PedigreeNode:
    horse: Horse
    sire: Optional["PedigreeNode"] = None
    dam: Optional["PedigreeNode"] = None
    row_start: int = 0
    row_end: int = 0
    depth: int = 0


UNKNOWN_HORSE = Horse(
    key="UNKNOWN",
    sire=None,
    dam=None,
    sex="",
    color="",
    year="",
    details="",
    url="",
    name="Unknown",
)


def placeholder_horse() -> Horse:
    return Horse(
        key="",
        sire=None,
        dam=None,
        sex="",
        color="",
        year="",
        details="",
        url="",
        name="",
        is_placeholder=True,
    )


def load_horses(csv_path: Path) -> Dict[str, Horse]:
    horses: Dict[str, Horse] = {}
    with csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            key = (row.get("PrimaryKey") or "").strip()
            if not key:
                continue
            horses[key] = Horse(
                key=key,
                sire=(row.get("Sire") or "").strip() or None,
                dam=(row.get("Dam") or "").strip() or None,
                sex=(row.get("Sex") or "").strip(),
                color=(row.get("Color") or "").strip(),
                year=(row.get("Year") or "").strip(),
                details=(row.get("Details") or "").strip(),
                url=(row.get("URL") or "").strip(),
                name=(row.get("Horse Name") or "").strip() or key,
            )
    return horses


def build_tree(horses: Dict[str, Horse], key: str, generations: int) -> PedigreeNode:
    root_horse = horses.get(key, UNKNOWN_HORSE)

    def build_node(horse: Horse, depth: int) -> PedigreeNode:
        node = PedigreeNode(horse=horse, depth=depth)
        if depth < generations - 1:
            sire_key = horse.sire
            dam_key = horse.dam
            sire_horse = horses.get(sire_key) if sire_key else None
            dam_horse = horses.get(dam_key) if dam_key else None
            if sire_horse is None:
                sire_horse = placeholder_horse()
            if dam_horse is None:
                dam_horse = placeholder_horse()
            node.sire = build_node(sire_horse, depth + 1)
            node.dam = build_node(dam_horse, depth + 1)
        return node

    return build_node(root_horse, 0)


def assign_rows(node: PedigreeNode, leaf_index: int, generations: int) -> int:
    if node.depth >= generations - 1 or (node.sire is None and node.dam is None):
        node.row_start = leaf_index
        node.row_end = leaf_index
        return leaf_index + 1

    if node.sire is not None:
        leaf_index = assign_rows(node.sire, leaf_index, generations)
    if node.dam is not None:
        leaf_index = assign_rows(node.dam, leaf_index, generations)

    node.row_start = node.sire.row_start if node.sire else node.dam.row_start
    node.row_end = node.dam.row_end if node.dam else node.sire.row_end
    return leaf_index


def collect_cells(node: PedigreeNode, cells: Dict[tuple, PedigreeNode]) -> None:
    cells[(node.row_start, node.depth)] = node
    if node.sire:
        collect_cells(node.sire, cells)
    if node.dam:
        collect_cells(node.dam, cells)


def horse_label(horse: Horse) -> str:
    if horse.is_placeholder:
        return "&nbsp;"
    name = horse.name or horse.key
    extras = " ".join(part for part in [horse.year, horse.color] if part)
    if horse.details:
        extras = " ".join(part for part in [extras, horse.details] if part)
    if extras:
        return f"{name}<br><span class=\"meta\">{extras}</span>"
    return name


def horse_class(horse: Horse) -> str:
    if horse.is_placeholder:
        return "b_empty"
    if horse.sex == "M":
        return "b_ml"
    if horse.sex == "F":
        return "b_fml"
    return "b_unknown"


def render_pedigree(root: PedigreeNode, generations: int) -> str:
    leaf_rows = 2 ** (generations - 1)
    cells: Dict[tuple, PedigreeNode] = {}
    collect_cells(root, cells)
    coverage = [0] * generations
    rows_html = []

    for row in range(leaf_rows):
        row_cells = []
        for col in range(generations):
            if coverage[col] > 0:
                coverage[col] -= 1
                continue
            node = cells.get((row, col))
            if node is None:
                row_cells.append("<td class=\"b_empty\">&nbsp;</td>")
                continue
            rowspan = node.row_end - node.row_start + 1
            coverage[col] = rowspan - 1
            label = horse_label(node.horse)
            if node.horse.url:
                label = f"<a href=\"{node.horse.url}\">{label}</a>"
            row_cells.append(
                f"<td class=\"{horse_class(node.horse)}\" rowspan=\"{rowspan}\">{label}</td>"
            )
        rows_html.append("<tr>" + "".join(row_cells) + "</tr>")

    return "\n".join(rows_html)


def build_html(root: PedigreeNode, generations: int) -> str:
    table_rows = render_pedigree(root, generations)
    return f"""<!DOCTYPE html>
<html lang=\"ja\">
<head>
  <meta charset=\"utf-8\">
  <title>血統表</title>
  <style>
    body {{ font-family: "Hiragino Kaku Gothic ProN", "Meiryo", sans-serif; background: #f5f6f8; }}
    .pedigree {{ border-collapse: separate; border-spacing: 2px; width: 100%; max-width: 1100px; margin: 12px auto 24px; }}
    .pedigree td {{ background: #fff; border: 1px solid #d7d7d7; padding: 8px; font-size: 13px; line-height: 1.3; vertical-align: middle; width: {100 / generations:.2f}%; }}
    .pedigree .b_ml {{ background: #eef4ff; }}
    .pedigree .b_fml {{ background: #fff1f4; }}
    .pedigree .b_unknown {{ background: #f2f2f2; color: #777; }}
    .pedigree .b_empty {{ background: #fafafa; color: #bbb; }}
    .pedigree .meta {{ color: #666; font-size: 11px; }}
    .pedigree a {{ color: #1a4fb4; text-decoration: none; }}
    .pedigree a:hover {{ text-decoration: underline; }}
    .title {{ max-width: 1100px; margin: 24px auto 0; font-size: 18px; font-weight: bold; }}
  </style>
</head>
<body>
  <div class="title">{generations}代血統表</div>
  <table class=\"pedigree\">
    {table_rows}
  </table>
</body>
</html>
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="CSVから血統表HTMLを生成します")
    parser.add_argument("csv", type=Path, help="CSVファイルのパス")
    parser.add_argument("root", help="対象馬のPrimaryKey")
    parser.add_argument("--generations", type=int, default=5, help="世代数(既定: 5)")
    parser.add_argument("--output", type=Path, default=Path("pedigree.html"), help="出力HTML")
    args = parser.parse_args()

    horses = load_horses(args.csv)
    root = build_tree(horses, args.root, args.generations)
    assign_rows(root, 0, args.generations)
    html = build_html(root, args.generations)
    args.output.write_text(html, encoding="utf-8")
    print(f"Generated {args.output}")


if __name__ == "__main__":
    main()
