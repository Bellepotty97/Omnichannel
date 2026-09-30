"""Build omni-dashboard.html from the PDEM forecast workbook.

Usage: python3 tools/build_omni_dashboard.py <workbook.xlsx>

Reads sheet "Revenue _ R.1", keeps rows where Revenue Stream Type is Omni Channel
and Responsible section / Person starts with PEM105, and uses column Q
(PO Receive, MB) and column R (GP, MB). Years 2026-2033 = B.E. 2569-2576.
"""
import base64
import collections
import json
import pathlib
import sys

import openpyxl

ROOT = pathlib.Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "tools" / "omni-dashboard.template.html"
OUT = ROOT / "omni-dashboard.html"
METHOD_IMG = ROOT / "assets" / "forecasting-methodology.jpg"

SHEET = "Revenue _ R.1"
YEARS_CE = list(range(2026, 2034))  # 2569-2576
COL = dict(year=0, cat=2, ptype=3, status=4, stream=6, segment=9, person=14, po=16, gp=17)

# Product Category -> core product group (same groups as the previous dashboard;
# B.E. 2569 totals per group match it exactly).
GROUP_OF_CATEGORY = {
    "Disconnecting Switch": "Disconnecting Switches",
    "Distribution Transformer": "Distribution Transformers",
    "Fuse": "Fuses",
    "GIS SF6-free": "HV Switchgear",
    "LBS SF6-free": "HV Switching Devices",
    "Instrument Transformer (Oil Type)": "Instrument Transformers",
    "Instrument Transformer (Dry Type)": "Instrument Transformers",
    "LED Products & system": "LED Lighting Systems",
    "LED Street light": "LED Lighting Systems",
    "Load Break Switch": "MV Switching Devices",
    "Recloser": "MV Switching Devices",
    "Switchboard": "MV and LV Switchgear",
    "Switchgear": "MV and LV Switchgear",
    "Compact substation": "MV and LV Switchgear",
    "Solid Ring Main Unit": "MV and LV Switchgear",
    "Low Voltage Switchboard": "MV and LV Switchgear",
    "Medium Voltage Switchgear": "MV and LV Switchgear",
    "Protection Relay": "MV and LV Switchgear",
    "Solution On Demand & Service": "Metal Enclosures and Cubicles",
    "Metal enclosure": "Metal Enclosures and Cubicles",
    "Power Capacitor": "Power Capacitors",
    "Switch Capacitor": "Power Capacitors",
    "Surge Arrester": "Surge Arresters",
}
# Product types whose group differs from their category's group.
GROUP_OF_TYPE = {
    "IOT Transformer Box": "Metal Enclosures and Cubicles",
    "24kV GIS SF6-free": "MV and LV Switchgear",
    "36kV GIS SF6-free": "MV and LV Switchgear",
}

# B.E. 2566-2568 actuals per group, carried over from the previous dashboard
# (Omni_Forcast_Data_2573_adj.xlsx). [PO/Revenue], [GP]
HISTORY = {
    "Disconnecting Switches": ([119.86, 57.38, 27.5], [18.4, 11.01, 5.88]),
    "Distribution Transformers": ([240.25, 249.92, 135.32], [47.57, 49.94, 22.46]),
    "Fuses": ([188.24, 195.71, 193.92], [60.37, 62.26, 72.35]),
    "Installation, Testing and Commissioning": ([23.87, 2.19, 0.29], [3.7, 0.46, 0.09]),
    "Instrument Transformers": ([310.26, 214.33, 186.89], [118.28, 91.26, 79.87]),
    "Insulators and Connectors": ([4.06, 2.26, 1.81], [1.29, 0.62, 0.13]),
    "LED Lighting Systems": ([38.81, 41.81, 36.99], [12.36, 15.38, 10.88]),
    "MV Switching Devices": ([246.37, 92.71, 242.58], [48.86, 41.06, 79.09]),
    "MV and LV Switchgear": ([242.99, 183.5, 82.86], [51.66, 46.32, 18.44]),
    "Metal Enclosures and Cubicles": ([9.55, 14.49, 29.04], [2.06, 2.92, 7.39]),
    "Power Capacitors": ([5.29, 6.16, 7.77], [1.92, 1.91, 2.35]),
    "Surge Arresters": ([98.0, 87.19, 89.53], [46.05, 45.9, 44.61]),
    "Others": ([20.45, 11.52, 1.67], [6.32, 2.42, 0.64]),
}
# Actual totals (MB) that override the history above; each group is scaled
# pro rata so the stacked chart still sums to the given total.
# B.E. 2566 (2023) is dropped. {B.E. year: (PO/Revenue, GP)}
HISTORY_YEARS = [2567, 2568]
HISTORY_TOTALS = {2567: (955.0, 318.0), 2568: (870.0, 317.0)}


def scaled_history():
    out = {g: ([], []) for g in HISTORY}
    for k, year in enumerate(HISTORY_YEARS):
        i = year - 2566
        for m in (0, 1):
            tot = sum(v[m][i] for v in HISTORY.values())
            f = HISTORY_TOTALS[year][m] / tot if year in HISTORY_TOTALS else 1.0
            for g, v in HISTORY.items():
                out[g][m].append(round(v[m][i] * f, 4))
    return out


def clean(s):
    s = " ".join(str(s).split())
    return s[2:] if s.startswith("- ") else s


def group_of(cat, ptype):
    return GROUP_OF_TYPE.get(ptype) or GROUP_OF_CATEGORY.get(cat) or "Others"


def tree():
    return collections.defaultdict(lambda: {"po": [0.0] * len(YEARS_CE), "gp": [0.0] * len(YEARS_CE)})


def add(node, yi, po, gp):
    node["po"][yi] += po
    node["gp"][yi] += gp


def rounded(node):
    return {"po": [round(v, 4) for v in node["po"]], "gp": [round(v, 4) for v in node["gp"]]}


def main(path):
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True)[SHEET]
    rows = [
        r for r in ws.iter_rows(min_row=2, max_col=25, values_only=True)
        if r[COL["stream"]] and "Omni" in str(r[COL["stream"]])
        and r[COL["person"]] and str(r[COL["person"]]).strip().startswith("PEM105")
        and r[COL["year"]] in YEARS_CE
    ]

    groups, cats = tree(), collections.defaultdict(tree)
    npd, npd_types = tree(), collections.defaultdict(tree)
    segs, seg_groups = tree(), collections.defaultdict(tree)
    for r in rows:
        yi = YEARS_CE.index(r[COL["year"]])
        po, gp = float(r[COL["po"]] or 0), float(r[COL["gp"]] or 0)
        cat, ptype = clean(r[COL["cat"]]), clean(r[COL["ptype"]])
        g = group_of(cat, ptype)
        add(groups[g], yi, po, gp)
        add(cats[g][cat], yi, po, gp)
        if str(r[COL["status"]]).strip() == "New":
            add(npd[g], yi, po, gp)
            add(npd_types[g][ptype], yi, po, gp)
        seg = clean(r[COL["segment"]])
        add(segs[seg], yi, po, gp)
        add(seg_groups[seg][g], yi, po, gp)

    def children(sub):
        return [{"name": k, **rounded(v)} for k, v in sorted(sub.items())]

    omni, hist_all = [], scaled_history()
    nh = len(HISTORY_YEARS)
    for g in sorted(set(groups) | set(HISTORY)):
        hist = hist_all.get(g, ([0.0] * nh, [0.0] * nh))
        cur = rounded(groups[g]) if g in groups else {"po": [0.0] * 8, "gp": [0.0] * 8}
        omni.append({"name": g, "po": hist[0] + cur["po"], "gp": hist[1] + cur["gp"],
                     "children": children(cats[g]) if g in cats else []})

    data = {
        "source": pathlib.Path(path).name,
        "rows": len(rows),
        "omni": omni,
        "npd": [{"name": g, **rounded(npd[g]), "children": children(npd_types[g])} for g in sorted(npd)],
        "channel": [{"name": s, **rounded(segs[s]), "children": children(seg_groups[s])} for s in segs],
    }
    img = "data:image/jpeg;base64," + base64.b64encode(METHOD_IMG.read_bytes()).decode()
    html = (TEMPLATE.read_text(encoding="utf-8")
            .replace("/*__DATA__*/null", json.dumps(data, ensure_ascii=False))
            .replace("__METHOD_IMG__", img))
    OUT.write_text(html, encoding="utf-8")

    n = len(omni[0]["po"])
    print(f"{len(rows)} rows -> {OUT.name}")
    print(f"PO total B.E. {HISTORY_YEARS[0]}-2576:", [round(sum(g["po"][i] for g in omni), 1) for i in range(n)])
    print(f"GP total B.E. {HISTORY_YEARS[0]}-2576:", [round(sum(g["gp"][i] for g in omni), 1) for i in range(n)])


if __name__ == "__main__":
    main(sys.argv[1])
