#!/usr/bin/env python3
"""harvest_template.py — 所定の draw.io テンプレートから、スタイルと寸法を抜き出す（build_drawio.py・build_activity.py が使う）。

  python3 scripts/harvest_template.py                         # assets/ のテンプレート 2 つ → assets/template_styles.json・activity_styles.json
  python3 scripts/harvest_template.py <テンプレート.drawio> -o <出力.json>                # 業務フローのテンプレート
  python3 scripts/harvest_template.py <テンプレート.drawio> -o <出力.json> --activity     # アクティビティ図のテンプレート

テンプレートを差し替えたら、これを実行し直す。図のスタイルはここで抜いたものだけを使う
（check_flow.py が、生成した図のスタイルがここに無いものを FAIL にする）。
要素は id ではなく形で見分ける: 見出し（太字 14pt の四角）と背景（opacity 付き）の組＝レーン、楕円＝開始・終了、
角丸＝プロセス、角丸＋破線＝システム処理、ひし形＝判断、shape=note＝ファイル、線は破線の種類で 通常／システム連携／参照。
テンプレートに無い表現（推定・確認事項の番号・別ページへのリンク）は、テンプレートの色だけを使って extensions に定義する。
アクティビティ図のテンプレート（assets/アクティビティ図_テンプレート.drawio）は、同じ形の決まりに、次を足したもの:
  レーンは左から 操作者／プログラム／外部（角丸＋破線のノードがあるレーン）。開始＝shape=startState、終了＝shape=endState、
  別ページの処理の呼び出し＝角丸＋shape=process、凡例＝id が legend で始まる文字の箱（文は、ここに書いたものがそのまま出る。1行を <br> で区切る）。
  処理の箱の角の丸み（absoluteArcSize・arcSize）は、処理の箱だけに掛かる。分かれた道を横に並べる配置は、機械が決める（テンプレートでは決まらない）。
"""
import argparse
import pathlib
import re
import sys
import xml.etree.ElementTree as ET

from common import ASSETS, save, style_dict as sdict


def geom(cell):
    g = cell.find("mxGeometry")
    if g is None:
        return None
    return {k: float(g.get(k, 0)) for k in ("x", "y", "width", "height")}


def in_lane(g, hg):
    """図形の中心が、そのレーン（見出しの幅）の中にあるか"""
    cx = g["x"] + g["width"] / 2
    return hg["x"] <= cx <= hg["x"] + hg["width"]


def find_lanes(V):
    """レーンの見出し（太字の四角）と背景（opacity 付きの四角）。見出しは左から順"""
    headers, containers = [], []
    for c in V:
        st, g = sdict(c.get("style")), geom(c)
        if "swimlane" in st or "text" in st or g is None:
            continue
        if "opacity" in st and not (c.get("value") or "").strip():
            containers.append((c, st, g))
        elif st.get("fontStyle") == "1" and st.get("rounded") == "0" and (c.get("value") or "").strip() and g["height"] <= 60:
            headers.append((c, st, g))
    if not headers or not containers:
        sys.exit("レーンの見出し（太字の四角）と背景（opacity 付きの四角）が見つからない。テンプレートの形を確かめる")
    headers.sort(key=lambda x: x[2]["x"])
    return headers, containers


def lane_palettes(headers, containers, rounded):
    """レーンごとの配色（見出し・背景・プロセスのスタイル）。角丸＋破線のノードがあるレーンは、システムのレーン"""
    palettes, system = [], None
    for hc, hst, hg in headers:
        cont = next((c for c, _s, g in containers if abs(g["x"] - hg["x"]) < 1), None)
        nodes = [(c, st, g) for c, st, g in rounded if in_lane(g, hg)]
        proc = next((c for c, st, _g in nodes if st.get("dashed") != "1"), None)
        sysn = next((c for c, st, _g in nodes if st.get("dashed") == "1"), None)
        pal = {"name": hc.get("value"), "fill": hst.get("fillColor"), "stroke": hst.get("strokeColor"),
               "header_style": hc.get("style"), "container_style": cont.get("style") if cont is not None else ""}
        if proc is not None:
            pal["process_style"] = proc.get("style")
            palettes.append(pal)
        elif sysn is not None:
            pal["system_style"] = sysn.get("style")
            system = pal
        else:
            palettes.append(pal)
    if not palettes:
        sys.exit("プロセス（角丸の四角）のあるレーンが見つからない")
    base_proc = next(p["process_style"] for p in palettes if p.get("process_style"))
    for p in palettes:                                               # プロセスの見本が無いレーンは、色だけ差し替えて作る
        if not p.get("process_style"):
            st = sdict(base_proc)
            st["fillColor"], st["strokeColor"] = p["fill"], p["stroke"]
            p["process_style"] = ";".join(k if v == "" else f"{k}={v}" for k, v in st.items()) + ";"
    if system is None:
        sys.exit("システムのレーン（角丸＋破線のノードがあるレーン）が見つからない")
    return palettes, system


def edge_styles(E):
    """線の見本を、破線の種類で見分ける: 通常／ラベルつき／システム連携（破線）／参照（点線）"""
    flow = flow_label = system_edge = ref = None
    for c in E:
        st = sdict(c.get("style"))
        if st.get("dashed") == "1":
            if st.get("endArrow") == "open" or st.get("dashPattern", "").startswith("3"):
                ref = ref or c.get("style")
            else:
                system_edge = system_edge or c.get("style")
        elif (c.get("value") or "").strip():
            flow_label = flow_label or c.get("style")
        else:
            flow = flow or c.get("style")
    if None in (flow, system_edge, ref):
        sys.exit("線の見本（通常・システム連携の破線・参照の点線）のどれかがテンプレートに無い")
    return {"flow": flow, "flow_label": flow_label or flow, "system": system_edge, "ref": ref}


def legend_cells(cells):
    """凡例の部品（id が legend で始まるセル）を、そのまま置き直せる形で"""
    out = []
    for c in cells:
        if (c.get("id") or "").startswith("legend"):
            g = c.find("mxGeometry")
            out.append({"id": c.get("id"), "value": c.get("value") or "", "style": c.get("style") or "",
                        "vertex": c.get("vertex") == "1", "edge": c.get("edge") == "1", "parent": c.get("parent"),
                        "source": c.get("source"), "target": c.get("target"),
                        "geometry": {k: g.get(k) for k in ("x", "y", "width", "height", "relative") if g.get(k) is not None}})
    return out


def extensions(red, yellow, blue):
    """テンプレートに無い表現。色はテンプレートのもの（参照線の赤・ファイルの黄・システム連携の線の青）だけを使う"""
    return {
        "inferred": f"dashed=1;dashPattern=4 3;strokeColor={red};",
        "badge": f"ellipse;whiteSpace=wrap;html=1;fillColor={yellow};strokeColor={red};fontSize=9;fontStyle=1;fontColor={red};",
        "subprocess": "shape=process;size=0.07;",
        "edge_label_bg": "labelBackgroundColor=#FFFFFF;",
        "anchor": "rounded=0;fillColor=none;strokeColor=none;",
        "title": "fontSize=16;fontStyle=1;",
        "link_text": "fontSize=11;fontColor=#0066CC;fontStyle=4;",
        "source_font": "font-size:9px;color:#666666",
        "system_tag_font": f"font-size:10px;color:{blue}",
    }


ARC_KEYS = ("absoluteArcSize", "arcSize")                            # 処理の箱だけに掛ける、角の丸み
TEMPLATES = {False: ("業務フロー図_テンプレート.drawio", "template_styles.json"), True: ("アクティビティ図_テンプレート.drawio", "activity_styles.json")}


def join(st):
    return ";".join(k if v == "" else f"{k}={v}" for k, v in st.items()) + ";"


def activity_extras(out, V, cells):
    """アクティビティ図だけの形: 開始・終了・処理の箱の角・別ページの処理の呼び出し・凡例の文"""
    shape = lambda name: next((c for c in V if sdict(c.get("style")).get("shape") == name), None)
    start, end, sub = shape("startState"), shape("endState"), shape("process")
    legend = next((c for c in cells if (c.get("id") or "").startswith("legend") and "text" in sdict(c.get("style")) and (c.get("value") or "").strip()), None)
    if None in (start, end, sub, legend):
        sys.exit("開始（shape=startState）／終了（shape=endState）／別ページの処理の呼び出し（角丸＋shape=process）／凡例（id が legend で始まる文字の箱）の"
                 "どれかが、アクティビティ図のテンプレートに無い")
    arc = {}
    for pal in out["palettes"] + [out["system_palette"]]:            # 角の丸みを、処理の箱のスタイルから分ける（呼び出しの箱・外部の箱には掛けない）
        for key in ("process_style", "system_style"):
            if pal.get(key):
                st = sdict(pal[key])
                arc.update({k: st.pop(k) for k in ARC_KEYS if k in st})
                pal[key] = join(st)
    out["process_extra"] = "".join(f"{k}={arc[k]};" for k in ARC_KEYS if k in arc)
    out["start"], out["end"] = start.get("style"), end.get("style")
    out["start_end_size"] = [geom(start)["width"], geom(start)["height"]]
    out["extensions"]["subprocess"] = f"shape=process;size={sdict(sub.get('style')).get('size', '0.07')};"
    out["legend_lines"] = [ln for ln in re.split(r"<br\s*/?>", legend.get("value")) if ln.strip()]
    out["legend_style"], out["legend_width"] = legend.get("style"), geom(legend)["width"]
    out["legend"] = []                                               # 業務フローの凡例（部品を置き直す形）は、使わない


def harvest(template, activity=False):
    """テンプレート1つを読んで、スタイルと寸法の一式にする"""
    model = ET.parse(template).getroot().find(".//mxGraphModel")
    cells = [c for c in model.iter("mxCell") if c.get("id") not in ("0", "1")]
    top = [c for c in cells if c.get("parent") == "1"]
    V = [c for c in top if c.get("vertex") == "1"]

    headers, containers = find_lanes(V)
    rounded = [(c, sdict(c.get("style")), geom(c)) for c in V if sdict(c.get("style")).get("rounded") == "1" and geom(c)
               and sdict(c.get("style")).get("shape") != "process"]     # 別ページの処理の呼び出しの見本は、処理の箱の見本にしない
    palettes, system = lane_palettes(headers, containers, rounded)

    def first(pred):
        return next((c for c in V if pred(sdict(c.get("style")))), None)

    start = first(lambda s: "ellipse" in s)
    decision = first(lambda s: "rhombus" in s)
    note = first(lambda s: s.get("shape") == "note")
    proc_c = next(c for c, st, g in rounded if st.get("dashed") != "1")
    sys_c = next(c for c, st, g in rounded if st.get("dashed") == "1")
    info = next((c for c in V if "text" in sdict(c.get("style")) and sdict(c.get("style")).get("fillColor", "none") != "none"
                 and not (c.get("id") or "").startswith("legend")), None)
    if None in (start, decision, note):
        sys.exit("開始・終了（楕円）／判断（ひし形）／ファイル（shape=note）のどれかがテンプレートに無い")
    edges = edge_styles([c for c in top if c.get("edge") == "1"])

    hg0 = headers[0][2]
    cg0 = next(g for c, _s, g in containers if abs(g["x"] - hg0["x"]) < 1)
    pg, ng = geom(proc_c), geom(note)
    lane_x = lambda g: next(h[2] for h in headers if in_lane(g, h[2]))["x"]
    proc_ys = sorted(g["y"] for c, st, g in rounded if st.get("dashed") != "1")
    pitches = [b - a_ for a_, b in zip(proc_ys, proc_ys[1:]) if b - a_ > 0]
    legend = legend_cells(cells)
    legend_root = next((x for x in legend if x["parent"] == "1"), None)
    red = sdict(edges["ref"]).get("strokeColor", "#b85450")
    yellow = sdict(note.get("style")).get("fillColor", "#fff2cc")
    text_style = next((c.get("style") for c in cells if "text" in sdict(c.get("style")) and sdict(c.get("style")).get("fillColor") == "none"),
                      "text;html=1;strokeColor=none;fillColor=none;align=left;verticalAlign=middle;whiteSpace=wrap;fontSize=10;")
    size = lambda c: [geom(c)["width"], geom(c)["height"]]

    out = {
        "template": pathlib.Path(template).name,
        "model": {k: model.get(k) for k in ("dx", "dy", "grid", "gridSize", "guides", "tooltips", "connect", "arrows", "fold",
                                            "page", "pageScale", "pageWidth", "pageHeight", "math", "shadow") if model.get(k) is not None},
        "lane": {"x0": hg0["x"], "width": hg0["width"], "header_y": hg0["y"], "header_h": hg0["height"], "content_y": cg0["y"],
                 "node_dx": pg["x"] - lane_x(pg), "file_dx": ng["x"] - lane_x(ng),
                 "first_y": min(geom(c)["y"] for c in V if sdict(c.get("style")).get("rounded") == "1" or "ellipse" in sdict(c.get("style"))),
                 "row_pitch": min(pitches) if pitches else 110},
        "palettes": palettes,
        "system_palette": system,
        "node": {"start_end": start.get("style"), "decision": decision.get("style"), "file": note.get("style"),
                 "system": sys_c.get("style")},
        "size": {"process": size(proc_c), "system": size(sys_c), "decision": size(decision), "start_end": size(start), "file": size(note)},
        "edge": edges,
        "legend": legend,
        "legend_gap": (float(legend_root["geometry"]["x"]) - (headers[-1][2]["x"] + headers[-1][2]["width"])) if legend_root else 60,
        "info_style": info.get("style") if info is not None else text_style,
        "text_style": text_style,
        "extensions": extensions(red, yellow, sdict(edges.get("system", "")).get("strokeColor", "#0066CC")),
    }
    if activity:
        activity_extras(out, V, cells)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("template", nargs="?", help="テンプレートの .drawio（省くと、assets/ の 2 つを読む）")
    ap.add_argument("-o", "--out", help="書き出す先（テンプレートを指定したときは、必ず書く）")
    ap.add_argument("--activity", action="store_true", help="指定したテンプレートが、アクティビティ図のもの")
    a = ap.parse_args()
    if a.template and not a.out:
        sys.exit("テンプレートを指定するときは、-o <出力.json> も書く")
    jobs = [(a.template, a.out, a.activity)] if a.template else [(str(ASSETS / t), str(ASSETS / o), act) for act, (t, o) in TEMPLATES.items()]
    for template, dest, activity in jobs:
        out = harvest(template, activity)
        save(dest, out)
        print(f"テンプレート {out['template']} → {dest}")
        print(f"  レーン: 幅 {out['lane']['width']:.0f}・左端 {out['lane']['x0']:.0f}・行の間隔 {out['lane']['row_pitch']:.0f}")
        print(f"  配色 {len(out['palettes'])} 種（{', '.join(p['fill'] for p in out['palettes'])}）＋{'外部' if activity else 'システム'} {out['system_palette']['fill']}")
        print(f"  凡例の{'行' if activity else '部品'} {len(out['legend_lines'] if activity else out['legend'])} 個")


if __name__ == "__main__":
    main()
