#!/usr/bin/env python3
"""sc_applies.py — 共通の資料の、拠点ごとの書き分けを拾い、applies.json の下書きを書く（工程 1b）。

  python3 scripts/sc_applies.py draft -w <比較の作業フォルダ>     # 候補を拾って applies.json に足す（決めた所には触らない）
  python3 scripts/sc_applies.py show -w <比較の作業フォルダ>      # いまの書き分けを、拠点ごとに見る

拾うもの（共通の作業フォルダのカードの手順・分岐）:
  - 文（作業・問い・引用・条件・小見出し）に、拠点の名前・別名が出てくる項目
  - 「拠点ごと」「支店によって」などの言い回しの出てくる項目
  - 上の項目と同じページ・同じ小見出しにある、ほかの手順・分岐（「大阪支店では支店長が承認する」の、ふだんの側「課長が承認する」を拾うため）
下書きの scope は「未定」。LLM が、項目ごとに当てはまる拠点を決める（all／only／except と sites・evidence）。
拠点ごとのフローは、これに従う（sc_site_ops.py が、外す直しを出す。点検 --stage applies が、フローと突き合わせる）。
"""
import argparse
import sys

from sc_common import A, SWITCH_WORDS, load, mentions, norm, save, shared_work, sites_of, try_load, work_dir


def candidates(cmp, sites):
    """共通の資料のカードから拾った項目（鍵の順）"""
    sh = shared_work(cmp, sites)
    if sh is None:
        return []
    m = load(sh / "manifest.json", None)
    if not m:
        sys.exit(f"共通の資料の棚卸しが無い（{sh}）")
    out = []
    for d in m["docs"]:
        if d.get("duplicate_of"):
            continue
        for p in d.get("page_list") or []:
            c = A.load_card(sh, d["doc_id"], p["page"]) or {}
            items = [(x, x.get("action", "")) for x in c.get("steps") or []] + [(x, x.get("question", "")) for x in c.get("decisions") or []]
            hits = []
            for x, text in items:
                blob = " ".join(str(v or "") for v in (text, x.get("quote"), x.get("condition"), x.get("section")))
                who = mentions(blob, sites)
                hits.append((x, text, who, bool(who) or any(norm(w) in norm(blob) for w in SWITCH_WORDS)))
            secs = {x.get("section", "") for x, _t, _w, hit in hits if hit}
            for x, text, who, hit in hits:
                if hit or x.get("section", "") in secs:
                    out.append({"key": f"{d['doc_id']}.{p['page']}.{x['id']}", "doc": d["sha1"], "text": text or x.get("quote", ""),
                                "mentions": who, "scope": "未定", "sites": [], "evidence": ""})
    return out


def applies_to(item, sid):
    """書き分けの項目が、その拠点に当てはまるか（未定は None）"""
    sc = item.get("scope")
    if sc == "all":
        return True
    if sc == "only":
        return sid in (item.get("sites") or [])
    if sc == "except":
        return sid not in (item.get("sites") or [])
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("cmd", choices=["draft", "show"])
    ap.add_argument("-w", "--work", required=True)
    a = ap.parse_args()
    cmp = work_dir(a.work)
    sites = sites_of(cmp)
    old, err = try_load(cmp / "applies.json")
    if err:
        sys.exit(f"applies.json が読めない（{err}）")
    old = old or {"items": [], "questions": []}
    if a.cmd == "show":
        names = {x["id"]: x["name"] for x in sites["sites"] if x.get("uses_shared", True)}
        for it in old["items"]:
            to = [names[s] for s in names if applies_to(it, s)] if it.get("scope") != "未定" else ["未定"]
            print(f"{it['key']}  {it['text'][:40]}  → {'・'.join(to) or 'どの拠点にも当てはまらない'}")
        return
    cand = candidates(cmp, sites)
    have = {it["key"]: it for it in old["items"]}
    keys = {c["key"] for c in cand}
    items = []
    for c in cand:
        it = have.get(c["key"])
        if it and it.get("doc") == c["doc"]:
            it.update(text=c["text"], mentions=c["mentions"])
            items.append(it)
        else:
            items.append(c)
    dropped = [k for k in have if k not in keys]
    save(cmp / "applies.json", {"items": items, "questions": old.get("questions") or []})
    new = sum(1 for c in cand if c["key"] not in have)
    print(f"書き分けの候補 {len(cand)} 件（新しく足した {new}・決めた所は引き継いだ）→ {cmp / 'applies.json'}")
    if dropped:
        print(f"拾われなくなった項目 {len(dropped)} 件を外した（{', '.join(dropped[:5])}）")
    left = sum(1 for it in items if it["scope"] == "未定")
    if left:
        print(f"未定 {left} 件: 項目ごとに、当てはまる拠点（all・only・except と sites）と根拠（evidence）を書く（references/site-guide.md）")


if __name__ == "__main__":
    main()
