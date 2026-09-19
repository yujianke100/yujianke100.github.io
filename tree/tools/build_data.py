#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把星型设计实验里的类型树打成浏览器可直接 <script src> 的 JS 数据文件。

输入（只读，不改）:
  深树（交付物）: out/deep_controlled/<dom>_deep_tree_none+votestrength.json
  三层对照:      out/final_merge/<dom>_final_tree.json
输出:
  ../data/<dom>_<deep|final>.js

序列化格式（紧简，省体积）:
  window.TREES["arxiv_deep"]={"nodes":["a","b",...],"parent":[父节点下标或-1,...],"stats":{...}};
  只存父指针（每个非根节点恰好一个父），不存完整 edges 对象数组。
"""
import json, os, sys, difflib

DEEP_DIR = "/home/hanw/ontology_learning/experiments/star_design/out/deep_controlled"
FM_DIR = "/home/hanw/ontology_learning/experiments/star_design/out/final_merge"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(os.path.dirname(HERE), "data")

DOMAINS = [("arxiv", "arXiv"), ("cfr", "CFR"), ("me", "MultiEURLEX")]


def discover(directory, domain, family):
    """按前缀匹配发现真实文件名（不手打配置名）。"""
    if family == "deep":
        prefix = domain + "_deep_tree_none"
    else:
        prefix = domain + "_final_tree"
    cands = sorted(f for f in os.listdir(directory)
                   if f.startswith(prefix) and f.endswith(".json"))
    if not cands:
        allf = sorted(f for f in os.listdir(directory) if f.endswith(".json"))
        near = difflib.get_close_matches(prefix + ".json", allf, n=5, cutoff=0.3)
        raise SystemExit("找不到 %s 下的 %s* ：候选 %s" % (directory, prefix, near))
    if len(cands) > 1:
        print("  [warn] 多个候选，取第一个:", cands)
    return os.path.join(directory, cands[0])


def build(path, key):
    with open(path, "r", encoding="utf-8") as fh:
        raw = json.load(fh)
    nodes = raw["nodes"]
    edges = raw["edges"]
    n = len(nodes)
    assert len(set(nodes)) == n, "节点名有重复: %s" % path

    index = {name: i for i, name in enumerate(nodes)}
    parent = [-1] * n
    extra_parents = 0
    for e in edges:
        c = index.get(e.get("child"))
        p = index.get(e.get("parent"))
        if c is None:
            continue
        if p is None:
            # 父节点不在 nodes 里（理论上不该出现），挂到 root
            p = index.get("root_concept", -1)
            if p == c:
                p = -1
        if parent[c] == -1 and p != c:
            parent[c] = p
        elif parent[c] != p and p != c:
            extra_parents += 1  # 多父：保留第一个，其余丢弃（渲染成树）

    if "root_concept" not in index:
        raise SystemExit("没有 root_concept: %s" % path)
    root = index["root_concept"]
    parent[root] = -1

    # 校验：从 root 出发能覆盖多少节点
    kids = [[] for _ in range(n)]
    for c, p in enumerate(parent):
        if p >= 0:
            kids[p].append(c)
    seen = [False] * n
    stack = [root]
    seen[root] = True
    cnt = 0
    while stack:
        v = stack.pop()
        cnt += 1
        for c in kids[v]:
            if not seen[c]:
                seen[c] = True
                stack.append(c)
    unreachable = n - cnt

    st = raw.get("stats") or {}
    stats = {
        "n_nodes": n,
        "n_edges": len(edges),
        "max_depth": st.get("max_depth"),
        "cycles": st.get("cycles"),
        "chain_ge2": st.get("chain_ge2"),
        "chain_ge3": st.get("chain_ge3"),
        "deep_nodes_ge5": st.get("deep_nodes_ge5"),
        "config": st.get("config"),
        "is_one_component": st.get("is_one_component"),
        "unreachable": unreachable,
        "extra_parents": extra_parents,
        "root_children": len(kids[root]),
    }

    payload = {"nodes": nodes, "parent": parent, "stats": stats}
    os.makedirs(OUT_DIR, exist_ok=True)
    out = os.path.join(OUT_DIR, key + ".js")
    with open(out, "w", encoding="utf-8") as fh:
        fh.write('window.TREES=window.TREES||{};window.TREES["%s"]=' % key)
        json.dump(payload, fh, ensure_ascii=False, separators=(",", ":"))
        fh.write(";\n")
    size = os.path.getsize(out)
    print("  %-14s <- %s" % (key, os.path.basename(path)))
    print("     nodes=%d edges=%d max_depth=%s root_children=%d unreachable=%d multi_parent=%d -> %s (%.1f KB)"
          % (n, len(edges), stats["max_depth"], stats["root_children"], unreachable,
             extra_parents, out, size / 1024.0))
    return out, size


def main():
    total = 0
    print("== 深树（交付物） deep_controlled ==")
    for dom, disp in DOMAINS:
        p = discover(DEEP_DIR, dom, "deep")
        _, s = build(p, "%s_deep" % dom)
        total += s
    print("== 三层中间件（对照） final_merge ==")
    for dom, disp in DOMAINS:
        p = discover(FM_DIR, dom, "final")
        _, s = build(p, "%s_final" % dom)
        total += s
    print("合计 %.2f MB" % (total / 1048576.0))


if __name__ == "__main__":
    main()
