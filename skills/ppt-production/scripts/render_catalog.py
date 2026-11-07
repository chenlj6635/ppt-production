#!/usr/bin/env python3
"""Build an offline searchable gallery and Markdown index from a verified catalog."""
import argparse
import html
import json
from pathlib import Path
from check_library import validate


def build(library):
    root = Path(library).resolve()
    validate(root)
    catalog = json.loads((root / 'catalog.json').read_text(encoding='utf-8'))
    entries = catalog['entries']
    esc = html.escape
    md = ['# 区域结构素材库', '', '先按页面逻辑选择结构，再统一目标稿样式。样例拼盘不代表完整论证链。', '',
          '[打开可搜索图库](gallery.html) · [原生素材](source.pptx) · [区域结构化原则](../../references/regional-structure.md)', '',
          f"源 SHA-256：`{catalog['source_sha256']}`；共 {len(entries)} 页。", '',
          '使用前运行 scripts/check_library.py；源变化后先更新提取、预览和人工标签。', '']
    cards = []
    for entry in entries:
        e = entry
        md.extend([f"## {e['page']:02d} · {e['name']}", '',
                   f"- 类型：{e['level']}；标签：{' / '.join(e['tags'])}",
                   f"- 可取用区域：{'；'.join(e['regions'])}",
                   f"- 用途：{e['use_for']}", f"- 注意：{e['caution']}",
                   f"- 原生定位：slide_id=`{e['slide_id']}`，`{e['slide_part']}`",
                   f"- [打开本页预览]({e['preview']})", ''])
        terms = ' '.join([e['name'], e['level'], *e['tags'], *e['regions'], e['use_for']])
        preview = esc(e['preview'], quote=True)
        cards.append(f'<article data-search="{esc(terms, quote=True)}">'
                     f'<h2>{e["page"]:02d} · {esc(e["name"])}</h2>'
                     f'<p class="tags">{esc(e["level"])} · {esc(" / ".join(e["tags"]))}</p>'
                     f'<a href="{preview}"><img loading="lazy" src="{preview}" alt="第{e["page"]}页素材"></a>'
                     f'<p><b>可取用：</b>{esc("；".join(e["regions"]))}</p>'
                     f'<p><b>用途：</b>{esc(e["use_for"])}</p>'
                     f'<p class="note">{esc(e["caution"])}</p>'
                     f'<details><summary>原生定位</summary><p>slide_id={esc(e["slide_id"])} · {esc(e["slide_part"])}</p></details></article>')
    start = '''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>区域结构素材库</title><style>body{margin:0;font:16px/1.65 "Microsoft YaHei",sans-serif;background:#f3f4f6;color:#202635}header{padding:26px 4vw;background:#202635;color:white}h1{margin:0;font-size:28px}input{font:inherit;padding:12px;width:min(650px,90%)}main{padding:24px 4vw;display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:24px}article{background:white;padding:22px;border:1px solid #d8dde5;border-radius:8px}h2{font-size:21px;margin:0}img{width:100%;height:auto;border:1px solid #e5e7eb}.tags{color:#48617d;font-size:14px}.note{background:#faf6ef;color:#77552b;padding:10px}footer{padding:20px 4vw}a{color:#376bb0}[hidden]{display:none!important}@media(max-width:500px){main{grid-template-columns:1fr}}</style><header><h1>区域结构素材库</h1><p>先组织有职责的小区域，再通过区域关系组成整页逻辑。素材可按内容选取，样式随目标稿统一。</p><input id="q" type="search" aria-label="筛选素材" placeholder="搜索：标题框 / 流程 / 双证据 / 图注…"><p id="n">'''
    script = r'''</main><footer><a href="source.pptx">原生素材</a> · <a href="index.md">文字索引</a> · <a href="catalog.json">语义目录</a> · <a href="inventory.json">原生对象</a><p>科研内容和引用需核实来源；本页不在后台自动更新。</p></footer><script>const q=document.getElementById('q'),a=[...document.querySelectorAll('article')];q.addEventListener('input',()=>{const terms=q.value.trim().toLowerCase().split(/\s+/).filter(Boolean);let n=0;a.forEach(x=>{x.hidden=!terms.every(t=>x.dataset.search.toLowerCase().includes(t));if(!x.hidden)n++});document.getElementById('n').textContent=`显示 ${n} / ${a.length} 页`});</script></html>'''
    page = start + f'显示 {len(entries)} / {len(entries)} 页</p></header><main>' + ''.join(cards) + script
    (root / 'index.md').write_bytes(('\n'.join(md).rstrip() + '\n').encode('utf-8'))
    (root / 'gallery.html').write_bytes(page.encode('utf-8'))
    return len(entries)


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('library')
    args = ap.parse_args()
    print(json.dumps({'pages': build(args.library), 'generated': ['index.md', 'gallery.html']}))