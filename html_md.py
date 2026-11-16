"""HTML (editado) -> Markdown. Extrae imágenes como archivos asociados."""
from __future__ import annotations
import base64, re
from .html_docx import parse_html, Node

def _inline_md(node: Node) -> str:
    out=[]
    for ch in node.children:
        if ch.tag is None:
            out.append(ch.text or ""); continue
        t=ch.tag
        if t in ("strong","b"): out.append(f"**{_inline_md(ch).strip()}**")
        elif t in ("em","i"): out.append(f"*{_inline_md(ch).strip()}*")
        elif t in ("s","del","strike"): out.append(f"~~{_inline_md(ch).strip()}~~")
        elif t=="code": out.append(f"`{_inline_md(ch).strip()}`")
        elif t=="sup": out.append(f"<sup>{_inline_md(ch).strip()}</sup>")
        elif t=="sub": out.append(f"<sub>{_inline_md(ch).strip()}</sub>")
        elif t=="a":
            href=ch.attrs.get("href",""); txt=_inline_md(ch).strip()
            out.append(f"[{txt}]({href})" if href and txt else txt)
        elif t=="br": out.append("  \n")
        elif t=="img": out.append(f"![{_inline_md(ch) or 'imagen'}]({ch.attrs.get('src','')})")
        elif t=="u": out.append(f"<u>{_inline_md(ch)}</u>")
        else: out.append(_inline_md(ch))
    return "".join(out)

def _cell_text(node: Node) -> str:
    txt=_inline_md(node).strip().replace("|","\\|").replace("\n"," ")
    return re.sub(r"\s+"," ",txt)

def html_to_markdown(html: str, meta: dict | None=None, title: str="", image_prefix: str="images"):
    """Devuelve (markdown, {nombre_archivo: bytes})"""
    root=parse_html(html)
    lines=[]; images={}
    counter=[0]
    def img_ref(src: str, alt: str="") -> str:
        m=re.match(r"data:image/(\w+);base64,(.*)$", src or "", re.S)
        if not m: return f"![{alt}]({src})"
        ext=m.group(1).lower().replace("jpeg","jpg")
        counter[0]+=1
        name=f"{image_prefix}/imagen-{counter[0]}.{ext}"
        try: images[name]=base64.b64decode(m.group(2))
        except Exception: return ""
        return f"![{alt}]({name})"
    def walk_block(node: Node):
        t=node.tag
        if t is None: return
        if t=="section" and "pdfpage" in node.attrs.get("class",""):
            pn=node.attrs.get("data-page","")
            if lines and pn: lines.extend([f"<!-- Página {pn} -->",""])
            for c in node.children: walk_block(c)
            return
        if re.fullmatch(r"h[1-6]", t or ""):
            lvl=int(t[1]); txt=_inline_md(node).strip().replace("**","")
            if txt: lines.extend([f"{'#'*lvl} {txt}",""]); return
        if t=="p":
            txt=_inline_md(node).strip()
            if txt: lines.extend([txt,""]); return
        if t in ("ul","ol"):
            num=1
            for li in [c for c in node.children if c.tag=="li"]:
                txt=_inline_md(li).strip()
                lines.append(f"{num}. {txt}" if t=="ol" else f"- {txt}")
                num+=1
            lines.append(""); return
        if t=="table":
            rows=[]
            for tr in node.find_all("tr"):
                cells=[_cell_text(c) for c in tr.children if c.tag in ("td","th")]
                if cells: rows.append(cells)
            if rows:
                nc=max(len(r) for r in rows); rows=[r+[""]*(nc-len(r)) for r in rows]
                lines.append("| "+" | ".join(rows[0])+" |")
                lines.append("| "+" | ".join(["---"]*nc)+" |")
                for r in rows[1:]: lines.append("| "+" | ".join(r)+" |")
                lines.append(""); return
        if t=="figure":
            for im in node.find_all("img"): lines.extend([img_ref(im.attrs.get("src",""), im.attrs.get("alt","")),""]); return
        if t=="img":
            lines.extend([img_ref(node.attrs.get("src",""), node.attrs.get("alt","")),""]); return
        if t=="hr":
            if "pagebreak" in node.attrs.get("class",""): lines.extend(["---",""])
            return
        if t=="blockquote":
            for c in node.children: walk_block(c)
            return
        for c in node.children: walk_block(c)
    for c in root.children: walk_block(c)
    md="\n".join(lines).strip()
    md=re.sub(r"\n{3,}","\n\n",md)
    if title:
        clean_title=title.replace("#","").strip()
        if clean_title: md=f"# {clean_title}\n\n"+md if not md.startswith("# ") else md
    return md+"\n", images
