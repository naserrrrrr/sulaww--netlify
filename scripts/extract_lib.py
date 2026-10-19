"""Vector element extraction from the brand guideline PDF.

Groups a page's vector drawings into spatially connected clusters and
re-draws each cluster on its own transparent page, so every illustration
can be exported as a clean PNG (and SVG) without background removal.
"""
import pymupdf as fitz


def page_drawings(page, skip_bg=True):
    W, H = page.rect.width, page.rect.height
    out = []
    for i, d in enumerate(page.get_drawings()):
        r = d["rect"]
        if skip_bg and r.width > W * 0.9 and r.height > H * 0.9:
            continue  # page background
        if d["type"] == "s" and (r.height < 1 or r.width < 1) and r.width > W * 0.5:
            continue  # header/footer rule
        if r.y0 > H * 0.88 or r.y1 < H * 0.07:
            continue  # running header / footer (logo + page number)
        d["_i"] = i
        out.append(d)
    return out


def cluster(drawings, pad=6):
    rects = [fitz.Rect(d["rect"]) + (-pad, -pad, pad, pad) for d in drawings]
    parent = list(range(len(drawings)))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    for a in range(len(rects)):
        for b in range(a + 1, len(rects)):
            if rects[a].intersects(rects[b]):
                parent[find(a)] = find(b)
    groups = {}
    for i in range(len(drawings)):
        groups.setdefault(find(i), []).append(drawings[i])
    res = []
    for g in groups.values():
        g.sort(key=lambda d: d["_i"])
        bb = fitz.Rect(g[0]["rect"])
        for d in g[1:]:
            bb |= d["rect"]
        res.append((bb, g))
    res.sort(key=lambda t: (round(t[0].y0 / 40), t[0].x0))
    return res


def draw_into(page, drawings, offset):
    for d in drawings:
        shape = page.new_shape()
        for it in d["items"]:
            op = it[0]
            if op == "l":
                shape.draw_line(it[1] - offset, it[2] - offset)
            elif op == "re":
                shape.draw_rect(fitz.Rect(it[1]) - (offset.x, offset.y, offset.x, offset.y))
            elif op == "qu":
                q = it[1]
                shape.draw_quad(fitz.Quad(q.ul - offset, q.ur - offset, q.ll - offset, q.lr - offset))
            elif op == "c":
                shape.draw_bezier(it[1] - offset, it[2] - offset, it[3] - offset, it[4] - offset)
        shape.finish(
            fill=d.get("fill"),
            color=d.get("color"),
            dashes=d.get("dashes"),
            even_odd=d.get("even_odd", False),
            closePath=d.get("closePath", False),
            lineJoin=d.get("lineJoin") or 0,
            lineCap=max(d.get("lineCap") or (0,)) if isinstance(d.get("lineCap"), (tuple, list)) else (d.get("lineCap") or 0),
            width=d.get("width") or 1,
            stroke_opacity=d.get("stroke_opacity") if d.get("stroke_opacity") is not None else 1,
            fill_opacity=d.get("fill_opacity") if d.get("fill_opacity") is not None else 1,
        )
        shape.commit()


def isolate(drawings, bbox, margin=2):
    """Return a one-page PDF document holding only `drawings`, cropped to bbox."""
    bb = fitz.Rect(bbox) + (-margin, -margin, margin, margin)
    doc = fitz.open()
    page = doc.new_page(width=bb.width, height=bb.height)
    draw_into(page, drawings, fitz.Point(bb.x0, bb.y0))
    return doc


def render_png(doc, path, target_px=1024):
    page = doc[0]
    zoom = target_px / max(page.rect.width, page.rect.height)
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=True)
    pix.save(path)
    return pix.width, pix.height


def write_svg(doc, path):
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc[0].get_svg_image(text_as_path=True))
