# app/laudo_pptx.py
import io, os, re, math, json, shutil, subprocess, tempfile, random, string
from pathlib import Path
from typing import List, Dict, Any, Optional, Sequence, Union, Tuple

# gráficos
import matplotlib.pyplot as plt

# pptx / imagens
from pptx import Presentation
from pptx.dml.color import RGBColor, MSO_COLOR_TYPE
from pptx.util import Inches
from PIL import Image

import uuid
LAUDO_PREFIX = "laudo_"

REX = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*\}\}")
BASE_DIR = Path(__file__).resolve().parent.parent
TPL_DIR = BASE_DIR / "templates"

def _rid(n=6): 
    import string, random
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=n))

# ---------- helpers numéricos/labels ----------
PT_BR_SHORT = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def _periodo_to_label(periodo: str) -> str:
    s = str(periodo).strip()
    if "-" not in s: return s
    ano, tri = s.split("-", 1)
    try:
        q = int(tri)
    except ValueError:
        return s
    return f"{q}º Tri\n{ano}"

def _to_float(v) -> float:
    if v is None or str(v).strip() == "":
        return float("nan")
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return float("nan")

def _coerce_float(x) -> float:
    if x is None:
        return float("nan")
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        parts = s.split(".")
        if len(parts) > 2:
            s = "".join(parts[:-1]) + "." + parts[-1]
        elif len(parts) == 2 and len(parts[1]) <= 2:
            s = ".".join(parts)
        else:
            s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        return float("nan")

def _meses_labels_2linhas(inicio_ym: str, n: int = 12):
    ano, mes = inicio_ym.split("-")
    a = int(ano); m = int(mes)
    return [f"{PT_BR_SHORT[(m-1+i)%12]}\n{a + (m-1+i)//12}" for i in range(n)]

# ---------- gráficos (PNG transparente) ----------
def grafico_trimestral_png(
    data: List[Dict[str, Any]],
    label_key="periodo",
    anunciados_key="anuncios",
    vendidos_key="vendidos",
    ylim: Optional[Sequence[float]] = None,
    dpi=200,
    figsize=(9,4),
) -> bytes:
    labels_raw = [str(r.get(label_key, "")) for r in data]
    x_labels = [_periodo_to_label(s) for s in labels_raw]
    X = list(range(len(x_labels)))

    y_anun = [_to_float(r.get(anunciados_key)) for r in data]
    y_vend = [_to_float(r.get(vendidos_key)) for r in data]

    if ylim is None:
        vals = [v for v in (y_anun + y_vend) if not (isinstance(v, float) and math.isnan(v))]
        if not vals: raise ValueError("Sem valores numéricos.")
        ymax = max(80, math.ceil(max(vals)/20)*20)
        y_min, y_max = 0, ymax
    else:
        y_min, y_max = ylim

    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    ax.grid(True, which="major", axis="both", linestyle="-", linewidth=1, color="#D9D9D9", alpha=0.35)
    for s in ax.spines.values(): s.set_visible(False)
    ax.tick_params(axis="x", colors="#46484C", labelsize=9)
    ax.tick_params(axis="y", colors="#46484C", labelsize=9)
    ax.set_ylim(y_min, y_max)
    ax.set_yticks(list(range(int(y_min), int(y_max)+1, 20)))
    ax.set_xticks(X); ax.set_xticklabels(x_labels)
    ax.plot(X, y_vend, marker="o", markersize=6, linewidth=2.5, color="#7a2be2")
    ax.plot(X, y_anun, marker="o", markersize=6, linewidth=2.5, color="#f39c12")
    plt.tight_layout(pad=0)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0, transparent=True)
    plt.close(fig)
    return buf.getvalue()

def grafico_area_mensal_png(
    valores: List[Union[int,float,str]],
    inicio_ym="2023-08",
    moeda_prefix="R$ ",
    dpi=200,
    figsize=(12,3.2),
) -> bytes:
    if len(valores) != 12:
        raise ValueError("Forneça 12 valores (um por mês).")
    ys = [_coerce_float(v) for v in valores]
    xs = list(range(12))
    labels = _meses_labels_2linhas(inicio_ym, 12)
    valid = [v for v in ys if not math.isnan(v)]
    if not valid:
        raise ValueError("Sem valores numéricos válidos.")
    step = 5000.0
    ymax = max(step, math.ceil(max(valid)/step)*step)

    purple = "#720d83"
    grid_c = "#46484C"
    axis_c = "#e6e6e6"
    txt_c  = "#46484C"

    fig, ax = plt.subplots(figsize=figsize, dpi=dpi)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    ax.set_xlim(0, 11)
    ax.margins(x=0)
    ax.grid(True, which="major", axis="both", linestyle="-", linewidth=1, color=grid_c, alpha=0.25)
    for s in ax.spines.values(): s.set_color(axis_c); s.set_linewidth(1)
    ax.tick_params(axis="x", colors=txt_c, labelsize=10)
    ax.tick_params(axis="y", colors=txt_c, labelsize=10)
    ax.set_ylim(0, ymax)
    yticks = list(range(0, int(ymax)+1, int(step)))
    ax.set_yticks(yticks)
    ax.set_yticklabels([f"{moeda_prefix}{int(v):,}".replace(",", ".") for v in yticks])
    ax.set_xticks(xs); ax.set_xticklabels(labels)
    ax.fill_between(xs, ys, 0, color=purple, alpha=0.85, linewidth=0)
    ax.plot(xs, ys, color=purple, linewidth=2.5, marker="o", markersize=5)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", transparent=True)
    plt.close(fig)
    return buf.getvalue()

# ---------- pptx helpers ----------
def _snapshot_font(run_font):
    name = run_font.name
    size = run_font.size
    bold = run_font.bold
    italic = run_font.italic
    color_info = None
    cf = run_font.color
    if cf is not None and cf.type is not None:
        try:
            if cf.type == MSO_COLOR_TYPE.RGB and cf.rgb:
                color_info = ("rgb", cf.rgb, None)
            elif cf.type == MSO_COLOR_TYPE.SCHEME and cf.theme_color is not None:
                br = None
                try: br = cf.brightness
                except Exception: pass
                color_info = ("scheme", cf.theme_color, br)
        except Exception:
            pass
    return (name, size, bold, italic, color_info)

def _restore_font(run_font, snap):
    name, size, bold, italic, color_info = snap
    if name: run_font.name = name
    if size: run_font.size = size
    if bold is not None: run_font.bold = bold
    if italic is not None: run_font.italic = italic
    if color_info:
        kind, val, br = color_info
        try:
            if kind == "rgb" and val: run_font.color.rgb = RGBColor(val[0],val[1],val[2])
            elif kind == "scheme" and val is not None:
                run_font.color.theme_color = val
                if br is not None: run_font.color.brightness = br
        except Exception:
            pass

def _replace_text_in_textframe(text_frame, get_value_fn, img_keys_set):
    for p in text_frame.paragraphs:
        for r in p.runs:
            txt = r.text
            if "{{" not in txt: continue
            full = REX.fullmatch(txt.strip())
            if full and full.group(1) in img_keys_set:
                continue
            snap = _snapshot_font(r.font)
            def repl(m):
                k = m.group(1)
                v = get_value_fn(k)
                return v if v is not None else m.group(0)
            new_txt = REX.sub(repl, txt)
            if new_txt != txt:
                r.text = new_txt
                _restore_font(r.font, snap)

def _find_img_keys(shape, img_keys_set):
    keys=[]
    if hasattr(shape,"text_frame") and shape.text_frame:
        text = "".join(r.text for p in shape.text_frame.paragraphs for r in p.runs).strip()
        m = REX.fullmatch(text)
        if m:
            k = m.group(1)
            if k in img_keys_set: keys.append(k)
    return keys

def _add_image(slide, image_path, box, target_size=None):
    image_path = str(image_path)
    if not Path(image_path).exists():
        raise FileNotFoundError(f"Imagem não encontrada: {image_path}")
    l,t,w,h = box
    if target_size:
        tw, th = target_size
        left = int(l + (w - tw)/2)
        top  = int(t + (h - th)/2)
        slide.shapes.add_picture(image_path, left, top, width=int(tw), height=int(th))
        return
    # auto-fit contain
    img_w, img_h = Image.open(image_path).size
    scale = min(w/img_w, h/img_h)
    new_w = int(img_w*scale); new_h = int(img_h*scale)
    left = int(l + (w - new_w)/2); top = int(t + (h - new_h)/2)
    slide.shapes.add_picture(image_path, left, top, width=new_w, height=new_h)

def _walk(slide, shape, images_to_place, get_value_fn, img_keys_set):
    if getattr(shape,"has_table",False):
        for row in shape.table.rows:
            for cell in row.cells:
                if cell.text_frame:
                    _replace_text_in_textframe(cell.text_frame, get_value_fn, img_keys_set)
    if hasattr(shape,"text_frame") and shape.text_frame:
        img_keys = _find_img_keys(shape, img_keys_set)
        if img_keys:
            images_to_place.append((slide, shape, img_keys))
        else:
            _replace_text_in_textframe(shape.text_frame, get_value_fn, img_keys_set)
    if hasattr(shape,"shapes"):
        for s in shape.shapes:
            _walk(slide, s, images_to_place, get_value_fn, img_keys_set)

# ---------- conversão PDF ----------
def _candidate_fonts_dirs() -> list[Path]:
    """Tenta encontrar pastas de fontes dentro do repo (ex.: ./fonts/nunito/static)."""
    here = Path(__file__).resolve().parent
    repo_root = here.parent  # ajuste se precisar
    candidates = [
        repo_root / "fonts" / "nunito" / "static",
        repo_root / "fonts" / "nunito",
        repo_root / "fonts",
    ]
    return [p for p in candidates if p.exists() and any(p.glob("*.ttf"))]

def _write_fonts_conf(out_dir: Path, font_dirs: list[Path]) -> Path:
    """
    Gera um fonts.conf temporário apontando para os diretórios informados.
    Retorna o caminho do arquivo gerado.
    """
    fc_dir = out_dir / ".fontconfig"
    fc_dir.mkdir(parents=True, exist_ok=True)
    conf_path = fc_dir / "fonts.conf"
    # precisa ser caminho ABSOLUTO
    dirs_xml = "\n".join(f"  <dir>{str(d.resolve())}</dir>" for d in font_dirs)
    conf_xml = f"""<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
{dirs_xml}
</fontconfig>
"""
    conf_path.write_text(conf_xml, encoding="utf-8")
    return conf_path

def _libreoffice_pdf(pptx_path: Path, out_dir: Path):
    """
    Converte PPTX -> PDF usando LibreOffice (Impress) e força o uso de um fonts.conf
    que inclui as fontes do projeto (Nunito), se existirem.
    """
    import shutil, subprocess, os

    cmd = shutil.which("libreoffice") or shutil.which("soffice")
    if not cmd:
        raise RuntimeError("LibreOffice não encontrado no PATH.")

    # tenta achar diretórios de fontes no repo
    font_dirs = _candidate_fonts_dirs()
    env = os.environ.copy()
    env.setdefault("SAL_USE_VCLPLUGIN", "gen")  # render estável/headless

    # se achou fontes, cria um fonts.conf temporário e usa FONTCONFIG_FILE
    if font_dirs:
        conf_path = _write_fonts_conf(out_dir, font_dirs)
        env["FONTCONFIG_FILE"] = str(conf_path)
        # garante cache atualizado para esse config
        try:
            subprocess.run(["fc-cache", "-f", "-v"], check=True, env=env)
        except Exception:
            pass

    # use o filtro do IMPRESS para PPTX -> PDF
    subprocess.run(
        [
            cmd, "--headless",
            "--convert-to", "pdf:impress_pdf_Export",
            "--outdir", str(out_dir),
            str(pptx_path),
        ],
        check=True,
        env=env,
    )


# ---------- núcleo ----------
def render_laudo(
    template_path: Union[str, Path],
    text_vars: Dict[str,str],
    aliases: Optional[Dict[str,str]] = None,
    chart1: Optional[Dict[str,Any]] = None,
    chart2: Optional[Dict[str,Any]] = None,
    images_bindings: Optional[Dict[str,Union[str,Tuple[str,float,float]]]] = None,
    chart_slots: Optional[Dict[str,str]] = None,
    out_basename: Optional[str] = None
) -> Dict[str,str]:
    template_path = Path(template_path)
    if not template_path.exists():
        # fallback para templates/laudo-imogo.pptx
        candidate = (TPL_DIR / "laudo-imogo.pptx")
        if candidate.exists():
            template_path = candidate
        else:
            raise FileNotFoundError(f"Template não encontrado: {template_path}")

    rid = out_basename or str(uuid.uuid4())
    # pasta com prefixo
    work = Path(tempfile.gettempdir()) / f"{LAUDO_PREFIX}{rid}"
    work.mkdir(exist_ok=True)


    # 1) gera gráficos
    gen_imgs: Dict[str, Path] = {}
    if chart1:
        png = grafico_trimestral_png(
            data = chart1["data"],
            label_key = chart1.get("label_key","periodo"),
            anunciados_key = chart1.get("anunciados_key","anuncios"),
            vendidos_key   = chart1.get("vendidos_key","vendidos"),
            ylim = chart1.get("ylim")
        )
        p = work / "grafico_trimestres_transparente.png"
        p.write_bytes(png); gen_imgs["chart1"] = p

    if chart2:
        png = grafico_area_mensal_png(
            valores = chart2["valores"],
            inicio_ym = chart2.get("inicio_ym","2023-08"),
            moeda_prefix = chart2.get("moeda_prefix","R$ ")
        )
        p = work / "grafico2_transparente.png"
        p.write_bytes(png); gen_imgs["chart2"] = p

    # 2) vars e imagens
    VARS = dict(text_vars or {})
    ALIASES = dict(aliases or {})
    IMG_VARS: Dict[str,Union[str,Tuple[Path,float,float]]] = {}
    images_bindings = images_bindings or {}
    for k,v in images_bindings.items():
        if isinstance(v,(list,tuple)) and len(v)==3:
            IMG_VARS[k] = (Path(v[0]), float(Inches(v[1])), float(Inches(v[2])))
        elif isinstance(v,str):
            IMG_VARS[k] = v
        else:
            raise ValueError(f"Imagem inválida para '{k}'")

    slots = chart_slots or {"chart1":"grafico_01","chart2":"grafico_02"}
    if "chart1" in gen_imgs:
        varname = slots.get("chart1","grafico_01")
        IMG_VARS[varname] = (gen_imgs["chart1"], float(Inches(4)), float(Inches(1.8)))
    if "chart2" in gen_imgs:
        varname = slots.get("chart2","grafico_02")
        IMG_VARS[varname] = (gen_imgs["chart2"], float(Inches(6.5)), float(Inches(1.5)))

    def _get_value(k: str):
        if k in VARS: return VARS[k]
        if k in ALIASES and ALIASES[k] in VARS: return VARS[ALIASES[k]]
        return None

    img_keys = set(IMG_VARS.keys())

    # 3) carrega/varre pptx
    prs = Presentation(str(template_path))
    to_place = []
    for slide in prs.slides:
        for shape in slide.shapes:
            _walk(slide, shape, to_place, _get_value, img_keys)

    # 4) insere imagens nas caixas {{chave}}
    for slide, shape, keys in to_place:
        key = keys[0]
        info = IMG_VARS.get(key)
        if hasattr(shape,"text_frame") and shape.text_frame:
            for p in shape.text_frame.paragraphs:
                for r in p.runs: r.text = ""
        if info is None: 
            continue
        if isinstance(info, tuple):
            img_path, w_emus, h_emus = info
            target = (int(w_emus), int(h_emus))
        else:
            img_path = info
            target = None
        box = (shape.left, shape.top, shape.width, shape.height)
        _add_image(slide, img_path, box, target)

    out_pptx = work / f"{rid}.pptx"
    prs.save(out_pptx)

    # 5) pdf (libreoffice)
    pdf_path = work / f"{rid}.pdf"
    try:
        _libreoffice_pdf(out_pptx, work)
    except Exception as e:
        # sem PDF não quebra – retorna pptx ok
        pass

    return {
        "id": rid,
        "dir": str(work),
        "pptx_path": str(out_pptx),
        "pdf_path": str(pdf_path) if pdf_path.exists() else ""
    }
