# utils/pptx_pdf.py
# Conversão robusta PPTX -> PDF usando LibreOffice (headless)
# - Saneia DISPLAY / VCL
# - Usa profile isolado em <out_dir>/.lo_profile
# - Tenta filtro 'impress_pdf_Export', depois 'pdf' genérico
# - Se for preciso, re-tenta via xvfb-run -a (se instalado)
# - Fallback: exporta PNGs dos slides e junta num PDF raster
# - Expõe função convert_pptx_to_pdf(...) e CLI
# python3 utils/pptx_pdf.py /tmp/laudo_22210a1e-9058-41ef-9068-1aa2ca736e90/laudo_22210a1e-9058-41ef-9068-1aa2ca736e90.pptx

import os
import sys
import shutil
import subprocess
from pathlib import Path
from typing import Optional, List
from PIL import Image  # pip install Pillow


# ------------------------ utilidades de ambiente ------------------------

def _candidate_fonts_dirs() -> List[Path]:
    """Procura fontes do projeto p/ usar via FONTCONFIG_FILE."""
    here = Path(__file__).resolve().parent
    repo_root = here.parent
    candidates = [
        repo_root / "fonts" / "nunito" / "static",
        repo_root / "fonts" / "nunito",
        repo_root / "fonts",
    ]
    return [p for p in candidates if p.exists() and any(p.glob("*.ttf"))]


def _write_fonts_conf(out_dir: Path, font_dirs: List[Path]) -> Path:
    """Gera fonts.conf temporário apontando para as pastas de fontes."""
    fc_dir = out_dir / ".fontconfig"
    fc_dir.mkdir(parents=True, exist_ok=True)
    conf_path = fc_dir / "fonts.conf"
    dirs_xml = "\n".join(f"  <dir>{str(d.resolve())}</dir>" for d in font_dirs)
    conf_xml = f"""<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
{dirs_xml}
</fontconfig>
"""
    conf_path.write_text(conf_xml, encoding="utf-8")
    return conf_path


def _sanitized_env(out_dir: Path) -> dict:
    """Ambiente estável/headless p/ LibreOffice."""
    env = os.environ.copy()
    env.setdefault("SAL_USE_VCLPLUGIN", "gen")
    env.setdefault("HOME", "/tmp")
    env.setdefault("XDG_CACHE_HOME", "/tmp/.cache")
    # mata tentativa de X11
    env.pop("DISPLAY", None)
    # se já geramos fonts.conf, aproveita
    fc_conf = out_dir / ".fontconfig" / "fonts.conf"
    if fc_conf.exists():
        env["FONTCONFIG_FILE"] = str(fc_conf.resolve())
    return env


def _lo_bin() -> str:
    cmd = shutil.which("libreoffice") or shutil.which("soffice") or shutil.which("loffice")
    if not cmd:
        raise RuntimeError("LibreOffice (libreoffice/soffice/loffice) não encontrado no PATH.")
    return cmd


def _run(args: list, out_dir: Path, log_prefix: str):
    """Roda subprocesso, grava stdout/stderr em logs."""
    proc = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    (out_dir / f"{log_prefix}_stdout.log").write_text(proc.stdout or "", encoding="utf-8")
    (out_dir / f"{log_prefix}_stderr.log").write_text(proc.stderr or "", encoding="utf-8")
    return proc


def _run_lo(args: list, out_dir: Path, log_prefix: str):
    """Executa LO com ambiente saneado."""
    env = _sanitized_env(out_dir)
    proc = subprocess.run(args, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    (out_dir / f"{log_prefix}_stdout.log").write_text(proc.stdout or "", encoding="utf-8")
    (out_dir / f"{log_prefix}_stderr.log").write_text(proc.stderr or "", encoding="utf-8")
    return proc


def _run_lo_maybe_xvfb(args: list, out_dir: Path, log_prefix: str):
    """
    Roda LO com ambiente saneado; se detectar erro de X11 no stderr,
    re-tenta automaticamente com xvfb-run -a, mesmo quando RC==0.
    """
    proc = _run_lo(args, out_dir, log_prefix)
    err = (proc.stderr or "")

    # Detecta mensagens típicas de X11 em stderr
    saw_x11 = ("Can't open display" in err) or ("X11 error" in err) or (" DISPLAY" in err)

    xvfb = shutil.which("xvfb-run")
    if saw_x11 and xvfb:
        # re-tenta sob X virtual
        proc = _run_lo([xvfb, "-a"] + args, out_dir, log_prefix + "_xvfb")

    return proc


# ------------------------ passos de conversão ------------------------

def _ensure_fonts(out_dir: Path):
    """Se houver fontes no projeto, ativa FONTCONFIG_FILE e força fc-cache."""
    font_dirs = _candidate_fonts_dirs()
    if not font_dirs:
        return
    conf_path = _write_fonts_conf(out_dir, font_dirs)
    env = _sanitized_env(out_dir)
    env["FONTCONFIG_FILE"] = str(conf_path)
    try:
        subprocess.run(["fc-cache", "-f", "-v"], check=True, env=env,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except Exception:
        pass  # não falha a conversão só por cache de fonte


def _export_pdf_via_impress(pptx_path: Path, out_dir: Path) -> bool:
    """Tenta converter com filtro do Impress; se precisar, usa xvfb-run."""
    lo = _lo_bin()
    # profile isolado evita HOME "estranho"
    lo_profile = out_dir / ".lo_profile"
    lo_profile.mkdir(parents=True, exist_ok=True)
    lo_profile_uri = f"file://{lo_profile.resolve()}"

    args = [
        lo,
        "--headless", "--invisible", "--nologo", "--nolockcheck",
        "--norestore", "--nodefault", "--nofirststartwizard",
        f"-env:UserInstallation={lo_profile_uri}",
        "--convert-to", "pdf:impress_pdf_Export",
        "--outdir", str(out_dir),
        str(pptx_path),
    ]
    proc = _run_lo_maybe_xvfb(args, out_dir, "lo")
    pdf_out = out_dir / (pptx_path.with_suffix(".pdf").name)
    return proc.returncode == 0 and pdf_out.exists()


def _export_pdf_generic(pptx_path: Path, out_dir: Path) -> bool:
    """Segunda tentativa: --convert-to pdf genérico."""
    lo = _lo_bin()
    lo_profile = out_dir / ".lo_profile"
    lo_profile.mkdir(parents=True, exist_ok=True)
    lo_profile_uri = f"file://{lo_profile.resolve()}"

    args = [
        lo,
        "--headless", "--invisible", "--nologo", "--nolockcheck",
        "--norestore", "--nodefault", "--nofirststartwizard",
        f"-env:UserInstallation={lo_profile_uri}",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(pptx_path),
    ]
    proc = _run_lo_maybe_xvfb(args, out_dir, "lo_no_filter")
    pdf_out = out_dir / (pptx_path.with_suffix(".pdf").name)
    return proc.returncode == 0 and pdf_out.exists()


def _export_pngs_then_pdf(pptx_path: Path, out_dir: Path) -> bool:
    """Fallback: exporta slides como PNG e junta em PDF raster."""
    lo = _lo_bin()
    lo_profile = out_dir / ".lo_profile"
    lo_profile.mkdir(parents=True, exist_ok=True)
    lo_profile_uri = f"file://{lo_profile.resolve()}"

    # exporta PNGs
    args = [
        lo,
        "--headless", "--invisible", "--nologo", "--nolockcheck",
        "--norestore", "--nodefault", "--nofirststartwizard",
        f"-env:UserInstallation={lo_profile_uri}",
        "--convert-to", "png:impress_png_Export",
        "--outdir", str(out_dir),
        str(pptx_path),
    ]
    proc = _run_lo_maybe_xvfb(args, out_dir, "lo_png")
    if proc.returncode != 0:
        return False

    stem = pptx_path.stem
    pngs = sorted(out_dir.glob(f"{stem}*.png"), key=lambda p: p.name.lower())
    if not pngs:
        return False

    # junta PNGs num PDF
    images = [Image.open(str(p)).convert("RGB") for p in pngs]
    out_pdf = out_dir / f"{stem}.pdf"
    images[0].save(str(out_pdf), save_all=True, append_images=images[1:])
    return out_pdf.exists()


# ------------------------ API pública ------------------------

def convert_pptx_to_pdf(
    pptx_path: Path,
    out_dir: Path,
    try_xvfb: bool = True,           # mantido por compat, internamente já tentamos
    use_impress_filter: bool = True, # idem
    fallback_png_pdf: bool = True,   # faz raster se tudo falhar
    ensure_fonts: bool = True,       # ativa FONTCONFIG_FILE a partir de ./fonts
) -> bool:
    """
    Converte PPTX -> PDF de forma robusta.
    Retorna True se <out_dir>/<stem>.pdf existir no final.
    """
    print("convert_pptx_to_pdf: run")
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if ensure_fonts:
        _ensure_fonts(out_dir)

    ok = False
    if use_impress_filter:
        ok = _export_pdf_via_impress(pptx_path, out_dir)
    if not ok:
        ok = _export_pdf_generic(pptx_path, out_dir)
    if not ok and fallback_png_pdf:
        ok = _export_pngs_then_pdf(pptx_path, out_dir)
    return ok


# ------------------------ CLI ------------------------

def _cli():
    if len(sys.argv) != 2:
        print("uso: python utils/pptx_pdf.py /caminho/arquivo.pptx")
        sys.exit(2)
    pptx = Path(sys.argv[1]).resolve()
    if not pptx.exists():
        print(f"arquivo não encontrado: {pptx}")
        sys.exit(1)
    out_dir = pptx.parent
    print("RUN:", _lo_bin(),
          "--headless --invisible --nologo --nolockcheck --norestore --nodefault --nofirststartwizard",
          "--convert-to pdf:impress_pdf_Export",
          "--outdir", out_dir, pptx)
    ok = convert_pptx_to_pdf(pptx, out_dir=out_dir)
    print("OK?:", ok)
    pdf = out_dir / (pptx.stem + ".pdf")
    print("PDF:", pdf, "exists?", pdf.exists())
    # Mostra erro se existir
    for name in [
        "lo_stderr.log", "lo_stderr_xvfb.log",
        "lo_no_filter_stderr.log", "lo_no_filter_stderr_xvfb.log",
        "lo_png_stderr.log", "lo_png_stderr_xvfb.log",
    ]:
        p = out_dir / name
        if p.exists() and p.stat().st_size:
            print(f"\n--- {name} ---")
            try:
                print(p.read_text(encoding="utf-8")[:2000])
            except Exception:
                print("(não-utf8/binary)")

if __name__ == "__main__":
    _cli()
