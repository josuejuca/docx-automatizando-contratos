# utils/office_pdf.py
from __future__ import annotations
import os, shutil, subprocess, tempfile, signal
from pathlib import Path
from typing import Optional, List
from PIL import Image

def _candidate_fonts_dirs() -> list[Path]:
    here = Path(__file__).resolve().parent
    root = here.parent
    return [p for p in [
        root / "fonts" / "nunito" / "static",
        root / "fonts" / "nunito",
        root / "fonts",
    ] if p.exists() and any(p.glob("*.ttf"))]

def _write_fonts_conf(out_dir: Path, font_dirs: list[Path]) -> Path:
    fc_dir = out_dir / ".fontconfig"
    fc_dir.mkdir(parents=True, exist_ok=True)
    conf_path = fc_dir / "fonts.conf"
    dirs_xml = "\n".join(f"  <dir>{str(d.resolve())}</dir>" for d in font_dirs)
    conf_path.write_text(f"""<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
{dirs_xml}
</fontconfig>""", encoding="utf-8")
    return conf_path

def _run_lo(args: list[str], out_dir: Path, env_extra: dict, timeout: int, stdout_name: str, stderr_name: str) -> subprocess.CompletedProcess:
    env = os.environ.copy()
    env.setdefault("SAL_USE_VCLPLUGIN", "gen")
    env.setdefault("HOME", "/tmp")
    env.setdefault("XDG_CACHE_HOME", "/tmp/.cache")
    env.update(env_extra or {})
    out_dir.mkdir(parents=True, exist_ok=True)
    with open(out_dir / stdout_name, "w", encoding="utf-8") as f_out, open(out_dir / stderr_name, "w", encoding="utf-8") as f_err:
        # cria um process group para matar tudo em timeout
        proc = subprocess.Popen(args, env=env, stdout=f_out, stderr=f_err, preexec_fn=os.setsid)
        try:
            rc = proc.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            os.killpg(proc.pid, signal.SIGKILL)
            raise RuntimeError(f"LibreOffice timeout ({timeout}s) ao executar: {' '.join(args)}")
        return subprocess.CompletedProcess(args=args, returncode=rc)

def _base_cmd_and_env(out_dir: Path) -> tuple[list[str], dict]:
    cmd = shutil.which("libreoffice") or shutil.which("soffice") or shutil.which("loffice")
    if not cmd:
        raise RuntimeError("LibreOffice (libreoffice/soffice/loffice) não encontrado no PATH.")
    profile = out_dir / ".lo_profile"
    profile.mkdir(exist_ok=True, parents=True)
    profile_uri = f"file://{profile.resolve()}"
    env = {"SAL_USE_VCLPLUGIN": "gen"}  # reforço
    fonts = _candidate_fonts_dirs()
    if fonts:
        conf = _write_fonts_conf(out_dir, fonts)
        env["FONTCONFIG_FILE"] = str(conf)
        try:
            _run_lo(["fc-cache","-f","-v"], out_dir, env, timeout=30, stdout_name="fc_stdout.log", stderr_name="fc_stderr.log")
        except Exception:
            pass
    base = [cmd, "--headless", "--invisible", "--nologo", "--nolockcheck", "--norestore", "--nodefault", f"-env:UserInstallation={profile_uri}"]
    return base, env

def convert_docx_to_pdf(docx_path: Path, out_dir: Path, timeout: int = 60) -> bool:
    docx_path = Path(docx_path)
    out_dir = Path(out_dir)
    base, env = _base_cmd_and_env(out_dir)
    args = base + ["--convert-to", "pdf:writer_pdf_Export", "--outdir", str(out_dir), str(docx_path)]
    proc = _run_lo(args, out_dir, env, timeout, "lo_docx_stdout.log", "lo_docx_stderr.log")
    pdf = out_dir / (docx_path.with_suffix(".pdf").name)
    return proc.returncode == 0 and pdf.exists()

def convert_pptx_to_pdf(pptx_path: Path, out_dir: Path, timeout: int = 90) -> bool:
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    base, env = _base_cmd_and_env(out_dir)
    # 1) tenta PDF direto (vetorial)
    args = base + ["--convert-to", "pdf:impress_pdf_Export", "--outdir", str(out_dir), str(pptx_path)]
    proc = _run_lo(args, out_dir, env, timeout, "lo_pptx_stdout.log", "lo_pptx_stderr.log")
    pdf = out_dir / (pptx_path.with_suffix(".pdf").name)
    if proc.returncode == 0 and pdf.exists():
        return True
    # 2) fallback PNG->PDF (pixel-perfect)
    args_png = base + ["--convert-to", "png:impress_png_Export", "--outdir", str(out_dir), str(pptx_path)]
    _run_lo(args_png, out_dir, env, timeout, "lo_png_stdout.log", "lo_png_stderr.log")
    pngs = sorted(out_dir.glob(f"{pptx_path.stem}*.png"), key=lambda p: p.name)
    if not pngs:
        return False
    images = [Image.open(str(p)).convert("RGB") for p in pngs]
    images[0].save(str(pdf), save_all=True, append_images=images[1:])
    return pdf.exists()
