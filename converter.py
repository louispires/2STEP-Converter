import sys
import os
import json
import ctypes
import traceback
import argparse
import subprocess
import zipfile
import struct
import tempfile
import time
import threading
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from xml.etree import ElementTree as ET

_DATA_DIR = Path(__file__).parent / "data"
_DATA_DIR.mkdir(exist_ok=True)

_CONFIG_DEFAULTS = {
    "DEFAULT_TOLERANCE":    0.01,
    "DEFAULT_SIMPLIFY":     0,
    "SIMPLIFY_INTERACTIVE": False,
    "SKIP_EXISTING":        True,
    "ANGULAR_TOLERANCE":    1e-3,
    "NAME_TRIM_WIDTH":      62,
    "SEPARATOR_WIDTH":      32,
    "BYTES_PER_KB":         1024,
    "STD_OUTPUT_HANDLE":    -11,
    "CONSOLE_MODE_FLAGS":   7,
    "DEFAULT_FORMAT":       "ap203",
    "MODELS_DIR_NAME":      "models",
    "STL_EXT":              ".stl",
    "TMF_EXT":              ".3mf",
    "OBJ_EXT":              ".obj",
    "IGS_EXT":              ".igs",
    "AMF_EXT":              ".amf",
    "STP_EXT":              ".stp",
}

_CONFIG_PATH = _DATA_DIR / "config.json"
if not _CONFIG_PATH.exists():
    _CONFIG_PATH.write_text(json.dumps(_CONFIG_DEFAULTS, indent=4), encoding="utf-8")

_cfg = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
STD_OUTPUT_HANDLE    = _cfg.get("STD_OUTPUT_HANDLE",    _CONFIG_DEFAULTS["STD_OUTPUT_HANDLE"])
CONSOLE_MODE_FLAGS   = _cfg.get("CONSOLE_MODE_FLAGS",   _CONFIG_DEFAULTS["CONSOLE_MODE_FLAGS"])
DEFAULT_TOLERANCE    = _cfg.get("DEFAULT_TOLERANCE",    _CONFIG_DEFAULTS["DEFAULT_TOLERANCE"])
ANGULAR_TOLERANCE    = _cfg.get("ANGULAR_TOLERANCE",    _CONFIG_DEFAULTS["ANGULAR_TOLERANCE"])
NAME_TRIM_WIDTH      = _cfg.get("NAME_TRIM_WIDTH",      _CONFIG_DEFAULTS["NAME_TRIM_WIDTH"])
SEPARATOR_WIDTH      = _cfg.get("SEPARATOR_WIDTH",      _CONFIG_DEFAULTS["SEPARATOR_WIDTH"])
MODELS_DIR_NAME      = _cfg.get("MODELS_DIR_NAME",      _CONFIG_DEFAULTS["MODELS_DIR_NAME"])
STL_EXT              = _cfg.get("STL_EXT",              _CONFIG_DEFAULTS["STL_EXT"])
TMF_EXT              = _cfg.get("TMF_EXT",              _CONFIG_DEFAULTS["TMF_EXT"])
OBJ_EXT              = _cfg.get("OBJ_EXT",              _CONFIG_DEFAULTS["OBJ_EXT"])
IGS_EXT              = _cfg.get("IGS_EXT",              _CONFIG_DEFAULTS["IGS_EXT"])
AMF_EXT              = _cfg.get("AMF_EXT",              _CONFIG_DEFAULTS["AMF_EXT"])
STP_EXT              = _cfg.get("STP_EXT",              _CONFIG_DEFAULTS["STP_EXT"])
BYTES_PER_KB         = _cfg.get("BYTES_PER_KB",         _CONFIG_DEFAULTS["BYTES_PER_KB"])
DEFAULT_SIMPLIFY     = _cfg.get("DEFAULT_SIMPLIFY",     _CONFIG_DEFAULTS["DEFAULT_SIMPLIFY"])
SIMPLIFY_INTERACTIVE = _cfg.get("SIMPLIFY_INTERACTIVE", _CONFIG_DEFAULTS["SIMPLIFY_INTERACTIVE"])
SKIP_EXISTING        = _cfg.get("SKIP_EXISTING",        _CONFIG_DEFAULTS["SKIP_EXISTING"])
DEFAULT_FORMAT       = _cfg.get("DEFAULT_FORMAT",       _CONFIG_DEFAULTS["DEFAULT_FORMAT"])

_cfg_warnings: list[str] = []
for _k, _lo, _hi in [
    ("DEFAULT_TOLERANCE",  1e-9, None),
    ("ANGULAR_TOLERANCE",  1e-9, None),
    ("DEFAULT_SIMPLIFY",   0,    99),
]:
    _v = _cfg.get(_k)
    if _v is not None and not isinstance(_v, (int, float)):
        _cfg_warnings.append(f"config: {_k} must be a number (got {type(_v).__name__!r})")
    elif _v is not None:
        if _lo is not None and _v < _lo:
            _cfg_warnings.append(f"config: {_k} must be ≥ {_lo} (got {_v})")
        if _hi is not None and _v > _hi:
            _cfg_warnings.append(f"config: {_k} must be ≤ {_hi} (got {_v})")
for _bool_k in ("SIMPLIFY_INTERACTIVE", "SKIP_EXISTING"):
    _b = _cfg.get(_bool_k)
    if _b is not None and not isinstance(_b, bool):
        _cfg_warnings.append(f"config: {_bool_k} must be true or false (got {_b!r})")
_df = _cfg.get("DEFAULT_FORMAT")
if _df is not None and _df not in ("ap203", "ap214", "ap242"):
    _cfg_warnings.append(f"config: DEFAULT_FORMAT must be ap203, ap214, or ap242 (got {_df!r})")
    DEFAULT_FORMAT = "ap203"
for _ext_k in ("STL_EXT", "TMF_EXT", "OBJ_EXT", "IGS_EXT", "AMF_EXT", "STP_EXT"):
    _ev = _cfg.get(_ext_k)
    if _ev is not None and (not isinstance(_ev, str) or not _ev.startswith(".")):
        _cfg_warnings.append(f"config: {_ext_k} must start with '.' (got {_ev!r})")

try:
    ctypes.windll.kernel32.SetConsoleMode(
        ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE), CONSOLE_MODE_FLAGS)
except Exception:
    pass

G   = '\033[92m'
R   = '\033[91m'
Y   = '\033[93m'
C   = '\033[96m'
DIM = '\033[2m'
B   = '\033[1m'
X   = '\033[0m'

_real_stdout_fd = os.dup(1)


@contextmanager
def quiet():
    sys.stdout.flush()
    sys.stderr.flush()
    fd1, fd2 = os.dup(1), os.dup(2)
    nul = os.open(os.devnull, os.O_WRONLY)
    try:
        os.dup2(nul, 1)
        os.dup2(nul, 2)
    finally:
        os.close(nul)
    try:
        yield
    finally:
        sys.stdout.flush()
        sys.stderr.flush()
        os.dup2(fd1, 1)
        os.close(fd1)
        os.dup2(fd2, 2)
        os.close(fd2)


try:
    with quiet():
        from OCC.Core.StlAPI import StlAPI_Reader
        from OCC.Core.BRep import BRep_Builder
        from OCC.Core.BRepTools import breptools
        from OCC.Core.BRepBuilderAPI import BRepBuilderAPI_Sewing
        from OCC.Core.ShapeUpgrade import ShapeUpgrade_UnifySameDomain
        from OCC.Core.ShapeFix import ShapeFix_Shape
        from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_AsIs
        from OCC.Core.Interface import Interface_Static
        from OCC.Core.IGESControl import IGESControl_Reader
        from OCC.Core.IFSelect import IFSelect_RetDone, IFSelect_RetError, IFSelect_RetFail
        from OCC.Core.TopoDS import TopoDS_Shape, TopoDS_Compound
        from OCC.Core.TopExp import TopExp_Explorer
        from OCC.Core.TopAbs import TopAbs_FACE, TopAbs_EDGE, TopAbs_SHELL
except Exception as e:
    print(f"\n  {R}[ERROR]{X} Failed to load OpenCASCADE: {e}\n")
    traceback.print_exc()
    input("\n  Press Enter to exit...")
    sys.exit(1)


_3MF_NS = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02"
_STL_TRI = struct.Struct("<12fH")
_SUPPORTED_EXTS = {STL_EXT, TMF_EXT, OBJ_EXT, AMF_EXT, IGS_EXT, ".iges"}

_BOX_CONTENT = 72
_BOX_LABEL   = 12
_BOX_TIME    = 6
_BOX_DETAIL  = _BOX_CONTENT - _BOX_LABEL - 2 - _BOX_TIME


def _mesh_to_shape(verts: list, tris: list):
    if not tris:
        raise ValueError("no triangle data found")
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".stl")
    try:
        with os.fdopen(tmp_fd, "wb") as f:
            buf = bytearray(80 + 4 + len(tris) * _STL_TRI.size)
            struct.pack_into("<I", buf, 80, len(tris))
            for i, t in enumerate(tris):
                v0, v1, v2 = verts[t[0]], verts[t[1]], verts[t[2]]
                _STL_TRI.pack_into(buf, 84 + i * _STL_TRI.size, 0.0, 0.0, 0.0, *v0, *v1, *v2, 0)
            f.write(buf)
        shape = TopoDS_Shape()
        with quiet():
            StlAPI_Reader().Read(shape, tmp_path.replace("\\", "/"))
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
    return shape


def _find_3mf_model(zf: zipfile.ZipFile) -> str:
    names = zf.namelist()
    rels_path = "_rels/.rels"
    if rels_path in names:
        with zf.open(rels_path) as f:
            root = ET.parse(f).getroot()
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            if "3dmanufacturing" in rel.get("Type", ""):
                target = rel.get("Target", "").lstrip("/")
                if target:
                    return target
    for name in names:
        if name.endswith(".model"):
            return name
    raise ValueError("could not find 3D model document in 3MF archive")


def _read_3mf_shape(path: Path):
    verts_all = []
    tris_all  = []

    with zipfile.ZipFile(str(path), "r") as zf:
        model_file = _find_3mf_model(zf)
        with zf.open(model_file) as f:
            root = ET.parse(f).getroot()

    resources = root.find(f"{{{_3MF_NS}}}resources")
    if resources is None:
        raise ValueError("no <resources> element found in 3MF model")

    offset = 0
    for obj in resources.findall(f"{{{_3MF_NS}}}object"):
        mesh_el = obj.find(f"{{{_3MF_NS}}}mesh")
        if mesh_el is None:
            continue
        verts_el = mesh_el.find(f"{{{_3MF_NS}}}vertices")
        tris_el  = mesh_el.find(f"{{{_3MF_NS}}}triangles")
        if verts_el is None or tris_el is None:
            continue
        verts = [
            (float(v.get("x")), float(v.get("y")), float(v.get("z")))
            for v in verts_el.findall(f"{{{_3MF_NS}}}vertex")
        ]
        tris_all.extend(
            (int(t.get("v1")) + offset, int(t.get("v2")) + offset, int(t.get("v3")) + offset)
            for t in tris_el.findall(f"{{{_3MF_NS}}}triangle")
        )
        verts_all.extend(verts)
        offset += len(verts)

    if not tris_all:
        raise ValueError("no triangle data found in 3MF file")
    return _mesh_to_shape(verts_all, tris_all), len(verts_all), len(tris_all)


def _read_obj_shape(path: Path):
    verts = []
    tris  = []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            parts = line.split()
            if not parts:
                continue
            if parts[0] == "v":
                verts.append((float(parts[1]), float(parts[2]), float(parts[3])))
            elif parts[0] == "f":
                indices = []
                for p in parts[1:]:
                    raw = int(p.split("/")[0])
                    indices.append((len(verts) + raw) if raw < 0 else (raw - 1))
                for i in range(1, len(indices) - 1):
                    tris.append((indices[0], indices[i], indices[i + 1]))
    if not verts:
        raise ValueError("no vertex data found in OBJ file")
    return _mesh_to_shape(verts, tris), len(verts), len(tris)


def _read_amf_shape(path: Path):
    raw = path.read_bytes()
    if raw[:2] == b"PK":
        with zipfile.ZipFile(str(path), "r") as zf:
            names = zf.namelist()
            target = next((n for n in names if n.endswith(".amf")), names[0])
            with zf.open(target) as f:
                root = ET.parse(f).getroot()
    else:
        root = ET.fromstring(raw)

    verts_all = []
    tris_all  = []
    offset = 0
    for obj in root.findall("object"):
        mesh_el = obj.find("mesh")
        if mesh_el is None:
            continue
        verts_el = mesh_el.find("vertices")
        if verts_el is None:
            continue
        verts = []
        for vertex in verts_el.findall("vertex"):
            coords = vertex.find("coordinates")
            if coords is None:
                continue
            verts.append((
                float(coords.findtext("x", "0")),
                float(coords.findtext("y", "0")),
                float(coords.findtext("z", "0")),
            ))
        for volume in mesh_el.findall("volume"):
            for tri in volume.findall("triangle"):
                tris_all.append((
                    int(tri.findtext("v1")) + offset,
                    int(tri.findtext("v2")) + offset,
                    int(tri.findtext("v3")) + offset,
                ))
        verts_all.extend(verts)
        offset += len(verts)
    if not tris_all:
        raise ValueError("no triangle data found in AMF file")
    return _mesh_to_shape(verts_all, tris_all), len(verts_all), len(tris_all)


def _read_iges_shape(path: Path):
    reader = IGESControl_Reader()
    with quiet():
        status = reader.ReadFile(path.as_posix())
        if status == IFSelect_RetDone:
            reader.TransferRoots()
    if status != IFSelect_RetDone:
        raise ValueError(f"IGES reader failed with status {status}")
    shape = reader.OneShape()
    if shape.IsNull():
        raise ValueError("IGES file produced an empty shape")
    return shape, None, None


def _stl_tri_count(path: Path):
    try:
        with open(path, "rb") as f:
            header = f.read(84)
        if len(header) < 84:
            return None
        n = struct.unpack_from("<I", header, 80)[0]
        if 84 + n * 50 == path.stat().st_size:
            return n
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return sum(1 for line in f if line.lstrip().startswith("facet normal"))
    except Exception:
        return None


def _count_topo(shape, kind) -> int:
    exp = TopExp_Explorer(shape, kind)
    n = 0
    while exp.More():
        n += 1
        exp.Next()
    return n


def _topo_counts(shape) -> dict:
    return {
        "faces": _count_topo(shape, TopAbs_FACE),
        "edges": _count_topo(shape, TopAbs_EDGE),
    }


def _sew_chunk(args):
    faces, tolerance = args
    builder = BRep_Builder()
    compound = TopoDS_Compound()
    builder.MakeCompound(compound)
    for face in faces:
        builder.Add(compound, face)
    sew = BRepBuilderAPI_Sewing(tolerance)
    sew.Add(compound)
    sew.Perform()
    result = sew.SewedShape()
    return result if not result.IsNull() else compound


def _parallel_sew(shape, tolerance):
    faces = []
    exp = TopExp_Explorer(shape, TopAbs_FACE)
    while exp.More():
        faces.append(exp.Current())
        exp.Next()

    n_threads = os.cpu_count() or 1

    if len(faces) < 200 or n_threads < 2:
        sew = BRepBuilderAPI_Sewing(tolerance)
        sew.Add(shape)
        sew.Perform()
        sewn = sew.SewedShape()
        return sewn, sew.NbFreeEdges()

    chunk_size = max(1, (len(faces) + n_threads - 1) // n_threads)
    chunks = [faces[i:i + chunk_size] for i in range(0, len(faces), chunk_size)]

    with ThreadPoolExecutor(max_workers=n_threads) as executor:
        partial_shapes = list(executor.map(_sew_chunk, [(c, tolerance) for c in chunks]))

    partial_shapes = [s for s in partial_shapes if s is not None and not s.IsNull()]

    if not partial_shapes:
        return TopoDS_Shape(), 0

    final_sew = BRepBuilderAPI_Sewing(tolerance)
    for s in partial_shapes:
        final_sew.Add(s)
    final_sew.Perform()
    sewn = final_sew.SewedShape()
    return sewn, final_sew.NbFreeEdges()


def _run_brep_subprocess(shape, make_script):
    fd_in,  in_path  = tempfile.mkstemp(suffix='.brep')
    fd_out, out_path = tempfile.mkstemp(suffix='.brep')
    os.close(fd_in)
    os.close(fd_out)
    in_fwd  = in_path.replace('\\', '/')
    out_fwd = out_path.replace('\\', '/')
    try:
        breptools.Write(shape, in_fwd)
        proc = subprocess.run(
            [sys.executable, '-c', make_script(in_fwd, out_fwd)],
            capture_output=True,
            timeout=300,
        )
        if proc.returncode == 0 and os.path.getsize(out_path) > 0:
            result = TopoDS_Shape()
            breptools.Read(result, out_fwd, BRep_Builder())
            if not result.IsNull():
                return result
    except Exception:
        pass
    finally:
        for p in (in_path, out_path):
            try:
                os.unlink(p)
            except Exception:
                pass
    return shape


def _parallel_fix(shape):
    def make_script(in_fwd, out_fwd):
        return (
            "from OCC.Core.BRepTools import breptools;"
            "from OCC.Core.BRep import BRep_Builder;"
            "from OCC.Core.TopoDS import TopoDS_Shape;"
            "from OCC.Core.ShapeFix import ShapeFix_Shape;"
            "b=BRep_Builder();s=TopoDS_Shape();"
            f"breptools.Read(s,'{in_fwd}',b);"
            "f=ShapeFix_Shape(s);f.Perform();r=f.Shape();"
            f"breptools.Write(r if not r.IsNull() else s,'{out_fwd}')"
        )
    return _run_brep_subprocess(shape, make_script)


def _parallel_refine(shape, tolerance):
    def make_script(in_fwd, out_fwd):
        return (
            "from OCC.Core.BRepTools import breptools;"
            "from OCC.Core.BRep import BRep_Builder;"
            "from OCC.Core.TopoDS import TopoDS_Shape;"
            "from OCC.Core.ShapeUpgrade import ShapeUpgrade_UnifySameDomain;"
            "b=BRep_Builder();s=TopoDS_Shape();"
            f"breptools.Read(s,'{in_fwd}',b);"
            f"u=ShapeUpgrade_UnifySameDomain(s,True,True,True);"
            f"u.SetLinearTolerance({tolerance});"
            f"u.SetAngularTolerance({ANGULAR_TOLERANCE});"
            "u.Build();"
            "r=u.Shape();"
            f"breptools.Write(r if not r.IsNull() else s,'{out_fwd}')"
        )
    return _run_brep_subprocess(shape, make_script)


_step_open   = [False]
_timer_label = [""]
_timer_t0    = [0.0]
_timer_stop  = threading.Event()
_timer_th    = [None]
_TIMER_PAD   = _BOX_DETAIL + _BOX_TIME + 2


def _box_top():
    print(f"  {DIM}┌{'─' * (_BOX_CONTENT + 4)}┐{X}")


def _box_sep():
    print(f"  {DIM}├{'─' * (_BOX_CONTENT + 4)}┤{X}")


def _box_bot():
    print(f"  {DIM}└{'─' * (_BOX_CONTENT + 4)}┘{X}")


def _box_row(left: str, right: str = "", lc: str = "", rc: str = "") -> None:
    max_left = _BOX_CONTENT - len(right) - (1 if right else 0)
    if len(left) > max_left:
        left = left[:max_left - 3] + "..."
    gap = _BOX_CONTENT - len(left) - len(right)
    print(f"  {DIM}│{X}  {lc}{left}{X}{' ' * gap}{rc}{right}{X}  {DIM}│{X}")


def _stop_timer() -> None:
    _timer_stop.set()
    t = _timer_th[0]
    if t is not None:
        t.join(timeout=1.0)
        _timer_th[0] = None


def _step_start(label: str) -> float:
    _step_open[0]   = True
    _timer_label[0] = label
    t0 = time.perf_counter()
    _timer_t0[0]    = t0
    _timer_stop.clear()
    print(f"  {DIM}│{X}  {DIM}{label:<{_BOX_LABEL}}", end="", flush=True)

    def _tick():
        while not _timer_stop.wait(0.5):
            t_str = _fmt_time(time.perf_counter() - _timer_t0[0])
            line = f"\r  {DIM}│{X}  {DIM}{label:<{_BOX_LABEL}}{DIM}{t_str:>{_TIMER_PAD}}{X}\033[K"
            try:
                os.write(_real_stdout_fd, line.encode())
            except Exception:
                pass

    _timer_th[0] = threading.Thread(target=_tick, daemon=True)
    _timer_th[0].start()
    return t0


def _step_end(t0: float, detail: str = "") -> None:
    _stop_timer()
    _step_open[0] = False
    elapsed = time.perf_counter() - t0
    if len(detail) > _BOX_DETAIL:
        detail = detail[:_BOX_DETAIL - 3] + "..."
    time_str = _fmt_time(elapsed)
    print(f"\r  {DIM}│{X}  {DIM}{_timer_label[0]:<{_BOX_LABEL}}{X}{C}{detail:<{_BOX_DETAIL}}{X}  {Y}{time_str:>{_BOX_TIME}}{X}  {DIM}│{X}")


def _step_fail() -> None:
    if _step_open[0]:
        _stop_timer()
        _step_open[0] = False
        print(f"\r  {DIM}│{X}  {DIM}{_timer_label[0]:<{_BOX_LABEL}}{X}{R}{'error':<{_BOX_DETAIL}}{X}  {R}{'failed':>{_BOX_TIME}}{X}  {DIM}│{X}")


def _trim(name: str, width: int = NAME_TRIM_WIDTH) -> str:
    return name if len(name) <= width else name[:width - 3] + "..."


_EST_HISTORY = _DATA_DIR / "estimator.json"
_old_est = Path(__file__).parent / "estimator.json"
if not _EST_HISTORY.exists() and _old_est.exists():
    _EST_HISTORY.write_bytes(_old_est.read_bytes())
_EST_MIN     = 5
_EST_POWERS  = (2, 1, 0)


def _est_load():
    if _EST_HISTORY.exists():
        try:
            data = json.loads(_EST_HISTORY.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except Exception:
            pass
    return {}


def _est_save(data):
    try:
        _EST_HISTORY.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


def _bucket_add(bucket, x, y):
    scale  = bucket.get("scale",  x) or 1.0
    XtX    = bucket.get("XtX",    [[0.0] * 3 for _ in range(3)])
    Xty    = bucket.get("Xty",    [0.0] * 3)
    sum_y2 = bucket.get("sum_y2", 0.0)
    if x > scale:
        r   = scale / x
        XtX = [[XtX[i][j] * r ** (_EST_POWERS[i] + _EST_POWERS[j])
                for j in range(3)] for i in range(3)]
        Xty = [Xty[i] * r ** _EST_POWERS[i] for i in range(3)]
        scale = x
    xn  = x / scale
    row = (xn * xn, xn, 1.0)
    for i in range(3):
        Xty[i] += row[i] * y
        for j in range(3):
            XtX[i][j] += row[i] * row[j]
    return {"scale": scale, "XtX": XtX, "Xty": Xty, "sum_y2": sum_y2 + y * y}


def _bucket_predict(bucket, x):
    XtX    = bucket.get("XtX",    [[0.0] * 3 for _ in range(3)])
    Xty    = bucket.get("Xty",    [0.0] * 3)
    sum_y2 = bucket.get("sum_y2", 0.0)
    scale  = bucket.get("scale",  1.0) or 1.0
    n = int(round(XtX[2][2]))
    if n < _EST_MIN:
        return None, n, 0.0
    cs = _solve3(XtX, Xty)
    if cs is None:
        return None, n, 0.0
    a2, a1, a0 = cs
    xn   = x / scale
    pred   = max(0.5, a2 * xn * xn + a1 * xn + a0)
    sum_y  = Xty[2]
    ss_tot = sum_y2 - sum_y * sum_y / n
    ss_res = (sum_y2
              - 2 * (a2 * Xty[0] + a1 * Xty[1] + a0 * Xty[2])
              + a2 * a2 * XtX[0][0] + a1 * a1 * XtX[1][1] + a0 * a0 * XtX[2][2]
              + 2 * a2 * a1 * XtX[0][1] + 2 * a2 * a0 * XtX[0][2] + 2 * a1 * a0 * XtX[1][2])
    r2 = max(0.0, 1.0 - ss_res / ss_tot) if ss_tot > 1e-10 else 1.0
    return pred, n, r2


def _solve3(A, b):
    M = [A[i][:] + [b[i]] for i in range(3)]
    for col in range(3):
        pivot = max(range(col, 3), key=lambda r: abs(M[r][col]))
        if abs(M[pivot][col]) < 1e-12:
            return None
        M[col], M[pivot] = M[pivot], M[col]
        inv = 1.0 / M[col][col]
        M[col] = [v * inv for v in M[col]]
        for row in range(3):
            if row != col:
                f = M[row][col]
                M[row] = [M[row][j] - f * M[col][j] for j in range(4)]
    return [M[i][3] for i in range(3)]


def _rec_post_sew(fmt: str, n_faces: int, post_sew_seconds: float):
    data = _est_load()
    x, y = float(n_faces), float(post_sew_seconds)
    data[f"f:{fmt}"]  = _bucket_add(data.get(f"f:{fmt}",  {}), x, y)
    data["f:_all"]    = _bucket_add(data.get("f:_all",    {}), x, y)
    _est_save(data)


def _est_time_faces(n_faces: int, fmt=None):
    data  = _est_load()
    x     = float(n_faces)
    n_all = int(round(data.get("f:_all", {}).get("XtX", [[0]*3]*3)[2][2]))
    if fmt:
        key = f"f:{fmt}"
        if key in data:
            pred, n, r2 = _bucket_predict(data[key], x)
            if pred is not None:
                return pred, n, r2
    if "f:_all" in data:
        pred, n, r2 = _bucket_predict(data["f:_all"], x)
        if pred is not None:
            return pred, n, r2
    return None, n_all, 0.0


def _fmt_time(seconds: float) -> str:
    if seconds < 1.0:
        return f"{int(seconds * 1000)}ms"
    s = int(round(seconds))
    if s < 60:
        return f"{s}s"
    return f"{s // 60}m {s % 60:02d}s"


def _show_post_sew_estimate(n_faces: int, fmt: str = None) -> None:
    est, n_samples, r2 = _est_time_faces(n_faces, fmt=fmt)
    _box_sep()
    if est is not None:
        _box_row(
            f"~{_fmt_time(est)}  estimated",
            f"{n_samples} records · {int(r2 * 100)}%",
            lc=Y, rc=DIM,
        )
    else:
        _box_row(
            f"learning...  {n_samples} of {_EST_MIN} conversions recorded",
            lc=DIM,
        )
    _box_sep()


def _simplify_prompt(default_fraction, n_tris=None, batch=False):
    pct = int(round((1.0 - default_fraction) * 100)) if default_fraction else 0
    hint = ""
    if n_tris is not None and default_fraction:
        hint = f"  {DIM}({n_tris:,} → ~{max(1, int(n_tris * default_fraction)):,}){X}"
    batch_hint = f"  {DIM}[!N = all files]{X}" if batch else ""
    raw = input(f"  {DIM}│{X}  {DIM}{'simplify':<{_BOX_LABEL}}{X}{Y}{pct}%{X}{hint}{batch_hint}  ").strip()
    _box_sep()
    lock_all = raw.startswith("!")
    if lock_all:
        raw = raw[1:].strip()
    fraction = _parse_simplify(raw) if raw else default_fraction
    return fraction, lock_all


def _parse_simplify(value: str):
    if not value:
        return None
    try:
        pct = float(value.strip().rstrip('%'))
        if pct <= 0 or pct >= 100:
            return None
        return (100.0 - min(99.0, pct)) / 100.0
    except ValueError:
        return None


def _load_mesh_arrays(path: Path):
    import numpy as np
    ext = path.suffix.lower()

    if ext == STL_EXT:
        with open(path, 'rb') as f:
            f.seek(80)
            n = struct.unpack('<I', f.read(4))[0]
            data = f.read()
        if len(data) == n * 50:
            stl_dt = np.dtype([('n', np.float32, (3,)), ('v0', np.float32, (3,)),
                                ('v1', np.float32, (3,)), ('v2', np.float32, (3,)),
                                ('attr', np.uint16)])
            tris = np.frombuffer(data, dtype=stl_dt)
            verts = np.stack([tris['v0'], tris['v1'], tris['v2']], axis=1).reshape(-1, 3)
            return verts.astype(np.float32), np.arange(n * 3, dtype=np.int32).reshape(n, 3)
        verts = []
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                p = line.split()
                if p and p[0] == 'vertex':
                    verts.append([float(p[1]), float(p[2]), float(p[3])])
        v = np.array(verts, dtype=np.float32)
        return v, np.arange(len(v), dtype=np.int32).reshape(-1, 3)

    if ext == TMF_EXT:
        with zipfile.ZipFile(str(path)) as zf:
            with zf.open(_find_3mf_model(zf)) as f:
                root = ET.parse(f).getroot()
        resources = root.find(f"{{{_3MF_NS}}}resources")
        verts_all, tris_all, offset = [], [], 0
        for obj in resources.findall(f"{{{_3MF_NS}}}object"):
            mesh_el = obj.find(f"{{{_3MF_NS}}}mesh")
            if mesh_el is None:
                continue
            ve = mesh_el.find(f"{{{_3MF_NS}}}vertices")
            te = mesh_el.find(f"{{{_3MF_NS}}}triangles")
            if ve is None or te is None:
                continue
            verts = [(float(v.get('x')), float(v.get('y')), float(v.get('z')))
                     for v in ve.findall(f"{{{_3MF_NS}}}vertex")]
            tris_all += [(int(t.get('v1')) + offset, int(t.get('v2')) + offset,
                          int(t.get('v3')) + offset)
                         for t in te.findall(f"{{{_3MF_NS}}}triangle")]
            verts_all += verts
            offset += len(verts)
        return np.array(verts_all, dtype=np.float32), np.array(tris_all, dtype=np.int32)

    if ext == OBJ_EXT:
        verts, tris = [], []
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                p = line.split()
                if not p:
                    continue
                if p[0] == 'v':
                    verts.append([float(p[1]), float(p[2]), float(p[3])])
                elif p[0] == 'f':
                    idx = [(len(verts) + int(x.split('/')[0])) if int(x.split('/')[0]) < 0
                           else (int(x.split('/')[0]) - 1) for x in p[1:]]
                    for i in range(1, len(idx) - 1):
                        tris.append([idx[0], idx[i], idx[i + 1]])
        return np.array(verts, dtype=np.float32), np.array(tris, dtype=np.int32)

    if ext == AMF_EXT:
        raw = path.read_bytes()
        if raw[:2] == b'PK':
            with zipfile.ZipFile(str(path)) as zf:
                names = zf.namelist()
                target = next((n for n in names if n.endswith('.amf')), names[0])
                with zf.open(target) as f:
                    root = ET.parse(f).getroot()
        else:
            root = ET.fromstring(raw)
        verts_all, tris_all, offset = [], [], 0
        for obj in root.findall('object'):
            mesh_el = obj.find('mesh')
            if mesh_el is None:
                continue
            ve = mesh_el.find('vertices')
            if ve is None:
                continue
            verts = []
            for vertex in ve.findall('vertex'):
                coords = vertex.find('coordinates')
                if coords is None:
                    continue
                verts.append([float(coords.findtext('x', '0')),
                               float(coords.findtext('y', '0')),
                               float(coords.findtext('z', '0'))])
            for volume in mesh_el.findall('volume'):
                for tri in volume.findall('triangle'):
                    tris_all.append([int(tri.findtext('v1')) + offset,
                                     int(tri.findtext('v2')) + offset,
                                     int(tri.findtext('v3')) + offset])
            verts_all += verts
            offset += len(verts)
        return np.array(verts_all, dtype=np.float32), np.array(tris_all, dtype=np.int32)

    raise ValueError(f"unsupported format for simplification: {path.suffix}")


def _write_simplified_stl(path: Path, verts, faces):
    n = len(faces)
    buf = bytearray(80 + 4 + n * _STL_TRI.size)
    struct.pack_into('<I', buf, 80, n)
    for i, tri in enumerate(faces):
        v0, v1, v2 = verts[tri[0]], verts[tri[1]], verts[tri[2]]
        _STL_TRI.pack_into(buf, 84 + i * _STL_TRI.size, 0.0, 0.0, 0.0, *v0, *v1, *v2, 0)
    path.write_bytes(bytes(buf))


def _simplify_mesh(path: Path, keep_fraction: float):
    import fast_simplification
    import numpy as np

    verts, faces = _load_mesh_arrays(path)
    n_before = len(faces)
    if n_before == 0:
        raise ValueError("mesh is empty")
    target_reduction = max(0.01, min(0.99, 1.0 - keep_fraction))
    verts_out, faces_out = fast_simplification.simplify(
        verts.astype(np.float32),
        faces.astype(np.int32),
        target_reduction,
    )
    n_after = len(faces_out)
    fd, tmp = tempfile.mkstemp(suffix='.stl')
    os.close(fd)
    _write_simplified_stl(Path(tmp), verts_out, faces_out)
    return Path(tmp), n_before, n_after


_STEP_SCHEMAS = {"ap203": "AP203", "ap214": "AP214IS", "ap242": "AP242DIS"}


def convert(input_path: Path, output_path: Path, tolerance: float = DEFAULT_TOLERANCE,
            simplify_fraction=None, interactive: bool = False,
            batch: bool = False, _chosen_out: dict = None,
            step_schema: str = "AP203", _suppress_read_step: bool = False):
    n_verts = n_tris = None
    _tmp_simplified = None
    original_path = input_path

    try:
        ext = input_path.suffix.lower()
        original_ext = ext

        if not _suppress_read_step:
            t = _step_start("reading")
        if ext == TMF_EXT:
            shape, n_verts, n_tris = _read_3mf_shape(input_path)
        elif ext == OBJ_EXT:
            shape, n_verts, n_tris = _read_obj_shape(input_path)
        elif ext == AMF_EXT:
            shape, n_verts, n_tris = _read_amf_shape(input_path)
        elif ext in {IGS_EXT, ".iges"}:
            shape, n_verts, n_tris = _read_iges_shape(input_path)
        else:
            shape = TopoDS_Shape()
            with quiet():
                StlAPI_Reader().Read(shape, input_path.as_posix())
            n_tris = _stl_tri_count(input_path)
        if shape.IsNull():
            if not _suppress_read_step:
                _step_fail()
            return False, "input produced an empty shape"
        if not _suppress_read_step:
            read_parts = []
            if n_verts is not None:
                read_parts.append(f"{n_verts:,} vertices")
            if n_tris is not None:
                read_parts.append(f"{n_tris:,} triangles")
            if not read_parts:
                read_parts.append(f"{_count_topo(shape, TopAbs_FACE):,} faces")
            _step_end(t, " · ".join(read_parts))

        if interactive and original_ext not in {IGS_EXT, ".iges"}:
            _box_sep()
            simplify_fraction, _lock_all = _simplify_prompt(
                simplify_fraction, n_tris=n_tris, batch=batch)
            if _chosen_out is not None:
                _chosen_out['fraction'] = simplify_fraction
                _chosen_out['lock'] = _lock_all

        if simplify_fraction is not None and original_ext not in {IGS_EXT, ".iges"}:
            t = _step_start("simplify")
            try:
                _tmp_simplified, n_before, n_after = _simplify_mesh(original_path, simplify_fraction)
                red_pct = int(round((1.0 - n_after / n_before) * 100))
                _new_shape = TopoDS_Shape()
                with quiet():
                    StlAPI_Reader().Read(_new_shape, _tmp_simplified.as_posix())
                if _new_shape.IsNull():
                    raise ValueError("simplified STL produced an empty shape")
                _step_end(t, f"{n_before:,} → {n_after:,} triangles  (-{red_pct}%)")
                shape = _new_shape
                n_tris = n_after
            except ImportError as e:
                _step_end(t, f"skipped  ({e})")
            except Exception as e:
                _step_end(t, f"skipped  ({e})")

        t = _step_start("sewing")
        with quiet():
            sewn, n_free = _parallel_sew(shape, tolerance)
        if sewn.IsNull():
            _step_fail()
            return False, "sewing failed, try a larger tolerance"
        n_shells     = _count_topo(sewn, TopAbs_SHELL)
        n_faces_sewn = _count_topo(sewn, TopAbs_FACE)
        sew_parts = [f"{n_shells:,} shell{'s' if n_shells != 1 else ''}"]
        if n_free:
            sew_parts.append(f"{n_free:,} free edge{'s' if n_free != 1 else ''}")
        _step_end(t, " · ".join(sew_parts))

        t_post_sew = time.perf_counter()
        _show_post_sew_estimate(n_faces_sewn, fmt=original_ext)

        t = _step_start("fixing")
        with quiet():
            fixed = _parallel_fix(sewn)
        n_faces_out = _count_topo(fixed, TopAbs_FACE)
        _step_end(t, f"{n_faces_sewn:,} to {n_faces_out:,} faces")

        t = _step_start("refining")
        with quiet():
            refined = _parallel_refine(fixed, tolerance)
        n_faces_after = _count_topo(refined, TopAbs_FACE)
        _step_end(t, f"{n_faces_out:,} to {n_faces_after:,} faces")

        t = _step_start("writing")
        Interface_Static.SetCVal("write.step.schema", step_schema)
        writer = STEPControl_Writer()
        with quiet():
            ts = writer.Transfer(refined, STEPControl_AsIs)
            ws = writer.Write(output_path.as_posix())
        if ts in (IFSelect_RetError, IFSelect_RetFail) or \
           ws in (IFSelect_RetError, IFSelect_RetFail):
            _step_fail()
            return False, "STEP writer failed"
        if not output_path.exists() or output_path.stat().st_size == 0:
            _step_fail()
            return False, "output file is missing or empty"
        out_kb = output_path.stat().st_size // BYTES_PER_KB
        _step_end(t, f"{out_kb:,} KB")

        topo = _topo_counts(refined)
        _rec_post_sew(original_ext, n_faces_sewn, time.perf_counter() - t_post_sew)
        return True, {
            "kb":    out_kb,
            "verts": n_verts,
            "tris":  n_tris,
            "faces": topo["faces"],
            "edges": topo["edges"],
        }

    except Exception:
        _step_fail()
        return False, traceback.format_exc().strip()

    finally:
        if _tmp_simplified is not None:
            try:
                _tmp_simplified.unlink()
            except Exception:
                pass


def models_dir() -> Path:
    return Path(__file__).parent / MODELS_DIR_NAME


def _err_line(tb: str) -> str:
    lines = [l for l in tb.splitlines() if l.strip()]
    return lines[-1] if lines else tb


def _quick_tri_count(path: Path):
    ext = path.suffix.lower()
    if ext == STL_EXT:
        return _stl_tri_count(path)
    if ext == TMF_EXT:
        try:
            with zipfile.ZipFile(str(path)) as zf:
                model_file = _find_3mf_model(zf)
                with zf.open(model_file) as f:
                    root = ET.parse(f).getroot()
            return sum(1 for _ in root.iter(f"{{{_3MF_NS}}}triangle")) or None
        except Exception:
            return None
    if ext == OBJ_EXT:
        try:
            count = 0
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    if line.startswith("f ") or line.startswith("f\t"):
                        count += max(1, len(line.split()) - 3)
            return count or None
        except Exception:
            return None
    if ext == AMF_EXT:
        try:
            raw = path.read_bytes()
            if raw[:2] == b"PK":
                with zipfile.ZipFile(str(path)) as zf:
                    names = zf.namelist()
                    target = next((n for n in names if n.endswith(".amf")), names[0])
                    with zf.open(target) as f:
                        root = ET.parse(f).getroot()
            else:
                root = ET.fromstring(raw)
            return sum(1 for _ in root.iter("triangle")) or None
        except Exception:
            return None
    return None


def _make_output_path(src: Path, base_dir: Path, fraction, ext: str) -> Path:
    stem = src.stem
    if fraction:
        reduction_pct = int(round((1.0 - fraction) * 100))
        stem = f"{stem} [{reduction_pct}]"
    return base_dir / (stem + ext)


def _run_batch(files, n, out_dir, args, simplify_fraction, step_schema, label_prefix=True):
    ok_n = fail_n = skip_n = 0
    _lock_frac = None

    for i, src_file in enumerate(files):
        base_dir = out_dir or src_file.parent
        _is_interactive = SIMPLIFY_INTERACTIVE and not args.simplify and _lock_frac is None
        eff_fraction = _lock_frac if _lock_frac is not None else simplify_fraction

        src_kb = src_file.stat().st_size // BYTES_PER_KB
        size_str = f"{src_kb:,} KB"
        prefix = f"[{i+1}/{n}]  " if label_prefix else ""
        name_trim = _BOX_CONTENT - len(size_str) - len(prefix) - 1

        if _is_interactive and not getattr(args, 'dry_run', False):
            _box_top()
            _box_row(f"{prefix}{_trim(src_file.name, name_trim)}", size_str, lc=B, rc=DIM)
            _box_sep()
            _t_read = _step_start("reading")
            n_tris_preview = _quick_tri_count(src_file)
            _read_detail = f"{n_tris_preview:,} triangles" if n_tris_preview is not None else ""
            _step_end(_t_read, _read_detail)
            _box_sep()
            chosen_frac, lock_all = _simplify_prompt(eff_fraction, n_tris=n_tris_preview, batch=True)
            if lock_all:
                _lock_frac = chosen_frac
            out_file = _make_output_path(src_file, base_dir, chosen_frac, STP_EXT)
            _up_to_date = (SKIP_EXISTING and not args.force
                           and out_file.exists()
                           and out_file.stat().st_mtime >= src_file.stat().st_mtime)
            if _up_to_date:
                _skip_kb = out_file.stat().st_size // BYTES_PER_KB
                _skip_str = f"{_skip_kb:,} KB"
                _box_row(f"✓  {_trim(out_file.name, _BOX_CONTENT - len(_skip_str) - 3)}", _skip_str, lc=f"{G}{B}", rc=G)
                _box_bot()
                print()
                skip_n += 1
                continue
            _t0 = time.perf_counter()
            success, info = convert(src_file, out_file, args.tolerance, chosen_frac,
                                    step_schema=step_schema, _suppress_read_step=True)
            _elapsed = time.perf_counter() - _t0
            _box_sep()
            if success:
                out_str = f"{info['kb']:,} KB · {_fmt_time(_elapsed)}"
                _box_row(f"✓  {_trim(out_file.name, _BOX_CONTENT - len(out_str) - 3)}", out_str, lc=f"{G}{B}", rc=G)
                ok_n += 1
            else:
                _box_row(f"✗  {_err_line(info)}", lc=R)
                fail_n += 1
            _box_bot()
            print()
            continue

        out_file = _make_output_path(src_file, base_dir, eff_fraction, STP_EXT)
        _up_to_date = (SKIP_EXISTING and not args.force
                       and out_file.exists()
                       and out_file.stat().st_mtime >= src_file.stat().st_mtime)

        if getattr(args, 'dry_run', False):
            if _up_to_date:
                print(f"  {DIM}↷  {_trim(src_file.name)}  up-to-date{X}")
                skip_n += 1
            else:
                print(f"  {C}→  {_trim(src_file.name)}  {src_kb:,} KB → {out_file.name}{X}")
                ok_n += 1
            continue

        _box_top()
        _box_row(f"{prefix}{_trim(src_file.name, name_trim)}", size_str, lc=B, rc=DIM)
        _box_sep()

        if _up_to_date:
            _skip_kb = out_file.stat().st_size // BYTES_PER_KB
            _skip_str = f"{_skip_kb:,} KB"
            _box_row(f"✓  {_trim(out_file.name, _BOX_CONTENT - len(_skip_str) - 3)}", _skip_str, lc=f"{G}{B}", rc=G)
            _box_bot()
            print()
            skip_n += 1
            continue

        _t0 = time.perf_counter()
        success, info = convert(src_file, out_file, args.tolerance, eff_fraction,
                                step_schema=step_schema)
        _elapsed = time.perf_counter() - _t0
        _box_sep()
        if success:
            out_str = f"{info['kb']:,} KB · {_fmt_time(_elapsed)}"
            _box_row(f"✓  {_trim(out_file.name, _BOX_CONTENT - len(out_str) - 3)}", out_str, lc=f"{G}{B}", rc=G)
            ok_n += 1
        else:
            _box_row(f"✗  {_err_line(info)}", lc=R)
            fail_n += 1
        _box_bot()
        print()

    return ok_n, fail_n, skip_n


def _print_summary(ok_n, fail_n, skip_n, dry_run=False):
    print(f"  {DIM}{'─' * SEPARATOR_WIDTH}{X}")
    if fail_n == 0 and skip_n == 0:
        verb = "would convert" if dry_run else "converted"
        print(f"  {G}{B}✓  All {ok_n} file{'s' if ok_n > 1 else ''} {verb} successfully{X}")
    else:
        _parts = []
        if ok_n:   _parts.append(f"{G}✓  {ok_n} {'would convert' if dry_run else 'converted'}{X}")
        if skip_n: _parts.append(f"{DIM}↷  {skip_n} skipped{X}")
        if fail_n: _parts.append(f"{R}✗  {fail_n} failed{X}")
        print(f"  {'    '.join(_parts)}")


def main():
    parser = argparse.ArgumentParser(description="Convert to STEP.")
    parser.add_argument("input", nargs="*",
        help="input file(s); omit to convert everything in the models/ folder")
    parser.add_argument("--output", "-o", metavar="FILE",
        help="output path (single-file mode only)")
    parser.add_argument("--output-dir", "-d", metavar="DIR",
        help="write output files to this directory instead of alongside the source")
    parser.add_argument("--tolerance", "-t", type=float, default=DEFAULT_TOLERANCE)
    parser.add_argument("--simplify", "-s", metavar="PCT",
        help="keep this %% of triangles before converting (0 = off)")
    parser.add_argument("--format", metavar="SCHEMA", default=DEFAULT_FORMAT,
        choices=["ap203", "ap214", "ap242"],
        help=f"STEP schema: ap203, ap214, ap242 (default: {DEFAULT_FORMAT})")
    parser.add_argument("--force", "-f", action="store_true",
        help="re-convert files even if the output is already up-to-date")
    parser.add_argument("--dry-run", "--dry", action="store_true",
        help="show what would be converted without actually converting")
    parser.add_argument("--watch", "-w", action="store_true",
        help="after batch conversion, watch the folder and convert new files automatically")
    args = parser.parse_args()

    simplify_fraction = (_parse_simplify(args.simplify) if args.simplify
                         else ((100.0 - DEFAULT_SIMPLIFY) / 100.0 if 0 < DEFAULT_SIMPLIFY < 100 else None))
    step_schema = _STEP_SCHEMAS.get(args.format, "AP203")

    print()
    _w = _BOX_CONTENT + 4
    print(f"  {C}{B}╔{'═' * _w}╗{X}")
    print(f"  {C}{B}║{'2STEP-Converter':^{_w}}║{X}")
    print(f"  {C}{B}╚{'═' * _w}╝{X}")
    print()

    if _cfg_warnings:
        for _w_msg in _cfg_warnings:
            print(f"  {Y}⚠  {_w_msg}{X}")
        print()

    out_dir = Path(args.output_dir).resolve() if args.output_dir else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)

    if len(args.input) > 1:
        files = []
        for p in args.input:
            fp = Path(p).resolve()
            if not fp.exists():
                print(f"  {R}[ERROR]{X} File not found: {fp}")
            elif fp.suffix.lower() not in _SUPPORTED_EXTS:
                print(f"  {R}[ERROR]{X} Unsupported format: {fp.name}")
            else:
                files.append(fp)
        if not files:
            sys.exit(1)
        n = len(files)
        print(f"  {n} file{'s' if n > 1 else ''} to convert\n")
        ok_n, fail_n, skip_n = _run_batch(files, n, out_dir, args, simplify_fraction, step_schema)
        _print_summary(ok_n, fail_n, skip_n, dry_run=args.dry_run)
        print()
        input("  Press Enter to exit...")
        sys.exit(0 if fail_n == 0 else 1)

    if len(args.input) == 1:
        input_path = Path(args.input[0]).resolve()
        if not input_path.exists():
            print(f"  {R}[ERROR]{X} File not found: {input_path}\n")
            input("  Press Enter to exit...")
            sys.exit(1)

        _single_interactive = SIMPLIFY_INTERACTIVE and not args.simplify
        _single_base_dir = out_dir or input_path.parent

        if args.output:
            out = Path(args.output)
            output_path = input_path.parent / out if not out.is_absolute() else out
        else:
            output_path = _make_output_path(input_path, _single_base_dir, simplify_fraction, STP_EXT)

        src_kb = input_path.stat().st_size // BYTES_PER_KB
        size_str = f"{src_kb:,} KB"

        if args.dry_run:
            if SKIP_EXISTING and not args.force and output_path.exists() and output_path.stat().st_mtime >= input_path.stat().st_mtime:
                print(f"  {DIM}↷  {_trim(input_path.name)}  up-to-date{X}\n")
            else:
                print(f"  {C}→  {_trim(input_path.name)}  {src_kb:,} KB → {output_path.name}{X}\n")
            input("  Press Enter to exit...")
            sys.exit(0)

        _box_top()
        _box_row(f"[1/1]  {_trim(input_path.name, _BOX_CONTENT - len(size_str) - 3)}", size_str, lc=B, rc=DIM)
        _box_sep()
        _chosen = {}
        _t0 = time.perf_counter()
        success, info = convert(input_path, output_path, args.tolerance, simplify_fraction,
                                interactive=_single_interactive,
                                _chosen_out=_chosen if _single_interactive else None,
                                step_schema=step_schema)
        if _single_interactive and success and _chosen and not args.output:
            chosen_frac = _chosen.get('fraction')
            new_out = _make_output_path(input_path, _single_base_dir, chosen_frac, STP_EXT)
            if new_out != output_path:
                try:
                    output_path.rename(new_out)
                    output_path = new_out
                except Exception:
                    pass
        _elapsed = time.perf_counter() - _t0
        _box_sep()
        if success:
            out_str = f"{info['kb']:,} KB · {_fmt_time(_elapsed)}"
            _box_row(f"✓  {_trim(output_path.name, _BOX_CONTENT - len(out_str) - 3)}", out_str, lc=f"{G}{B}", rc=G)
        else:
            _box_row(f"✗  {_err_line(info)}", lc=R)
        _box_bot()
        print()
        if success:
            print(f"  {G}{B}✓  Done{X}")
        print()
        input("  Press Enter to exit...")
        sys.exit(0 if success else 1)

    folder = models_dir()
    folder.mkdir(exist_ok=True)

    files = sorted(f for f in folder.iterdir() if f.suffix.lower() in _SUPPORTED_EXTS)
    if not files:
        print(f"  No supported files found in {MODELS_DIR_NAME}\\\n")
        input("  Press Enter to exit...")
        sys.exit(0)

    n = len(files)
    print(f"  {n} file{'s' if n > 1 else ''} found in {C}{MODELS_DIR_NAME}\\{X}\n")

    ok_n, fail_n, skip_n = _run_batch(files, n, out_dir, args, simplify_fraction, step_schema)
    _print_summary(ok_n, fail_n, skip_n, dry_run=args.dry_run)
    print()

    if args.watch and not args.dry_run:
        known = {f.name for f in folder.iterdir() if f.suffix.lower() in _SUPPORTED_EXTS}
        print(f"  {DIM}Watching {C}{MODELS_DIR_NAME}\\{X}{DIM}  Ctrl+C to stop{X}\n")
        try:
            while True:
                time.sleep(2)
                current = {f.name for f in folder.iterdir() if f.suffix.lower() in _SUPPORTED_EXTS}
                new_names = current - known
                known = current
                for name in sorted(new_names):
                    src = folder / name
                    dst = _make_output_path(src, out_dir or folder, simplify_fraction, STP_EXT)
                    src_kb2 = src.stat().st_size // BYTES_PER_KB
                    size_str2 = f"{src_kb2:,} KB"
                    _box_top()
                    _box_row(f"[new]  {_trim(src.name, _BOX_CONTENT - len(size_str2) - 3)}", size_str2, lc=B, rc=DIM)
                    _box_sep()
                    _t0 = time.perf_counter()
                    success2, info2 = convert(src, dst, args.tolerance, simplify_fraction,
                                              step_schema=step_schema)
                    _elapsed2 = time.perf_counter() - _t0
                    _box_sep()
                    if success2:
                        out_str2 = f"{info2['kb']:,} KB · {_fmt_time(_elapsed2)}"
                        _box_row(f"✓  {_trim(dst.name, _BOX_CONTENT - len(out_str2) - 3)}", out_str2, lc=f"{G}{B}", rc=G)
                    else:
                        _box_row(f"✗  {_err_line(info2)}", lc=R)
                    _box_bot()
                    print()
        except KeyboardInterrupt:
            print(f"\n  {DIM}Watch stopped.{X}\n")
        sys.exit(0 if fail_n == 0 else 1)
    else:
        input("  Press Enter to exit...")
        sys.exit(0 if fail_n == 0 else 1)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\n  Press Enter to exit...")
        sys.exit(1)
