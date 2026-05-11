import sys
import os
import json
import ctypes
import traceback
import argparse
import zipfile
import struct
import tempfile
import time
from contextlib import contextmanager
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from xml.etree import ElementTree as ET

_CONFIG_PATH = Path(__file__).parent / "config.py"
if not _CONFIG_PATH.exists():
    _CONFIG_PATH.write_text(
        "DEFAULT_TOLERANCE   = 0.01\n"
        "ANGULAR_TOLERANCE   = 1e-3\n"
        "NAME_TRIM_WIDTH     = 62\n"
        "SEPARATOR_WIDTH     = 32\n"
        "BYTES_PER_KB        = 1024\n"
        "STD_OUTPUT_HANDLE   = -11\n"
        "CONSOLE_MODE_FLAGS  = 7\n"
        'MODELS_DIR_NAME     = "models"\n'
        'STL_EXT             = ".stl"\n'
        'TMF_EXT             = ".3mf"\n'
        'OBJ_EXT             = ".obj"\n'
        'IGS_EXT             = ".igs"\n'
        'AMF_EXT             = ".amf"\n'
        'STP_EXT             = ".stp"\n',
        encoding="utf-8",
    )

from config import (
    STD_OUTPUT_HANDLE, CONSOLE_MODE_FLAGS,
    DEFAULT_TOLERANCE, ANGULAR_TOLERANCE,
    NAME_TRIM_WIDTH, SEPARATOR_WIDTH,
    MODELS_DIR_NAME, STL_EXT, TMF_EXT, OBJ_EXT, IGS_EXT, AMF_EXT, STP_EXT, BYTES_PER_KB,
)

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
        from OCC.Core.BRepBuilderAPI import BRepBuilderAPI_Sewing
        from OCC.Core.ShapeUpgrade import ShapeUpgrade_UnifySameDomain
        from OCC.Core.ShapeFix import ShapeFix_Shape
        from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_AsIs
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
_BOX_LABEL   = 10
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


def _collect_shells(shape):
    shells = []
    exp = TopExp_Explorer(shape, TopAbs_SHELL)
    while exp.More():
        shells.append(exp.Current())
        exp.Next()
    return shells


def _combine_shapes(shapes):
    builder = BRep_Builder()
    compound = TopoDS_Compound()
    builder.MakeCompound(compound)
    for s in shapes:
        builder.Add(compound, s)
    return compound


def _fix_shell(shell):
    try:
        fix = ShapeFix_Shape(shell)
        fix.Perform()
        result = fix.Shape()
        return result if not result.IsNull() else shell
    except Exception:
        return shell


def _refine_shell(args):
    shell, tolerance = args
    try:
        u = ShapeUpgrade_UnifySameDomain(shell, True, True, True)
        u.SetLinearTolerance(tolerance)
        u.SetAngularTolerance(ANGULAR_TOLERANCE)
        u.Build()
        result = u.Shape()
        return result if not result.IsNull() else shell
    except Exception:
        return shell


def _parallel_fix(shape):
    shells = _collect_shells(shape)
    n_threads = os.cpu_count() or 1
    if len(shells) < 2 or n_threads < 2:
        try:
            fix = ShapeFix_Shape(shape)
            fix.Perform()
            result = fix.Shape()
            return result if not result.IsNull() else shape
        except Exception:
            return shape
    with ThreadPoolExecutor(max_workers=min(n_threads, len(shells))) as executor:
        results = list(executor.map(_fix_shell, shells))
    results = [r for r in results if r is not None and not r.IsNull()]
    return _combine_shapes(results) if results else shape


def _parallel_refine(shape, tolerance):
    shells = _collect_shells(shape)
    if len(shells) < 2:
        try:
            u = ShapeUpgrade_UnifySameDomain(shape, True, True, True)
            u.SetLinearTolerance(tolerance)
            u.SetAngularTolerance(ANGULAR_TOLERANCE)
            u.Build()
            result = u.Shape()
            return result if not result.IsNull() else shape
        except Exception:
            return shape
    results = [_refine_shell((s, tolerance)) for s in shells]
    results = [r for r in results if r is not None and not r.IsNull()]
    return _combine_shapes(results) if results else shape


_step_open = [False]


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


def _step_start(label: str) -> float:
    _step_open[0] = True
    print(f"  {DIM}│{X}  {DIM}{label:<{_BOX_LABEL}}", end="", flush=True)
    return time.perf_counter()


def _step_end(t0: float, detail: str = "") -> None:
    _step_open[0] = False
    elapsed = time.perf_counter() - t0
    if len(detail) > _BOX_DETAIL:
        detail = detail[:_BOX_DETAIL - 3] + "..."
    time_str = _fmt_time(elapsed)
    print(f"{X}{C}{detail:<{_BOX_DETAIL}}{X}  {Y}{time_str:>{_BOX_TIME}}{X}  {DIM}│{X}")


def _step_fail() -> None:
    if _step_open[0]:
        _step_open[0] = False
        print(f"{X}{R}{'error':<{_BOX_DETAIL}}{X}  {R}{'failed':>{_BOX_TIME}}{X}  {DIM}│{X}")


def _trim(name: str, width: int = NAME_TRIM_WIDTH) -> str:
    return name if len(name) <= width else name[:width - 3] + "..."


_EST_HISTORY = Path(__file__).parent / "estimator.json"
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


def convert(input_path: Path, output_path: Path, tolerance: float = DEFAULT_TOLERANCE):
    n_verts = n_tris = None

    try:
        t = _step_start("reading")
        ext = input_path.suffix.lower()
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
            _step_fail()
            return False, "input produced an empty shape"
        read_parts = []
        if n_verts is not None:
            read_parts.append(f"{n_verts:,} vertices")
        if n_tris is not None:
            read_parts.append(f"{n_tris:,} triangles")
        if not read_parts:
            read_parts.append(f"{_count_topo(shape, TopAbs_FACE):,} faces")
        _step_end(t, "  ·  ".join(read_parts))

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
        _step_end(t, "  ·  ".join(sew_parts))

        t_post_sew = time.perf_counter()
        _show_post_sew_estimate(n_faces_sewn, fmt=ext)

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
        _rec_post_sew(ext, n_faces_sewn, time.perf_counter() - t_post_sew)
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


def models_dir() -> Path:
    return Path(__file__).parent / MODELS_DIR_NAME


def main():
    parser = argparse.ArgumentParser(description="Convert to STEP.")
    parser.add_argument("input",  nargs="?")
    parser.add_argument("output", nargs="?")
    parser.add_argument("--tolerance", "-t", type=float, default=DEFAULT_TOLERANCE)
    args = parser.parse_args()

    print()
    _w = _BOX_CONTENT + 4
    print(f"  {C}{B}╔{'═' * _w}╗{X}")
    print(f"  {C}{B}║{'2STEP-Converter':^{_w}}║{X}")
    print(f"  {C}{B}╚{'═' * _w}╝{X}")
    print()

    if args.input is None:
        folder = models_dir()
        folder.mkdir(exist_ok=True)

        files = sorted(f for f in folder.iterdir() if f.suffix.lower() in _SUPPORTED_EXTS)
        if not files:
            print(f"  No supported files found in {MODELS_DIR_NAME}\\\n")
            input("  Press Enter to exit...")
            sys.exit(0)

        n = len(files)
        print(f"  {n} file{'s' if n > 1 else ''} found in {C}{MODELS_DIR_NAME}\\{X}\n")

        tasks = [(f, f.with_suffix(STP_EXT), args.tolerance) for f in files]
        ok_n = fail_n = 0

        for i, (src_file, out_file, tol) in enumerate(tasks):
            src_kb = src_file.stat().st_size // BYTES_PER_KB
            size_str = f"{src_kb:,} KB"
            _box_top()
            _box_row(f"[{i+1}/{n}]  {_trim(src_file.name, _BOX_CONTENT - len(size_str) - 3)}", size_str, lc=B, rc=DIM)
            _box_sep()
            _t0 = time.perf_counter()
            success, info = convert(src_file, out_file, tol)
            _elapsed = time.perf_counter() - _t0
            _box_sep()
            if success:
                out_str = f"{info['kb']:,} KB  ·  {_fmt_time(_elapsed)}"
                _box_row(f"✓  {_trim(out_file.name, _BOX_CONTENT - len(out_str) - 3)}", out_str, lc=f"{G}{B}", rc=G)
                ok_n += 1
            else:
                _box_row(f"✗  {info.split(chr(10))[-1]}", lc=R)
                fail_n += 1
            _box_bot()
            print()

        print(f"  {DIM}{'─' * SEPARATOR_WIDTH}{X}")
        if fail_n == 0:
            print(f"  {G}{B}✓  All {ok_n} file{'s' if ok_n > 1 else ''} converted successfully{X}")
        else:
            print(f"  {G}✓  {ok_n} converted{X}    {R}✗  {fail_n} failed{X}")
        print()
        input("  Press Enter to exit...")
        sys.exit(0 if fail_n == 0 else 1)

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"  {R}[ERROR]{X} File not found: {input_path}\n")
        input("  Press Enter to exit...")
        sys.exit(1)

    if args.output:
        out = Path(args.output)
        output_path = input_path.parent / out if not out.is_absolute() else out
    else:
        output_path = input_path.with_suffix(STP_EXT)

    src_kb = input_path.stat().st_size // BYTES_PER_KB
    size_str = f"{src_kb:,} KB"
    _box_top()
    _box_row(f"[1/1]  {_trim(input_path.name, _BOX_CONTENT - len(size_str) - 3)}", size_str, lc=B, rc=DIM)
    _box_sep()
    _t0 = time.perf_counter()
    success, info = convert(input_path, output_path, args.tolerance)
    _elapsed = time.perf_counter() - _t0
    _box_sep()
    if success:
        out_str = f"{info['kb']:,} KB  ·  {_fmt_time(_elapsed)}"
        _box_row(f"✓  {_trim(output_path.name, _BOX_CONTENT - len(out_str) - 3)}", out_str, lc=f"{G}{B}", rc=G)
    else:
        _box_row(f"✗  {info.split(chr(10))[-1]}", lc=R)
    _box_bot()
    print()
    if success:
        print(f"  {G}{B}✓  Done{X}")
    print()
    input("  Press Enter to exit...")
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("\n  Press Enter to exit...")
        sys.exit(1)
