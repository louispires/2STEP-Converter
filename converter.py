import sys
import os
import ctypes
import traceback
import argparse
import zipfile
import struct
import tempfile
from contextlib import contextmanager
from pathlib import Path
from xml.etree import ElementTree as ET

_CONFIG_PATH = Path(__file__).parent / "config.py"
if not _CONFIG_PATH.exists():
    _CONFIG_PATH.write_text(
        "DEFAULT_TOLERANCE   = 0.01\n"
        "ANGULAR_TOLERANCE   = 1e-3\n"
        "NAME_TRIM_WIDTH     = 50\n"
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
    os.dup2(nul, 1)
    os.dup2(nul, 2)
    os.close(nul)
    try:
        yield
    finally:
        sys.stdout.flush()
        sys.stderr.flush()
        os.dup2(fd1, 1); os.close(fd1)
        os.dup2(fd2, 2); os.close(fd2)


_3MF_NS = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02"


def _mesh_to_shape(verts: list, tris: list):
    if not tris:
        raise ValueError("no triangle data found")
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".stl")
    try:
        with os.fdopen(tmp_fd, "wb") as f:
            f.write(b"\x00" * 80)
            f.write(struct.pack("<I", len(tris)))
            for t in tris:
                v0, v1, v2 = verts[t[0]], verts[t[1]], verts[t[2]]
                f.write(struct.pack("<fff", 0.0, 0.0, 0.0))
                f.write(struct.pack("<fff", *v0))
                f.write(struct.pack("<fff", *v1))
                f.write(struct.pack("<fff", *v2))
                f.write(struct.pack("<H", 0))
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
    rels_path = "_rels/.rels"
    if rels_path in zf.namelist():
        with zf.open(rels_path) as f:
            root = ET.parse(f).getroot()
        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        for rel in root.findall(f"{{{rels_ns}}}Relationship"):
            if "3dmanufacturing" in rel.get("Type", ""):
                target = rel.get("Target", "").lstrip("/")
                if target:
                    return target
    for name in zf.namelist():
        if name.endswith(".model"):
            return name
    raise ValueError("could not find 3D model document in 3MF archive")


def _read_3mf_shape(path: Path):
    verts_all: list = []
    tris_all: list = []

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
        tris = [
            (int(t.get("v1")), int(t.get("v2")), int(t.get("v3")))
            for t in tris_el.findall(f"{{{_3MF_NS}}}triangle")
        ]
        for t in tris:
            tris_all.append((t[0] + offset, t[1] + offset, t[2] + offset))
        verts_all.extend(verts)
        offset += len(verts)

    if not tris_all:
        raise ValueError("no triangle data found in 3MF file")
    return _mesh_to_shape(verts_all, tris_all)


def _read_obj_shape(path: Path):
    verts: list = []
    tris: list = []
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
    return _mesh_to_shape(verts, tris)


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

    verts_all: list = []
    tris_all: list = []
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
    return _mesh_to_shape(verts_all, tris_all)


def _read_iges_shape(path: Path):
    reader = IGESControl_Reader()
    with quiet():
        status = reader.ReadFile(path.as_posix())
    if status != IFSelect_RetDone:
        raise ValueError(f"IGES reader failed with status {status}")
    with quiet():
        reader.TransferRoots()
    shape = reader.OneShape()
    if shape.IsNull():
        raise ValueError("IGES file produced an empty shape")
    return shape


try:
    with quiet():
        from OCC.Core.StlAPI import StlAPI_Reader
        from OCC.Core.BRepBuilderAPI import BRepBuilderAPI_Sewing
        from OCC.Core.ShapeUpgrade import ShapeUpgrade_UnifySameDomain
        from OCC.Core.ShapeFix import ShapeFix_Shape
        from OCC.Core.STEPControl import STEPControl_Writer, STEPControl_AsIs
        from OCC.Core.IGESControl import IGESControl_Reader
        from OCC.Core.IFSelect import IFSelect_RetDone, IFSelect_RetError, IFSelect_RetFail
        from OCC.Core.TopoDS import TopoDS_Shape
except Exception as e:
    print(f"\n  {R}[ERROR]{X} Failed to load OpenCASCADE: {e}\n")
    traceback.print_exc()
    input("\n  Press Enter to exit...")
    sys.exit(1)



def convert(input_path: Path, output_path: Path, tolerance: float = DEFAULT_TOLERANCE):
    src = input_path.as_posix()
    dst = output_path.as_posix()

    try:
        _dot("reading")
        ext = input_path.suffix.lower()
        if ext == TMF_EXT:
            shape = _read_3mf_shape(input_path)
        elif ext == OBJ_EXT:
            shape = _read_obj_shape(input_path)
        elif ext == AMF_EXT:
            shape = _read_amf_shape(input_path)
        elif ext in {IGS_EXT, ".iges"}:
            shape = _read_iges_shape(input_path)
        else:
            shape = TopoDS_Shape()
            with quiet():
                StlAPI_Reader().Read(shape, src)
        if shape.IsNull():
            return False, "input produced an empty shape"

        _dot("sewing")
        sew = BRepBuilderAPI_Sewing(tolerance)
        sew.Add(shape)
        with quiet():
            sew.Perform()
        sewn = sew.SewedShape()
        if sewn.IsNull():
            return False, "sewing failed - try a larger tolerance"

        _dot("fixing")
        try:
            with quiet():
                fix = ShapeFix_Shape(sewn)
                fix.Perform()
            fixed = fix.Shape()
            if fixed.IsNull():
                fixed = sewn
        except Exception:
            fixed = sewn

        _dot("refining")
        try:
            with quiet():
                u = ShapeUpgrade_UnifySameDomain(fixed, True, True, True)
                u.SetLinearTolerance(tolerance)
                u.SetAngularTolerance(ANGULAR_TOLERANCE)
                u.Build()
            refined = u.Shape()
            if refined.IsNull():
                raise RuntimeError
        except Exception:
            refined = fixed

        _dot("writing")
        writer = STEPControl_Writer()
        with quiet():
            ts = writer.Transfer(refined, STEPControl_AsIs)
            ws = writer.Write(dst)
        if ts in (IFSelect_RetError, IFSelect_RetFail) or \
           ws in (IFSelect_RetError, IFSelect_RetFail):
            return False, "STEP writer failed"

        if not output_path.exists() or output_path.stat().st_size == 0:
            return False, "output file is missing or empty"

        return True, output_path.stat().st_size // BYTES_PER_KB

    except Exception:
        return False, traceback.format_exc().strip()


def _dot(label: str):
    print(f"{DIM}{label}{X}", end="  ", flush=True)


def _trim(name: str, width: int = NAME_TRIM_WIDTH) -> str:
    return name if len(name) <= width else name[:width - 3] + "..."


def models_dir() -> Path:
    return Path(__file__).parent / MODELS_DIR_NAME


def main():
    parser = argparse.ArgumentParser(description="Convert to STEP.")
    parser.add_argument("input",  nargs="?")
    parser.add_argument("output", nargs="?")
    parser.add_argument("--tolerance", "-t", type=float, default=DEFAULT_TOLERANCE)
    args = parser.parse_args()

    print()
    print(f"  {C}{B}╔══════════════════════════════╗{X}")
    print(f"  {C}{B}║       2STEP-Converter        ║{X}")
    print(f"  {C}{B}╚══════════════════════════════╝{X}")
    print()

    if args.input is None:
        folder = models_dir()
        folder.mkdir(exist_ok=True)

        _supported = {STL_EXT, TMF_EXT, OBJ_EXT, AMF_EXT, IGS_EXT, ".iges"}
        files = sorted(f for f in folder.iterdir() if f.suffix.lower() in _supported)
        if not files:
            print(f"  No supported files found in {MODELS_DIR_NAME}\\\n")
            input("  Press Enter to exit...")
            sys.exit(0)

        n = len(files)
        print(f"  {n} file{'s' if n > 1 else ''} found in {C}{MODELS_DIR_NAME}\\{X}\n")

        ok_n = fail_n = 0
        for i, src_file in enumerate(files, 1):
            out_file = src_file.with_suffix(STP_EXT)
            src_kb = src_file.stat().st_size // BYTES_PER_KB
            print(f"  {B}[{i}/{n}]{X}  {_trim(src_file.name)}  {DIM}{src_kb} KB{X}")
            print(f"         ", end="", flush=True)
            success, info = convert(src_file, out_file, args.tolerance)
            print()
            if success:
                print(f"         {G}{_trim(out_file.name)}  {info} KB{X}")
                ok_n += 1
            else:
                print(f"         {R}✗  {info}{X}")
                fail_n += 1
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
    print(f"  {B}[1/1]{X}  {_trim(input_path.name)}  {DIM}{src_kb} KB{X}")
    print(f"         ", end="", flush=True)
    success, info = convert(input_path, output_path, args.tolerance)
    print()
    if success:
        print(f"         {G}{_trim(output_path.name)}  {info} KB{X}")
        print()
        print(f"  {G}{B}✓  Done{X}")
    else:
        print(f"         {R}✗  {info}{X}")
    print()
    input("  Press Enter to exit...")
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
