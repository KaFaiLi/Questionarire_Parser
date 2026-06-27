# tests/test_longpath.py
# Long-path helpers: no-op on posix; on Windows they apply the \\?\ extended-length
# prefix so paths over MAX_PATH (~260 chars) don't raise "file not found".
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_kyd_json as akj


def test_posix_noop():
    assert akj.longpath("/a/b/c") == "/a/b/c"
    assert akj.longpath(Path("/a/b/c")) == "/a/b/c"
    print("OK posix no-op")


def test_windows_prefixing():
    orig_name, orig_abspath = os.name, os.path.abspath
    try:
        os.name = "nt"
        # fake abspath: leave drive/UNC paths, prepend a cwd otherwise
        os.path.abspath = lambda s: (
            s if (len(s) > 1 and s[1] == ":") or s.startswith("\\\\")
            else "C:\\cwd\\" + s)
        assert akj.longpath(r"C:\very\long\path") == r"\\?\C:\very\long\path"
        assert akj.longpath(r"\\server\share\f") == r"\\?\UNC\server\share\f"
        assert akj.longpath(r"rel\path") == r"\\?\C:\cwd\rel\path"
        assert akj.longpath(r"\\?\C:\already") == r"\\?\C:\already"  # idempotent
    finally:
        os.name, os.path.abspath = orig_name, orig_abspath
    print("OK windows prefixing")


def test_roundtrip_read_write(tmp=None):
    import tempfile
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "sub" / "deep" / "f.txt"   # parent does not exist yet
        akj.write_text(p, "hello")               # creates parents
        assert akj.path_exists(p)
        assert akj.read_text(p) == "hello"
    print("OK read/write roundtrip")


if __name__ == "__main__":
    test_posix_noop()
    test_windows_prefixing()
    test_roundtrip_read_write()
