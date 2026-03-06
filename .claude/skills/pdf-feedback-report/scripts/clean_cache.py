"""
清理 pdf-feedback-report/cache 目录：删除所有 .txt 和 .md 文件，仅保留 .csv。
应在报告生成并发布后执行，便于下次执行时 cache 仅保留数据文件。
"""
import os
import sys


def _ensure_utf8_stdio():
    """在 Windows 下将 stdout/stderr 设为 UTF-8，避免打印中文或 emoji 时 UnicodeEncodeError。"""
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            if hasattr(stream, "reconfigure"):
                try:
                    stream.reconfigure(encoding="utf-8", errors="replace")
                except (AttributeError, OSError):
                    pass


def main():
    _ensure_utf8_stdio()
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    cache_dir = os.path.join(base_dir, "cache")
    if not os.path.isdir(cache_dir):
        return 0
    removed = []
    for name in os.listdir(cache_dir):
        if name.startswith("."):
            continue
        lower = name.lower()
        if lower.endswith(".txt") or lower.endswith(".md"):
            path = os.path.join(cache_dir, name)
            try:
                os.remove(path)
                removed.append(name)
            except OSError as e:
                print(f"⚠️ 删除失败 {name}: {e}", file=sys.stderr)
    if removed:
        print(f"✅ 已清理 cache：删除 {len(removed)} 个文件（.txt/.md），仅保留 .csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
