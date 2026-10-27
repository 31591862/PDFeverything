"""启动加速 — 阻止 PyMuPDF 2.x 在 import 时急切加载 AI 分析栈。

PyMuPDF 2.0 的 ``__init__`` 会无条件执行 ``import pymupdf.layout`` 和
``import pymupdf4llm``（用于 to_markdown 等可选功能）。这条链会把 numpy、
onnxruntime、tabulate、yaml、multiprocessing 等全部拖进来，实测约 1.2 s ——
而本项目只用 PyMuPDF 的核心 API（渲染 / 文本 / 表格 / 页面编辑），从不调用
to_markdown。把这两个模块拦下后，GUI 模块链的导入时间从 ~1.4 s 降到 ~0.2 s。

拦截方式是插入一个 meta_path finder，抛 ModuleNotFoundError（与"未安装"无法
区分，因此 pymupdf 自己的 ``try/except ImportError`` 会干净降级：
``pymupdf._get_layout`` 保持 None，``Document.to_markdown`` 抛 NotImplementedError）。
pymupdf 自带的 pymupdf.table / pymupdf.extra 等子模块不受影响，find_tables()
仍可用。若将来需要 to_markdown，调用 :func:`unblock` 即可恢复。

必须在**任何** ``import pymupdf`` / ``import fitz`` 之前调用 :func:`install`。
"""

import sys
from importlib.abc import MetaPathFinder

_BLOCKED_PREFIXES = (
    "pymupdf.layout",
    "pymupdf4llm",
)

_BLOCKED_EXACT = {"pymupdf.layout", "pymupdf4llm"}

_finder: "MetaPathFinder | None" = None


class _LayoutBlocker(MetaPathFinder):
    """Make the optional AI-analysis modules look "not installed"."""

    def find_spec(self, fullname, path=None, target=None):
        if fullname in _BLOCKED_EXACT or any(
                fullname.startswith(p + ".") for p in _BLOCKED_PREFIXES):
            raise ModuleNotFoundError(
                f"{fullname} is disabled to keep PDFeverything startup fast",
                name=fullname)
        return None


def install() -> None:
    """Idempotently install the blocker (no-op when already installed)."""
    global _finder
    if _finder is None:
        _finder = _LayoutBlocker()
        sys.meta_path.insert(0, _finder)
    # The project's code uses the historic `import fitz` alias throughout.
    # PyMuPDF 2.x turns every such import into a FutureWarning on stderr;
    # on the MCP stdio channel stderr noise is also undesirable. Silence just
    # that one category — other warnings keep working.
    import warnings

    warnings.filterwarnings(
        "ignore", category=FutureWarning,
        message=r"The `fitz` API is deprecated.*")


def unblock() -> None:
    """Remove the blocker so lazily-imported heavy modules work again."""
    global _finder
    if _finder is not None:
        try:
            sys.meta_path.remove(_finder)
        except ValueError:
            pass
        _finder = None
