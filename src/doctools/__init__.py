from importlib.metadata import version, PackageNotFoundError

try:
    __version__ = version("doctools-mcp")
except PackageNotFoundError:
    __version__ = "unknown"
