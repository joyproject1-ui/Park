"""Download EU (EC/EMA) and US FDA regulations on AI in GMP-regulated operations."""

from .catalog import Catalog, CatalogError, Document, load_catalog

__version__ = "1.0.0"
__all__ = ["Catalog", "CatalogError", "Document", "load_catalog", "__version__"]
