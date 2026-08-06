"""civil3d-mcp — FastMCP server for Autodesk Civil 3D via COM automation."""

__version__ = "1.0.0"
__all__ = ["Civil3DClient", "Civil3DError"]

from .client import Civil3DClient, Civil3DError
