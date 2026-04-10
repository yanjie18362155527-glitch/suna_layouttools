import gdstk


class BaseEngine:
    """Shared GDS read/write helpers."""

    def __init__(self):
        self.lib = None
        self.cells_map: dict[str, gdstk.Cell] = {}

    def load_lib(self, path) -> list[str]:
        self.lib = gdstk.read_gds(str(path))
        self.cells_map = {cell.name: cell for cell in self.lib.cells}
        return sorted(self.cells_map.keys())

    def save_lib(self, path) -> None:
        if self.lib is not None:
            self.lib.write_gds(str(path))
