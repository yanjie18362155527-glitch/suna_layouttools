import gdstk

from .base import BaseEngine
from .helpers import collect_unique_coordinates


class LensEngine(BaseEngine):
    def process(
        self,
        parent_name: str,
        child_name: str,
        layer: int,
        datatype: int,
        size: float,
        offset: tuple[float, float],
        tolerance: float,
        mode: str,
        sort_dir: str,
        out_path,
        digit_width: int = 4,
    ) -> None:
        if parent_name not in self.cells_map or child_name not in self.cells_map:
            raise ValueError("未找到指定的 cell")

        parent_cell = self.cells_map[parent_name]
        child_cell = self.cells_map[child_name]
        bbox = child_cell.bounding_box()
        local_center = ((bbox[0][0] + bbox[1][0]) / 2, (bbox[0][1] + bbox[1][1]) / 2) if bbox else (0, 0)

        instances = []
        for reference in parent_cell.references:
            reference_name = reference.cell.name if isinstance(reference.cell, gdstk.Cell) else reference.cell
            if reference_name == child_name:
                abs_x = reference.origin[0] + local_center[0]
                abs_y = reference.origin[1] + local_center[1]
                instances.append({"x": abs_x, "y": abs_y, "label": ""})

        if not instances:
            raise ValueError("未找到目标实例")

        if mode == "row_col":
            unique_x = collect_unique_coordinates(sorted(item["x"] for item in instances), tolerance)
            unique_y = collect_unique_coordinates(sorted(item["y"] for item in instances), tolerance)
            for instance in instances:
                row_index = next(
                    (index + 1 for index, value in enumerate(unique_y) if abs(instance["y"] - value) <= tolerance),
                    -1,
                )
                col_index = next(
                    (index + 1 for index, value in enumerate(unique_x) if abs(instance["x"] - value) <= tolerance),
                    -1,
                )
                instance["label"] = f"{row_index}-{col_index}"
        else:
            if sort_dir == "y_first":
                instances.sort(key=lambda item: (round(item["y"] / tolerance), item["x"]))
            else:
                instances.sort(key=lambda item: (round(item["x"] / tolerance), item["y"]))
            for index, instance in enumerate(instances, start=1):
                instance["label"] = f"{index:0{digit_width}d}"

        for instance in instances:
            target_x = instance["x"] + offset[0]
            target_y = instance["y"] + offset[1]
            text_polygons = gdstk.text(instance["label"], size=size, position=(0, 0), layer=layer, datatype=datatype)

            min_tx = float("inf")
            min_ty = float("inf")
            max_tx = float("-inf")
            max_ty = float("-inf")
            valid = False
            for polygon in text_polygons:
                bbox = polygon.bounding_box()
                if bbox:
                    valid = True
                    min_tx = min(min_tx, bbox[0][0])
                    min_ty = min(min_ty, bbox[0][1])
                    max_tx = max(max_tx, bbox[1][0])
                    max_ty = max(max_ty, bbox[1][1])

            if not valid:
                continue

            shift_x = target_x - (min_tx + max_tx) / 2
            shift_y = target_y - (min_ty + max_ty) / 2
            for polygon in text_polygons:
                polygon.translate(shift_x, shift_y)
                parent_cell.add(polygon)

        self.save_lib(out_path)
