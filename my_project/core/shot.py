import gdstk

from .base import BaseEngine
from .helpers import transform_reference_point


class ShotEngine(BaseEngine):
    def process(self, parent_name: str, child_name: str, text_anchor, text_area, layer: int, datatype: int, out_path):
        if parent_name not in self.cells_map:
            raise ValueError(f"{parent_name} 不存在")

        parent_cell = self.cells_map[parent_name]
        instances = []
        for reference in parent_cell.references:
            reference_name = reference.cell.name if isinstance(reference.cell, gdstk.Cell) else reference.cell
            if reference_name == child_name:
                sort_x, sort_y = transform_reference_point((0, 0), reference)
                text_x, text_y = transform_reference_point(text_anchor, reference)
                instances.append(
                    {
                        "ref": reference,
                        "sort_x": sort_x,
                        "sort_y": sort_y,
                        "text_x": text_x,
                        "text_y": text_y,
                    }
                )

        if not instances:
            raise ValueError("未找到匹配的子 cell 实例")

        all_x = [item["sort_x"] for item in instances]
        all_y = [item["sort_y"] for item in instances]
        average_x = sum(all_x) / len(all_x)
        average_y = sum(all_y) / len(all_y)
        center_instance = min(
            instances,
            key=lambda item: (item["sort_x"] - average_x) ** 2 + (item["sort_y"] - average_y) ** 2,
        )

        precision = 3
        unique_x = sorted({round(value, precision) for value in all_x})
        unique_y = sorted({round(value, precision) for value in all_y})
        base_x = unique_x.index(round(center_instance["sort_x"], precision))
        base_y = unique_y.index(round(center_instance["sort_y"], precision))

        label_width, label_height = text_area
        char_aspect = 0.6
        for instance in instances:
            index_x = unique_x.index(round(instance["sort_x"], precision)) - base_x
            index_y = unique_y.index(round(instance["sort_y"], precision)) - base_y
            label = f"({index_x},{index_y})"
            magnification = instance["ref"].magnification or 1.0
            text_size = min(label_height * 0.9, (label_width * 0.9) / (len(label) * char_aspect)) * magnification
            text_polygons = gdstk.text(label, text_size, (0, 0), layer=layer, datatype=datatype)
            temp_cell = gdstk.Cell("temp")
            temp_cell.add(*text_polygons)
            bbox = temp_cell.bounding_box()
            if bbox:
                center_x = (bbox[0][0] + bbox[1][0]) / 2
                center_y = (bbox[0][1] + bbox[1][1]) / 2
                dx = instance["text_x"] - center_x
                dy = instance["text_y"] - center_y
                for polygon in text_polygons:
                    polygon.translate(dx, dy)
                    parent_cell.add(polygon)

        self.save_lib(out_path)
