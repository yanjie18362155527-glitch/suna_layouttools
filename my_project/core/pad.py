import io

import gdstk
import matplotlib
import pandas as pd


matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPolygon

from .base import BaseEngine


class PadEngine(BaseEngine):
    def extract_pads(
        self,
        gds_path,
        cell_name: str,
        layer: int,
        datatype: int,
        aux_layer: int | None = None,
        aux_datatype: int | None = None,
    ):
        lib = gdstk.read_gds(str(gds_path))
        cell_map = {cell.name: cell for cell in lib.cells}

        if cell_name not in cell_map:
            raise ValueError(f"Cell '{cell_name}' 未找到")

        flat_cell = cell_map[cell_name].flatten()
        bbox = flat_cell.bounding_box()
        if bbox is not None:
            min_x, min_y = bbox[0]
            dx, dy = -min_x, -min_y
            for polygon in flat_cell.polygons:
                polygon.translate(dx, dy)
            for path in flat_cell.paths:
                path.translate(dx, dy)
            for label in flat_cell.labels:
                label.origin = (label.origin[0] + dx, label.origin[1] + dy)

        pads = []
        aux_polygons = []
        for polygon in flat_cell.polygons:
            if polygon.layer == layer and polygon.datatype == datatype:
                bbox = polygon.bounding_box()
                if bbox:
                    min_x, min_y = bbox[0]
                    max_x, max_y = bbox[1]
                    pads.append(
                        {
                            "data": {
                                "cx": (min_x + max_x) / 2,
                                "cy": (min_y + max_y) / 2,
                                "w": max_x - min_x,
                                "h": max_y - min_y,
                            },
                            "points": polygon.points,
                        }
                    )
            elif (
                aux_layer is not None
                and aux_datatype is not None
                and polygon.layer == aux_layer
                and polygon.datatype == aux_datatype
            ):
                aux_polygons.append(polygon.points)

        return pads, aux_polygons

    def generate_preview(self, pads, aux_polygons):
        fig, ax = plt.subplots(figsize=(10, 10))
        min_x = float("inf")
        min_y = float("inf")
        max_x = float("-inf")
        max_y = float("-inf")

        for points in aux_polygons or []:
            ax.add_patch(MplPolygon(points, closed=True, facecolor="black", edgecolor="none", alpha=0.15))
            for x, y in points:
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

        for item in pads:
            ax.add_patch(
                MplPolygon(
                    item["points"],
                    closed=True,
                    facecolor="red",
                    edgecolor="darkred",
                    alpha=0.3,
                    linewidth=0.5,
                )
            )
            for x, y in item["points"]:
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

        if min_x == float("inf"):
            min_x, min_y, max_x, max_y = 0, 0, 100, 100

        width = max_x - min_x
        height = max_y - min_y
        min_x -= width * 0.05
        max_x += width * 0.05
        min_y -= height * 0.05
        max_y += height * 0.05

        ax.set_xlim(min_x, max_x)
        ax.set_ylim(min_y, max_y)
        ax.set_aspect("equal")
        ax.axis("off")
        plt.subplots_adjust(left=0, right=1, top=1, bottom=0)

        buffer = io.BytesIO()
        plt.savefig(buffer, format="png")
        plt.close(fig)
        buffer.seek(0)

        map_info = {
            "bbox": (min_x, min_y, max_x, max_y),
            "gds_w": max_x - min_x,
            "gds_h": max_y - min_y,
        }
        return buffer, map_info

    def generate_report(self, ordered_pads, aux_polygons, output_path):
        if not ordered_pads:
            raise ValueError("没有可导出的 pad 数据")

        data_list = [item["data"] for item in ordered_pads]
        polygon_list = [item["points"] for item in ordered_pads]
        image_buffer = self._generate_plot(data_list, polygon_list, aux_polygons)
        self._write_excel(data_list, output_path, image_buffer)
        return len(data_list)

    def _generate_plot(self, data_list, polygons, aux_polygons=None):
        fig, ax = plt.subplots(figsize=(10, 10))

        for points in aux_polygons or []:
            ax.add_patch(MplPolygon(points, closed=True, facecolor="black", edgecolor="none", alpha=0.15))

        for points in polygons:
            ax.add_patch(
                MplPolygon(points, closed=True, facecolor="red", edgecolor="darkred", alpha=0.3, linewidth=0.5)
            )

        xs = [item["cx"] for item in data_list]
        ys = [item["cy"] for item in data_list]
        ax.scatter(xs, ys, c="red", s=60, marker="o", zorder=10)

        if len(data_list) < 3000:
            for index, (x, y) in enumerate(zip(xs, ys), start=1):
                ax.text(x, y, str(index), fontsize=16, color="black", weight="bold", zorder=15)

        ax.set_title(f"Pad Inspection (Count: {len(data_list)})")
        ax.set_aspect("equal")
        ax.autoscale_view()
        ax.grid(True, linestyle="--", alpha=0.3)

        buffer = io.BytesIO()
        plt.savefig(buffer, format="png")
        plt.close(fig)
        buffer.seek(0)
        return buffer

    def _write_excel(self, data_list, output_path, img_buffer):
        dataframe = pd.DataFrame(data_list)
        dataframe.index = dataframe.index + 1
        dataframe.index.name = "Index"
        try:
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                dataframe.to_excel(writer, sheet_name="Sheet1")
                worksheet = writer.sheets["Sheet1"]
                worksheet.insert_image(
                    1,
                    8,
                    "pad_view.png",
                    {"image_data": img_buffer, "x_scale": 0.8, "y_scale": 0.8},
                )
        except PermissionError as exc:
            raise IOError(f"无法写入文件 '{output_path}'，请先关闭已打开的 Excel 文件") from exc
