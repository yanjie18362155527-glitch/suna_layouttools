import io

import gdstk
import matplotlib
import pandas as pd


matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import ConnectionPatch
from matplotlib.patches import Rectangle as MplRect

from .base import BaseEngine


class CellInfoEngine(BaseEngine):
    def get_child_names(self, parent_name: str) -> list[str]:
        if parent_name not in self.cells_map:
            return []

        parent_cell = self.cells_map[parent_name]
        return sorted(
            {
                reference.cell.name if isinstance(reference.cell, gdstk.Cell) else reference.cell
                for reference in parent_cell.references
            }
        )

    def process(self, parent_name: str, target_children: list[str], output_path, temp_img_path=None):
        del temp_img_path

        if parent_name not in self.cells_map:
            raise ValueError("Parent cell 不存在")

        parent_cell = self.cells_map[parent_name]
        parent_bbox = parent_cell.bounding_box() or [[0, 0], [0, 0]]

        all_instances = []
        for reference in parent_cell.references:
            reference_name = reference.cell.name if isinstance(reference.cell, gdstk.Cell) else reference.cell
            if reference_name in target_children:
                bbox = reference.bounding_box()
                if bbox is None:
                    continue
                center_x = (bbox[0][0] + bbox[1][0]) / 2
                center_y = (bbox[0][1] + bbox[1][1]) / 2
                all_instances.append(
                    {
                        "CellName": reference_name,
                        "Center_X": center_x,
                        "Center_Y": center_y,
                        "Width": bbox[1][0] - bbox[0][0],
                        "Height": bbox[1][1] - bbox[0][1],
                    }
                )

        all_instances.sort(key=lambda item: (-round(item["Center_Y"], 3), item["Center_X"]))

        seen_names = set()
        data_list = []
        for item in all_instances:
            if item["CellName"] not in seen_names:
                seen_names.add(item["CellName"])
                data_list.append(item)

        if not data_list:
            raise ValueError("未找到指定子 cell 实例")

        plot_items = []
        for index, item in enumerate(data_list, start=1):
            item["Index"] = index
            plot_items.append(
                {
                    "id": index,
                    "name": item["CellName"],
                    "cx": item["Center_X"],
                    "cy": item["Center_Y"],
                    "w": item["Width"],
                    "h": item["Height"],
                }
            )

        row_height_pt = 30
        image_buffer = self._generate_aligned_plot(plot_items, parent_bbox, row_height_pt)
        self._write_excel(data_list, output_path, image_buffer, row_height_pt)
        return len(data_list)

    def _generate_aligned_plot(self, plot_items, parent_bbox, row_height_pt):
        row_count = len(plot_items)
        dpi = 100
        data_height_inch = (row_count * row_height_pt) / 72.0
        axis_margin_inch = 0.8
        total_height_inch = data_height_inch + axis_margin_inch

        fig = plt.figure(figsize=(12, total_height_inch), dpi=dpi)
        bottom_ratio = axis_margin_inch / total_height_inch
        plt.subplots_adjust(left=0, right=0.92, top=1.0, bottom=bottom_ratio, wspace=0.0)

        grid = fig.add_gridspec(1, 2, width_ratios=[1, 25])
        ax_left = fig.add_subplot(grid[0])
        ax_layout = fig.add_subplot(grid[1])

        ax_left.set_axis_off()
        ax_left.set_ylim(row_count, 0)
        ax_left.set_xlim(0, 1)

        unique_names = list({item["name"] for item in plot_items})
        colors = plt.cm.get_cmap("tab10", len(unique_names))
        color_map = {name: colors(index) for index, name in enumerate(unique_names)}

        parent_min_x, parent_min_y = parent_bbox[0]
        parent_max_x, parent_max_y = parent_bbox[1]
        parent_width = parent_max_x - parent_min_x
        parent_height = parent_max_y - parent_min_y

        ax_layout.add_patch(
            MplRect(
                (parent_min_x, parent_min_y),
                parent_width,
                parent_height,
                linewidth=1.0,
                edgecolor="black",
                linestyle="--",
                facecolor="none",
            )
        )

        shrink_amount = 10.0
        for index, item in enumerate(plot_items):
            anchor_y = index + 0.5
            color = color_map[item["name"]]

            draw_width = item["w"] - 2 * shrink_amount if item["w"] > 2.5 * shrink_amount else item["w"]
            draw_height = item["h"] - 2 * shrink_amount if item["h"] > 2.5 * shrink_amount else item["h"]
            min_x = item["cx"] - draw_width / 2
            min_y = item["cy"] - draw_height / 2

            ax_layout.add_patch(
                MplRect((min_x, min_y), draw_width, draw_height, linewidth=2, edgecolor=color, facecolor="none")
            )
            ax_layout.scatter([item["cx"]], [item["cy"]], color=color, s=40, zorder=10)
            ax_left.scatter([0.9], [anchor_y], color=color, s=40, zorder=10)

            fig.add_artist(
                ConnectionPatch(
                    xyA=(item["cx"], item["cy"]),
                    coordsA=ax_layout.transData,
                    xyB=(0.9, anchor_y),
                    coordsB=ax_left.transData,
                    color=color,
                    arrowstyle="-",
                    linestyle="--",
                    linewidth=0.8,
                )
            )
            ax_layout.text(min_x, max(min_y, item["cy"]), item["name"], color=color, fontsize=8, alpha=0.7)

        ax_layout.set_aspect("equal")
        ax_layout.grid(True, which="both", linestyle=":", alpha=0.5, color="gray")
        ax_layout.set_xlabel("X Coordinate (um)", fontsize=10)
        ax_layout.set_ylabel("Y Coordinate (um)", fontsize=10)
        ax_layout.yaxis.tick_right()
        ax_layout.yaxis.set_label_position("right")

        margin_x = parent_width * 0.1 if parent_width > 0 else 100
        margin_y = parent_height * 0.1 if parent_height > 0 else 100
        ax_layout.set_xlim(parent_min_x - margin_x, parent_max_x + margin_x)
        ax_layout.set_ylim(parent_min_y - margin_y, parent_max_y + margin_y)

        buffer = io.BytesIO()
        plt.savefig(buffer, format="png", dpi=dpi)
        plt.close(fig)
        buffer.seek(0)
        return buffer

    def _write_excel(self, data, output_path, image_buffer, row_height_pt):
        dataframe = pd.DataFrame(data)
        dataframe = dataframe[["Index", "CellName", "Center_X", "Center_Y", "Width", "Height"]]

        try:
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                sheet_name = "CellInfo"
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]

                header_fmt = workbook.add_format(
                    {"bold": True, "bg_color": "#D7E4BC", "border": 1, "align": "center", "valign": "vcenter"}
                )
                cell_fmt = workbook.add_format({"align": "center", "valign": "vcenter"})

                for column_index, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, column_index, value, header_fmt)

                for row_index in range(len(data)):
                    worksheet.set_row(row_index + 1, row_height_pt, cell_fmt)

                worksheet.set_column("A:A", 6)
                worksheet.set_column("B:B", 25)
                worksheet.set_column("C:F", 12)
                worksheet.insert_image(
                    1,
                    7,
                    "plot.png",
                    {
                        "image_data": image_buffer,
                        "x_offset": 0,
                        "y_offset": 0,
                        "x_scale": 1.0,
                        "y_scale": 1.0,
                        "object_position": 1,
                    },
                )
                worksheet.write("G1", "Visual ->", header_fmt)
        except PermissionError as exc:
            raise IOError(f"无法写入文件 '{output_path}'，请先关闭已打开的 Excel 文件") from exc
