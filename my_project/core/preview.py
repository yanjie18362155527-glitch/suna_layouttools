import io

import matplotlib


matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPolygon
from matplotlib.patches import Rectangle as MplRect


class GDSPreviewer:
    @staticmethod
    def render_cell_to_image(cell, width_px: int = 800, height_px: int = 600, view_bbox=None):
        full_bbox = cell.bounding_box()
        if full_bbox is None:
            return None, None

        if view_bbox:
            min_x, min_y, max_x, max_y = view_bbox
        else:
            min_x, min_y = full_bbox[0]
            max_x, max_y = full_bbox[1]
            width = max_x - min_x
            height = max_y - min_y
            min_x -= width * 0.05
            max_x += width * 0.05
            min_y -= height * 0.05
            max_y += height * 0.05

        fig, ax = plt.subplots(figsize=(width_px / 100, height_px / 100), dpi=100)

        poly_count = 0
        for polygon in cell.polygons:
            bbox = polygon.bounding_box()
            if not bbox:
                continue
            if bbox[1][0] > min_x and bbox[0][0] < max_x and bbox[1][1] > min_y and bbox[0][1] < max_y:
                ax.add_patch(
                    MplPolygon(
                        polygon.points,
                        closed=True,
                        facecolor="#66b3ff",
                        edgecolor="none",
                        alpha=0.5,
                    )
                )
                poly_count += 1
                if poly_count >= 5000:
                    break

        ref_count = 0
        for reference in cell.references:
            bbox = reference.bounding_box()
            if not bbox:
                continue
            if bbox[1][0] > min_x and bbox[0][0] < max_x and bbox[1][1] > min_y and bbox[0][1] < max_y:
                rx, ry = bbox[0]
                width = bbox[1][0] - rx
                height = bbox[1][1] - ry
                ax.add_patch(
                    MplRect(
                        (rx, ry),
                        width,
                        height,
                        linewidth=1,
                        edgecolor="#ffa500",
                        facecolor="none",
                        linestyle="--",
                    )
                )
                ref_count += 1
                if ref_count >= 5000:
                    break

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

    @staticmethod
    def calculate_coords(canvas_result, map_info, canvas_width: int, canvas_height: int):
        if not canvas_result.json_data:
            return None

        objects = canvas_result.json_data["objects"]
        if not objects:
            return None

        obj = objects[-1]
        x_px, y_px = obj["left"], obj["top"]
        w_px, h_px = obj["width"], obj["height"]

        gds_min_x, _, _, gds_max_y = map_info["bbox"]
        gds_w, gds_h = map_info["gds_w"], map_info["gds_h"]

        img_aspect = canvas_width / canvas_height
        gds_aspect = gds_w / gds_h

        if gds_aspect > img_aspect:
            scale = gds_w / canvas_width
            valid_h_px = gds_h / scale
            x_offset_px = 0
            y_offset_px = (canvas_height - valid_h_px) / 2
        else:
            scale = gds_h / canvas_height
            valid_w_px = gds_w / scale
            x_offset_px = (canvas_width - valid_w_px) / 2
            y_offset_px = 0

        center_x_px = x_px + w_px / 2
        center_y_px = y_px + h_px / 2

        gds_center_x = gds_min_x + (center_x_px - x_offset_px) * scale
        gds_center_y = gds_max_y - (center_y_px - y_offset_px) * scale

        gds_w_real = w_px * scale
        gds_h_real = h_px * scale

        rect_bbox = (
            gds_center_x - gds_w_real / 2,
            gds_center_y - gds_h_real / 2,
            gds_center_x + gds_w_real / 2,
            gds_center_y + gds_h_real / 2,
        )
        return gds_center_x, gds_center_y, gds_w_real, gds_h_real, rect_bbox

    @staticmethod
    def map_point(x_px: float, y_px: float, map_info, canvas_width: int, canvas_height: int):
        gds_min_x, _, _, gds_max_y = map_info["bbox"]
        gds_w, gds_h = map_info["gds_w"], map_info["gds_h"]

        img_aspect = canvas_width / canvas_height
        gds_aspect = gds_w / gds_h

        if gds_aspect > img_aspect:
            scale = gds_w / canvas_width
            valid_h_px = gds_h / scale
            x_offset_px = 0
            y_offset_px = (canvas_height - valid_h_px) / 2
        else:
            scale = gds_h / canvas_height
            valid_w_px = gds_w / scale
            x_offset_px = (canvas_width - valid_w_px) / 2
            y_offset_px = 0

        gds_x = gds_min_x + (x_px - x_offset_px) * scale
        gds_y = gds_max_y - (y_px - y_offset_px) * scale
        return gds_x, gds_y
