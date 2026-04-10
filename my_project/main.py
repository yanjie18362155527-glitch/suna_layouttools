import sys
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.elements.image as st_image
from PIL import Image


def _patch_streamlit_drawable_canvas_compat() -> None:
    if hasattr(st_image, "image_to_url"):
        return

    from streamlit.elements.lib.image_utils import image_to_url as modern_image_to_url
    from streamlit.elements.lib.layout_utils import LayoutConfig

    def legacy_image_to_url(image, width, clamp, channels, output_format, image_id):
        return modern_image_to_url(
            image=image,
            layout_config=LayoutConfig(width=width),
            clamp=clamp,
            channels=channels,
            output_format=output_format,
            image_id=image_id,
        )

    st_image.image_to_url = legacy_image_to_url


_patch_streamlit_drawable_canvas_compat()

from streamlit_drawable_canvas import st_canvas

if __package__ in {None, ""}:
    sys.path.append(str(Path(__file__).resolve().parent.parent))

from my_project.config import (
    APP_LAYOUT,
    APP_MODES,
    APP_TITLE,
    CELL_INFO_MODE,
    CELL_INFO_OUTPUT_NAME,
    ETCH_DUTY_DEFAULT_SHEET,
    ETCH_DUTY_DEFAULT_START_CELL,
    ETCH_DUTY_MODE,
    ETCH_DUTY_OUTPUT_NAME,
    FILM_VOLUME_DEFAULT_SHEET,
    FILM_VOLUME_DEFAULT_START_CELL,
    FILM_VOLUME_MODE,
    FILM_VOLUME_OUTPUT_NAME,
    LENS_DEFAULT_CHILD,
    LENS_DEFAULT_PARENT,
    LENS_DEFAULTS,
    LENS_MODE,
    LENS_OUTPUT_NAME,
    PAD_CANVAS_SIZE,
    PAD_DEFAULT_AUX_DATATYPE,
    PAD_DEFAULT_AUX_LAYER,
    PAD_DEFAULT_CELL,
    PAD_DEFAULT_DATATYPE,
    PAD_DEFAULT_LAYER,
    PAD_MODE,
    PAD_OUTPUT_NAME,
    PREVIEW_CANVAS_HEIGHT,
    PREVIEW_CANVAS_WIDTH,
    SHOT_DEFAULT_CHILD,
    SHOT_DEFAULT_PARENT,
    SHOT_DEFAULTS,
    SHOT_MODE,
    SHOT_OUTPUT_NAME,
    SUMMARY_LAYOUT_DEFAULT_STREET_WIDTH,
    SUMMARY_LAYOUT_DEFAULT_TOP,
    SUMMARY_LAYOUT_MODE,
    SUMMARY_LAYOUT_OUTPUT_NAME,
    USAGE_GUIDE_SECTIONS,
)
from my_project.core import (
    CellInfoEngine,
    EtchDutyEngine,
    FilmVolumeEngine,
    GDSPreviewer,
    LensEngine,
    PadEngine,
    ShotEngine,
    SummaryLayoutEngine,
)
from my_project.utils.common import build_temp_output_path, parse_int_set, parse_pair, save_uploaded_file


def configure_page() -> None:
    st.set_page_config(page_title=APP_TITLE, layout=APP_LAYOUT)


def initialize_state() -> None:
    for key, value in LENS_DEFAULTS.items():
        st.session_state.setdefault(f"lens_{key}", value)

    for key, value in SHOT_DEFAULTS.items():
        st.session_state.setdefault(f"shot_{key}", value)

    st.session_state.setdefault("pad_parsed_data", None)
    st.session_state.setdefault("pad_aux_data", None)
    st.session_state.setdefault("pad_map_info", None)
    st.session_state.setdefault("pad_img_buf", None)


def render_usage_guide() -> None:
    st.markdown(
        """
        <style>
        section[data-testid="stSidebar"] div[data-testid="stSidebarContent"] {
            display: flex;
            flex-direction: column;
            height: 100%;
        }
        section[data-testid="stSidebar"] div[data-testid="stSidebarContent"] > div:last-child {
            margin-top: auto;
            padding-top: 1rem;
            padding-bottom: 1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        with st.popover("使用说明", use_container_width=True):
            st.markdown("### 七大功能使用说明")
            for index, section in enumerate(USAGE_GUIDE_SECTIONS):
                st.markdown(f"#### {section['title']}")
                st.markdown("**操作步骤**")
                for step_index, step in enumerate(section["steps"], start=1):
                    st.markdown(f"{step_index}. {step}")
                st.markdown("**使用提示**")
                for tip in section["tips"]:
                    st.markdown(f"- {tip}")
                if index != len(USAGE_GUIDE_SECTIONS) - 1:
                    st.markdown("---")


def main() -> None:
    configure_page()
    initialize_state()

    st.sidebar.title("工具菜单")
    app_mode = st.sidebar.radio("选择功能:", APP_MODES)
    render_usage_guide()

    st.title(app_mode.split("\n", 1)[0])
    st.markdown("---")

    if app_mode == LENS_MODE:
        render_lens_page()
    elif app_mode == PAD_MODE:
        render_pad_page()
    elif app_mode == SHOT_MODE:
        render_shot_page()
    elif app_mode == CELL_INFO_MODE:
        render_cell_info_page()
    elif app_mode == SUMMARY_LAYOUT_MODE:
        render_summary_layout_page()
    elif app_mode == FILM_VOLUME_MODE:
        render_film_volume_page()
    elif app_mode == ETCH_DUTY_MODE:
        st.warning("备注：shot级 trench 图层占空比无法计算，会引起卡顿。")
        render_etch_duty_page()


def render_lens_page() -> None:
    uploaded_file = st.file_uploader("上传 GDS 文件", type=["gds"], key="lens_file")
    if not uploaded_file:
        return

    gds_path = save_uploaded_file(uploaded_file)
    engine = LensEngine()
    cell_names = engine.load_lib(gds_path)

    parent_index = cell_names.index(LENS_DEFAULT_PARENT) if LENS_DEFAULT_PARENT in cell_names else 0
    child_index = cell_names.index(LENS_DEFAULT_CHILD) if LENS_DEFAULT_CHILD in cell_names else 0

    col_left, col_right = st.columns(2)
    parent_cell = col_left.selectbox("父 Cell（阵列）", cell_names, index=parent_index)
    child_cell = col_right.selectbox("子 Cell（单元）", cell_names, index=child_index)

    st.subheader("参数配置")
    use_visual = st.checkbox("启用可视化框选", value=True)

    if use_visual:
        render_lens_selector(engine, child_cell)

    col_size, col_x, col_y = st.columns(3)
    size = col_size.number_input("字号（Size）", key="lens_size")
    offset_x = col_x.number_input("Offset X", key="lens_offset_x")
    offset_y = col_y.number_input("Offset Y", key="lens_offset_y")

    with st.expander("高级设置", expanded=True):
        col_mode, col_sort, col_digits = st.columns(3)
        mode_label = col_mode.selectbox("编号模式", ["行列坐标（Row-Col）", "顺序索引（1, 2...）"])
        mode = "row_col" if "Row-Col" in mode_label else "index"
        sort_disabled = mode == "row_col"
        sort_label = col_sort.selectbox("排序方向", ["Y 优先", "X 优先"], disabled=sort_disabled)
        sort_dir = "y_first" if sort_label.startswith("Y") else "x_first"
        digits = col_digits.number_input("位数（Digits）", value=4, min_value=1, disabled=sort_disabled)

        col_tol, col_layer, col_datatype = st.columns(3)
        tolerance = col_tol.number_input("容差", value=1.0)
        layer = col_layer.number_input("Layer", value=66, step=1)
        datatype = col_datatype.number_input("Datatype", value=0, step=1)

    if st.button("开始生成 GDS", type="primary"):
        output_path = build_temp_output_path(LENS_OUTPUT_NAME)
        try:
            engine.process(
                parent_cell,
                child_cell,
                int(layer),
                int(datatype),
                size,
                (offset_x, offset_y),
                tolerance,
                mode,
                sort_dir,
                output_path,
                int(digits),
            )
            st.success("Lens 自动编号完成")
            with output_path.open("rb") as file_obj:
                st.download_button("下载结果 GDS", file_obj, file_name="lens_output.gds")
        except Exception as exc:
            st.error(f"处理失败: {exc}")


def render_lens_selector(engine: LensEngine, child_cell: str) -> None:
    if st.session_state["lens_view_bbox"]:
        st.info("当前处于局部放大视图，仅显示当前视野内的本层图形和子 Cell 边框。")
    else:
        st.info("蓝色表示本层 polygon，橙色虚线表示子 Cell 边框。框选后可继续放大。")

    with st.spinner("正在渲染预览..."):
        image_buffer, map_info = GDSPreviewer.render_cell_to_image(
            engine.cells_map[child_cell],
            view_bbox=st.session_state["lens_view_bbox"],
        )

    if not image_buffer:
        st.warning("当前 Cell 没有可预览的几何数据。")
        return

    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.0)",
        stroke_color="#FF0000",
        stroke_width=2,
        background_image=Image.open(image_buffer),
        update_streamlit=True,
        height=PREVIEW_CANVAS_HEIGHT,
        width=PREVIEW_CANVAS_WIDTH,
        drawing_mode="rect",
        key=f"lens_canvas_{st.session_state['lens_zoom_version']}",
    )

    selected_bbox = None
    if canvas_result.json_data:
        coords = GDSPreviewer.calculate_coords(canvas_result, map_info, PREVIEW_CANVAS_WIDTH, PREVIEW_CANVAS_HEIGHT)
        if coords:
            center_x, center_y, _, height, selected_bbox = coords
            st.session_state["lens_offset_x"] = float(f"{center_x:.3f}")
            st.session_state["lens_offset_y"] = float(f"{center_y:.3f}")
            st.session_state["lens_size"] = float(f"{height:.3f}")
            st.success(f"当前选区中心: ({center_x:.2f}, {center_y:.2f})，推荐字号: {height:.2f}")

    col_zoom, col_reset = st.columns(2)
    if col_zoom.button("缩放到选区"):
        if selected_bbox:
            st.session_state["lens_view_bbox"] = selected_bbox
            st.session_state["lens_zoom_version"] += 1
            st.rerun()
        else:
            st.warning("请先框选一个区域。")

    if col_reset.button("重置视图"):
        st.session_state["lens_view_bbox"] = None
        st.session_state["lens_zoom_version"] += 1
        st.rerun()


def render_pad_page() -> None:
    uploaded_file = st.file_uploader("上传 GDS 文件", type=["gds"], key="pad_file")
    if not uploaded_file:
        return

    gds_path = save_uploaded_file(uploaded_file)
    engine = PadEngine()
    cell_names = engine.load_lib(gds_path)
    target_index = cell_names.index(PAD_DEFAULT_CELL) if PAD_DEFAULT_CELL in cell_names else 0
    target_cell = st.selectbox("选择目标 Cell", cell_names, index=target_index)

    st.markdown("##### 目标信息提取图层")
    col_layer, col_datatype = st.columns(2)
    layer = col_layer.number_input("Layer", value=PAD_DEFAULT_LAYER, step=1, key="pad_layer")
    datatype = col_datatype.number_input("Datatype", value=PAD_DEFAULT_DATATYPE, step=1, key="pad_datatype")

    st.markdown("##### 辅助可视化图层（可选）")
    use_auxiliary = st.checkbox("在 Excel 预览图中显示辅助图层")
    auxiliary_layer = None
    auxiliary_datatype = None
    if use_auxiliary:
        col_aux_layer, col_aux_datatype = st.columns(2)
        auxiliary_layer = col_aux_layer.number_input("辅助 Layer", value=PAD_DEFAULT_AUX_LAYER, step=1)
        auxiliary_datatype = col_aux_datatype.number_input("辅助 Datatype", value=PAD_DEFAULT_AUX_DATATYPE, step=1)

    st.markdown("---")
    if st.button("1. 解析 GDS（提取候选 Pad）"):
        with st.spinner("正在解析..."):
            try:
                pads, aux_polygons = engine.extract_pads(
                    gds_path,
                    target_cell,
                    int(layer),
                    int(datatype),
                    int(auxiliary_layer) if use_auxiliary else None,
                    int(auxiliary_datatype) if use_auxiliary else None,
                )
                if not pads:
                    st.error("未找到指定图层上的 Pad 数据，请检查 Layer 和 Datatype 设置。")
                else:
                    preview_buffer, map_info = engine.generate_preview(pads, aux_polygons)
                    st.session_state["pad_parsed_data"] = pads
                    st.session_state["pad_aux_data"] = aux_polygons
                    st.session_state["pad_img_buf"] = preview_buffer
                    st.session_state["pad_map_info"] = map_info
                    st.success(f"解析完成，共找到 {len(pads)} 个候选 Pad。")
            except Exception as exc:
                st.error(f"解析失败: {exc}")

    if st.session_state["pad_parsed_data"]:
        render_pad_selector()


def render_pad_selector() -> None:
    st.markdown("### 交互式排序")
    st.info("请按目标顺序逐个点击 Pad，点击顺序就是最终导出的编号顺序。")

    canvas_result = st_canvas(
        fill_color="rgba(255, 0, 0, 1)",
        stroke_color="rgba(255, 0, 0, 1)",
        stroke_width=3,
        background_image=Image.open(st.session_state["pad_img_buf"]),
        update_streamlit=True,
        height=PAD_CANVAS_SIZE,
        width=PAD_CANVAS_SIZE,
        drawing_mode="point",
        point_display_radius=5,
        key="pad_canvas",
    )

    ordered_pads = []
    if canvas_result.json_data and canvas_result.json_data.get("objects"):
        for obj in canvas_result.json_data["objects"]:
            if obj.get("type") not in {"circle", "rect", "path"}:
                continue
            radius = obj.get("radius", 0)
            x_px = obj["left"] + radius
            y_px = obj["top"] + radius
            gds_x, gds_y = GDSPreviewer.map_point(
                x_px,
                y_px,
                st.session_state["pad_map_info"],
                PAD_CANVAS_SIZE,
                PAD_CANVAS_SIZE,
            )
            closest_pad = min(
                st.session_state["pad_parsed_data"],
                key=lambda item: (item["data"]["cx"] - gds_x) ** 2 + (item["data"]["cy"] - gds_y) ** 2,
            )
            if closest_pad not in ordered_pads:
                ordered_pads.append(closest_pad)

    if ordered_pads:
        st.write(f"当前已选 Pad 数量: {len(ordered_pads)}")

    if st.button("2. 导出顺序提取报告", type="primary"):
        if not ordered_pads:
            st.warning("尚未选择任何 Pad。")
            return

        output_path = build_temp_output_path(PAD_OUTPUT_NAME)
        try:
            count = PadEngine().generate_report(ordered_pads, st.session_state["pad_aux_data"], output_path)
            st.success(f"导出完成，共输出 {count} 个目标。")
            with output_path.open("rb") as file_obj:
                st.download_button("下载 Excel 报告", file_obj, file_name="pad_report.xlsx")
        except Exception as exc:
            st.error(f"报告生成失败: {exc}")


def render_shot_page() -> None:
    uploaded_file = st.file_uploader("上传 GDS 文件", type=["gds"], key="shot_file")
    if not uploaded_file:
        return

    gds_path = save_uploaded_file(uploaded_file)
    engine = ShotEngine()
    cell_names = engine.load_lib(gds_path)

    parent_index = cell_names.index(SHOT_DEFAULT_PARENT) if SHOT_DEFAULT_PARENT in cell_names else 0
    child_index = cell_names.index(SHOT_DEFAULT_CHILD) if SHOT_DEFAULT_CHILD in cell_names else 0

    col_left, col_right = st.columns(2)
    top_cell = col_left.selectbox("Top Cell（父）", cell_names, index=parent_index)
    unit_cell = col_right.selectbox("Shot Cell（子）", cell_names, index=child_index)

    if st.checkbox("启用可视化框选", value=True):
        render_shot_selector(engine, unit_cell)

    col_anchor, col_area = st.columns(2)
    anchor_value = col_anchor.text_input("Anchor (x, y)", key="shot_anchor")
    area_value = col_area.text_input("Area (w, h)", key="shot_area")

    col_layer, col_datatype = st.columns(2)
    layer = col_layer.number_input("Layer", value=100, step=1)
    datatype = col_datatype.number_input("Datatype", value=0, step=1)

    if st.button("运行编号", type="primary"):
        output_path = build_temp_output_path(SHOT_OUTPUT_NAME)
        try:
            anchor = parse_pair(anchor_value, "Anchor")
            area = parse_pair(area_value, "Area")
            engine.process(top_cell, unit_cell, anchor, area, int(layer), int(datatype), output_path)
            st.success("Shot 自动编号完成")
            with output_path.open("rb") as file_obj:
                st.download_button("下载结果 GDS", file_obj, file_name="shot_output.gds")
        except Exception as exc:
            st.error(f"处理失败: {exc}")


def render_shot_selector(engine: ShotEngine, unit_cell: str) -> None:
    if st.session_state["shot_view_bbox"]:
        st.info("当前处于局部放大视图。")
    else:
        st.info("蓝色表示本层 polygon，橙色虚线表示子 Cell 边框。框选后可继续放大。")

    with st.spinner("正在渲染预览..."):
        image_buffer, map_info = GDSPreviewer.render_cell_to_image(
            engine.cells_map[unit_cell],
            view_bbox=st.session_state["shot_view_bbox"],
        )

    if not image_buffer:
        st.warning("当前 Cell 没有可预览的几何数据。")
        return

    canvas_result = st_canvas(
        fill_color="rgba(0, 255, 0, 0.0)",
        stroke_color="#00AA00",
        stroke_width=2,
        background_image=Image.open(image_buffer),
        update_streamlit=True,
        height=PREVIEW_CANVAS_HEIGHT,
        width=PREVIEW_CANVAS_WIDTH,
        drawing_mode="rect",
        key=f"shot_canvas_{st.session_state['shot_zoom_version']}",
    )

    selected_bbox = None
    if canvas_result.json_data:
        coords = GDSPreviewer.calculate_coords(canvas_result, map_info, PREVIEW_CANVAS_WIDTH, PREVIEW_CANVAS_HEIGHT)
        if coords:
            center_x, center_y, width, height, selected_bbox = coords
            st.session_state["shot_anchor"] = f"{center_x:.3f}, {center_y:.3f}"
            st.session_state["shot_area"] = f"{width:.3f}, {height:.3f}"
            st.success(f"已捕获 Anchor=({center_x:.2f}, {center_y:.2f}), Area={width:.2f} x {height:.2f}")

    col_zoom, col_reset = st.columns(2)
    if col_zoom.button("缩放到选区", key="btn_zoom_shot"):
        if selected_bbox:
            st.session_state["shot_view_bbox"] = selected_bbox
            st.session_state["shot_zoom_version"] += 1
            st.rerun()
        else:
            st.warning("请先框选一个区域。")

    if col_reset.button("重置视图", key="btn_reset_shot"):
        st.session_state["shot_view_bbox"] = None
        st.session_state["shot_zoom_version"] += 1
        st.rerun()


def render_cell_info_page() -> None:
    uploaded_file = st.file_uploader("上传 GDS 文件", type=["gds"], key="cell_info_file")
    if not uploaded_file:
        return

    gds_path = save_uploaded_file(uploaded_file)
    engine = CellInfoEngine()
    cell_names = engine.load_lib(gds_path)

    parent_cell = st.selectbox("选择 Top Cell", cell_names)
    children = engine.get_child_names(parent_cell)
    targets = st.multiselect("勾选需要提取的子 Cell", children)

    if st.button("生成报告", type="primary"):
        if not targets:
            st.warning("请至少选择一个子 Cell。")
            return

        output_path = build_temp_output_path(CELL_INFO_OUTPUT_NAME)
        try:
            count = engine.process(parent_cell, targets, output_path)
            st.success(f"处理完成，共导出 {count} 条记录。")
            with output_path.open("rb") as file_obj:
                st.download_button("下载 Excel", file_obj, file_name="cell_structure.xlsx")
        except Exception as exc:
            st.error(f"提取失败: {exc}")


def render_summary_layout_page() -> None:
    uploaded_file = st.file_uploader("上传 Excel 汇总表", type=["xlsx"], key="summary_layout_excel")
    if not uploaded_file:
        return

    workbook_path = save_uploaded_file(uploaded_file)
    engine = SummaryLayoutEngine()

    try:
        sheet_names = engine.list_sheet_names(workbook_path)
    except Exception as exc:
        st.error(f"读取 Excel 失败: {exc}")
        return

    default_sheet_index = sheet_names.index("tooling") if "tooling" in sheet_names else 0
    sheet_name = st.selectbox("选择用于排版的 Sheet", sheet_names, index=default_sheet_index)

    default_top_name = SummaryLayoutEngine.sanitize_cell_name(
        f"{Path(uploaded_file.name).stem}_{sheet_name or SUMMARY_LAYOUT_DEFAULT_TOP}"
    )
    uploaded_gds_files = st.file_uploader(
        "上传对应的 GDS 文件（可多选）",
        type=["gds"],
        accept_multiple_files=True,
        key="summary_layout_gds_files",
    )

    st.markdown("##### 排版参数")
    output_top_name = st.text_input(
        "输出 Top Cell 名称",
        value=default_top_name or SUMMARY_LAYOUT_DEFAULT_TOP,
        key=f"summary_layout_top_{sheet_name}",
    )
    street_width = st.number_input(
        "切割道宽度",
        value=float(SUMMARY_LAYOUT_DEFAULT_STREET_WIDTH),
        key=f"summary_layout_street_width_{sheet_name}",
        help="同时加在阵列列间距和行间距上，列 pitch = Cell 外框宽度 + 切割道宽度，行 pitch = Cell 外框高度 + 切割道宽度。",
    )

    st.info(
        "说明：`Chip array` 的 x/y 表示阵列列数和行数；"
        "`4X / 5X LBC center location` 的 x/y 表示阵列最左下角第一个 Cell 的中心位置；"
        "列 pitch = Cell 外框宽度 + 切割道宽度，行 pitch = Cell 外框高度 + 切割道宽度；"
        "程序只会在本页已上传的 GDS 文件中按 Database 匹配。"
    )
    st.caption(f"当前已上传 {len(uploaded_gds_files or [])} 个 GDS 文件。")

    try:
        parsed_entries = engine.parse_sheet(workbook_path, sheet_name)
    except Exception as exc:
        st.error(f"解析 Sheet 失败: {exc}")
        return

    preview_df = pd.DataFrame(engine.entries_to_rows(parsed_entries))
    st.caption(f"当前 Sheet 共解析到 {len(parsed_entries)} 条可排版记录。")
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

    if st.button("生成排版 GDS", type="primary"):
        if not uploaded_gds_files:
            st.warning("请至少上传一个用于排版的 GDS 文件。")
            return

        safe_sheet_name = SummaryLayoutEngine.sanitize_cell_name(sheet_name)
        output_filename = f"{safe_sheet_name}_{SUMMARY_LAYOUT_OUTPUT_NAME}"
        output_path = build_temp_output_path(output_filename)

        try:
            uploaded_gds_paths = [save_uploaded_file(item, preserve_name=True) for item in uploaded_gds_files]
            result = engine.process(
                workbook_path,
                sheet_name,
                uploaded_gds_paths,
                output_path,
                output_top_name,
                street_width=street_width,
                entries=parsed_entries,
            )
            st.success(
                f"排版完成：成功放置 {result['placed_count']} / {result['total_entries']} 个条目，"
                f"共生成 {result['placed_instance_count']} 个实例，输出 Top Cell 为 {result['output_top_name']}。"
            )

            metric_total, metric_placed, metric_missing_file, metric_missing_top = st.columns(4)
            metric_total.metric("总条目", result["total_entries"])
            metric_placed.metric("成功排版", result["placed_count"])
            metric_missing_file.metric("缺失 GDS", result["skipped_missing_file"])
            metric_missing_top.metric("缺失 Top Cell", result["skipped_missing_top"])

            if result["skipped_empty_top"]:
                st.info(f"另有 {result['skipped_empty_top']} 个条目因为目标 Top Cell 没有边界框而被跳过。")

            for warning in result["duplicate_name_warnings"]:
                st.warning(warning)

            if result["unit_mismatch_files"]:
                st.warning(
                    "以下文件的 unit/precision 与首个导入库不同，请额外确认单位一致性："
                    + ", ".join(result["unit_mismatch_files"])
                )

            result_df = pd.DataFrame(result["results"])
            with st.expander("查看排版明细", expanded=False):
                st.dataframe(result_df, use_container_width=True, hide_index=True)

            with output_path.open("rb") as file_obj:
                st.download_button("下载排版后的 GDS", file_obj, file_name=output_filename)
        except Exception as exc:
            st.error(f"排版失败: {exc}")


def render_film_volume_page() -> None:
    uploaded_gds = st.file_uploader("上传 GDS 文件", type=["gds"], key="film_volume_gds")
    if not uploaded_gds:
        return

    gds_path = save_uploaded_file(uploaded_gds)
    engine = FilmVolumeEngine()

    try:
        cell_names = engine.load_lib(gds_path)
    except Exception as exc:
        st.error(f"读取 GDS 失败: {exc}")
        return

    cell_name = st.selectbox("选择需要统计的 Cell", cell_names, key="film_volume_cell")

    st.info(
        "说明：程序会直接递归读取选中的 Cell 几何，不先整体打散；"
        "再按 layer/datatype 合并同层图形后汇总面积；"
        "导出时会使用默认 Excel 参数，自动写入 `AreaSummary` Sheet 的 `A1` 起始位置，并固定带表头。"
    )

    if st.button("计算膜层面积并导出 Excel", type="primary"):
        safe_gds_name = SummaryLayoutEngine.sanitize_cell_name(Path(uploaded_gds.name).stem)
        safe_cell_name = SummaryLayoutEngine.sanitize_cell_name(cell_name)
        output_basename = f"{safe_gds_name}_{safe_cell_name}_{FILM_VOLUME_OUTPUT_NAME}"
        output_path = build_temp_output_path(output_basename)

        try:
            rows = engine.summarize_cell_areas(gds_path, cell_name)
            write_result = engine.write_area_table(
                rows,
                output_path,
                sheet_name=FILM_VOLUME_DEFAULT_SHEET,
                start_cell=FILM_VOLUME_DEFAULT_START_CELL,
            )

            result_df = pd.DataFrame(engine.rows_to_dicts(rows))
            result_df["total_area"] = result_df["total_area"].map(lambda value: f"{value:.6f}")

            st.success(
                f"统计完成：共汇总 {len(rows)} 个 layer/datatype 项，结果已写入 Sheet "
                f"`{write_result['sheet_name']}` 的 `{write_result['start_cell']}` 起始位置。"
            )

            st.dataframe(result_df, use_container_width=True, hide_index=True)

            with output_path.open("rb") as file_obj:
                st.download_button("下载 Excel 结果", file_obj, file_name=output_basename)
        except Exception as exc:
            st.error(f"处理失败: {exc}")


def render_etch_duty_page() -> None:
    uploaded_gds = st.file_uploader("上传 GDS 文件", type=["gds"], key="etch_duty_gds")
    if not uploaded_gds:
        return

    gds_path = save_uploaded_file(uploaded_gds)
    engine = EtchDutyEngine()

    try:
        cell_names = engine.load_lib(gds_path)
    except Exception as exc:
        st.error(f"读取 GDS 失败: {exc}")
        return

    cell_name = st.selectbox("选择需要统计的 Shot Cell", cell_names, key="etch_duty_cell")
    layer_filter_raw = st.text_input(
        "指定 Layer（可选）",
        value="",
        key="etch_duty_layer_filter",
        help="不填写则统计全部图层。可输入逗号分隔整数，例如 `1, 2, 5`，也支持范围写法 `1-4, 8`。",
    )

    st.info(
        "说明：程序会直接递归读取选中的 Shot Cell 几何，不先整体打散；"
        "再按 layer/datatype 合并同层图形后统计面积；"
        "如果填写了指定 Layer，则只统计这些层；"
        "占空比默认按 Shot Cell 的边界框面积计算，结果会写入 `DutyCycleSummary` Sheet 的 `A1`。"
    )

    if st.button("统计刻蚀占空比并导出 Excel", type="primary"):
        safe_gds_name = SummaryLayoutEngine.sanitize_cell_name(Path(uploaded_gds.name).stem)
        safe_cell_name = SummaryLayoutEngine.sanitize_cell_name(cell_name)
        output_basename = f"{safe_gds_name}_{safe_cell_name}_{ETCH_DUTY_OUTPUT_NAME}"
        output_path = build_temp_output_path(output_basename)

        try:
            included_layers = parse_int_set(layer_filter_raw, "指定 Layer")
            rows, cell_total_area = engine.summarize_shot_duty(
                gds_path,
                cell_name,
                included_layers=included_layers or None,
            )
            write_result = engine.write_duty_table(
                rows,
                output_path,
                sheet_name=ETCH_DUTY_DEFAULT_SHEET,
                start_cell=ETCH_DUTY_DEFAULT_START_CELL,
            )

            result_df = pd.DataFrame(engine.rows_to_dicts(rows))
            result_df["total_area"] = result_df["total_area"].map(lambda value: f"{value:.6f}")
            result_df["duty_cycle_percent"] = result_df["duty_cycle_percent"].map(lambda value: f"{value:.6f}")

            st.success(
                f"统计完成：共汇总 {len(rows)} 个 layer/datatype 项，Shot Cell 总面积为 "
                f"`{cell_total_area:.6f}`，结果已写入 Sheet `{write_result['sheet_name']}` 的 "
                f"`{write_result['start_cell']}` 起始位置。"
            )

            if included_layers:
                st.caption("本次仅统计 Layer: " + ", ".join(str(layer) for layer in sorted(included_layers)))

            st.dataframe(result_df, use_container_width=True, hide_index=True)

            with output_path.open("rb") as file_obj:
                st.download_button("下载 Excel 结果", file_obj, file_name=output_basename)
        except Exception as exc:
            st.error(f"处理失败: {exc}")


if __name__ == "__main__":
    main()
