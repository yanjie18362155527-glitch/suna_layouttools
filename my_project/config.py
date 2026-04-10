from pathlib import Path


PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent
DATA_DIR = PROJECT_ROOT / "data"

APP_TITLE = "版图自动化工具平台"
APP_LAYOUT = "wide"

LENS_MODE = "1. Lens 自动编号"
PAD_MODE = "2. Pad 信息提取"
SHOT_MODE = "3. Shot 自动编号"
CELL_INFO_MODE = "4. Shot 内 DB 信息提取"
SUMMARY_LAYOUT_MODE = "5. 根据 DB 信息排Shot"
FILM_VOLUME_MODE = "6. 各图层面积（体积）统计"
ETCH_DUTY_MODE = "7. 刻蚀占空比统计"
APP_MODES = [LENS_MODE, PAD_MODE, SHOT_MODE, CELL_INFO_MODE, SUMMARY_LAYOUT_MODE, FILM_VOLUME_MODE, ETCH_DUTY_MODE]

LENS_DEFAULT_PARENT = "D53Z_V1"
LENS_DEFAULT_CHILD = "lens_fan"
LENS_DEFAULTS = {
    "offset_x": 0.0,
    "offset_y": 0.0,
    "size": 50.0,
    "view_bbox": None,
    "zoom_version": 0,
}

PAD_DEFAULT_CELL = "DIFF_OPT_V2_1"
PAD_DEFAULT_LAYER = 9
PAD_DEFAULT_DATATYPE = 0
PAD_DEFAULT_AUX_LAYER = 10
PAD_DEFAULT_AUX_DATATYPE = 0

SHOT_DEFAULT_PARENT = "0MA8_9CUN"
SHOT_DEFAULT_CHILD = "00A_Shot"
SHOT_DEFAULTS = {
    "anchor": "0, 0",
    "area": "100, 100",
    "view_bbox": None,
    "zoom_version": 0,
}

PREVIEW_CANVAS_WIDTH = 700
PREVIEW_CANVAS_HEIGHT = 500
PAD_CANVAS_SIZE = 800

LENS_OUTPUT_NAME = "lens_result.gds"
PAD_OUTPUT_NAME = "pad_result.xlsx"
SHOT_OUTPUT_NAME = "shot_result.gds"
CELL_INFO_OUTPUT_NAME = "cell_info.xlsx"
SUMMARY_LAYOUT_OUTPUT_NAME = "summary_layout.gds"
SUMMARY_LAYOUT_DEFAULT_TOP = "summary_layout"
SUMMARY_LAYOUT_DEFAULT_STREET_WIDTH = 0.0
FILM_VOLUME_OUTPUT_NAME = "layer_area_summary.xlsx"
FILM_VOLUME_DEFAULT_SHEET = "AreaSummary"
FILM_VOLUME_DEFAULT_START_CELL = "A1"
ETCH_DUTY_OUTPUT_NAME = "etch_duty_summary.xlsx"
ETCH_DUTY_DEFAULT_SHEET = "DutyCycleSummary"
ETCH_DUTY_DEFAULT_START_CELL = "A1"

USAGE_GUIDE_SECTIONS = [
    {
        "title": "1. Lens 自动编号",
        "steps": [
            "上传 GDS 文件后，先选择父 Cell（排好lens_fan的wafercell） 和需要编号的子 Cell。",
            "建议开启“可视化框选”，在预览图中框选文字放置区域，系统会自动回填 Offset 和 Size。",
            "在高级设置中选择编号模式：行列坐标模式会生成类似“1-1”的编号，顺序索引模式会按排序方向连续编号。",
            "确认 Layer、Datatype、容差和位数后，点击“开始生成 GDS”，下载生成后的结果文件。",
        ],
        "tips": [
            "如果阵列排布不完全规则，适当调大容差，避免同一行或同一列被拆成多个编号组。",
            "如果文字位置偏了，优先微调 Offset X、Offset Y 和 Size。",
        ],
    },
    {
        "title": "2. Pad 信息提取",
        "steps": [
            "上传 GDS 后选择目标 Cell，并填写需要提取的 Layer 和 Datatype。",
            "如果希望在导出的 Excel 预览图里叠加参考图层（如pad下面的M2层），可以勾选辅助可视化图层并填写M2层辅助 Layer/Datatype。",
            "点击“解析 GDS”后，系统会提取候选 Pad，并生成可点击预览图。",
            "在预览图中按实际需求顺序逐个点击 Pad，点击顺序就是最终导出的编号顺序。",
            "确认无误后点击“导出顺序提取报告”，下载 Excel 报告。",
        ],
        "tips": [
            "重复点击同一个 Pad 不会重复计入结果。",
            "如果没有提取到 Pad，通常是 Layer 或 Datatype 设置不对。",
        ],
    },
    {
        "title": "3. Shot 自动编号",
        "steps": [
            "上传 GDS 后选择 Top Cell 和 Shot Cell。",
            "建议开启“可视化框选”，在 Shot 单元预览图中框选文字锚点区域，系统会自动更新 Anchor 和 Area。",
            "设置输出文字所在的 Layer 和 Datatype。",
            "点击“运行编号”后，系统会根据阵列中心建立相对坐标，并生成类似“(0,0)”的 Shot 编号。",
            "完成后下载结果 GDS 文件。",
        ],
        "tips": [
            "Anchor 决定文字中心位置，Area 决定文字可用宽高，直接影响字号和排版。",
            "如果某些 Shot 存在旋转或镜像，程序会按 reference 变换后再放置文字。",
        ],
    },
    {
        "title": "4. Shot 内 DB 信息提取",
        "steps": [
            "上传 GDS 后选择 Top Cell。",
            "在子 Cell 列表中勾选需要提取的目标类型。",
            "点击“生成报告”，系统会统计每类目标的中心坐标和尺寸，并生成带示意图的 Excel。",
            "下载 Excel 后可以直接查看子 Cell 名称、位置、尺寸和可视化布局。",
        ],
        "tips": [
            "当前导出逻辑会按自上而下、从左到右的顺序整理结果。",
            "如果选中的子 Cell 在 Top Cell 中不存在，会提示提取失败。",
        ],
    },
    {
        "title": "5. 根据 DB 信息排Shot",
        "steps": [
            "上传汇总 Excel，选择包含 Database、Top cell、Chip array、4X / 5X LBC center location 信息的 Sheet。",
            "在页面中多选上传本次需要参与排版的 GDS 文件，程序只会在这些已上传文件里按 Database 名称匹配。",
            "程序会读取 Top cell 行指定的顶层 Cell，并读取 Chip array 后的 x、y 作为阵列列数和行数。",
            "程序会读取 4X / 5X LBC center location 后的 x、y，作为阵列中最左下角第一个 Cell 的中心位置。",
            "页面中的“切割道宽度”会同时加到阵列列间距和行间距上，因此列 pitch = Cell 外框宽度 + 切割道宽度，行 pitch = Cell 外框高度 + 切割道宽度。",
            "若某个 GDS 未上传，或对应 Top cell 不存在，该条目会自动跳过，不影响其他产品继续排版。",
            "点击“生成排版 GDS”后下载结果文件，在输出的 Top Cell 中查看总装排版结果。",
        ],
        "tips": [
            "如果重复上传了同名 GDS，程序会默认使用上传列表中的第一个文件。",
            "如果 Database 行填写的是 .db 名称，程序也会尝试用同名 .gds 做一次兜底匹配。",
            "如果 Chip array 的 x、y 为空，程序会按 1x1 单实例处理。",
            "切割道宽度会同时影响列与列、行与行之间的距离。",
        ],
    },
    {
        "title": "6. 各图层面积（体积）统计",
        "steps": [
            "上传 GDS 文件后，选择需要统计的目标 Cell。",
            "程序会直接递归读取该 Cell 及其子层级中的几何，不先整体打散，再按 layer/datatype 合并同层图形。",
            "程序会自动新建一个 Excel 文件，默认写入 `AreaSummary` Sheet，并从 `A1` 开始输出。",
            "写入内容固定为 Layer、Datatype、Total Area 三列表头，面积保留小数点后六位。",
            "点击“计算膜层面积并导出 Excel”后，程序会把统计结果写入指定位置，并提供下载。",
        ],
        "tips": [
            "当前功能输出的是各图层合并后的面积汇总，适合后续再结合膜厚换算体积。",
            "相较于先 flatten 再统计，这个实现通常会更省一些时间和内存。",
            "如果目标 Cell 及其子层级里没有图形，程序会提示无法统计面积。",
            "页面不会再显示 Excel 输出参数，统一使用默认 Sheet 和起始单元格。",
        ],
    },
    {
        "title": "7. 刻蚀占空比统计",
        "steps": [
            "上传本地 GDS 文件后，选择需要统计的 Shot Cell。",
            "如需提速，可在页面中填写指定 Layer，只统计这些图层；不填写则统计全部图层。",
            "程序会直接递归读取 Shot Cell 及其子层级中的几何，不先整体打散，再按 layer/datatype 合并同层图形。",
            "程序会以 Shot Cell 的边界框宽 × 高作为 Cell 总面积，并据此计算各层占空比。",
            "程序会自动新建一个 Excel 文件，默认写入 `DutyCycleSummary` Sheet，并从 `A1` 开始输出。",
            "写入内容固定为 Layer、Datatype、Total Area、Duty Cycle (%) 四列表头，面积和占空比都保留小数点后六位。",
            "点击“统计刻蚀占空比并导出 Excel”后，程序会把结果写入默认位置并提供下载。",
        ],
        "tips": [
            "如果同一层存在重叠图形，程序会先做合并，因此不会重复累计重叠区域面积。",
            "指定 Layer 支持逗号分隔和范围写法，例如 `1, 2, 5-8`。",
            "相较于先 flatten 再统计，这个实现通常会更省一些时间和内存。",
            "当前占空比的分母是 Shot Cell 的边界框面积；如果你后续要换成其他口径，可以再调整。",
            "如果目标 Shot Cell 及其子层级里没有图形，程序会提示无法统计占空比。",
        ],
    },
]
