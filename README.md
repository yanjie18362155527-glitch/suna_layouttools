# Layout Auto Tools

基于 Streamlit 的 GDS 版图自动化工具，包含七个功能：

- Lens 自动编号
- Pad 信息提取并导出 Excel
- Shot 自动编号
- Shot 内子 cell 信息提取并导出 Excel
- 根据 DB 信息排Shot
- 各图层面积（体积）统计
- 递归统计指定 shot cell 各图层面积及刻蚀占空比并写入 Excel

## 项目结构

```text
my_project/
├── data/
├── my_project/
│   ├── __init__.py
│   ├── config.py
│   ├── main.py
│   ├── core/
│   │   ├── __init__.py
│   │   ├── base.py
│   │   ├── cell_info.py
│   │   ├── etch_duty.py
│   │   ├── film_volume.py
│   │   ├── helpers.py
│   │   ├── lens.py
│   │   ├── logic.py
│   │   ├── pad.py
│   │   ├── preview.py
│   │   └── shot.py
│   └── utils/
│       ├── __init__.py
│       └── common.py
├── requirements.txt
└── README.md
```

## 环境准备

```bash
pip install -r requirements.txt
```

## 启动方式

```bash
streamlit run my_project/main.py
```

也可以使用仓库内的 `run_server.bat`。

## 模块说明

- `my_project/main.py`: Streamlit 页面入口和交互编排。
- `my_project/config.py`: 页面模式、默认参数和路径配置。
- `my_project/utils/common.py`: 通用工具，例如临时文件保存、参数解析。
- `my_project/core/base.py`: GDS 读写基础引擎。
- `my_project/core/preview.py`: GDS 预览渲染与坐标映射。
- `my_project/core/lens.py`: Lens 自动编号引擎。
- `my_project/core/pad.py`: Pad 提取、预览和 Excel 导出。
- `my_project/core/shot.py`: Shot 自动编号引擎。
- `my_project/core/cell_info.py`: 子 cell 信息提取和可视化导出。
- `my_project/core/etch_duty.py`: 递归读取 shot cell 几何后按 layer/datatype 合并图形并统计刻蚀占空比。
- `my_project/core/film_volume.py`: 递归读取指定 cell 几何后按 layer/datatype 合并图形并汇总面积写入 Excel。
- `my_project/core/helpers.py`: 核心纯函数与测试辅助对象。
- `my_project/core/logic.py`: 兼容导出层，便于旧引用平滑过渡。

## 说明

- 程序仍然使用上传 GDS 文件后在线处理并下载结果的方式。
- 第六个功能当前不先 flatten cell，而是递归读取几何后合并统计；默认写入 `AreaSummary` Sheet 的 `A1`，面积保留小数点后六位。
- 第七个功能当前不先 flatten shot cell，而是递归读取几何后合并统计；分母取 shot cell 的边界框面积，并支持只统计指定 layer。
- 运行时结果文件会写入系统临时目录，不会污染项目目录。
- `data/` 目录保留给数据库、日志或后续静态数据文件使用。
