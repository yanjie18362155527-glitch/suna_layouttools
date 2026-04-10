import tempfile
from pathlib import Path


def ensure_directory(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_uploaded_file(uploaded_file, preserve_name: bool = False) -> Path:
    original_name = Path(uploaded_file.name).name
    suffix = Path(original_name).suffix
    if preserve_name:
        target_dir = Path(tempfile.mkdtemp(prefix="layout_auto_tools_"))
        target_path = target_dir / original_name
        target_path.write_bytes(uploaded_file.getvalue())
        return target_path

    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        return Path(tmp_file.name)


def parse_pair(raw_value: str, label: str) -> tuple[float, float]:
    parts = [part.strip() for part in raw_value.split(",")]
    if len(parts) != 2:
        raise ValueError(f"{label} 必须是 'x, y' 格式")
    try:
        return float(parts[0]), float(parts[1])
    except ValueError as exc:
        raise ValueError(f"{label} 必须是数字") from exc


def parse_int_set(raw_value: str, label: str) -> set[int]:
    normalized = (raw_value or "").replace("，", ",").strip()
    if not normalized:
        return set()

    values: set[int] = set()
    for part in normalized.split(","):
        token = part.strip()
        if not token:
            continue

        if "-" in token:
            bounds = [item.strip() for item in token.split("-", 1)]
            if len(bounds) != 2 or not bounds[0] or not bounds[1]:
                raise ValueError(f"{label} 中的范围格式无效: {token}")
            try:
                start = int(bounds[0])
                end = int(bounds[1])
            except ValueError as exc:
                raise ValueError(f"{label} 中的范围必须是整数: {token}") from exc
            if end < start:
                raise ValueError(f"{label} 中的范围必须从小到大: {token}")
            values.update(range(start, end + 1))
            continue

        try:
            values.add(int(token))
        except ValueError as exc:
            raise ValueError(f"{label} 中存在无效整数: {token}") from exc

    return values


def build_temp_output_path(filename: str) -> Path:
    return Path(tempfile.gettempdir()) / filename
