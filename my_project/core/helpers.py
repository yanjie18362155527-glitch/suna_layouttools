import math
from types import SimpleNamespace


def collect_unique_coordinates(coords: list[float], tolerance: float) -> list[float]:
    if not coords:
        return []

    unique_values = [coords[0]]
    for coord in coords[1:]:
        if coord - unique_values[-1] > tolerance:
            unique_values.append(coord)
    return unique_values


def transform_reference_point(point: tuple[float, float], reference) -> tuple[float, float]:
    px, py = point
    origin = reference.origin
    rotation = reference.rotation
    magnification = reference.magnification or 1.0
    x_reflection = reference.x_reflection

    if x_reflection:
        py = -py

    px *= magnification
    py *= magnification

    if rotation not in (None, 0):
        cos_v = math.cos(rotation)
        sin_v = math.sin(rotation)
        px, py = px * cos_v - py * sin_v, px * sin_v + py * cos_v

    return px + origin[0], py + origin[1]


def build_reference_stub(**overrides):
    values = {
        "origin": (0.0, 0.0),
        "rotation": 0.0,
        "magnification": 1.0,
        "x_reflection": False,
    }
    values.update(overrides)
    return SimpleNamespace(**values)
