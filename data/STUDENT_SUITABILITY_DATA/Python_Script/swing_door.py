"""DOOR-001: CAD-independent parametric swing-door generator.

Coordinates: millimetres. X = door width, Y = depth, Z = height.
Origin = bottom-left of rough opening, viewed from the designated front.
All parts are described in local coordinates and placed by translation.
This is a geometry/assembly IR, not manufacturing approval.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any
import json


DEFAULTS = {
    "product_id": "DOOR-001",
    "opening": {"width": None, "height": None, "hinge_side": None},
    "frame": {"width": 45, "depth": 90},
    "leaf": {"thickness": 40, "finish": "Walnut"},
    "clearance": {"side": 4, "top": 4, "bottom": 8},
    "hardware": {"hinge": {"model": "HINGE-B01", "quantity": 2}},
}

# Demonstration installation dimensions only; replace with verified manufacturer data.
HINGE_CATALOG = {
    "HINGE-B01": {"type": "butt", "leaf_pocket": [100, 28, 3], "frame_pocket": [100, 28, 3]},
    "HINGE-C01": {"type": "concealed", "leaf_pocket": [110, 24, 22], "frame_pocket": [110, 24, 20]},
}


class ConfigurationError(ValueError):
    def __init__(self, issues: list[dict[str, str]]):
        self.issues = issues
        super().__init__(json.dumps(issues, ensure_ascii=False))


def merge_defaults(defaults: dict, overrides: dict) -> dict:
    result = deepcopy(defaults)
    for key, value in overrides.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = merge_defaults(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def validate(c: dict) -> None:
    issues: list[dict[str, str]] = []

    def positive_number(path: str, value: Any, low: float, high: float):
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not low <= value <= high:
            issues.append({"path": path, "message": f"请输入 {low}～{high} mm 范围内的数值"})

    opening = c["opening"]
    positive_number("opening.width", opening.get("width"), 600, 1200)
    positive_number("opening.height", opening.get("height"), 1900, 3000)
    if opening.get("hinge_side") not in ("left", "right"):
        issues.append({"path": "opening.hinge_side", "message": "请选择 left 或 right"})

    for group, fields in {
        "frame": {"width": (20, 100), "depth": (40, 300)},
        "leaf": {"thickness": (20, 80)},
        "clearance": {"side": (1, 15), "top": (1, 15), "bottom": (1, 30)},
    }.items():
        for field, (lo, hi) in fields.items():
            positive_number(f"{group}.{field}", c[group].get(field), lo, hi)

    model = c["hardware"]["hinge"].get("model")
    if model not in HINGE_CATALOG:
        issues.append({"path": "hardware.hinge.model", "message": f"未知型号：{model}"})
    quantity = c["hardware"]["hinge"].get("quantity")
    if type(quantity) is not int or not 2 <= quantity <= 4:
        issues.append({"path": "hardware.hinge.quantity", "message": "演示版本支持 2～4 个合页"})

    if issues:
        raise ConfigurationError(issues)

    inner_width = opening["width"] - 2 * c["frame"]["width"]
    inner_height = opening["height"] - c["frame"]["width"]
    leaf_width = inner_width - 2 * c["clearance"]["side"]
    leaf_height = inner_height - c["clearance"]["top"] - c["clearance"]["bottom"]
    if leaf_width <= 0 or leaf_height <= 0:
        raise ConfigurationError([{"path": "opening", "message": "门框和间隙占用尺寸过大"}])
    hinge = HINGE_CATALOG[model]
    if c["leaf"]["thickness"] < hinge["leaf_pocket"][2]:
        raise ConfigurationError([{"path": "leaf.thickness", "message": "门扇厚度小于合页开槽深度"}])
    if c["frame"]["depth"] < hinge["frame_pocket"][2]:
        raise ConfigurationError([{"path": "frame.depth", "message": "门框深度小于合页开槽深度"}])


def box_part(part_id: str, size: list[float], position: list[float], material: str) -> dict:
    return {
        "id": part_id,
        "geometry": {"type": "extrude", "profile": [[0, 0], [size[0], 0], [size[0], size[1]], [0, size[1]]], "vector": [0, 0, size[2]]},
        "transform": {"translation": position, "rotation_deg": [0, 0, 0]},
        "material": material,
    }


def hinge_positions(height: float, quantity: int) -> list[float]:
    """Hinge center Z values in global door coordinates, illustrative only."""
    low, high = 300.0, height - 300.0
    return [low + i * (high - low) / (quantity - 1) for i in range(quantity)]


def pocket_feature(feature_id: str, target: str, x: float, y: float, z_center: float, pocket: list[float]) -> dict:
    length, width, depth = pocket
    # Cutter geometry is expressed in the target part's local coordinate frame.
    return {
        "id": feature_id,
        "target": target,
        "operation": "subtract",
        "geometry": {"type": "box", "size": [width, depth, length]},
        "local_transform": {"translation": [x, y, z_center - length / 2], "rotation_deg": [0, 0, 0]},
        "note": "Demo pocket only; mounting face and manufacturer specs require verification",
    }


def generate(config: dict) -> dict:
    """Generate a portable parts/features/assembly description, no bpy imports."""
    c = merge_defaults(DEFAULTS, config)
    validate(c)

    width = c["opening"]["width"]
    height = c["opening"]["height"]
    fw, fd = c["frame"]["width"], c["frame"]["depth"]
    side, top, bottom = (c["clearance"][k] for k in ("side", "top", "bottom"))
    leaf_w = width - 2 * fw - 2 * side
    leaf_h = height - fw - top - bottom
    leaf_t = c["leaf"]["thickness"]
    leaf_x = fw + side
    leaf_z = bottom

    parts = [
        box_part("FRAME-L", [fw, fd, height], [0, 0, 0], "Frame"),
        box_part("FRAME-R", [fw, fd, height], [width - fw, 0, 0], "Frame"),
        box_part("FRAME-T", [width - 2 * fw, fd, fw], [fw, 0, height - fw], "Frame"),
        box_part("LEAF", [leaf_w, leaf_t, leaf_h], [leaf_x, (fd - leaf_t) / 2, leaf_z], c["leaf"]["finish"]),
    ]

    hinge_config = c["hardware"]["hinge"]
    hinge_model = hinge_config["model"]
    hinge_spec = HINGE_CATALOG[hinge_model]
    hinge_side = c["opening"]["hinge_side"]
    hinge_x = leaf_x if hinge_side == "left" else leaf_x + leaf_w
    frame_id = "FRAME-L" if hinge_side == "left" else "FRAME-R"
    frame_global_x = 0 if hinge_side == "left" else width - fw
    leaf_global_y = (fd - leaf_t) / 2
    frame_pocket_x = fw - hinge_spec["frame_pocket"][1] if hinge_side == "left" else 0
    leaf_pocket_x = 0 if hinge_side == "left" else leaf_w - hinge_spec["leaf_pocket"][1]

    features = []
    assembly = [
        {"parent": "DOOR-001", "child": p["id"], "relation": "fixed_placement", "transform": p["transform"]}
        for p in parts if p["id"] != "LEAF"
    ]
    assembly.append({"parent": "DOOR-001", "child": "LEAF", "relation": "hinged_closed_pose", "transform": parts[-1]["transform"], "hinge_side": hinge_side})

    for index, z in enumerate(hinge_positions(height, hinge_config["quantity"]), start=1):
        hid = f"HINGE-{index:02d}"
        parts.append({
            "id": hid,
            "geometry": {"type": "asset", "asset_id": hinge_model},
            "transform": {"translation": [hinge_x, fd / 2, z], "rotation_deg": [0, 0, 0]},
            "material": "Metal",
        })
        features.append(pocket_feature(f"{hid}-LEAF-POCKET", "LEAF", leaf_pocket_x, 0, z - leaf_z, hinge_spec["leaf_pocket"]))
        features.append(pocket_feature(f"{hid}-FRAME-POCKET", frame_id, frame_pocket_x, 0, z, hinge_spec["frame_pocket"]))
        assembly.append({"parent": "DOOR-001", "child": hid, "relation": "hinge", "frame": frame_id, "leaf": "LEAF"})

    return {
        "schema_version": "0.1",
        "product_id": c["product_id"],
        "units": "mm",
        "coordinate_system": "X width, Y depth, Z height; origin bottom-left of opening",
        "resolved_config": c,
        "parts": parts,
        "features": features,
        "assembly": assembly,
        "derived": {"leaf_width": leaf_w, "leaf_height": leaf_h},
        "manufacturing_ready": False,
    }


if __name__ == "__main__":
    example = {
        "opening": {"width": 900, "height": 2200, "hinge_side": "left"},
        "hardware": {"hinge": {"model": "HINGE-C01", "quantity": 2}},
    }
    result = generate(example)
    print(json.dumps(result, ensure_ascii=False, indent=2))
