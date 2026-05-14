import json
import os
from collections import defaultdict
from multiprocessing import Pool, cpu_count


base_dir = "/pfs/xiabin/datasets/new_layers_all/elements"
output_dir = "/pfs/xiabin/datasets/new_layers_all/elements/tongjiall"
target_filename = "layers_caption.json"
ignored_keys = {"caption", "short_caption", "saved_path", "rgb", "text", "font_name","xfrm_rot_raw",}
NUM_WORKERS = 128
SKIP_VALUE = object()
DATASET_DIRS = [
    "d-0414-0417-part1-caption-merge",
    "d-0414-0417-part2-caption-merge",
    "d-0420-caption-merge",
    "d-0421-caption-merge",
    "d-0422-caption-merge",
    "d-0423-caption-merge",
    "d-0424-caption-merge",
    "d-0427-caption-merge",
    "d-0428-caption-merge",
    "d-0429-caption-merge",
    "d-0430-caption-merge",
]


def collect_files():
    all_files = []
    missing_dirs = []

    for dataset_dir in DATASET_DIRS:
        input_dir = os.path.join(base_dir, dataset_dir)
        if not os.path.isdir(input_dir):
            missing_dirs.append(input_dir)
            continue

        dir_file_count = 0
        for first_entry in os.scandir(input_dir):
            if not first_entry.is_dir():
                continue
            for second_entry in os.scandir(first_entry.path):
                if not second_entry.is_dir():
                    continue
                for third_entry in os.scandir(second_entry.path):
                    if not third_entry.is_dir():
                        continue
                    fpath = os.path.join(third_entry.path, target_filename)
                    if os.path.isfile(fpath):
                        all_files.append(fpath)
                        dir_file_count += 1

        print(f"[收集] {dataset_dir}: {dir_file_count} 个 {target_filename}", flush=True)

    if missing_dirs:
        print(f"[WARN] 跳过不存在的目录 {len(missing_dirs)} 个:")
        for missing_dir in missing_dirs:
            print(f"  - {missing_dir}")

    return sorted(all_files)


def normalize_items(data):
    if isinstance(data, dict):
        if isinstance(data.get("layers"), list):
            return normalize_items(data["layers"])
        return [data]
    if isinstance(data, list):
        items = []
        for item in data:
            items.extend(normalize_items(item))
        return items
    return []


def make_structure(value):
    if isinstance(value, dict):
        return {
            "type": "dict",
            "keys": {
                key: make_structure(value[key])
                for key in value.keys()
            },
        }
    if isinstance(value, list):
        element_structures = {}
        for item in value:
            item_structure = make_structure(item)
            signature = json.dumps(item_structure, ensure_ascii=False)
            element_structures.setdefault(signature, item_structure)
        return {
            "type": "list",
            "items": list(element_structures.values()),
        }
    return {"type": "value"}


def merge_structure_items(target_items, source_items):
    merged_items = [item for item in target_items]
    for source_item in source_items:
        for idx, target_item in enumerate(merged_items):
            if target_item["type"] == source_item["type"]:
                merged_items[idx] = merge_structures(target_item, source_item)
                break
        else:
            merged_items.append(source_item)
    return merged_items


def merge_structures(target, source):
    if target["type"] != source["type"]:
        target_items = target["items"] if target["type"] == "variants" else [target]
        source_items = source["items"] if source["type"] == "variants" else [source]
        merged_items = merge_structure_items(target_items, source_items)
        if len(merged_items) == 1:
            return merged_items[0]
        return {
            "type": "variants",
            "items": merged_items,
        }

    if target["type"] == "dict":
        merged_keys = dict(target["keys"])
        for key, source_child in source["keys"].items():
            if key in merged_keys:
                merged_keys[key] = merge_structures(merged_keys[key], source_child)
            else:
                merged_keys[key] = source_child
        return {
            "type": "dict",
            "keys": merged_keys,
        }

    if target["type"] == "list":
        return {
            "type": "list",
            "items": merge_structure_items(target["items"], source["items"]),
        }

    if target["type"] == "variants":
        return {
            "type": "variants",
            "items": merge_structure_items(target["items"], source["items"]),
        }

    return target


def iter_leaf_values(value, prefix=""):
    if isinstance(value, dict):
        for key, child in value.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            if key in ignored_keys:
                yield child_prefix, SKIP_VALUE
                continue
            yield from iter_leaf_values(child, child_prefix)
        return

    if isinstance(value, list):
        child_prefix = f"{prefix}[]"
        for child in value:
            yield from iter_leaf_values(child, child_prefix)
        return

    yield prefix, value


def value_key(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True)


def new_template_stats(kind, item):
    return {
        "kind": kind,
        "count": 0,
        "structure": make_structure(item),
        "values": defaultdict(dict),
    }


def add_item_to_template(stats, item):
    stats["count"] += 1
    stats["structure"] = merge_structures(stats["structure"], make_structure(item))
    for key_path, value in iter_leaf_values(item):
        if value is SKIP_VALUE:
            stats["values"].setdefault(key_path, {})
            continue
        stats["values"][key_path][value_key(value)] = value


def process_file(fpath):
    local_templates = {}
    total_items = 0

    try:
        with open(fpath, "r", encoding="utf-8") as fp:
            data = json.load(fp)
    except Exception as e:
        return None, f"{fpath}: {e}"

    for item in normalize_items(data):
        if not isinstance(item, dict):
            continue
        kind = item.get("kind")
        if kind is None:
            continue

        if kind not in local_templates:
            local_templates[kind] = new_template_stats(kind, item)
        add_item_to_template(local_templates[kind], item)
        total_items += 1

    serializable_templates = []
    for kind, stats in local_templates.items():
        serializable_templates.append(
            {
                "kind": kind,
                "signature": "__merged__",
                "count": stats["count"],
                "structure": stats["structure"],
                "values": {
                    key: list(values.values())
                    for key, values in stats["values"].items()
                },
            }
        )

    return {"templates": serializable_templates, "total_items": total_items}, None


def merge_values(target, source_values):
    for key, values in source_values.items():
        target.setdefault(key, {})
        for value in values:
            target[key][value_key(value)] = value


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def format_values(values):
    unique_values = list(values.values())
    if unique_values and all(is_number(value) for value in unique_values):
        sorted_values = sorted(unique_values)
        if len(sorted_values) > 20:
            return [{"min": sorted_values[0], "max": sorted_values[-1]}]
        return sorted_values
    return sorted(unique_values, key=lambda value: value_key(value))


def build_stat_template(structure, values, prefix=""):
    if structure["type"] == "dict":
        return {
            key: (
                []
                if key in ignored_keys
                else build_stat_template(child, values, f"{prefix}.{key}" if prefix else key)
            )
            for key, child in structure["keys"].items()
        }

    if structure["type"] == "list":
        item_structures = structure["items"]
        if len(item_structures) == 1 and item_structures[0]["type"] == "value":
            return format_values(values.get(f"{prefix}[]", {}))
        return [
            build_stat_template(item_structure, values, f"{prefix}[]")
            for item_structure in item_structures
        ]

    if structure["type"] == "variants":
        return [
            build_stat_template(item_structure, values, prefix)
            for item_structure in structure["items"]
        ]

    return format_values(values.get(prefix, {}))


def safe_filename(name):
    safe = "".join(ch if ch.isalnum() or ch in "-_." else "_" for ch in str(name))
    return safe or "unknown"


def write_kind_outputs(templates_by_kind):
    os.makedirs(output_dir, exist_ok=True)
    index = []

    for kind in sorted(templates_by_kind.keys(), key=str):
        templates = sorted(
            templates_by_kind[kind].values(),
            key=lambda item: item["count"],
            reverse=True,
        )
        output = {
            "kind": kind,
            "total_count": sum(template["count"] for template in templates),
            "template_count": len(templates),
            "templates": [],
        }

        for template in templates:
            stat_template = build_stat_template(template["structure"], template["values"])
            output["templates"].append(stat_template)

        output_path = os.path.join(output_dir, f"{safe_filename(kind)}.json")
        with open(output_path, "w", encoding="utf-8") as fp:
            json.dump(output, fp, ensure_ascii=False, indent=2)
        index.append(
            {
                "kind": kind,
                "total_count": output["total_count"],
                "template_count": output["template_count"],
                "output_file": output_path,
            }
        )

    index_path = os.path.join(output_dir, "_index.json")
    with open(index_path, "w", encoding="utf-8") as fp:
        json.dump(index, fp, ensure_ascii=False, indent=2)

    return index_path


def main():
    all_files = collect_files()
    total_file_count = len(all_files)
    worker_count = min(NUM_WORKERS, cpu_count(), total_file_count) if total_file_count else 0
    print(f"找到 {total_file_count} 个 layers_caption.json, 使用 {worker_count} 个进程读取并统计...")

    templates_by_kind = defaultdict(dict)
    total_files = 0
    total_items = 0
    errors = []
    progress_step = max(1, total_file_count // 100)

    if not all_files:
        print("[WARN] 没有找到待处理的 layers_caption.json")
        return

    chunksize = max(1, min(64, total_file_count // (worker_count * 4) or 1))
    with Pool(worker_count) as pool:
        worker_pids = [worker.pid for worker in pool._pool]
        print(f"worker 进程 PID: {worker_pids}", flush=True)

        for result, err in pool.imap_unordered(process_file, all_files, chunksize=chunksize):
            total_files += 1
            if err is not None:
                errors.append(err)
            else:
                total_items += result["total_items"]
                for template in result["templates"]:
                    kind = template["kind"]
                    signature = template["signature"]
                    if signature not in templates_by_kind[kind]:
                        templates_by_kind[kind][signature] = {
                            "count": 0,
                            "structure": template["structure"],
                            "values": {},
                        }
                    merged = templates_by_kind[kind][signature]
                    merged["count"] += template["count"]
                    merged["structure"] = merge_structures(
                        merged["structure"],
                        template["structure"],
                    )
                    merge_values(merged["values"], template["values"])

            if total_files == total_file_count or total_files % progress_step == 0:
                percent = (total_files / total_file_count * 100) if total_file_count else 100
                print(
                    f"[读取进度] {total_files}/{total_file_count} "
                    f"({percent:.1f}%), 进程数={worker_count}, "
                    f"已统计 {total_items} 个含 kind 的条目",
                    flush=True,
                )

    for e in errors:
        print(f"[ERROR] {e}")

    index_path = write_kind_outputs(templates_by_kind)
    sorted_kinds = sorted(
        (
            (
                kind,
                sum(template["count"] for template in templates.values()),
                len(templates),
            )
            for kind, templates in templates_by_kind.items()
        ),
        key=lambda x: -x[1],
    )

    print("=" * 80)
    print(f"扫描完成: 共 {total_files} 个 layers_caption.json, {total_items} 个含 kind 的条目")
    print(f"共发现 {len(sorted_kinds)} 种 kind 类别")
    print(f"统计结果已保存到: {output_dir}")
    print(f"索引文件: {index_path}")
    print("=" * 80)

    print("\n### 各 kind 类别统计 ###\n")
    for i, (kind, cnt, template_count) in enumerate(sorted_kinds, 1):
        print(f"  {i:3d}. {kind:40s}  count={cnt}  templates={template_count}")


if __name__ == "__main__":
    main()