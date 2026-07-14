#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import shutil
from pathlib import Path


def project_root() -> Path:
    return Path(__file__).resolve().parents[4]


def resolve_target(root: Path, target: str) -> Path:
    path = Path(target)
    if not path.is_absolute():
        path = root / path
    return path.resolve()


def assert_under_root(root: Path, target: Path) -> None:
    root = root.resolve()
    if target == root or root not in target.parents:
        raise SystemExit(f"拒绝删除非项目子目录: {target}")


def main() -> None:
    parser = argparse.ArgumentParser(description="清理 FieldServiceAgent BI 下载临时数据")
    parser.add_argument("--target", default="temp/data", help="要清理的临时下载目录，默认 temp/data")
    parser.add_argument("--dry-run", action="store_true", help="只打印将要删除的目录")
    parser.add_argument("--delete-valid", action="store_true", help="删除整个目录，包括已通过校验的 xlsx")
    args = parser.parse_args()

    root = project_root()
    target = resolve_target(root, args.target)
    assert_under_root(root, target)

    if not target.exists():
        print(f"临时下载目录不存在，无需清理: {target}")
        return
    if not target.is_dir():
        raise SystemExit(f"目标不是目录，拒绝删除: {target}")

    if args.delete_valid:
        if args.dry_run:
            print(f"将删除临时下载目录: {target}")
            return
        shutil.rmtree(target)
        print(f"已删除临时下载目录: {target}")
        return

    stale_files = [
        path
        for path in target.rglob("*")
        if path.is_file() and (path.name.endswith(".part") or path.name.endswith(".invalid"))
    ]
    if args.dry_run:
        if stale_files:
            print("将删除失败/中间文件:")
            for path in stale_files:
                print(path)
        else:
            print(f"没有需要清理的失败/中间文件；保留已校验数据: {target}")
        return

    for path in stale_files:
        path.unlink()
    print(f"已清理失败/中间文件 {len(stale_files)} 个；保留已校验数据: {target}")


if __name__ == "__main__":
    main()
