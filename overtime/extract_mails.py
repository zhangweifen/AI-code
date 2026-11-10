#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从源文件夹按文件名关键词复制文件到目标文件夹。

规则：文件名含「加班」或「服务时长」才复制（任意扩展名）。

去重：按文件内容 SHA256，不按文件名。
- 内容完全相同 → 只保留一份
- 文件名相同但内容不同 → 都会复制，目标侧自动改名
"""

from __future__ import annotations

import hashlib
import shutil
from pathlib import Path

# 【改动1】只保留文件名包含关键词；不再按邮件扩展名/日报月报排除
INCLUDE_KEYWORDS = ("加班", "服务时长")

CHUNK_SIZE = 1024 * 1024  # 1MB


def ask_dir(prompt: str, *, must_exist: bool) -> Path:
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            print("路径不能为空，请重新输入。")
            continue
        path = Path(raw).expanduser().resolve()
        if must_exist and not path.is_dir():
            print(f"目录不存在或不是文件夹：{path}")
            continue
        return path


def should_copy(path: Path) -> bool:
    """【改动2】仅当文件名包含关键词时才复制。"""
    return any(k in path.name for k in INCLUDE_KEYWORDS)


def file_sha256(path: Path) -> str:
    """按内容计算哈希，用于判断是否真重复。"""
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        while True:
            chunk = fh.read(CHUNK_SIZE)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def unique_dest(dest_dir: Path, name: str) -> Path:
    """目标已存在同名文件时，自动追加 _1, _2 ..."""
    candidate = dest_dir / name
    if not candidate.exists():
        return candidate
    stem = Path(name).stem
    suffix = Path(name).suffix
    index = 1
    while True:
        candidate = dest_dir / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def load_existing_hashes(dest_dir: Path) -> dict[str, str]:
    """扫描目标目录已有文件，建立 内容哈希 → 文件名 映射，支持重复跑脚本。"""
    seen: dict[str, str] = {}
    for path in dest_dir.iterdir():
        if not path.is_file():
            continue
        try:
            seen[file_sha256(path)] = path.name
        except OSError as exc:
            print(f"[警告] 无法读取目标文件 {path.name}: {exc}")
    return seen


def main() -> int:
    print("=" * 50)
    print("按文件名关键词提取（内容去重）")
    print("=" * 50)
    # 【改动3】启动提示只显示关键词
    print(f"文件名需包含：{', '.join(INCLUDE_KEYWORDS)}")
    print("去重方式：内容 SHA256（同名不同内容会保留）")
    print()

    src_dir = ask_dir("请输入源文件夹路径：", must_exist=True)
    dst_dir = ask_dir("请输入目标文件夹路径：", must_exist=False)

    recursive = input("是否递归子文件夹？(Y/n，默认 Y)：").strip().lower()
    recursive = recursive not in {"n", "no", "否"}

    dst_dir.mkdir(parents=True, exist_ok=True)

    print()
    print("正在读取目标目录已有文件，建立去重索引...")
    seen_hashes = load_existing_hashes(dst_dir)
    if seen_hashes:
        print(f"目标目录已有 {len(seen_hashes)} 个文件纳入去重。")

    files = src_dir.rglob("*") if recursive else src_dir.iterdir()
    copied = 0
    skipped_rule = 0
    skipped_dup = 0
    failed = 0

    print()
    print(f"开始扫描：{src_dir}")
    print(f"输出到：{dst_dir}")
    print("-" * 50)

    for path in files:
        if not path.is_file():
            continue

        try:
            path.resolve().relative_to(dst_dir)
            continue
        except ValueError:
            pass

        name = path.name
        if not should_copy(path):
            skipped_rule += 1
            continue

        try:
            digest = file_sha256(path)
        except OSError as exc:
            failed += 1
            print(f"[失败] 无法读取 {name}: {exc}")
            continue

        if digest in seen_hashes:
            skipped_dup += 1
            print(f"[重复跳过] {name}  已有：{seen_hashes[digest]}")
            continue

        target = unique_dest(dst_dir, name)
        try:
            shutil.copy2(path, target)
            seen_hashes[digest] = target.name
            copied += 1
            note = f" -> {target.name}" if target.name != name else ""
            print(f"[复制] {name}{note}")
        except OSError as exc:
            failed += 1
            print(f"[失败] {name}: {exc}")

    print("-" * 50)
    print(
        f"完成：复制 {copied} 个，规则跳过 {skipped_rule} 个，"
        f"内容重复跳过 {skipped_dup} 个，失败 {failed} 个"
    )
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n已取消。")
        raise SystemExit(130)
