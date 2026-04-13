# -*- coding: utf-8 -*-
"""
通用打包 + 数字签名脚本。

用法:
    python build_with_signing.py
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import importlib.util
from pathlib import Path


try:
    sys.path.insert(0, str(Path(__file__).parent))
    from code_signer.sign_exe_file import sign_exe_with_sha1, verify_exe_signature

    SIGNER_AVAILABLE = True
except ImportError as err:
    SIGNER_AVAILABLE = False
    print(f"[警告] 签名模块不可用: {err}")

    def sign_exe_with_sha1(exe_path):
        return False, "签名模块不可用"

    def verify_exe_signature(exe_path):
        return False, "签名模块不可用"


CONFIG = {
    "main_script": "New_Customer_Operate.py",
    "icon_file": "ch-2.ico",
    "exe_name": "New_Customer_Operate",
    "console": False,
}

PROJECT_ROOT = Path(__file__).resolve().parent


def print_header(text: str) -> None:
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)


def print_step(step_num: int, text: str) -> None:
    print(f"\n[步骤 {step_num}] {text}")
    print("-" * 60)


def check_files() -> bool:
    """检查打包所需文件是否存在。"""
    print_step(1, "检查必要文件")

    missing = []
    for relative_path in [CONFIG["main_script"], CONFIG["icon_file"]]:
        path = PROJECT_ROOT / relative_path
        exists = path.exists()
        print(f"  {'OK' if exists else 'NO'} {path}")
        if not exists:
            missing.append(str(path))

    print(
        f"  {'OK' if SIGNER_AVAILABLE else 'SKIP'} code_signer: "
        f"{'可用' if SIGNER_AVAILABLE else '不可用，将跳过签名'}"
    )

    if missing:
        print("\n缺少必要文件:")
        for path in missing:
            print(f"  - {path}")
        return False

    print("\n文件检查通过")
    return True


def clean_build_artifacts() -> None:
    """清理旧的打包产物。"""
    print_step(2, "清理旧的打包文件")

    removed = []
    for name in ["build", "dist", f"{CONFIG['exe_name']}.spec"]:
        path = PROJECT_ROOT / name
        if not path.exists():
            continue

        try:
            if path.is_dir():
                shutil.rmtree(path)
                print(f"  删除目录: {path}")
            else:
                path.unlink()
                print(f"  删除文件: {path}")
            removed.append(path)
        except Exception as exc:
            print(f"  删除失败: {path} ({exc})")

    if removed:
        print(f"\n已清理 {len(removed)} 项")
    else:
        print("\n无需清理")


def _candidate_exe_paths() -> list[Path]:
    exe_name = f"{CONFIG['exe_name']}.exe"
    return [
        PROJECT_ROOT / "dist" / exe_name,
        PROJECT_ROOT / "dist" / CONFIG["exe_name"] / exe_name,
        PROJECT_ROOT / "build" / exe_name,
        PROJECT_ROOT / "build" / CONFIG["exe_name"] / exe_name,
    ]


def find_built_exe() -> Path | None:
    """在常见输出目录中查找 PyInstaller 生成的 EXE。"""
    for path in _candidate_exe_paths():
        if path.exists() and path.is_file():
            return path

    search_roots = [PROJECT_ROOT / "dist", PROJECT_ROOT / "build"]
    discovered: list[Path] = []

    for root in search_roots:
        if not root.exists():
            continue
        discovered.extend(p for p in root.rglob("*.exe") if p.is_file())

    if not discovered:
        return None

    exact_matches = [
        path
        for path in discovered
        if path.name.lower() == f"{CONFIG['exe_name'].lower()}.exe"
    ]
    candidates = exact_matches or discovered
    candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
    return candidates[0]


def _print_recent_build_output(result: subprocess.CompletedProcess[str]) -> None:
    print("\nPyInstaller 输出摘要:")
    combined = "\n".join(
        part.strip() for part in [result.stdout, result.stderr] if part and part.strip()
    )
    if not combined:
        print("  无输出")
        return

    lines = [line for line in combined.splitlines() if line.strip()]
    for line in lines[-30:]:
        print(f"  {line}")


def _print_existing_exe_candidates() -> None:
    print("\n已扫描到的 EXE 文件:")
    found = False
    for root_name in ["dist", "build"]:
        root = PROJECT_ROOT / root_name
        if not root.exists():
            continue
        for path in root.rglob("*.exe"):
            if path.is_file():
                size_mb = path.stat().st_size / (1024 * 1024)
                print(f"  - {path} ({size_mb:.1f} MB)")
                found = True

    if not found:
        print("  未找到任何 EXE")


def build_exe() -> tuple[bool, str]:
    """使用 PyInstaller 打包 EXE。"""
    print_step(3, "开始打包")

    if importlib.util.find_spec("PyInstaller") is None:
        return False, "当前 Python 环境未安装 PyInstaller"

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--clean",
        "--noconfirm",
        f"--name={CONFIG['exe_name']}",
    ]

    cmd.append("--console" if CONFIG["console"] else "--windowed")

    icon_path = PROJECT_ROOT / CONFIG["icon_file"]
    if icon_path.exists():
        cmd.append(f"--icon={icon_path}")

    cmd.append(str(PROJECT_ROOT / CONFIG["main_script"]))

    print("  执行命令:")
    print(f"  {' '.join(str(item) for item in cmd)}")

    try:
        result = subprocess.run(
            cmd,
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception as exc:
        print(f"\n打包异常: {exc}")
        return False, str(exc)

    if result.returncode != 0:
        print("\n打包失败")
        _print_recent_build_output(result)
        return False, result.stderr.strip() or result.stdout.strip() or "PyInstaller 返回非 0"

    exe_path = find_built_exe()
    if exe_path is None:
        print("\n打包完成，但未能定位 EXE 文件")
        _print_recent_build_output(result)
        _print_existing_exe_candidates()
        return False, "打包完成但找不到 EXE 文件"

    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print("\n打包成功")
    print(f"  文件: {exe_path}")
    print(f"  大小: {size_mb:.1f} MB")
    return True, str(exe_path)


def main() -> bool:
    print_header("自动打包 + 数字签名")
    print(f"  主程序: {CONFIG['main_script']}")
    print(f"  图标:   {CONFIG['icon_file']}")
    print(f"  签名:   {'启用' if SIGNER_AVAILABLE else '跳过'}")

    start_time = time.time()

    try:
        if not check_files():
            input("\n按回车退出...")
            return False

        clean_build_artifacts()

        success, result = build_exe()
        if not success:
            print(f"\n打包失败: {result}")
            input("\n按回车退出...")
            return False

        exe_path = Path(result)

        print_step(4, "数字签名")
        sign_success = False
        sign_message = "签名模块不可用"
        if SIGNER_AVAILABLE:
            sign_success, sign_message = sign_exe_with_sha1(str(exe_path))
            if sign_success:
                verify_exe_signature(str(exe_path))
        else:
            print("  跳过签名")

        elapsed = time.time() - start_time
        print_header("打包完成")
        print(f"  文件: {exe_path}")
        print(f"  大小: {exe_path.stat().st_size / (1024 * 1024):.1f} MB")
        print(f"  签名: {'已签名' if sign_success else f'未签名（{sign_message}）'}")
        print(f"  耗时: {elapsed:.1f} 秒")

        try:
            if input("\n是否打开输出目录? (y/n): ").strip().lower() in {"y", "yes"}:
                output_dir = exe_path.parent
                if hasattr(os, "startfile"):
                    os.startfile(str(output_dir))
        except Exception:
            pass

        input("\n按回车退出...")
        return True

    except KeyboardInterrupt:
        print("\n\n用户取消操作")
        input("\n按回车退出...")
        return False
    except Exception as exc:
        import traceback

        print(f"\n\n发生错误: {exc}")
        traceback.print_exc()
        input("\n按回车退出...")
        return False


if __name__ == "__main__":
    main()
