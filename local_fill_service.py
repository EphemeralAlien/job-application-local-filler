#!/usr/bin/env python3
"""Unattended local filler for AI-produced, placeholder-only DOCX files.

The watcher accepts DOCX files from an AI-visible inbox and writes completed
files to a separate output directory. It never prints, logs, or returns local
profile values. The Vault password is supplied once on stdin by the launcher
and kept only in this process's memory.

This is an automation boundary, not an operating-system security boundary.
For a hard guarantee against an agent with arbitrary terminal access, run the
agent in a separate OS account, VM, or sandbox that cannot read the Vault or
the output directory.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import signal
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict


SAFE_NAME_RE = re.compile(r"[^A-Za-z0-9._-]+")
STOP = False


def stop_service(_signum, _frame) -> None:
    global STOP
    STOP = True


def safe_name(name: str) -> str:
    cleaned = SAFE_NAME_RE.sub("_", name).strip("._")
    return cleaned or "job"


def write_status(status_dir: Path, job_name: str, status: str, output_name: str | None = None) -> None:
    """Write status metadata only; never include command output or local values."""
    status_dir.mkdir(parents=True, exist_ok=True)
    payload: Dict[str, str] = {"job": job_name, "status": status}
    if output_name:
        payload["output"] = output_name
    target = status_dir / f"{safe_name(job_name)}.json"
    temporary = target.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, target)


def process_job(
    input_path: Path,
    output_dir: Path,
    failed_dir: Path,
    status_dir: Path,
    vault_path: Path,
    fill_script: Path,
    password: str,
) -> None:
    job_name = input_path.name
    base_name = input_path.stem
    if base_name.endswith(".ready"):
        base_name = base_name[:-6]
    output_name = f"{safe_name(base_name)}_完整信息表.docx"
    output_path = output_dir / output_name
    if output_path.exists():
        write_status(status_dir, job_name, "output_exists")
        input_path.replace(failed_dir / input_path.name)
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    failed_dir.mkdir(parents=True, exist_ok=True)
    command = [
        sys.executable,
        str(fill_script),
        "fill",
        "--input",
        str(input_path),
        "--output",
        str(output_path),
        "--vault",
        str(vault_path),
        "--password-stdin",
    ]
    try:
        result = subprocess.run(
            command,
            input=password + "\n",
            text=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=300,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        result = None

    if result is not None and result.returncode == 0 and output_path.exists():
        write_status(status_dir, job_name, "completed", output_name)
        input_path.unlink()
    else:
        write_status(status_dir, job_name, "needs_attention")
        input_path.replace(failed_dir / input_path.name)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="后台本地网申填表服务")
    parser.add_argument("--inbox", required=True, help="AI 可写入的占位符 DOCX 投递目录")
    parser.add_argument("--outbox", required=True, help="仅用户可读取的完整表格输出目录")
    parser.add_argument("--status", required=True, help="只写入处理状态的目录")
    parser.add_argument("--failed", required=True, help="处理失败文件目录")
    parser.add_argument("--vault", required=True, help="本地加密资料库；不会输出其内容")
    parser.add_argument("--fill-script", required=True, help="local_fill_form.py")
    parser.add_argument("--password-stdin", action="store_true", help="从标准输入读取一次 Vault 密码")
    parser.add_argument("--interval", type=float, default=2.0, help="扫描间隔秒数")
    parser.add_argument("--pid-file", help="服务进程 ID 文件")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if not args.password_stdin:
        print("服务必须使用 --password-stdin；不会在命令行参数中接收密码。", file=sys.stderr)
        return 2
    password = sys.stdin.readline().rstrip("\r\n")
    if not password:
        return 2

    inbox = Path(args.inbox).resolve()
    outbox = Path(args.outbox).resolve()
    status = Path(args.status).resolve()
    failed = Path(args.failed).resolve()
    vault = Path(args.vault).resolve()
    fill_script = Path(args.fill_script).resolve()
    inbox.mkdir(parents=True, exist_ok=True)
    status.mkdir(parents=True, exist_ok=True)
    failed.mkdir(parents=True, exist_ok=True)
    outbox.mkdir(parents=True, exist_ok=True)
    if args.pid_file:
        pid_path = Path(args.pid_file).resolve()
        pid_path.parent.mkdir(parents=True, exist_ok=True)
        pid_path.write_text(str(os.getpid()), encoding="ascii")
    else:
        pid_path = None

    signal.signal(signal.SIGINT, stop_service)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, stop_service)
    try:
        while not STOP:
            for candidate in sorted(inbox.glob("*.docx")):
                if STOP:
                    break
                if candidate.name.startswith("~$") or candidate.name.startswith("."):
                    continue
                process_job(candidate, outbox, failed, status, vault, fill_script, password)
            time.sleep(max(0.5, args.interval))
    finally:
        password = ""
        if pid_path and pid_path.exists():
            pid_path.unlink()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
