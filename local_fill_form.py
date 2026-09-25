#!/usr/bin/env python3
"""Local placeholder replacement for job-application DOCX files.

The script has three independent jobs:

* prepare: replace the sensitive/basic-information cells in the supplied
  template with placeholders before the DOCX is sent to an AI.
* check: list unresolved placeholders without revealing any local values.
* fill: replace placeholders from a local INI file or an encrypted vault.

The DOCX replacement is performed at the OOXML level so images, tables and
most existing formatting are copied unchanged. If a placeholder is split
across Word runs, the replacement is kept in the first run's formatting.
"""

from __future__ import annotations

import argparse
import base64
import configparser
import getpass
import json
import os
import re
import secrets
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Sequence, Tuple

try:
    from lxml import etree
except ImportError:  # pragma: no cover - gives a useful message on a new machine
    etree = None


TOKEN_RE = re.compile(r"\{\{([A-Z][A-Z0-9_]*)\}\}")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def die(message: str, code: int = 2) -> None:
    print(f"错误：{message}", file=sys.stderr)
    raise SystemExit(code)


def require_lxml() -> None:
    if etree is None:
        die("缺少 lxml。请运行：python -m pip install lxml")


def docx_xml_names(zf: zipfile.ZipFile) -> Iterable[str]:
    """Return Word XML parts that may contain visible text."""
    for name in zf.namelist():
        if name.startswith("word/") and name.endswith(".xml"):
            yield name


def load_xml(data: bytes):
    require_lxml()
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser)


def paragraph_text_nodes(paragraph) -> List[object]:
    """Get w:t nodes belonging directly to this w:p, excluding nested w:p."""
    result = []
    for node in paragraph.xpath(".//w:t", namespaces={"w": W_NS}):
        nearest_p = node.xpath("ancestor::w:p[1]", namespaces={"w": W_NS})
        if nearest_p and nearest_p[0] is paragraph:
            result.append(node)
    return result


def node_spans(nodes: Sequence[object]) -> List[Tuple[int, int]]:
    spans: List[Tuple[int, int]] = []
    cursor = 0
    for node in nodes:
        text = node.text or ""
        spans.append((cursor, cursor + len(text)))
        cursor += len(text)
    return spans


def locate_position(spans: Sequence[Tuple[int, int]], position: int) -> Tuple[int, int]:
    """Map a flattened text position to (node index, local offset)."""
    if not spans:
        raise ValueError("no text nodes")
    for index, (start, end) in enumerate(spans):
        if position < end or (position == end and start == end):
            return index, position - start
    index = len(spans) - 1
    return index, spans[index][1] - spans[index][0]


def replace_in_paragraph(paragraph, values: Mapping[str, str]) -> int:
    nodes = paragraph_text_nodes(paragraph)
    if not nodes:
        return 0
    text = "".join(node.text or "" for node in nodes)
    matches = list(TOKEN_RE.finditer(text))
    if not matches:
        return 0

    # Work from right to left. Positions before the current match are then
    # unchanged even when a replacement has a different length.
    replaced = 0
    for match in reversed(matches):
        key = match.group(1)
        if key not in values:
            continue
        replacement = str(values[key])
        spans = node_spans(nodes)
        start_i, start_o = locate_position(spans, match.start())
        end_i, end_o = locate_position(spans, match.end())

        if start_i == end_i:
            current = nodes[start_i].text or ""
            nodes[start_i].text = current[:start_o] + replacement + current[end_o:]
        else:
            start_text = nodes[start_i].text or ""
            end_text = nodes[end_i].text or ""
            nodes[start_i].text = start_text[:start_o] + replacement
            for index in range(start_i + 1, end_i):
                nodes[index].text = ""
            nodes[end_i].text = end_text[end_o:]
        replaced += 1
    return replaced


def scan_docx(path: Path) -> List[str]:
    require_lxml()
    found = set()
    with zipfile.ZipFile(path, "r") as zf:
        for name in docx_xml_names(zf):
            root = load_xml(zf.read(name))
            for paragraph in root.xpath(".//w:p", namespaces={"w": W_NS}):
                text = "".join(node.text or "" for node in paragraph_text_nodes(paragraph))
                found.update(match.group(1) for match in TOKEN_RE.finditer(text))
    return sorted(found)


def replace_docx(input_path: Path, output_path: Path, values: Mapping[str, str]) -> int:
    require_lxml()
    if input_path.resolve() == output_path.resolve():
        die("输入文件和输出文件不能相同；请使用新的输出文件名")
    if output_path.exists():
        die(f"输出文件已存在：{output_path}。如需覆盖，请先移走它")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    total = 0
    fd, temp_name = tempfile.mkstemp(prefix=".docx-fill-", suffix=".tmp", dir=output_path.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(input_path, "r") as source, zipfile.ZipFile(temp_path, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename in set(docx_xml_names(source)):
                    root = load_xml(data)
                    for paragraph in root.xpath(".//w:p", namespaces={"w": W_NS}):
                        total += replace_in_paragraph(paragraph, values)
                    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                target.writestr(info, data)
        os.replace(temp_path, output_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return total


def read_ini(path: Path) -> Dict[str, str]:
    if not path.exists():
        die(f"找不到资料库：{path}")
    parser = configparser.ConfigParser(interpolation=None, empty_lines_in_values=False)
    parser.optionxform = str
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            parser.read_file(handle)
    except (OSError, configparser.Error) as exc:
        die(f"无法读取资料库：{exc}")
    if not parser.has_section("profile"):
        die("资料库必须包含 [profile] 段")
    return {key.strip(): value.strip() for key, value in parser.items("profile")}


def sync_profile_file(path: Path, tokens: Sequence[str]) -> List[str]:
    """Append missing document variables without changing existing values."""
    wanted = sorted(set(tokens))
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("; 此文件由 local_fill_form.py 自动同步。不要上传。\n\n[profile]\n", encoding="utf-8")
    existing = read_ini(path)
    missing = [token for token in wanted if token not in existing]
    if not missing:
        return []
    try:
        original = path.read_text(encoding="utf-8-sig")
        with path.open("a", encoding="utf-8", newline="\n") as handle:
            if original and not original.endswith(("\n", "\r")):
                handle.write("\n")
            handle.write("\n; 以下字段由当前 Word 占位符自动补齐。\n")
            for token in missing:
                handle.write(f"{token} =\n")
    except OSError as exc:
        die(f"无法同步资料库：{exc}")
    return missing


def derive_key(password: str, salt: bytes, iterations: int):
    try:
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    except ImportError:
        die("加密资料库需要 cryptography。请运行：python -m pip install cryptography")
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iterations)
    return base64.urlsafe_b64encode(kdf.derive(password.encode("utf-8")))


def write_vault(profile: Mapping[str, str], output: Path, overwrite: bool) -> None:
    if output.exists() and not overwrite:
        die(f"加密资料库已存在：{output}。如需覆盖，请显式加 --overwrite")
    password = getpass.getpass("设置资料库密码（不会显示）：")
    confirm = getpass.getpass("再次输入资料库密码：")
    if not password or password != confirm:
        die("密码为空或两次输入不一致")
    try:
        from cryptography.fernet import Fernet
    except ImportError:
        die("加密资料库需要 cryptography。请运行：python -m pip install cryptography")
    salt = secrets.token_bytes(16)
    iterations = 600_000
    key = derive_key(password, salt, iterations)
    plaintext = json.dumps(dict(profile), ensure_ascii=False, sort_keys=True).encode("utf-8")
    payload = {
        "version": 1,
        "kdf": "PBKDF2-HMAC-SHA256",
        "iterations": iterations,
        "salt": base64.b64encode(salt).decode("ascii"),
        "ciphertext": Fernet(key).encrypt(plaintext).decode("ascii"),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已创建加密资料库：{output}")


def read_vault(
    path: Path,
    sync_tokens: Sequence[str] | None = None,
    password: str | None = None,
) -> Dict[str, str]:
    if not path.exists():
        die(f"找不到加密资料库：{path}")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("version") != 1:
            die("不支持的加密资料库版本")
        salt = base64.b64decode(payload["salt"])
        iterations = int(payload["iterations"])
        from cryptography.fernet import Fernet, InvalidToken
    except (OSError, ValueError, KeyError, ImportError) as exc:
        die(f"无法读取加密资料库：{exc}")
    if password is None:
        password = getpass.getpass("输入资料库密码（不会显示）：")
    try:
        plaintext = Fernet(derive_key(password, salt, iterations)).decrypt(
            payload["ciphertext"].encode("ascii")
        )
        values = json.loads(plaintext.decode("utf-8"))
    except (InvalidToken, ValueError, KeyError, UnicodeDecodeError, json.JSONDecodeError):
        die("密码错误或资料库已损坏")
    if not isinstance(values, dict):
        die("资料库内容格式错误")
    normalized = {str(key).strip(): str(value).strip() for key, value in values.items()}
    missing = sorted(set(sync_tokens or []) - set(normalized))
    if missing:
        normalized.update({token: "" for token in missing})
        payload["ciphertext"] = Fernet(derive_key(password, salt, iterations)).encrypt(
            json.dumps(normalized, ensure_ascii=False, sort_keys=True).encode("utf-8")
        ).decode("ascii")
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"已同步加密资料库：新增 {len(missing)} 个字段（未显示字段值）。")
    return normalized


def placeholders_from_values(values: Mapping[str, str]) -> Dict[str, str]:
    return {f"{{{{{key}}}}}": value for key, value in values.items() if key}


def prepare_docx(input_path: Path, output_path: Path, mapping_path: Path, include_yellow: bool) -> int:
    require_lxml()
    if output_path.exists():
        die(f"输出文件已存在：{output_path}。如需覆盖，请先移走它")
    try:
        mapping = json.loads(mapping_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        die(f"无法读取字段映射：{exc}")
    fields = mapping.get("fields")
    if not isinstance(fields, list):
        die("字段映射必须包含 fields 数组")

    changed = 0
    temp_fd, temp_name = tempfile.mkstemp(prefix=".docx-prepare-", suffix=".tmp", dir=output_path.parent or Path("."))
    os.close(temp_fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(input_path, "r") as source, zipfile.ZipFile(temp_path, "w") as target:
            document_name = "word/document.xml"
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == document_name:
                    root = load_xml(data)
                    tables = root.xpath(".//w:tbl", namespaces={"w": W_NS})
                    for field in fields:
                        if field.get("kind") == "yellow" and not include_yellow:
                            continue
                        try:
                            table_index = int(field.get("table", 0))
                            row_index = int(field["row"])
                            label_col = int(field["label_col"])
                            value_col = int(field["value_col"])
                            label = str(field["label"])
                            token = "{{" + str(field["token"]) + "}}"
                            table = tables[table_index]
                            rows = table.xpath("./w:tr", namespaces={"w": W_NS})
                            cells = rows[row_index].xpath("./w:tc", namespaces={"w": W_NS})
                            label_cell = cells[label_col]
                            value_cell = cells[value_col]
                            actual_label = "".join(
                                node.text or ""
                                for node in label_cell.xpath(".//w:t", namespaces={"w": W_NS})
                            ).strip()
                        except (KeyError, TypeError, ValueError, IndexError) as exc:
                            die(f"字段映射无效：{field!r}（{exc}）")
                        if label not in actual_label:
                            die(f"模板与字段映射不一致：第 {row_index} 行期望标签 {label!r}，实际为 {actual_label!r}")
                        value_nodes = value_cell.xpath(".//w:t", namespaces={"w": W_NS})
                        if not value_nodes:
                            # Empty Word cells often contain a styled w:r but
                            # no w:t. Add the text node inside that run so the
                            # cell keeps its original font and paragraph style.
                            runs = value_cell.xpath(".//w:r", namespaces={"w": W_NS})
                            if runs:
                                value_nodes = [etree.SubElement(runs[0], W + "t")]
                            else:
                                paragraphs = value_cell.xpath("./w:p", namespaces={"w": W_NS})
                                if not paragraphs:
                                    paragraph = etree.SubElement(value_cell, W + "p")
                                else:
                                    paragraph = paragraphs[0]
                                run = etree.SubElement(paragraph, W + "r")
                                value_nodes = [etree.SubElement(run, W + "t")]
                        value_nodes[0].text = token
                        for node in value_nodes[1:]:
                            node.text = ""
                        changed += 1
                    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                target.writestr(info, data)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        os.replace(temp_path, output_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return changed


def cmd_check(args: argparse.Namespace) -> None:
    tokens = scan_docx(Path(args.input))
    if args.map and args.kind:
        try:
            mapping = json.loads(Path(args.map).read_text(encoding="utf-8-sig"))
            kinds = {
                str(field["token"]): str(field.get("kind", ""))
                for field in mapping.get("fields", [])
                if "token" in field
            }
        except (OSError, json.JSONDecodeError, TypeError, KeyError) as exc:
            die(f"无法读取检查映射：{exc}")
        tokens = [token for token in tokens if kinds.get(token) == args.kind]
    if tokens:
        print("未解决的占位符：")
        for token in tokens:
            print(f"  {{{{{token}}}}}")
        raise SystemExit(1)
    print("检查通过：未发现占位符。")


def cmd_fill(args: argparse.Namespace) -> None:
    input_path = Path(args.input)
    output_path = Path(args.output)
    if args.profile and args.vault:
        die("--profile 和 --vault 只能二选一")
    if not args.profile and not args.vault:
        die("请提供 --profile 或 --vault")
    if args.password_stdin and args.profile:
        die("--password-stdin 只能配合 --vault 使用")
    tokens = scan_docx(input_path)
    if args.profile:
        added = sync_profile_file(Path(args.profile), tokens)
        if added:
            print(f"已同步资料库：新增 {len(added)} 个字段（未显示字段值）。")
        values = read_ini(Path(args.profile))
    else:
        password = sys.stdin.readline().rstrip("\r\n") if args.password_stdin else None
        if not password:
            die("未从标准输入收到资料库密码")
        values = read_vault(Path(args.vault), tokens, password=password)
    missing = [token for token in tokens if not values.get(token, "").strip()]
    if missing:
        print("缺少本地字段（未打印字段值）：", file=sys.stderr)
        for token in missing:
            print(f"  {token}", file=sys.stderr)
        raise SystemExit(1)
    replacements = {token: values[token] for token in tokens}
    count = replace_docx(input_path, output_path, replacements)
    leftovers = scan_docx(output_path)
    if leftovers:
        die("替换后仍有占位符：" + ", ".join(leftovers))
    print(f"已生成完整信息表：{output_path}（替换 {count} 处）")


def cmd_prepare(args: argparse.Namespace) -> None:
    output = Path(args.output)
    count = prepare_docx(Path(args.input), output, Path(args.mapping), not args.skip_yellow)
    profile_path = Path(args.profile) if args.profile else Path(args.mapping).with_name("profile.ini")
    added = sync_profile_file(profile_path, scan_docx(output))
    suffix = f"；已同步资料库并新增 {len(added)} 个字段" if added else "；资料库变量已是最新"
    print(f"已生成 AI 输入版：{output}（写入 {count} 个占位符{suffix}）")


def cmd_sync_profile(args: argparse.Namespace) -> None:
    tokens = scan_docx(Path(args.input))
    if args.profile and args.vault:
        die("--profile 和 --vault 只能二选一")
    if not args.profile and not args.vault:
        die("请提供 --profile 或 --vault")
    if args.profile:
        added = sync_profile_file(Path(args.profile), tokens)
        print(f"资料库已同步：新增 {len(added)} 个字段；已有值未改动。")
    else:
        read_vault(Path(args.vault), tokens)
        print("加密资料库变量已同步；已有值未改动。")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="本地处理网申 DOCX 占位符，不向 AI 发送敏感值")
    sub = parser.add_subparsers(dest="command", required=True)

    prepare = sub.add_parser("prepare", help="生成交给 AI 的占位符版本")
    prepare.add_argument("--input", required=True, help="原始 DOCX 模板")
    prepare.add_argument("--output", required=True, help="AI 输入版 DOCX")
    prepare.add_argument("--map", required=True, dest="mapping", help="field_map.json")
    prepare.add_argument("--profile", help="要同步的 profile.ini；默认使用 field_map.json 同目录的 profile.ini")
    prepare.add_argument("--skip-yellow", action="store_true", help="不替换黄色岗位字段")
    prepare.set_defaults(func=cmd_prepare)

    check = sub.add_parser("check", help="只扫描占位符，不读取本地资料库")
    check.add_argument("--input", required=True, help="AI 返回的 DOCX")
    check.add_argument("--map", help="字段映射；配合 --kind 只检查某一类字段")
    check.add_argument("--kind", choices=["local", "yellow"], help="只检查 local 或 yellow 字段")
    check.set_defaults(func=cmd_check)

    fill = sub.add_parser("fill", help="在本机替换占位符")
    fill.add_argument("--input", required=True, help="AI 返回的 DOCX")
    fill.add_argument("--output", required=True, help="最终 DOCX")
    fill.add_argument("--profile", help="明文 profile.ini（仅建议临时使用）")
    fill.add_argument("--vault", help="加密资料库 .vault")
    fill.add_argument(
        "--password-stdin",
        action="store_true",
        help="从标准输入读取 Vault 密码；仅供受限本地服务调用，不要在命令行中写密码",
    )
    fill.set_defaults(func=cmd_fill)

    sync = sub.add_parser("sync-profile", help="按 Word 占位符补齐本地资料库变量")
    sync.add_argument("--input", required=True, help="包含占位符的 DOCX")
    sync.add_argument("--profile", help="明文 profile.ini")
    sync.add_argument("--vault", help="加密资料库 .vault")
    sync.set_defaults(func=cmd_sync_profile)

    vault = sub.add_parser("vault", help="把 profile.ini 加密为本地资料库")
    vault.add_argument("--profile", required=True, help="明文 profile.ini")
    vault.add_argument("--output", required=True, help="输出 .vault 文件")
    vault.add_argument("--overwrite", action="store_true", help="允许覆盖已有 vault")
    vault.set_defaults(func=lambda a: write_vault(read_ini(Path(a.profile)), Path(a.output), a.overwrite))
    return parser


if __name__ == "__main__":
    parser = build_parser()
    arguments = parser.parse_args()
    arguments.func(arguments)
