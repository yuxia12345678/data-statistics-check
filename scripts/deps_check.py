# -*- coding: utf-8 -*-
"""
运行前系统依赖自检 / 自动补齐（离线优先）。

职责
----
在 Skill 真正跑起来之前，检测 `requirements.txt` 声明的第三方依赖是否齐全：
  - **齐全** → 什么都不装，直接放行（打一行状态即返回）；
  - **缺失** → 自动补齐：优先用工程内 `vendor/` 离线安装（`--no-index`，断网可用），
    失败再退到联网 `pip install -r requirements.txt`；
  - **补齐后仍缺失** → 抛 `DependencyError`，由调用方给出人话提示并终止。

被 `run_skill.py` / `verify.py` 在 `import pandas`、`import openpyxl` **之前**调用，
因此本模块自身**只依赖标准库**，绝不 import 任何第三方包。

命令行
------
    python scripts/deps_check.py            # 检测；缺失则自动安装（离线优先）
    python scripts/deps_check.py --check    # 只检测不安装（退出码 1 = 有缺失）
    python scripts/deps_check.py --offline  # 只允许离线安装（vendor/）
    python scripts/deps_check.py --list     # 打印依赖清单与当前状态
    python scripts/deps_check.py --quiet    # 静默，仅出错时输出

环境变量
--------
    DSC_NO_AUTO_INSTALL=1   等价于 --check（只检测、不安装）
"""

import argparse
import importlib
import importlib.metadata as md
import os
import re
import subprocess
import sys

# --------------------------------------------------------------------------- 路径
_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.normpath(os.path.join(_HERE, ".."))
_VENDOR = os.path.join(_ROOT, "vendor")
_REQUIREMENTS = os.path.join(_ROOT, "requirements.txt")
_LOCK = os.path.join(_VENDOR, "requirements.lock.txt")

# 发行包名（requirements 里写的）→ 顶层导入名
_TOP_MODULE = {
    "numpy": "numpy",
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "python-dateutil": "dateutil",
    "six": "six",
    "tzdata": "tzdata",
    "et-xmlfile": "et_xmlfile",
}

# requirements.txt 读不到时的兜底声明（与文件内容保持一致）
_FALLBACK_REQS = [("numpy", ">=", "1.24"), ("pandas", ">=", "2.0"), ("openpyxl", ">=", "3.1")]

_REQ_RE = re.compile(r"^\s*([A-Za-z0-9][A-Za-z0-9_.\-]*)\s*(==|>=|<=|~=|>|<)?\s*([0-9][^\s;#]*)?")

PREFIX = "[deps]"


class DependencyError(RuntimeError):
    """依赖无法满足（尝试安装后仍然缺失）。"""


# --------------------------------------------------------------------------- 输出
def _safe_stdout():
    """Windows 控制台/管道可能是 GBK，非 GBK 字符会抛 UnicodeEncodeError。"""
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass


def _log(msg, quiet=False):
    if quiet:
        return
    try:
        print(f"{PREFIX} {msg}")
    except Exception:
        try:
            sys.stderr.write(f"{PREFIX} {msg}\n")
        except Exception:
            pass


# --------------------------------------------------------------------------- 声明解析
def _parse_requirements(path=_REQUIREMENTS):
    """把 requirements.txt 解析成 [(发行名, 比较符, 版本), ...]。"""
    if not os.path.isfile(path):
        return list(_FALLBACK_REQS)
    reqs = []
    with open(path, "r", encoding="utf-8") as fh:
        for line in fh:
            line = line.split("#", 1)[0].strip()
            if not line or line.startswith("-"):
                continue
            m = _REQ_RE.match(line)
            if m:
                reqs.append((m.group(1), m.group(2) or "", m.group(3) or ""))
    return reqs or list(_FALLBACK_REQS)


def _top_module(dist):
    return _TOP_MODULE.get(dist.lower(), dist.replace("-", "_"))


# --------------------------------------------------------------------------- 版本比较
def _ver_tuple(text):
    """'2.9.0.post0' -> (2, 9, 0)；容忍非数字段（遇到即停）。"""
    out = []
    for part in re.split(r"[._\-+]", str(text).strip()):
        m = re.match(r"^(\d+)", part)
        if not m:
            break
        out.append(int(m.group(1)))
    return tuple(out)


def _satisfies(have, op, want):
    if not op or not want:
        return True
    a, b = _ver_tuple(have), _ver_tuple(want)
    n = max(len(a), len(b))
    a = a + (0,) * (n - len(a))
    b = b + (0,) * (n - len(b))
    return {
        "==": a == b,
        ">=": a >= b,
        "<=": a <= b,
        ">": a > b,
        "<": a < b,
        "~=": a >= b and a[: len(b)] == b[: len(b)],
    }.get(op, True)


def _installed_version(dist):
    """已安装版本；未安装返回 None。"""
    for name in (dist, dist.replace("_", "-"), dist.replace("-", "_")):
        try:
            return md.version(name)
        except Exception:
            continue
    return None


def _probe_import(top):
    """真正 import 一次，捕获「装了但加载不了」（DLL 缺失、二进制不匹配等）。"""
    try:
        importlib.import_module(top)
        return None
    except Exception as exc:
        return "%s: %s" % (type(exc).__name__, exc)


# --------------------------------------------------------------------------- 检测
def collect_missing(probe=True):
    """返回未满足的依赖列表，元素形如：

    {"name": "pandas", "top": "pandas", "have": "1.5.3", "need": ">=2.0", "reason": "版本不满足"}
    """
    missing = []
    for name, op, want in _parse_requirements():
        top = _top_module(name)
        have = _installed_version(name)
        need = "%s%s" % (op, want)
        if have is None:
            missing.append({"name": name, "top": top, "have": None,
                            "need": need, "reason": "未安装"})
            continue
        if not _satisfies(have, op, want):
            missing.append({"name": name, "top": top, "have": have,
                            "need": need, "reason": "版本不满足"})
            continue
        if probe:
            err = _probe_import(top)
            if err:
                missing.append({"name": name, "top": top, "have": have,
                                "need": need, "reason": "导入失败（%s）" % err})
    return missing


def status_lines(probe=True):
    """人类可读的依赖状态清单。"""
    bad = {m["name"] for m in collect_missing(probe=probe)}
    lines = []
    for name, op, want in _parse_requirements():
        have = _installed_version(name)
        mark = "MISS" if name in bad else "OK"
        lines.append("%-4s  %-16s 需要 %-8s 当前 %s"
                     % (mark, name, ("%s%s" % (op, want)) or "-", have or "未安装"))
    return lines


def _summary():
    """形如 'numpy 2.4.6 / pandas 3.0.6 / openpyxl 3.1.5'（只列直接依赖）。"""
    parts = []
    for name, _op, _want in _parse_requirements():
        parts.append("%s %s" % (name, _installed_version(name) or "未安装"))
    return " / ".join(parts)


# --------------------------------------------------------------------------- 安装
def _reset_modules():
    """把已导入（尤其失败了一半）的顶层模块清出 sys.modules，让重装后的包能被重新发现。"""
    tops = {_top_module(n) for n, _o, _w in _parse_requirements()}
    for mod in list(sys.modules):
        if mod.split(".")[0] in tops:
            sys.modules.pop(mod, None)


def _pip(extra, quiet=False, offline=False):
    cmd = [sys.executable, "-m", "pip", "install", "--disable-pip-version-check"]
    if quiet:
        cmd.append("--quiet")
    if offline:
        cmd += ["--no-index", "--find-links", _VENDOR]
    cmd += list(extra)
    try:
        return subprocess.call(cmd)
    except OSError as exc:
        _log("无法调用 pip：%s（请确认 %s 自带 pip）" % (exc, sys.executable), quiet)
        return 127


def ensure_dependencies(auto_install=True, offline_only=False, quiet=False, probe=True):
    """检测并按需补齐依赖。返回缺失列表（空列表 = 依赖齐全）。

    失败时抛 `DependencyError`（调用方负责转成人话并终止）。
    """
    if os.environ.get("DSC_NO_AUTO_INSTALL") == "1":
        auto_install = False

    missing = collect_missing(probe=probe)
    if not missing:
        _log("依赖齐全，跳过安装 -> %s" % _summary(), quiet)
        return []

    brief = ", ".join("%s(%s)" % (m["name"], m["reason"]) for m in missing)
    if not auto_install:
        _log("依赖不齐全：%s（已指定不自动安装）" % brief, quiet)
        raise DependencyError("缺少依赖：" + brief)

    _log("依赖不齐全：%s -> 开始自动补齐" % brief, quiet)

    attempts = []

    # 1) 离线优先：工程内 vendor/
    if os.path.isdir(_VENDOR):
        req_file = _LOCK if os.path.isfile(_LOCK) else _REQUIREMENTS
        _log("离线安装：pip install --no-index --find-links vendor -r %s"
             % os.path.relpath(req_file, _ROOT).replace("\\", "/"), quiet)
        rc = _pip(["-r", req_file], quiet=quiet, offline=True)
        attempts.append(("离线 vendor/", rc))
        if rc == 0:
            _reset_modules()
            if not collect_missing(probe=probe):
                _log("离线安装完成，依赖已齐全 -> %s" % _summary(), quiet)
                return []
        else:
            _log("离线安装失败（rc=%d）" % rc, quiet)
    else:
        _log("未找到 vendor/ 目录，跳过离线安装", quiet)

    # 2) 联网兜底
    if not offline_only:
        _log("联网安装：pip install -r requirements.txt", quiet)
        rc = _pip(["-r", _REQUIREMENTS], quiet=quiet)
        attempts.append(("联网", rc))
        _reset_modules()
        if rc == 0 and not collect_missing(probe=probe):
            _log("联网安装完成，依赖已齐全 -> %s" % _summary(), quiet)
            return []
    elif not quiet:
        _log("已指定 --offline，跳过联网安装", quiet)

    still = collect_missing(probe=probe)
    detail = "；".join("%s %s" % (k, "成功" if v == 0 else "失败(rc=%s)" % v)
                       for k, v in attempts) or "未执行任何安装"
    raise DependencyError(
        "安装后仍缺少依赖：%s\n"
        "      安装尝试：%s\n"
        "      排查建议：① 确认 Python 为 3.11；② 确认 vendor/ 中 wheel 完整；"
        "③ 手动执行\n"
        "        %s -m pip install --no-index --find-links vendor -r vendor/requirements.lock.txt"
        % (", ".join("%s(%s)" % (m["name"], m["reason"]) for m in still), detail, sys.executable)
    )


# --------------------------------------------------------------------------- CLI
def main(argv=None):
    _safe_stdout()
    ap = argparse.ArgumentParser(description="运行前系统依赖自检 / 自动安装（离线优先）")
    ap.add_argument("--check", action="store_true", help="只检测，不安装依赖")
    ap.add_argument("--offline", action="store_true", help="只允许离线安装（vendor/）")
    ap.add_argument("--quiet", action="store_true", help="静默模式，仅出错时输出")
    ap.add_argument("--no-probe", action="store_true", help="跳过导入探测，只比对已安装版本")
    ap.add_argument("--list", action="store_true", help="打印依赖清单与当前状态后退出")
    args = ap.parse_args(argv)

    probe = not args.no_probe

    if args.list:
        for line in status_lines(probe=probe):
            print(line)
        return 0 if not collect_missing(probe=probe) else 1

    try:
        ensure_dependencies(auto_install=not args.check, offline_only=args.offline,
                            quiet=args.quiet, probe=probe)
    except DependencyError as exc:
        print("❌ 依赖检查未通过：%s" % exc)
        return 1
    if args.check:
        _log("依赖检查通过 -> %s" % _summary(), args.quiet)
    return 0


if __name__ == "__main__":
    sys.exit(main())
