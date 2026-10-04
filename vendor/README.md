# vendor/ —— 离线依赖包

本目录存放本 Skill 运行所需的**全部第三方依赖 wheel**，用于**完全离线**的机器上安装。
不含 Python 解释器本身（需自行准备 **CPython 3.11**，pip 随解释器自带）。

> **通常不需要手动执行本目录下的任何命令。**
> `scripts/run_skill.py` / `scripts/verify.py` 启动时会先做依赖自检：齐全就跳过安装，
> 缺失就自动从本目录离线补齐（失败再退联网），补齐后继续执行。

## 平台约束

| 项 | 值 |
|---|---|
| 目标 Python | CPython **3.11**（cp311） |
| 目标平台 | **Windows x86_64**（win_amd64） |
| 轮子类型 | 全部为预编译 wheel，安装时**无需联网、无需编译** |

> 纯 Python 包（openpyxl / python-dateutil / six / et-xmlfile / tzdata）跨平台通用；
> `numpy` / `pandas` 是平台相关二进制轮，本目录为 **Windows 3.11** 版本。
> 若目标机为 Linux / macOS 或其它 Python 版本，见文末「其它平台」。

## 内容

| wheel | 版本 | 说明 |
|---|---|---|
| numpy | 2.4.6 | pandas 底层依赖 |
| pandas | 3.0.6 | 分组聚合与排序键 |
| openpyxl | 3.1.5 | 读写 xlsx、建表与美化 |
| python-dateutil | 2.9.0.post0 | pandas 依赖 |
| tzdata | 2026.5 | pandas 时区数据 |
| six | 1.17.0 | python-dateutil 依赖 |
| et-xmlfile | 2.0.0 | openpyxl 依赖 |

精确版本见 `requirements.lock.txt`。

## 安装（四选一）

```bash
# A. 什么都不做
#     run_skill.py / verify.py 启动时自动自检并按需从本目录离线补齐

# B. 工程提供的一键脚本（幂等，已装则跳过）
./install_offline.sh          # bash / Git Bash
install_offline.bat           # Windows 双击或 cmd
#    等价于：python scripts/deps_check.py --offline

# C. 手动，锁定版本
python -m pip install --no-index --find-links vendor -r vendor/requirements.lock.txt

# D. 手动，只按工程宽松约束（>=）从本地解析
python -m pip install --no-index --find-links vendor -r requirements.txt
```

关键参数：`--no-index` 关闭 PyPI 索引、`--find-links vendor` 只从本地目录找包 ⇒ 断网可装。

只检测不安装（环境巡检，退出码 1 = 有缺失）：

```bash
python scripts/deps_check.py --list
python scripts/deps_check.py --check
```

## 验证

```bash
python -c "import pandas, numpy, openpyxl; print(pandas.__version__, numpy.__version__, openpyxl.__version__)"
# 期望：3.0.6 2.4.6 3.1.5
```

## 其它平台

在**目标平台**的联网机器上重新拉取（示例：Linux x86_64 + Py3.11）：

```bash
pip download -r requirements.txt -d vendor \
    --only-binary=:all: --platform manylinux2014_x86_64 \
    --python-version 3.11 --implementation cp
```

Windows + Py3.11（复现本目录）：

```bat
pip download -r requirements.txt -d vendor --only-binary=:all: ^
    --platform win_amd64 --python-version 3.11 --implementation cp
```

> `pip download` 会自动带上全部**传递依赖**（如 pandas → python-dateutil / tzdata），
> 无需手工列出。升级依赖后请同步更新 `requirements.lock.txt`。
