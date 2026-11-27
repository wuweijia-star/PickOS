<a id="top"></a>

**简体中文** | [English](#english)

# PickOS 班级管理系统

> 一款面向中学班主任 / 班干部的班级积分与成绩管理工具。



![Python](https://img.shields.io/badge/Python-3.12-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-6.5+-green)
![License](https://img.shields.io/badge/License-GPLv3-orange)

[![Download](https://img.shields.io/github/v/release/wuweijia-star/PickOS?label=下载最新版&style=for-the-badge&color=2ea44f)](https://github.com/wuweijia-star/PickOS/releases/latest)

<div align="center">
  <img src="assets/demo.gif" width="80%" alt="PickOS 演示动图" />
</div>

PickOS 完全离线运行，所有数据均保存在本机 `user/` 目录中，不上传任何服务器。

## ✨ 功能简介

- 📝 **积分录入**：支持单条 / 连续批量录入，提供事由管理、折扣与赋分
- 📊 **排行榜**：实时排名、搜索高亮、单科王
- 🎯 **成绩总览**：本次与上次成绩对比、进步 / 退步榜、考试加分规则
- 📅 **首页看板**：便签、日历、倒计时、设备监控、签到
- 🏆 **等级与成就**：20 级经验体系 + 40+ 成就系统
- 🔐 **密码保护 / 智能解锁**：支持 U 盘插入自动解锁
- 🧭 **首次启动引导**：Hello 动画、用户协议 / 隐私政策确认、快捷方式 / 名单 / 科目 / 密码一键配置
- 🎨 **主题切换**：深色 / 浅色 / 自定义
- 💾 **数据管理**：导出 Excel、自动备份、存储空间分析
- 🔄 **在线更新**：内置一键检查更新，支持多镜像加速下载安装包

## 🖥️ 运行环境

- **Python 3.12+**（源码运行）
- **操作系统**：Windows 10 / 11（主要支持）/ macOS / Linux

## 🚀 快速开始

### 方式一：下载安装包（推荐普通用户）

前往 [Releases](https://github.com/wuweijia-star/PickOS/releases) 页面，下载最新安装包：

```text
PickOS_Setup_vX.Y.Z_win64.exe
```

> 实际文件名以 Releases 页面显示为准。

安装与升级说明：

1. 双击运行安装包，按提示点击“下一步”即可。
2. **安装过程不需要管理员权限**，默认安装到当前用户目录，不会触发 UAC 权限拦截。
3. 从旧版本升级时，安装包会自动覆盖旧程序与运行库，**不会删除 `user` 数据文件夹**。
4. 自动更新会保留 `user/data`、`user/exports`、`user/backups` 等用户数据。
5. 新版本发布后，打开软件并点击「设置 -> 检查更新」，即可一键下载并静默升级。

自动更新调用安装包时使用以下静默参数：

```text
/VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

### 方式二：直接运行源码（推荐开发者）

```bash
git clone https://github.com/wuweijia-star/PickOS.git
cd PickOS
pip install -r requirements.txt
python main.pyw
```

> 也可以直接双击 `build.bat` 快速打包生成免安装绿色版 exe。

## 📁 项目结构

```text
PickOS/
├── main.pyw              # 主程序
├── build.bat             # Windows 本地打包脚本
├── requirements.txt      # 依赖清单
├── LICENSE               # GPLv3 协议
├── README.md
├── dist/                 # 打包后生成的运行目录（包含所有依赖）
├── Output/               # Inno Setup 生成的 .exe 安装包目录
└── user/                 # 首次启动后自动生成（已加入 .gitignore）
    ├── data/             # 名单、积分、配置、成绩
    ├── exports/          # 导出的 Excel 表格
    └── backups/          # 自动备份的数据
```

## 📄 数据存储

所有数据均存储在**程序所在目录**的 `user/` 文件夹内，不上传任何服务器，**完全离线**。

| 文件 | 说明 |
|------|------|
| `user/data/config.json` | 系统配置、等级经验、主题设置 |
| `user/data/班级名单.xlsx` | 学生名单 |
| `user/data/班级积分记录.csv` | 积分原始记录 |
| `user/data/notes.json` | 首页便签 |

## ❓ 常见问题 (FAQ)

**1. 换电脑了，怎么把数据迁移过去？**
答：直接把你旧电脑上整个 `user` 文件夹（包含所有名单、成绩和积分）复制到新电脑的 PickOS 目录下即可。

**2. 为什么 Windows 系统/杀毒软件提示风险或误报？**
答：本软件由开源代码和 PyQt6 打包，无任何恶意行为。由于未购买昂贵的商业数字签名证书，部分安全软件（如 360、火绒）可能会误报。请在杀毒软件中将本软件添加为信任即可。

**3. 软件需要联网吗？会泄露学生隐私吗？**
答：完全离线运行！除了您主动点击“检查更新”时会访问 GitHub 接口外，软件没有任何网络请求，所有数据仅保存在您本地电脑中。

## ⚖️ 开源协议

本项目采用 **GNU General Public License v3.0** 开源。

### 声明

- 作者以**个人学习与自用**为目的开发和维护本项目，**不进行任何商业活动**。
- 依据 GPLv3，**任何第三方均可在遵守协议的前提下自由使用本项目，包括商业用途**，但修改后的代码必须以 GPLv3 开源。
- 本项目依赖 [PyQt6](https://pypi.org/project/PyQt6/)（GPLv3）与 [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets)（GPLv3），因此必须遵循 GPLv3 分发。

完整协议见 [LICENSE](LICENSE)。

## 📜 用户协议与隐私政策

- 用户协议：[查看 EULA.md](https://github.com/wuweijia-star/PickOS/blob/main/EULA.md)
- 隐私政策：[查看 PRIVACY.md](https://github.com/wuweijia-star/PickOS/blob/main/PRIVACY.md)

## 🙏 致谢

- [PyQt6](https://www.riverbankcomputing.com/software/pyqt/) —— Qt for Python
- [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) —— Windows 11 风格组件库
- [pandas](https://pandas.pydata.org/) / [openpyxl](https://openpyxl.readthedocs.io/) —— 数据处理
- [Inno Setup](https://jrsoftware.org/isinfo.php) —— Windows 安装包制作工具
- [SignPath Foundation](https://signpath.org/) —— 免费提供 Windows 安装包的数字签名

## 🐛 反馈与贡献

- 提交 Issue：[GitHub Issues](https://github.com/wuweijia-star/PickOS/issues)
- 贡献指南：[CONTRIBUTING.md](CONTRIBUTING.md)
- 安全策略：[SECURITY.md](SECURITY.md)
- 提交 PR：欢迎，但请遵循 GPLv3
- 提交前建议先搜索是否已有相似 Issue

## 📮 联系方式

- 作者：wuweijia-star
- 邮箱：2892065430@qq.com
- 版权：Copyright © 2026 wuweijia-star · GPLv3
- 仓库：[https://github.com/wuweijia-star/PickOS](https://github.com/wuweijia-star/PickOS)

---

<a id="english"></a>

[简体中文](#top) | **English**

# PickOS Class Management System

> A class points & grade management tool tailored for middle school homeroom teachers / class officers.

![Python](https://img.shields.io/badge/Python-3.12-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-6.5+-green)
![License](https://img.shields.io/badge/License-GPLv3-orange)

[![Download](https://img.shields.io/github/v/release/wuweijia-star/PickOS?label=Download%20Latest&style=for-the-badge&color=2ea44f)](https://github.com/wuweijia-star/PickOS/releases/latest)

PickOS runs fully offline. All data is stored in the local `user/` directory and is never uploaded to any server.

## ✨ Features

- 📝 **Points Entry**: Single / continuous batch entry with reason management, discounts, and scoring
- 📊 **Leaderboard**: Real-time ranking, search highlighting, subject champions
- 🎯 **Grade Overview**: Current vs. previous comparison, progress / regression lists, exam bonus rules
- 📅 **Dashboard**: Sticky notes, calendar, countdown, device monitoring, check-in
- 🏆 **Levels & Achievements**: 20-level XP system + 40+ achievements
- 🔐 **Password Protection / Smart Unlock**: Auto-unlock when a USB drive is plugged in
- 🧭 **First-Run Guide**: Hello animation, EULA / privacy policy confirmation, one-click setup for shortcuts / roster / subjects / password
- 🎨 **Theme Switching**: Dark / Light / Custom
- 💾 **Data Management**: Export to Excel, auto backup, storage usage analysis
- 🔄 **Online Updates**: Built-in one-click update check with multi-mirror accelerated downloads

## 🖥️ Requirements

- **Python 3.12+** (running from source)
- **OS**: Windows 10 / 11 (primary support) / macOS / Linux

## 🚀 Getting Started

### Option 1: Download the Installer (recommended for regular users)

Go to the [Releases](https://github.com/wuweijia-star/PickOS/releases) page and download the latest installer:

```text
PickOS_Setup_vX.Y.Z_win64.exe
```

> The actual filename may differ; refer to the Releases page.

Installation & upgrade notes:

1. Double-click the installer and follow the "Next" prompts.
2. **No administrator privileges are required**; it installs to the current user directory by default and does not trigger UAC prompts.
3. When upgrading from an older version, the installer automatically overwrites the old program and runtime libraries, but **does not delete your `user` data folder**.
4. Auto-update preserves user data such as `user/data`, `user/exports`, and `user/backups`.
5. When a new version is released, open the app and click "Settings -> Check for Updates" to download and silently upgrade in one click.

The auto-updater invokes the installer with the following silent parameters:

```text
/VERYSILENT /SUPPRESSMSGBOXES /NORESTART
```

### Option 2: Run from Source (recommended for developers)

```bash
git clone https://github.com/wuweijia-star/PickOS.git
cd PickOS
pip install -r requirements.txt
python main.pyw
```

> You can also double-click `build.bat` to quickly build a portable green-version exe.

## 📁 Project Structure

```text
PickOS/
├── main.pyw              # Main program
├── build.bat             # Windows local build script
├── requirements.txt      # Dependency list
├── LICENSE               # GPLv3 license
├── README.md
├── dist/                 # Build output directory (contains all dependencies)
├── Output/               # Inno Setup .exe installer output directory
└── user/                 # Auto-generated on first launch (already in .gitignore)
    ├── data/             # Roster, points, config, grades
    ├── exports/          # Exported Excel files
    └── backups/          # Auto-backed-up data
```

## 📄 Data Storage

All data is stored in the `user/` folder inside the **program directory**, and is never uploaded to any server — **fully offline**.

| File | Description |
|------|-------------|
| `user/data/config.json` | System config, level XP, theme settings |
| `user/data/班级名单.xlsx` | Student roster |
| `user/data/班级积分记录.csv` | Raw points records |
| `user/data/notes.json` | Dashboard sticky notes |

## ⚖️ License

This project is open-sourced under the **GNU General Public License v3.0**.

### Statement

- The author develops and maintains this project for **personal learning and self-use**, and **does not engage in any commercial activity**.
- Under GPLv3, **any third party may freely use this project, including for commercial purposes**, provided they comply with the license; modified code must be open-sourced under GPLv3.
- This project depends on [PyQt6](https://pypi.org/project/PyQt6/) (GPLv3) and [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) (GPLv3), and therefore must be distributed under GPLv3.

See [LICENSE](LICENSE) for the full license.

## 📜 User Agreement & Privacy Policy

- User Agreement: [View EULA.md](https://github.com/wuweijia-star/PickOS/blob/main/EULA.md)
- Privacy Policy: [View PRIVACY.md](https://github.com/wuweijia-star/PickOS/blob/main/PRIVACY.md)

## 🙏 Acknowledgements

- [PyQt6](https://www.riverbankcomputing.com/software/pyqt/) — Qt for Python
- [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) — Windows 11 style widget library
- [pandas](https://pandas.pydata.org/) / [openpyxl](https://openpyxl.readthedocs.io/) — Data processing
- [Inno Setup](https://jrsoftware.org/isinfo.php) — Windows installer builder
- [SignPath Foundation](https://signpath.org/) — Free code signing for Windows installers

## 🐛 Feedback & Contributing

- Submit an Issue: [GitHub Issues](https://github.com/wuweijia-star/PickOS/issues)
- Submit a PR: Welcome, but please comply with GPLv3
- Please search for existing similar issues before submitting

## 📮 Contact

- Author: wuweijia-star
- Email: 2892065430@qq.com
- Copyright: Copyright © 2026 wuweijia-star · GPLv3
- Repository: [https://github.com/wuweijia-star/PickOS](https://github.com/wuweijia-star/PickOS)