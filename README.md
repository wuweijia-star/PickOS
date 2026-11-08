# PickOS 班级管理系统

> 一个为中学班主任 / 班干部量身定制的班级积分与成绩管理工具。

![Python](https://img.shields.io/badge/Python-3.12-blue)
![PyQt6](https://img.shields.io/badge/PyQt6-6.5+-green)
![License](https://img.shields.io/badge/License-GPLv3-orange)

[![Download](https://img.shields.io/github/v/release/wuweijia-star/PickOS?label=下载最新版&style=for-the-badge&color=2ea44f)](https://github.com/wuweijia-star/PickOS/releases/latest)

## ✨ 功能简介

- 📝 **积分录入**：单条 / 连续批量录入，支持事由管理、折扣、赋分
- 📊 **排行榜**：实时排名、搜索高亮、单科王
- 🎯 **成绩总览**：本次 vs 上次对比、进步 / 退步榜、考试加分规则
- 📅 **首页看板**：便签、日历、倒计时、设备监控、签到
- 🏆 **等级与成就**：20 级经验体系 + 40+ 成就系统
- 🔐 **密码保护 / 智能解锁**：支持 U 盘插入自动解锁
- 🎨 **主题切换**：深色 / 浅色 / 自定义
- 💾 **数据管理**：导出 Excel、自动备份、存储空间分析
- 🔄 **在线更新**：内置一键检查更新，支持多镜像加速下载安装包

## 🖥️ 运行环境

- **Python 3.12+**（源码运行）
- **操作系统**：Windows 10/11（完美支持）/ macOS / Linux

## 🚀 快速开始

### 方式一：下载安装包（推荐普通用户）

前往 [Releases](https://github.com/wuweijia-star/PickOS/releases) 页面，下载最新的安装包：

```text
PickOS_Setup_vX.Y.Z_win64.exe
```

> 实际文件名以 Releases 页面显示为准。

1. 双击运行安装包，按照提示点击“下一步”即可。
2. **安装过程不需要管理员权限**，默认安装在当前用户目录下，不会触发 UAC 权限拦截。
3. 从旧版本升级时，**安装包会自动覆盖旧程序与运行库，但不会删除您的 `user` 数据文件夹**。
4. 自动更新会保留 `user/data`、`user/exports`、`user/backups` 等用户数据。
5. 以后有新版本发布时，只需打开软件，点击「设置 -> 检查更新」即可一键下载并静默升级。

自动更新调用安装包时使用静默参数：

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

*(注：也可以直接双击 `build.bat` 快速打包生成免安装绿色版 exe)*

## 📁 项目结构

```text
PickOS/
├── main.pyw              # 主程序
├── build.bat             # Windows 本地打包脚本
├── requirements.txt      # 依赖清单
├── LICENSE               # GPLv3 协议
├── README.md
├── dist/                 # 打包后生成的运行目录 (包含所有依赖)
├── Output/               # Inno Setup 生成的 .exe 安装包目录
└── user/                 # 首次启动后自动生成（已加入 .gitignore）
    ├── data/             # 名单、积分、配置、成绩
    ├── exports/          # 导出的 Excel 表格
    └── backups/          # 自动备份的数据
```

## 📄 数据存储

所有数据均存储在**程序所在目录**的 `user\` 文件夹内，不上传任何服务器，**完全离线**。

| 文件 | 说明 |
|------|------|
| `user/data/config.json` | 系统配置、等级经验、主题设置 |
| `user/data/班级名单.xlsx` | 学生名单 |
| `user/data/班级积分记录.csv` | 积分原始记录 |
| `user/data/notes.json` | 首页便签 |

## ⚖️ 开源协议

本项目采用 **GNU General Public License v3.0** 开源。

### 声明

- 作者**个人**将此项目作为**学习与自用工具**开发和维护，**不进行任何商业活动**（不售卖、不接定制）。
- 但依据 GPLv3，**任何第三方可在遵守协议的前提下自由使用本项目，包括商业用途**，前提是修改后的代码必须以 GPLv3 开源。
- 本项目依赖 [PyQt6](https://pypi.org/project/PyQt6/)（GPLv3）与 [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets)（GPLv3），因此必须遵循 GPLv3 分发。

完整协议见 [LICENSE](LICENSE)。

## 🙏 致谢

- [PyQt6](https://www.riverbankcomputing.com/software/pyqt/) —— Qt for Python
- [PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) —— Windows 11 风格组件库
- [pandas](https://pandas.pydata.org/) / [openpyxl](https://openpyxl.readthedocs.io/) —— 数据处理
- [Inno Setup](https://jrsoftware.org/isinfo.php) —— Windows 专业安装包制作工具

## 🐛 反馈与贡献

- 提交 Issue：[点击这里](https://github.com/wuweijia-star/PickOS/issues)
- 提交 PR：欢迎，但请遵循 GPLv3

## 📮 联系方式

- 邮箱：2892065430@qq.com
- 版权：Copyright © 2026 wuweijia-star · GPLv3
- 仓库：[https://github.com/wuweijia-star/PickOS](https://github.com/wuweijia-star/PickOS)