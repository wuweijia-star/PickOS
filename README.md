\# PickOS 班级管理系统



> 一个为中学班主任 / 班干部量身定制的班级积分与成绩管理工具。



!\[Python](https://img.shields.io/badge/Python-3.12-blue)

!\[PyQt6](https://img.shields.io/badge/PyQt6-6.5+-green)

!\[License](https://img.shields.io/badge/License-GPLv3-orange)



\## ✨ 功能简介



\- 📝 \*\*积分录入\*\*：单条 / 连续批量录入，支持事由管理、折扣、赋分

\- 📊 \*\*排行榜\*\*：实时排名、搜索高亮、单科王

\- 🎯 \*\*成绩总览\*\*：本次 vs 上次对比、进步 / 退步榜、考试加分规则

\- 📅 \*\*首页看板\*\*：便签、日历、倒计时、设备监控、签到

\- 🏆 \*\*等级与成就\*\*：20 级经验体系 + 40+ 成就系统

\- 🔐 \*\*密码保护 / 智能解锁\*\*：支持 U 盘插入自动解锁

\- 🎨 \*\*主题切换\*\*：深色 / 浅色 / 自定义

\- 💾 \*\*数据管理\*\*：导出 Excel、自动备份、存储空间分析



\## 🖥️ 运行环境



\- \*\*Python 3.12+\*\*

\- \*\*操作系统\*\*：Windows 10/11（主要支持）/ macOS / Linux



\## 🚀 快速开始



\### 方式一：直接运行源码



```bash

git clone https://github.com/wuweijia-star/PickOS.git

cd PickOS

pip install -r requirements.txt

python PickOS 班级管理系统v6.3.0.135源代码.pyw

```



\### 方式二：使用预打包程序



前往 \[Releases](https://github.com/wuweijia-star/PickOS/releases) 页面下载最新的压缩包：



1\. 解压到\*\*非系统目录\*\*（例如 `D:\\Apps\\PickOS\\`，\*\*不要\*\*放 `C:\\Program Files\\`）

2\. 双击 `PickOS.exe` 启动

3\. 首次启动会在 exe 同目录创建 `user\\` 文件夹



\## 📁 项目结构



```

PickOS/

├── PickOS 班级管理系统v6.3.0.135源代码.pyw              # 主程序（单文件）

├── PickOS班级管理系统批量打包程序.bat             # Windows 打包脚本

├── requirements.txt      # 依赖清单

├── LICENSE               # GPLv3 协议

├── README.md

└── user/                 # 首次启动后自动生成（已加入 .gitignore）

&#x20;   ├── data/

&#x20;   ├── exports/

&#x20;   └── backups/

```



\## 📄 数据存储



所有数据均存储在\*\*程序所在目录\*\*的 `user\\` 文件夹内，不上传任何服务器，\*\*完全离线\*\*。



| 文件 | 说明 |

|------|------|

| `user/data/config.json` | 系统配置 |

| `user/data/班级名单.xlsx` | 学生名单 |

| `user/data/班级积分记录.csv` | 积分原始记录 |

| `user/data/notes.json` | 便签 |



\## ⚖️ 开源协议



本项目采用 \*\*GNU General Public License v3.0\*\* 开源。



\### 声明



\- 作者\*\*个人\*\*将此项目作为\*\*学习与自用工具\*\*开发和维护，\*\*不进行任何商业活动\*\*（不售卖、不接定制）。

\- 但依据 GPLv3，\*\*任何第三方可在遵守协议的前提下自由使用本项目，包括商业用途\*\*，前提是修改后的代码必须以 GPLv3 开源。

\- 本项目依赖 \[PyQt6](https://pypi.org/project/PyQt6/)（GPLv3）与 \[PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets)（GPLv3），因此必须遵循 GPLv3 分发。



完整协议见 \[LICENSE](LICENSE)。



\## 🙏 致谢



\- \[PyQt6](https://www.riverbankcomputing.com/software/pyqt/) —— Qt for Python

\- \[PyQt-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) —— Windows 11 风格组件库

\- \[pandas](https://pandas.pydata.org/) / \[openpyxl](https://openpyxl.readthedocs.io/) —— 数据处理



\## 🐛 反馈与贡献



\- 提交 Issue：<https://github.com/wuweijia-star/PickOS/issues>

\- 提交 PR：欢迎，但请遵循 GPLv3



\## 📮 联系方式



\- 邮箱：2892065430@qq.com

\- 版权：Copyright © 2026 wuweijia-star  ·  GPLv3

\- 仓库：<https://github.com/wuweijia-star/PickOS>

