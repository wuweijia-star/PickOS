# -*- coding: utf-8 -*-
#
# PickOS 班级管理系统
# Copyright (C) 2026 wuweijia-star  ·  GPLv3
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
import sys
import calendar
import base64
import hashlib
from PyQt6.QtGui import QIcon, QPixmap
import os
import json
import ssl
import certifi
import math
import time
import random
import shutil
import glob
import subprocess
import urllib.request
import zipfile
import re
import threading
from datetime import datetime
from datetime import timedelta
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

from PyQt6.QtCore import (Qt, QTimer, QPoint, QRect, QRectF, QSize, QPointF, 
                          QPropertyAnimation, QVariantAnimation,
                          QEasingCurve, QEventLoop, pyqtProperty, pyqtSignal, QObject, QEvent)
from PyQt6.QtGui import (QColor, QFont, QPalette, QBrush, QPainter,
                         QPen, QIcon, QPixmap, QTextCursor, QImage, QPainterPath)
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QLineEdit, QComboBox, QPushButton, QCheckBox, QRadioButton,
    QButtonGroup, QGroupBox, QFormLayout, QFileDialog, QMessageBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QDialog, QDialogButtonBox,
    QFrame, QVBoxLayout, QHBoxLayout,
    QProgressBar, QInputDialog,
    QSpinBox, QSizePolicy,QScrollArea,QTextBrowser,QTextEdit,QGraphicsOpacityEffect,
    QGraphicsScene, QGraphicsPixmapItem, QGraphicsBlurEffect,
)
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtWidgets import QSystemTrayIcon, QMenu
from PyQt6.QtCore import QUrl
from PyQt6.QtCore import QMetaObject, Q_ARG, pyqtSlot
from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, FluentIcon,
    setTheme, Theme, setThemeColor, InfoBar, InfoBarPosition,
    PushButton, CardWidget, BodyLabel, TitleLabel, ComboBox,
    SwitchButton
)


# ======================== 全局配置 ========================
USER_DIR = "user"                     # 用户数据根目录
DATA_DIR = os.path.join(USER_DIR, "data")       # 数据文件夹
EXPORT_DIR = os.path.join(USER_DIR, "exports")  # 导出文件夹
BACKUP_DIR = os.path.join(USER_DIR, "backups")  # 备份文件夹
# 成绩导入数据持久化文件（多考试管理）
CURRENT_SCORES_FILE = os.path.join(DATA_DIR, "current_scores.csv")   # 本次考试
PREVIOUS_SCORES_FILE = os.path.join(DATA_DIR, "previous_scores.csv") # 上次考试

CONFIG_JSON = os.path.join(DATA_DIR, "config.json")
CUSTOM_LISTS_FILE = os.path.join(DATA_DIR, "custom_lists.json")
DEFAULT_STUDENT_EXCEL = os.path.join(DATA_DIR, "全班名单.xlsx")
CLASS_LIST_FILE = os.path.join(DATA_DIR, "班级名单.xlsx")
DATA_CSV = os.path.join(DATA_DIR, "班级积分记录.csv")
INSTALL_TIME_STR = "2026-05-01 18:00:00"
INSTALL_TIMESTAMP = time.mktime(time.strptime(INSTALL_TIME_STR, "%Y-%m-%d %H:%M:%S"))
SYSTEM_VERSION = "v6.4.0.110"

DIALOG_STYLE = """
    QDialog {
        background: #16161E;
        border-radius: 18px;
    }
    QWidget {
        font-family: "Inter", "Segoe UI", "Microsoft YaHei", sans-serif;
        font-size: 15px;
        color: #E0E0E0;
    }
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
        border-radius: 8px;
        padding: 8px 12px;
        border: 1px solid #3A3A4A;
        background: #1A1A24;
        color: #F0F0F0;
        selection-background-color: #64B5F6;
    }
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
        border: 2px solid #64B5F6;
        background: #1E1E2A;
    }
    QComboBox::drop-down {
        border: none;
        width: 24px;
    }
    QComboBox::down-arrow {
        image: none;
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 6px solid #A0A0B0;
        margin-right: 6px;
    }
    QComboBox QAbstractItemView {
        border-radius: 8px;
        background: #1E1E2A;
        border: 1px solid #3A3A4A;
        padding: 4px;
    }
    QComboBox QAbstractItemView::item {
        padding: 6px 10px;
        border-radius: 4px;
        color: #E0E0E0;
    }
    QComboBox QAbstractItemView::item:selected {
        background: #64B5F6;
        color: white;
    }
    QPushButton {
        border-radius: 10px;
        padding: 8px 20px;
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
        color: white;
        font-weight: 600;
        border: 1px solid rgba(100, 181, 246, 0.3);
        font-size: 15px;
    }
    QPushButton:hover {
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
        border-color: #64B5F6;
    }
    QPushButton:pressed {
        background: #1E88E5;
    }
    QPushButton:disabled {
        background: #555555;
        color: #888888;
        border-color: #555555;
    }
    QTableWidget {
        background: #16161E;
        alternate-background-color: #1E1E2A;
        gridline-color: #2A2A3A;
        border-radius: 12px;
        border: 1px solid #2A2A3A;
        selection-background-color: #64B5F6;
        selection-color: white;
        font-size: 14px;
    }
    QTableWidget::item {
        padding: 8px;
        color: #E0E0E0;
        border: none;
        background: transparent;  /* 关键：使item背景透明，显示表格背景 */
    }
    QTableWidget::item:selected {
        background: #64B5F6;
        color: white;
    }
    QTableWidget::item:hover {
        background: #2A2A3A;
    }
    QTableCornerButton::section {
        background: #1E1E2A;
        border: none;
    }
    QHeaderView::section {
        background: #1E1E2A;
        padding: 8px;
        border: none;
        font-weight: 600;
        color: #C0C0D0;
        border-bottom: 2px solid #2A2A3A;
        font-size: 14px;
    }
    QLabel {
        color: #E0E0E0;
    }
    QGroupBox {
        border: 1px solid #3A3A4A;
        border-radius: 12px;
        margin-top: 16px;
        padding-top: 12px;
        font-weight: 600;
        color: #D0D0E0;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 12px;
        padding: 0 8px;
        color: #C0C0D0;
        background: transparent;
    }
    QRadioButton, QCheckBox {
        color: #E0E0E0;
        spacing: 8px;
    }
    QRadioButton::indicator, QCheckBox::indicator {
        width: 18px;
        height: 18px;
        border-radius: 4px;
        border: 2px solid #5A5A6A;
        background: #1A1A24;
    }
    QRadioButton::indicator:checked, QCheckBox::indicator:checked {
        background: #64B5F6;
        border-color: #64B5F6;
    }
    QScrollBar:vertical {
        width: 6px;
        background: transparent;
        margin: 0px;
    }
    QScrollBar::handle:vertical {
        background: #4A4A5A;
        border-radius: 4px;
        min-height: 30px;
    }
    QScrollBar::handle:vertical:hover {
        background: #6A6A7A;
    }
    QScrollBar::add-line, QScrollBar::sub-line {
        height: 0px;
    }
"""

# ======================== 自动更新系统 ========================
CURRENT_VERSION = "v6.4.0.110"
GITHUB_API_URL = "https://api.github.com/repos/wuweijia-star/PickOS/releases/latest"

class UpdateManager:
    import ssl
    import certifi
    @staticmethod
    def parse_version(v_str):
        """将 v6.3.0.135 转换为 (6, 3, 0, 135) 用于比较"""
        try:
            return tuple(map(int, re.findall(r'\d+', v_str)))
        except:
            return (0,)

    @staticmethod
    def check_update():
        """返回 (是否有更新, 最新版本号, 更新日志, 下载链接)"""
        try:
            # 1. 构造 SSL 上下文，加载 certifi 提供的根证书
            context = ssl.create_default_context(cafile=certifi.where())
            
            req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': 'PickOS-Updater'})
            # 2. 传入 context
            with urllib.request.urlopen(req, timeout=10, context=context) as response:
                data = json.loads(response.read().decode('utf-8'))
                
            latest_version = data.get("tag_name", "v0.0.0")
            changelog = data.get("body", "暂无更新日志")
            
            # 比较版本
            if UpdateManager.parse_version(latest_version) > UpdateManager.parse_version(CURRENT_VERSION):
                # ===== 优先找 .exe 安装包 =====
                installer_url = None
                fallback_zip = None
                for asset in data.get("assets", []):
                    name = asset["name"].lower()
                    url = asset["browser_download_url"]
                    # 安装包判定：后缀 .exe 且文件名含 setup/installer/install/win64
                    if name.endswith(".exe") and any(
                        k in name for k in ("setup", "installer", "install", "win64")
                    ):
                        installer_url = url
                        break
                    # 记录一下 zip 作为兜底信息
                    if "win64.zip" in name:
                        fallback_zip = url

                download_url = installer_url or fallback_zip
                if download_url:
                    if installer_url:
                        print(f"[更新] 找到安装包：{installer_url}")
                    else:
                        print(f"[更新] ⚠️ 未找到 .exe 安装包，使用 zip 兜底")
                    return True, latest_version, changelog, download_url
            return False, latest_version, changelog, None
        except Exception as e:
            print(f"检查更新失败: {e}")
            err_msg = str(e)
            # 如果包含 SSL 错误，尝试忽略证书重试
            if "SSL" in err_msg or "CERTIFICATE" in err_msg:
                try:
                    print("⚠️ 证书验证失败，尝试忽略 SSL 验证重试...")
                    context = ssl._create_unverified_context()
                    req = urllib.request.Request(GITHUB_API_URL, headers={'User-Agent': 'PickOS-Updater'})
                    with urllib.request.urlopen(req, timeout=10, context=context) as response:
                        data = json.loads(response.read().decode('utf-8'))
                    latest_version = data.get("tag_name", "v0.0.0")
                    changelog = data.get("body", "暂无更新日志")
                    if UpdateManager.parse_version(latest_version) > UpdateManager.parse_version(CURRENT_VERSION):
                        installer_url = None
                        fallback_zip = None
                        for asset in data.get("assets", []):
                            name = asset["name"].lower()
                            url = asset["browser_download_url"]
                            if name.endswith(".exe") and any(
                                k in name for k in ("setup", "installer", "install", "win64")
                            ):
                                installer_url = url
                                break
                            if "win64.zip" in name:
                                fallback_zip = url
                        download_url = installer_url or fallback_zip
                        if download_url:
                            return True, latest_version, changelog, download_url
                except Exception as e2:
                    print(f"忽略证书重试依然失败: {e2}")
            
            # 如果超时或被墙，抛出具体的报错文字给 UI 处理
            if "timeout" in err_msg.lower() or "timed out" in err_msg.lower():
                raise Exception("连接 GitHub 超时，请开启加速器或代理后再试。")
            elif "403" in err_msg:
                raise Exception("GitHub API 次数受限，请稍后再试。")
            else:
                raise Exception(f"网络错误: {err_msg}")

    @staticmethod
    def _mirror_urls(original_url):
        """
        给 GitHub 原始下载链接生成一组候选：
        先试国内镜像（大概率快），最后才试 GitHub 原站（兜底）。
        """
        if "github.com" not in (original_url or ""):
            return [original_url]

        # 按推荐度排序，越靠前越优先尝试
        prefixes = [
            "https://ghfast.top/",
            "https://gh-proxy.com/",
            "https://ghproxy.net/",
            "https://ghproxy.homeboyc.cn/",
        ]
        urls = [p + original_url for p in prefixes]
        urls.append(original_url)          # 最后一个，原站兜底
        return urls

    @staticmethod
    def download_installer(url, progress_callback=None):
        """
        下载安装包（.exe）到临时目录。
        :param url: 安装包 URL
        :param progress_callback: 进度回调，接收 0-100 整数
        :return: 本地安装包路径，失败返回 None
        """
        try:
            # 1. 准备临时目录
            temp_dir = os.path.join(os.environ["TEMP"], "PickOS_Update")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            # 2. 从 URL 解析文件名
            filename = url.split("/")[-1].split("?")[0]
            if not filename.lower().endswith(".exe"):
                filename = "PickOS_Setup.exe"
            installer_path = os.path.join(temp_dir, filename)

            # 3. 多镜像依次尝试
            response = None
            last_err = None
            for try_url in UpdateManager._mirror_urls(url):
                try:
                    print(f"[更新] 尝试下载源：{try_url}")
                    context = ssl.create_default_context(cafile=certifi.where())
                    req = urllib.request.Request(
                        try_url, headers={'User-Agent': 'PickOS-Updater'}
                    )
                    response = urllib.request.urlopen(req, timeout=20, context=context)
                    print(f"[更新] 已连接：{try_url}")
                    break
                except Exception as e:
                    print(f"[更新] 此源失败，换下一个：{e}")
                    last_err = e
                    if "SSL" in str(e) or "CERTIFICATE" in str(e):
                        try:
                            context = ssl._create_unverified_context()
                            req = urllib.request.Request(
                                try_url, headers={'User-Agent': 'PickOS-Updater'}
                            )
                            response = urllib.request.urlopen(
                                req, timeout=20, context=context
                            )
                            print(f"[更新] 忽略证书后连接成功：{try_url}")
                            break
                        except Exception as e2:
                            last_err = e2
                    continue

            if response is None:
                raise last_err if last_err else Exception("所有下载源均不可用")

            # 4. 写盘
            with response as res, open(installer_path, "wb") as f:
                total_size = int(res.info().get('Content-Length', 0))
                downloaded = 0
                chunk_size = 256 * 1024
                while True:
                    chunk = res.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size > 0:
                        progress_callback(int(downloaded / total_size * 100))

            print(f"[更新] 安装包已保存：{installer_path}")
            return installer_path
        except Exception as e:
            print(f"[更新] 下载安装包失败：{e}")
            return None

    @staticmethod
    def run_installer_and_exit(installer_path):
        """
        启动安装程序并退出当前程序。
        安装程序自己处理文件覆盖、提示用户下一步，比手搓 bat 稳。
        """
        try:
            if not installer_path or not os.path.exists(installer_path):
                print(f"[更新] 安装包不存在：{installer_path}")
                return False

            app = QApplication.instance()

            # 1. 隐藏所有托盘图标
            try:
                if app:
                    for w in app.topLevelWidgets():
                        if hasattr(w, 'tray_icon') and w.tray_icon:
                            try:
                                w.tray_icon.hide()
                            except Exception:
                                pass
            except Exception:
                pass

            # 2. 启动安装程序
            print(f"[更新] 启动安装包：{installer_path}")
            try:
                os.startfile(installer_path)   # Windows 最干净，无黑窗
            except AttributeError:
                # 跨平台兜底
                if sys.platform == "darwin":
                    subprocess.Popen(["open", installer_path])
                else:
                    subprocess.Popen(["xdg-open", installer_path])

            # 3. 延迟 800ms 退出本程序，让安装程序先启动
            #    用户点"安装"时文件锁已经释放
            if app:
                QTimer.singleShot(800, app.quit)
            return True
        except Exception as e:
            print(f"[更新] 启动安装包失败：{e}")
            if QApplication.instance():
                QTimer.singleShot(500, QApplication.instance().quit)
            return False

    @staticmethod
    def download_and_apply(url, progress_callback=None):
        """下载、解压并生成覆盖脚本"""
        try:
            temp_dir = os.path.join(os.environ["TEMP"], "PickOS_Update")
            zip_path = os.path.join(temp_dir, "update.zip")
            
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            # ========== 1. 下载：多镜像依次尝试 ==========
            response = None
            last_err = None
            used_url = None

            for try_url in UpdateManager._mirror_urls(url):
                try:
                    print(f"[更新] 尝试下载源：{try_url}")
                    context = ssl.create_default_context(cafile=certifi.where())
                    req = urllib.request.Request(
                        try_url,
                        headers={'User-Agent': 'PickOS-Updater'}
                    )
                    response = urllib.request.urlopen(
                        req, timeout=20, context=context
                    )
                    used_url = try_url
                    print(f"[更新] 已连接：{try_url}")
                    break
                except Exception as e:
                    print(f"[更新] 此源失败，换下一个：{e}")
                    last_err = e
                    # SSL 证书问题，用不校验的上下文再试一次这个源
                    if "SSL" in str(e) or "CERTIFICATE" in str(e):
                        try:
                            context = ssl._create_unverified_context()
                            req = urllib.request.Request(
                                try_url,
                                headers={'User-Agent': 'PickOS-Updater'}
                            )
                            response = urllib.request.urlopen(
                                req, timeout=20, context=context
                            )
                            used_url = try_url
                            print(f"[更新] 忽略证书后连接成功：{try_url}")
                            break
                        except Exception as e2:
                            print(f"[更新] 忽略证书仍失败：{e2}")
                            last_err = e2
                    continue

            if response is None:
                # 所有源全部失败
                raise last_err if last_err else Exception("所有下载源均不可用")

            with response as res, open(zip_path, 'wb') as out_file:
                total_size = int(res.info().get('Content-Length', 0))
                downloaded = 0
                chunk_size = 1024 * 1024  # 1MB
                while True:
                    chunk = res.read(chunk_size)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size > 0:
                        progress_callback(int(downloaded / total_size * 100))
            
            # 2. 解压
            extract_dir = os.path.join(temp_dir, "extracted")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            # 找到解压后的实际 exe 目录
            app_dir = extract_dir
            for root, dirs, files in os.walk(extract_dir):
                if "PickOS.exe" in files:
                    app_dir = root
                    break

            # 3. 生成更新脚本
            current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
            bat_path = os.path.join(temp_dir, "updater.bat")
            
            bat_content = f"""@echo off
chcp 65001 >nul
timeout /t 2 /nobreak >nul
:waitloop
tasklist | find /i "PickOS.exe" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto waitloop
)

cd /d "{current_dir}"
for /d %%i in ("*") do (
    if /i not "%%~ni"=="user" rd /s /q "%%i"
)
for %%i in ("*") do (
    if /i not "%%~nxi"=="PickOS.exe" del /q "%%i"
)

xcopy "{app_dir}\\*" "{current_dir}\\" /E /H /Y /Q
start "" "{current_dir}\\PickOS.exe"
timeout /t 2 /nobreak >nul
rd /s /q "{temp_dir}"
exit
"""
            with open(bat_path, "w", encoding="gbk") as f:
                f.write(bat_content)

            return bat_path
        except Exception as e:
            print(f"下载更新失败: {e}")
            return None

    @staticmethod
    def restart_and_apply(bat_path):
        """执行批处理脚本并退出程序"""
        try:
            subprocess.Popen(["cmd.exe", "/c", bat_path], creationflags=0x08000000) # CREATE_NO_WINDOW
            QApplication.quit()
            sys.exit(0)
        except Exception as e:
            print(f"启动更新脚本失败: {e}")

# ======================== 主题系统 ========================
def _hex_to_rgb(hex_color):
    h = str(hex_color).strip().lstrip('#')
    if len(h) == 3:
        h = ''.join(ch * 2 for ch in h)
    if len(h) != 6:
        return (22, 22, 30)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return (22, 22, 30)

def _rgb_to_hex(rgb):
    return '#{:02X}{:02X}{:02X}'.format(
        max(0, min(255, int(rgb[0]))),
        max(0, min(255, int(rgb[1]))),
        max(0, min(255, int(rgb[2]))),
    )

def _shift_color(hex_color, delta):
    r, g, b = _hex_to_rgb(hex_color)
    return _rgb_to_hex((r + delta, g + delta, b + delta))

def _is_dark_color(hex_color):
    r, g, b = _hex_to_rgb(hex_color)
    return (r * 0.299 + g * 0.587 + b * 0.114) < 128

THEME_PRESETS = {
    "dark": {
        "bg":          "#16161E",
        "bg_alt":      "#1A1A24",
        "bg_card":     "#1E1E2A",
        "border":      "#2A2A3A",
        "border_light":"#3A3A4A",
        "fg":          "#E0E0E0",
        "fg_bright":   "#F0F0F0",
        "fg_dim":      "#A0A0B0",
    },
    "light": {
        "bg":          "#F0F2F5",
        "bg_alt":      "#FFFFFF",
        "bg_card":     "#E8EAEF",
        "border":      "#C8CAD0",
        "border_light":"#A8AAB0",
        "fg":          "#1A1A24",
        "fg_bright":   "#000000",
        "fg_dim":      "#606070",
    },
}

def get_theme_colors():
    """读取当前主题颜色（含自定义），返回 dict。"""
    try:
        cfg = load_config()
    except Exception:
        return dict(THEME_PRESETS["light"])
    mode = cfg.get("theme_mode", "light")
    if mode == "custom":
        bg = cfg.get("theme_bg", "#16161E") or "#16161E"
        fg = cfg.get("theme_fg", "#E0E0E0") or "#E0E0E0"
        is_dark = _is_dark_color(bg)
        if is_dark:
            bg_alt      = _shift_color(bg, 10)
            bg_card     = _shift_color(bg, 16)
            border      = _shift_color(bg, 25)
            border_light= _shift_color(bg, 40)
            fg_dim      = _shift_color(fg, -60)
            fg_bright   = _shift_color(fg, 20)
        else:
            bg_alt      = _shift_color(bg, -10)
            bg_card     = _shift_color(bg, -18)
            border      = _shift_color(bg, -35)
            border_light= _shift_color(bg, -55)
            fg_dim      = _shift_color(fg, 80)
            fg_bright   = _shift_color(fg, -20)
        return {
            "bg": bg, "bg_alt": bg_alt, "bg_card": bg_card,
            "border": border, "border_light": border_light,
            "fg": fg, "fg_bright": fg_bright, "fg_dim": fg_dim,
        }
    return dict(THEME_PRESETS.get(mode, THEME_PRESETS["light"]))

def get_qfluent_theme():
    """根据当前主题配置返回 qfluentwidgets 的 Theme 枚举"""
    try:
        cfg = load_config()
    except Exception:
        return Theme.LIGHT
    mode = cfg.get("theme_mode", "light")
    if mode == "light":
        return Theme.LIGHT
    if mode == "custom":
        bg = cfg.get("theme_bg", "#16161E") or "#16161E"
        return Theme.DARK if _is_dark_color(bg) else Theme.LIGHT
    return Theme.DARK

def _replace_theme_colors(template_str):
    """把样式表里常见的深色硬编码颜色替换为当前主题色。"""
    c = get_theme_colors()
    replacements = [
        ("#16161E", c["bg"]),
        ("#1A1A24", c["bg_alt"]),
        ("#1E1E2A", c["bg_card"]),
        ("#2A2A3A", c["border"]),
        ("#3A3A4A", c["border_light"]),
        ("#5A5A6A", c["border_light"]),
        ("#4A4A5A", c["border_light"]),
        ("#6A6A7A", c["fg_dim"]),
        ("#E0E0E0", c["fg"]),
        ("#F0F0F0", c["fg_bright"]),
        ("#A0A0B0", c["fg_dim"]),
        ("#C0C0D0", c["fg_dim"]),
        ("#D0D0E0", c["fg_dim"]),
        ("#141a2b", c["bg"]),
        ("#0a0d15", c["bg_alt"]),
    ]
    for old, new in replacements:
        template_str = template_str.replace(old, new)
    return template_str


# ======================== 可滚动表头 ========================
class ScrollingHeaderView(QHeaderView):
    STEP_PX = 1
    INTERVAL_MS = 80
    GAP_PX = 40
    PADDING_PX = 6
    BG_COLOR = "#1E1E2A"
    TEXT_COLOR = "#C0C0D0"

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setHighlightSections(False)

        # ★ 新增：根据当前主题刷新表头滚动文字的颜色，浅色主题下不再出现"深底白字"的错位效果
        try:
            _c = get_theme_colors()
            self.BG_COLOR = _c["bg_card"]
            self.TEXT_COLOR = _c["fg_dim"]
        except Exception:
            pass

        self._offset = 0
        self._timer = QTimer(self)
        self._timer.setInterval(self.INTERVAL_MS)
        self._timer.timeout.connect(self._on_tick)
        self._timer.start()

    def _on_tick(self):
        if self._has_overflow():
            self._offset += self.STEP_PX
            self.viewport().update()

    def _has_overflow(self):
        """检查是否存在放不下的表头文字"""
        model = self.model()
        if model is None:
            return False
        fm = self.fontMetrics()
        for i in range(self.count()):
            text = model.headerData(i, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
            if text is None:
                continue
            text = str(text)
            if not text:
                continue
            if fm.horizontalAdvance(text) > self.sectionSize(i) - self.PADDING_PX:
                return True
        return False

    def paintSection(self, painter, rect, logicalIndex):
        # 先让默认实现绘制背景、边框、排序箭头等
        super().paintSection(painter, rect, logicalIndex)

        model = self.model()
        if model is None:
            return
        text = model.headerData(logicalIndex, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        if text is None:
            return
        text = str(text)
        if not text:
            return

        fm = painter.fontMetrics()
        text_width = fm.horizontalAdvance(text)
        if text_width <= rect.width() - self.PADDING_PX:
            return  # 能完整显示，沿用默认绘制

        # 无法完整显示 —— 覆盖原文字并绘制滚动文本
        painter.save()
        painter.fillRect(rect, QColor(self.BG_COLOR))
        painter.setClipRect(rect)
        painter.setPen(QColor(self.TEXT_COLOR))

        cycle = text_width + self.GAP_PX
        offset = self._offset % cycle
        text_top = rect.y() + (rect.height() - fm.height()) // 2
        baseline_y = text_top + fm.ascent()

        x = rect.x() + self.PADDING_PX // 2 - offset
        right = rect.x() + rect.width()
        while x < right:
            painter.drawText(x, baseline_y, text)
            x += cycle

        painter.restore()


def install_scrolling_header(table_widget):
    """为 QTableWidget 的横向表头启用文字滚动效果"""
    old = table_widget.horizontalHeader()

    header = ScrollingHeaderView(Qt.Orientation.Horizontal, table_widget)

    # 在 setHorizontalHeader 之前，把旧表头的属性拷贝过来
    if old is not None:
        header.setSectionsClickable(old.sectionsClickable())
        header.setHighlightSections(False)          # ★ 强制关闭，避免选中行时表头列高亮
        header.setStretchLastSection(old.stretchLastSection())
        header.setDefaultAlignment(old.defaultAlignment())
        header.setSectionsMovable(old.sectionsMovable())
    else:
        header.setHighlightSections(False)          # ★ 兜底

    table_widget.setHorizontalHeader(header)
    return header

# ======================== 等级系统配置 ========================
# 等级 -> 四字称号
LEVEL_TITLES = {
    1: "微光初现", 2: "渐露锋芒", 3: "拾光前行",
    4: "星芒见习", 5: "逐光实践", 6: "耀光实干",
    7: "聚光执行", 8: "凝光统筹", 9: "流光协领", 10: "明光骨干",
    11: "星途进阶", 12: "灿星专务", 13: "熠星荣誉", 14: "繁星资深", 15: "慧星筹策",
    16: "星河开拓", 17: "星宇共建", 18: "瀚星构筑", 19: "极星领航",
    20: "万象星穹"
}

# 等级区间名称（仅用于展示，可省略）
LEVEL_RANGES = {
    1: "启行", 2: "启行", 3: "启行",
    4: "历练", 5: "历练", 6: "历练",
    7: "笃行", 8: "笃行", 9: "笃行", 10: "笃行",
    11: "深耕", 12: "深耕", 13: "深耕", 14: "深耕", 15: "深耕",
    16: "拓越", 17: "拓越", 18: "拓越", 19: "拓越",
    20: "顶峰"
}

def calculate_level(total_count=None):
    """根据累计经验计算等级（兼容旧调用）"""
    if total_count is None:
        cfg = load_config()
        total_exp = cfg.get("total_experience", 0)
    else:
        # 如果传入旧参数，视为经验值（兼容）
        total_exp = total_count
    return calculate_level_by_exp(total_exp)

# ======================== 经验系统配置 ========================
# 每日各来源上限
DAILY_LIMIT_TOTAL = 120          # 每日全局上限
DAILY_LIMIT_ONLINE = 20          # 在线计时每日上限
DAILY_LIMIT_RECORD = 60          # 录入记录每日上限

# 在线计时：每2分钟 +1 经验
ONLINE_INTERVAL_SECONDS = 120    # 2分钟
ONLINE_EXP_PER_INTERVAL = 1

# 签到经验阶梯（连续天数 -> 经验值）
SIGN_IN_REWARDS = {
    1: 2, 2: 6, 3: 10, 4: 10, 5: 15, 6: 15,
}
# 第7天及以后固定 +20
SIGN_IN_DAILY_BONUS = 20

# 录入记录：每条 +2
RECORD_EXP_PER_ENTRY = 2

# 设备运行：每天 +50
DEVICE_RUNTIME_DAILY_EXP = 50

# ---- 升级所需经验值（1~20级） ----
LEVEL_EXP_REQUIREMENTS = {
    1: 0,      # 初始等级1，需要0经验
    2: 50,
    3: 150,
    4: 300,
    5: 500,
    6: 800,
    7: 1200,
    8: 1600,
    9: 2000,
    10: 2500,
    11: 3000,
    12: 4000,
    13: 6000,
    14: 8000,
    15: 10000,
    16: 13000,
    17: 16000,
    18: 20000,
    19: 25000,
    20: 30000,
}
MAX_LEVEL = 20

# ======================== 成就列表 ========================
ACHIEVEMENT_LIST = [
    # ========== 每级升级成就 Lv2‑Lv20 ==========
    {"aid":"lv_2","name":"锋芒初露","desc":"等级提升至 Lv.2","exp":20},
    {"aid":"lv_3","name":"拾光启程","desc":"等级提升至 Lv.3","exp":30},
    {"aid":"lv_4","name":"星芒见习","desc":"等级提升至 Lv.4","exp":40},
    {"aid":"lv_5","name":"逐光而行","desc":"等级提升至 Lv.5","exp":60},
    {"aid":"lv_6","name":"耀光实干","desc":"等级提升至 Lv.6","exp":80},
    {"aid":"lv_7","name":"聚光执行","desc":"等级提升至 Lv.7","exp":100},
    {"aid":"lv_8","name":"凝光统筹","desc":"等级提升至 Lv.8","exp":130},
    {"aid":"lv_9","name":"流光协领","desc":"等级提升至 Lv.9","exp":160},
    {"aid":"lv_10","name":"明光骨干","desc":"等级提升至 Lv.10","exp":200},
    {"aid":"lv_11","name":"星途进阶","desc":"等级提升至 Lv.11","exp":250},
    {"aid":"lv_12","name":"灿星专务","desc":"等级提升至 Lv.12","exp":300},
    {"aid":"lv_13","name":"熠星荣誉","desc":"等级提升至 Lv.13","exp":360},
    {"aid":"lv_14","name":"繁星资深","desc":"等级提升至 Lv.14","exp":420},
    {"aid":"lv_15","name":"慧星筹策","desc":"等级提升至 Lv.15","exp":500},
    {"aid":"lv_16","name":"星河开拓","desc":"等级提升至 Lv.16","exp":600},
    {"aid":"lv_17","name":"星宇共建","desc":"等级提升至 Lv.17","exp":660},
    {"aid":"lv_18","name":"瀚星构筑","desc":"等级提升至 Lv.18","exp":720},
    {"aid":"lv_19","name":"极星领航","desc":"等级提升至 Lv.19","exp":800},
    {"aid":"lv_20","name":"万象星穹","desc":"抵达满级 Lv.20","exp":1000},

    # ========== 使用历程 ==========
    {"aid":"first_launch","name":"初次启幕","desc":"第一次启动PickOS系统","exp":30},
    {"aid":"walk_10day","name":"十日相伴","desc":"累计使用系统满10天","exp":80},
    {"aid":"walk_30day","name":"月度同行","desc":"累计使用系统满30天","exp":180},
    {"aid":"walk_100day","name":"百秩征途","desc":"累计使用系统满100天","exp":400},

    # ========== 导入成绩（新增） ==========
    {"aid":"import_1","name":"初导名册","desc":"第一次导入成绩/学生名单","exp":40},
    {"aid":"import_3","name":"三番汇册","desc":"累计导入成绩/学生名单3次","exp":100},
    {"aid":"import_5","name":"五度整编","desc":"累计导入成绩/学生名单5次","exp":200},
    {"aid":"import_10","name":"十载汇录","desc":"累计导入成绩/学生名单10次","exp":350},

    # ========== 录入记录 ==========
    {"aid":"record_first","name":"落笔初记","desc":"录入第一条班级加减分记录","exp":40},
    {"aid":"record_100","name":"百册留痕","desc":"累计录入100条积分记录","exp":150},
    {"aid":"record_500","name":"台账盈积","desc":"累计录入500条积分记录","exp":350},
    {"aid":"record_2000","name":"宏卷成书","desc":"累计录入2000条积分记录","exp":800},
    {"aid":"record_5000","name":"千章叠卷","desc":"累计录入5000条积分记录","exp":1100},
    {"aid":"record_10000","name":"万序归档","desc":"累计录入10000条积分记录","exp":1600},
    {"aid":"record_50000","name":"瀚卷浩册","desc":"累计录入50000条积分记录","exp":2400},
    {"aid":"record_100000","name":"鸿篇巨录","desc":"累计录入100000条积分记录","exp":3200},

    # ========== 签到 ==========
    {"aid":"sign_7","name":"七日不渝","desc":"达成连续签到7天","exp":120},
    {"aid":"sign_30","name":"月余坚守","desc":"达成连续签到30天","exp":300},
    {"aid":"sign_50","name":"半岁恒心","desc":"达成连续签到50天","exp":450},

    # ========== 导出备份探索 ==========
    {"aid":"export_first","name":"数表初成","desc":"第一次导出Excel汇总表格","exp":60},
    {"aid":"backup_first","name":"存鉴留档","desc":"完成第一次数据备份","exp":60}
]

def get_today_date():
    """获取今天的日期字符串"""
    return datetime.now().strftime("%Y-%m-%d")

def reset_daily_experience_if_needed():
    cfg = load_config()
    today = get_today_date()
    last_date = cfg.get("last_experience_date", "")
    if last_date != today:
        cfg["today_experience"] = 0
        cfg["online_seconds_today"] = 0
        cfg["today_record_exp"] = 0        # 之前漏掉的
        cfg["today_online_exp"] = 0        # 之前漏掉的
        cfg["last_online_update"] = time.time()  # 重置为当前时间，避免跨天时用旧时间戳
        cfg["last_experience_date"] = today
        save_config(cfg)
        return True
    return False

def get_remaining_today_exp():
    """获取今日剩余可获取经验（全局上限）"""
    cfg = load_config()
    reset_daily_experience_if_needed()
    today_exp = cfg.get("today_experience", 0)
    remaining = DAILY_LIMIT_TOTAL - today_exp
    return max(0, remaining)

def can_add_experience(amount):
    """检查今日是否还能添加经验"""
    return amount <= get_remaining_today_exp()

def add_experience(amount, source="未知", check_limit=True, update_daily=True):
    """
    添加经验（核心函数）
    :param amount: 经验值（正数）
    :param source: 来源描述（用于日志）
    :param check_limit: 是否检查每日上限
    :param update_daily: 是否计入今日经验（不影响总经验）
    :return: (是否成功, 实际添加的经验值, 是否升级)
    """
    if amount <= 0:
        return False, 0, False
    
    cfg = load_config()
    reset_daily_experience_if_needed()
    
    # 检查每日上限（若需要）
    if check_limit:
        remaining = get_remaining_today_exp()
        if remaining <= 0:
            return False, 0, False
        actual_amount = min(amount, remaining)
    else:
        actual_amount = amount
    
    # 更新今日经验（只有当 update_daily=True 时）
    if update_daily:
        cfg["today_experience"] = cfg.get("today_experience", 0) + actual_amount
    
    # 更新总经验（始终更新）
    old_total = cfg.get("total_experience", 0)
    new_total = old_total + actual_amount
    cfg["total_experience"] = new_total
    
    # 计算等级变化
    old_level = calculate_level_by_exp(old_total)
    new_level = calculate_level_by_exp(new_total)
    is_level_up = new_level > old_level
    
    save_config(cfg)
    
    
    if is_level_up and POST_BOOT_READY.get("value", False):
        QTimer.singleShot(100, lambda: show_level_up_notification(new_level))
        QTimer.singleShot(200, check_and_unlock_achievements)

    
    
    return True, actual_amount, is_level_up

    
    
    

    

def calculate_level_by_exp(total_exp):
    """根据总经验计算等级（1~20）"""
    level = 1
    for lv, exp_req in LEVEL_EXP_REQUIREMENTS.items():
        if total_exp >= exp_req:
            level = lv
        else:
            break
    return min(level, MAX_LEVEL)

def get_level_progress(total_exp=None):
    """获取当前等级进度 (当前经验, 升级所需经验, 进度百分比)"""
    if total_exp is None:
        cfg = load_config()
        total_exp = cfg.get("total_experience", 0)
    
    current_level = calculate_level_by_exp(total_exp)
    if current_level >= MAX_LEVEL:
        return total_exp, 0, 100  # 满级无进度
    
    next_level = current_level + 1
    current_req = LEVEL_EXP_REQUIREMENTS.get(current_level, 0)
    next_req = LEVEL_EXP_REQUIREMENTS.get(next_level, 0)
    progress = (total_exp - current_req) / (next_req - current_req) * 100
    return total_exp, next_req - total_exp, min(100, progress)

# 主窗口是否已完成"启动后置任务"（新手引导期间不弹成就/升级通知）
POST_BOOT_READY = {"value": False}

def show_level_up_notification(new_level):
    """显示升级通知（自定义弹窗，跟随主题）"""
    from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
    title = LEVEL_TITLES.get(new_level, "未知")
    _c = get_theme_colors()

    dialog = QDialog()
    dialog.setWindowTitle("🎉 升级")
    dialog.setFixedSize(400, 200)
    dialog.setStyleSheet(DIALOG_STYLE)   # ✅ 跟随主题

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(10)

    title_label = QLabel(f"🎉 升级！Lv.{new_level}")
    title_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #64B5F6;")
    layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignCenter)

    content_label = QLabel(f"恭喜晋升「{title}」！")
    content_label.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
    layout.addWidget(content_label, alignment=Qt.AlignmentFlag.AlignCenter)

    btn_layout = QHBoxLayout()
    btn_layout.addStretch()
    ok_btn = QPushButton("知道了")
    ok_btn.setFixedWidth(120)
    ok_btn.clicked.connect(dialog.accept)
    btn_layout.addWidget(ok_btn)
    btn_layout.addStretch()
    layout.addLayout(btn_layout)

    screen = QApplication.primaryScreen().availableGeometry()
    dialog.move(screen.center() - dialog.rect().center())
    dialog.exec()

# ======================== 成就系统核心 ========================
def get_achievement_by_id(aid):
    for ach in ACHIEVEMENT_LIST:
        if ach["aid"] == aid:
            return ach
    return None

def is_achievement_unlocked(aid):
    cfg = load_config()
    return cfg.get("achievements", {}).get(aid, False)

def unlock_achievement(aid, show_notification=True):
    """解锁成就，发放经验，保存状态，弹窗"""
    if is_achievement_unlocked(aid):
        return False
    ach = get_achievement_by_id(aid)
    if not ach:
        return False
    exp = ach.get("exp", 0)
    if exp > 0:
        # 不检查每日上限，不计入今日经验
        add_experience(exp, source=f"成就-{ach['name']}", check_limit=False, update_daily=False)
    # 标记已解锁
    cfg = load_config()
    cfg["achievements"][aid] = True
    save_config(cfg)
    if show_notification and POST_BOOT_READY.get("value", False):
        show_achievement_notification(ach, exp)
    return True

def show_achievement_notification(ach, exp):
    """自定义通知弹窗（跟随主题）"""
    from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
    _c = get_theme_colors()

    dialog = QDialog()
    dialog.setWindowTitle("🎉 解锁成就")
    dialog.setFixedSize(420, 200)
    dialog.setStyleSheet(DIALOG_STYLE)   # ✅ 不再硬编码

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(10)

    title = QLabel(f"🎉 解锁成就【{ach['name']}】")
    title.setStyleSheet("font-size: 20px; font-weight: bold; color: #64B5F6;")
    layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)

    desc = QLabel(ach['desc'])
    desc.setStyleSheet(f"color: {_c['fg']}; font-size: 14px;")   # ✅ 跟随主题
    layout.addWidget(desc, alignment=Qt.AlignmentFlag.AlignCenter)

    exp_label = QLabel(f"获得 {exp} 经验")
    exp_label.setStyleSheet("color: #FFD700; font-size: 16px; font-weight: bold;")
    layout.addWidget(exp_label, alignment=Qt.AlignmentFlag.AlignCenter)

    btn_layout = QHBoxLayout()
    btn_layout.addStretch()
    view_btn = QPushButton("查看成就")
    view_btn.clicked.connect(lambda: (dialog.accept(), open_achievement_panel()))
    confirm_btn = QPushButton("确认")
    confirm_btn.clicked.connect(dialog.accept)
    btn_layout.addWidget(view_btn)
    btn_layout.addWidget(confirm_btn)
    layout.addLayout(btn_layout)

    screen = QApplication.primaryScreen().availableGeometry()
    dialog.move(screen.center() - dialog.rect().center())
    dialog.exec()

def open_achievement_panel():
    """从通知跳转到成就面板"""
    main = get_main_window()
    if main and hasattr(main, 'settingsInterface'):
        main.settingsInterface.open_achievement()

def check_and_unlock_achievements():
    """遍历所有成就，检测并解锁满足条件的成就"""
    cfg = load_config()
    unlocked = cfg.get("achievements", {})
    # 获取当前数据
    total_records = cfg.get("total_records_count", 0)
    sign_in_streak = cfg.get("sign_in_streak", 0)
    level = calculate_level()
    install_time = cfg.get("install_time", time.time())
    days_since_install = int((time.time() - install_time) // 86400)
    import_count = cfg.get("import_count", 0)
    export_count = cfg.get("export_count", 0)
    backup_count = cfg.get("backup_count", 0)

    for ach in ACHIEVEMENT_LIST:
        aid = ach["aid"]
        if unlocked.get(aid, False):
            continue
        # 条件判断
        if aid.startswith("lv_"):
            lv = int(aid.split("_")[1])
            if level >= lv:
                unlock_achievement(aid)
        elif aid == "first_launch":
            unlock_achievement(aid)
        elif aid.startswith("walk_"):
            days = int(aid.split("_")[1].replace("day", ""))
            if days_since_install >= days:
                unlock_achievement(aid)
        elif aid.startswith("import_"):
            count = int(aid.split("_")[1])
            if import_count >= count:
                unlock_achievement(aid)
        elif aid.startswith("record_"):
            if aid == "record_first":
                if total_records >= 1:
                    unlock_achievement(aid)
            else:
                count = int(aid.split("_")[1])
                if total_records >= count:
                    unlock_achievement(aid)
        elif aid.startswith("sign_"):
            count = int(aid.split("_")[1])
            if sign_in_streak >= count:
                unlock_achievement(aid)
        elif aid == "export_first":
            if export_count >= 1:
                unlock_achievement(aid)
        elif aid == "backup_first":
            if backup_count >= 1:
                unlock_achievement(aid)

def get_main_window():
    """获取主窗口实例"""
    from PyQt6.QtWidgets import QApplication
    return getattr(QApplication.instance(), 'main_window', None)
# ======================== 配置文件管理（与原版相同） ========================
def load_config():
    default = {
        "student_file": DEFAULT_STUDENT_EXCEL,
        "template_file": "",
        "last_export_month": "",
        "default_discount": 100,
        "default_category": "日常",
        "show_boot_animation": True,
        "install_time": time.time(),
        "backup_dir": "backups",
        "max_backups": 10,
        "last_backup_date": "",
        "custom_reasons": [
            {"name": "迟到", "type": "-", "score": 2},
            {"name": "早退", "type": "-", "score": 2},
            {"name": "卫生", "type": "-", "score": 1},
            {"name": "纪律", "type": "-", "score": 1},
            {"name": "作业", "type": "-", "score": 2},
            {"name": "课堂表现", "type": "+", "score": 2},
        ],
        "last_input_category": "日常",
        "last_input_reason": "",
        "last_input_discount": "100",
        "fix_params_serial": False,
        "enable_reason_number_shortcut": True,
        "countdown_events": [],          # 倒计时事件列表
        "selected_countdown_index": -1 ,  # -1 表示无选中
        "overview_top_n": 5,
        "overview_rank_by": "赋分班名",
        "total_records_count": 0,
        "total_experience": 0,              # 累计总经验（用于计算等级）
        "today_experience": 0,              # 今日已获得经验
        "last_experience_date": "",         # 上次获得经验的日期（用于重置每日计数）
        "sign_in_streak": 0,                # 连续签到天数
        "last_sign_in_date": "",            # 上次签到日期
        "online_seconds_today": 0,          # 今日在线秒数（用于在线计时经验）
        "last_online_update": 0.0,          # 上次更新在线时间的时间戳
        "last_device_day": "",              # 上次领取设备运行经验的天数
        "achievements": {},                 # 成就领取记录 {"成就ID": True}
        "device_total_days_recorded": 0,
        "import_count": 0,          # 导入成绩次数
        "export_count": 0,          # 导出表格次数
        "backup_count": 0,
        "recent_records_limit": 100,   # 最近录入记录显示条数
        "enable_progress_animation": True,
        "enable_particle_effect": True,      # ← 新增：弹窗关闭粒子效果
        "enable_click_particle_effect": True,   # ← 新增：鼠标点击粒子
        "enable_drag_particle_effect": True,   # ← 新增：拖动拖尾粒子
        "enable_scroll_particle_effect": True, # ← 新增：鼠标滚动粒子
        "enable_nav_animation": True,          # ← 新增：导航栏切换动画
        # ---- 主题 ----
        "theme_mode": "light",          # dark / light / custom
        "theme_bg": "#16161E",         # 自定义模式下的窗口颜色
        "theme_fg": "#E0E0E0",         # 自定义模式下的文字颜色
        # 按类型分别保留条数（每种类型各自独立计数）
        "export_keep_counts": {
            "班级积分月度汇总": 5,
            "成绩总览对比": 5,
            "其它": 5,
        },
        "backup_keep_counts": {
            "备份-积分记录": 5,
            "名单备份": 5,
            "积分记录": 5,
            "其它": 5,
        },
        # ---- 密码保护 ----
        "password_enabled": False,     # ← 新增
        "password_hash": "",           # ← 新增
        "anti_peek_enabled": False,    # ← 新增：防窥模式
        # ---- 智能解锁（U 盘） ----
        "smart_unlock_enabled": False,
        "smart_unlock_device": [],          
        "onboarding_completed": False,   # ← 新增：是否已完成新手引导
    }
    if os.path.exists(CONFIG_JSON):
        try:
            with open(CONFIG_JSON, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in default.items():
                if k not in cfg:
                    cfg[k] = v
            if cfg.get("install_time", 0) < 1700000000:
                cfg["install_time"] = time.time()

            # ---- 兼容旧版单个 dict 配置，并统一为 list（最多 3 个） ----
            dev = cfg.get("smart_unlock_device", [])
            if isinstance(dev, dict):
                dev = [dev] if dev and dev.get("serial") else []
            elif not isinstance(dev, list):
                dev = []

            clean = []
            seen = set()
            for d in dev:
                if not isinstance(d, dict) or not d.get("serial"):
                    continue
                s = str(d.get("serial", "")).strip().upper()
                if s and s not in seen:
                    seen.add(s)
                    clean.append(d)
            cfg["smart_unlock_device"] = clean[:3]
            return cfg
        except:
            return default
    return default

def save_config(cfg):
    """原子写入：先写临时文件，再替换，避免写一半被中断导致配置损坏"""
    tmp_path = CONFIG_JSON + ".tmp"
    try:
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, CONFIG_JSON)   # 原子替换
    except Exception as e:
        # 兜底：直接写，至少保证有配置
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
        with open(CONFIG_JSON, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)

def read_config():
    cfg = load_config()
    return cfg.get("student_file", DEFAULT_STUDENT_EXCEL)

def save_config_path(path):
    cfg = load_config()
    cfg["student_file"] = path
    save_config(cfg)

def read_template_config():
    cfg = load_config()
    return cfg.get("template_file", "")

def save_template_config(path):
    cfg = load_config()
    cfg["template_file"] = path
    save_config(cfg)

def load_custom_lists():
    """加载自定义名单数据，返回列表，每个元素 {"name": 名单名, "items": [{"category":类别, "members":[学号], "bonus":加分}]}"""
    if os.path.exists(CUSTOM_LISTS_FILE):
        try:
            with open(CUSTOM_LISTS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    # 检查是否为新格式（第一个元素有 "items" 键）
                    if data and "items" in data[0]:
                        return data
                    else:
                        # 旧格式转换：将旧名单转为只有一个类别的条目
                        new_data = []
                        for item in data:
                            if "name" in item and "members" in item:
                                new_item = {
                                    "name": item["name"],
                                    "items": [{"category": item["name"], "members": item["members"], "bonus": 0}]
                                }
                                new_data.append(new_item)
                        # 保存新格式
                        save_custom_lists(new_data)
                        return new_data
        except:
            pass
    return []

def save_custom_lists(lists):
    """保存自定义名单数据"""
    with open(CUSTOM_LISTS_FILE, "w", encoding="utf-8") as f:
        json.dump(lists, f, ensure_ascii=False, indent=2)

# ======================== 备份管理器（与原版相同） ========================
class BackupManager:
    def __init__(self, backup_dir="backups", max_backups=10):
        self.backup_dir = backup_dir
        self.max_backups = max_backups
        os.makedirs(self.backup_dir, exist_ok=True)

    def _cleanup_old_backups(self, pattern):
        files = sorted(glob.glob(os.path.join(self.backup_dir, pattern)))
        for old in files[:-self.max_backups]:
            try:
                os.remove(old)
            except Exception:
                pass

    def save_student_backup(self, stu_df):
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"名单备份_{ts}.csv"
            filepath = os.path.join(self.backup_dir, filename)
            backup = stu_df[["学号", "姓名"]].copy()
            backup.to_csv(filepath, index=False, encoding="utf-8-sig")
            self._cleanup_old_backups("名单备份_*.csv")
            return True
        except Exception as e:
            print(f"名单备份失败: {e}")
            return False

    def get_student_backups(self):
        return sorted(glob.glob(os.path.join(self.backup_dir, "名单备份_*.csv")))

    def load_student_backup(self, filepath=None):
        if filepath is None:
            backups = self.get_student_backups()
            if not backups:
                return None
            filepath = backups[-1]
        return self._load_csv_backup(filepath)

    @staticmethod
    def _load_csv_backup(filepath):
        try:
            df = pd.read_csv(filepath, encoding="utf-8-sig")
            if "学号" not in df.columns or "姓名" not in df.columns:
                return None
            df["学号"] = df["学号"].astype(str)
            return df
        except Exception:
            return None

    def backup_score_file(self, source_path):
        today = datetime.now().strftime("%Y%m%d")
        backup_name = f"积分记录_{today}.csv"
        backup_path = os.path.join(self.backup_dir, backup_name)
        if not os.path.exists(backup_path) and os.path.exists(source_path):
            shutil.copy2(source_path, backup_path)
            self._cleanup_old_backups("积分记录_*.csv")
            return True
        return False

    def get_score_backups(self):
        return sorted(glob.glob(os.path.join(self.backup_dir, "积分记录_*.csv")))

# ======================== 数据层函数（与原版相同） ========================
def load_student_list(file_path):
    try:
        df_stu = pd.read_excel(file_path, engine="openpyxl")
        if "学号" not in df_stu.columns or "姓名" not in df_stu.columns:
            return None, None, f" 文件 {file_path} 必须包含“学号”和“姓名”两列"
        df_stu = df_stu[["学号", "姓名"]].drop_duplicates(subset=["学号"])
        df_stu["学号"] = df_stu["学号"].astype(str)
        id_to_name = dict(zip(df_stu["学号"], df_stu["姓名"]))
        return df_stu, id_to_name, f"成功载入 {len(df_stu)} 名学生（来自 {file_path}）"
    except FileNotFoundError:
        return None, None, f" 找不到文件：{file_path}"
    except Exception as e:
        return None, None, f" 读取失败：{e}"

def init_score_file():
    if not os.path.exists(DATA_CSV):
        cols = ["学号", "姓名", "日期", "变动类型", "类别", "分数", "赋分总分", "备注"]
        pd.DataFrame(columns=cols).to_csv(DATA_CSV, index=False, encoding="utf-8-sig")

def save_score_record(stu_id, name, date_str, op, category, score, note, assign_score=None):
    new_line = {
        "学号": stu_id,
        "姓名": name,
        "日期": date_str,
        "变动类型": op,
        "类别": category,
        "分数": score,
        "赋分总分": assign_score if assign_score is not None else "",
        "备注": note
    }
    try:
        df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        df = pd.concat([df, pd.DataFrame([new_line])], ignore_index=True)
        df.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
        

        # ---- 新增：累计计数 +1 ----
        cfg = load_config()
        cfg["total_records_count"] = cfg.get("total_records_count", 0) + 1
        save_config(cfg)

        # ---- 新增：录入记录经验 +2（受每日上限和子上限限制） ----
        # 检查记录子上限（每日最多60经验，即30条）
        today = get_today_date()
        today_record_exp = cfg.get("today_record_exp", 0)
        if today_record_exp < DAILY_LIMIT_RECORD:
            # 计算本次可获得经验（每条+2，最多到子上限）
            exp_gain = min(RECORD_EXP_PER_ENTRY, DAILY_LIMIT_RECORD - today_record_exp)
            if exp_gain > 0:
                # 尝试添加经验（会自动检查全局上限）
                success, actual, is_level_up = add_experience(exp_gain, "录入记录")
                if success:
                    # 更新今日记录经验计数
                    cfg = load_config()
                    cfg["today_record_exp"] = cfg.get("today_record_exp", 0) + actual
                    save_config(cfg)
                    # 刷新主窗口等级
                    main_win = getattr(QApplication.instance(), 'main_window', None)
                    if main_win and hasattr(main_win, 'update_title_with_level'):
                        main_win.update_title_with_level()
        # 刷新主窗口标题
        main_win = getattr(QApplication.instance(), 'main_window', None)
        if main_win and hasattr(main_win, 'update_title_with_level'):
            main_win.update_title_with_level()
            check_and_unlock_achievements()
        return True, " 记录保存成功"
    except Exception as e:
        return False, f" 保存失败：{e}"

def calc_total_score(stu_df, month=None):
    try:
        df_log = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
    except FileNotFoundError:
        return None, "暂无积分记录"
    if df_log.empty:
        return None, "暂无积分记录"
    if month:
        df_log["日期"] = pd.to_datetime(df_log["日期"])
        df_log = df_log[df_log["日期"].dt.strftime("%Y-%m") == month]
        if df_log.empty:
            return None, f"{month} 无任何积分记录"

    if '类别' not in df_log.columns:
        df_log['类别'] = '日常'

    # 透视表（排除“赋分”类别）
    pivot = df_log[df_log['类别'] != '赋分'].pivot_table(
        index=["学号", "姓名"], columns="类别", values="分数", aggfunc='sum', fill_value=0
    ).reset_index()
    category_mapping = ['日常', '听写', '背诵', '班干加分', '考试加分']
    for cat in category_mapping:
        if cat not in pivot.columns:
            pivot[cat] = 0

    # 计算赋分总分
    assign_df = df_log.groupby(["学号", "姓名"])["赋分总分"].apply(
        lambda x: pd.to_numeric(x, errors='coerce').sum()
    ).reset_index()
    assign_df.columns = ["学号", "姓名", "赋分总分"]
    assign_df["学号"] = assign_df["学号"].astype(str)
    assign_df["赋分总分"] = assign_df["赋分总分"].fillna(0)

    # 合并所有学生
    pivot["学号"] = pivot["学号"].astype(str)
    all_stu = pd.merge(stu_df[["学号", "姓名"]], pivot, how="left", on=["学号", "姓名"])
    for cat in category_mapping:
        if cat not in all_stu.columns:
            all_stu[cat] = 0
        else:
            all_stu[cat] = all_stu[cat].fillna(0)

    all_stu = pd.merge(all_stu, assign_df, how="left", on=["学号", "姓名"])
    if '赋分总分_y' in all_stu.columns:
        all_stu['赋分总分'] = all_stu['赋分总分_y'].fillna(0)
        all_stu.drop(columns=['赋分总分_y'], inplace=True)
    elif '赋分总分' not in all_stu.columns:
        all_stu['赋分总分'] = 0
    else:
        all_stu['赋分总分'] = all_stu['赋分总分'].fillna(0)

    # 计算总积分，并将所有数值列转为 float（避免精度问题）
    all_stu['累计总积分'] = all_stu[category_mapping].sum(axis=1) + all_stu['赋分总分']
    # 强制转换为数值类型
    all_stu['累计总积分'] = pd.to_numeric(all_stu['累计总积分'], errors='coerce').fillna(0)

    # 先按总积分降序、学号升序排序，再计算排名
    all_stu = all_stu.sort_values(["累计总积分", "学号"], ascending=[False, True])
    # 重置索引，使排名与行号对应
    all_stu = all_stu.reset_index(drop=True)

    # 计算排名（密集排名，即相同积分排名相同，下一名次跳过）
    all_stu["排名"] = all_stu["累计总积分"].rank(method="min", ascending=False).astype(int)

    # 将整数列转换为 int（优化显示）
    for col in category_mapping + ['赋分总分', '累计总积分']:
        if col in all_stu.columns and all_stu[col].apply(lambda x: x.is_integer()).all():
            all_stu[col] = all_stu[col].astype(int)

    return all_stu, None

def import_assign_scores(stu_df, excel_path, discount=100):
    try:
        df_import = pd.read_excel(excel_path, engine="openpyxl")
        if "学号" not in df_import.columns:
            return 0, ["Excel中没有'学号'列"]
        score_col = None
        for col in df_import.columns:
            if str(col).strip() == "赋分总分":
                score_col = col
                break
        if score_col is None:
            return 0, ["Excel中没有找到'赋分总分'列"]
        df_import["学号"] = df_import["学号"].astype(str)
        df_import["分数"] = pd.to_numeric(df_import[score_col], errors='coerce').fillna(0)
        df_import["分数"] = df_import["分数"] * (discount / 100.0)
        df_import["分数"] = df_import["分数"].round(2)

        try:
            df_log = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        except:
            df_log = pd.DataFrame(columns=["学号", "姓名", "日期", "变动类型", "类别", "分数", "赋分总分", "备注"])
        if not df_log.empty:
            df_log["学号"] = df_log["学号"].astype(str)
            df_log["类别"] = df_log["类别"].astype(str).str.strip()

        import_ids = set(df_import["学号"].values)

        if not df_log.empty:
            df_log["类别_清理"] = df_log["类别"].str.strip()
            mask = ~((df_log["学号"].isin(import_ids)) & (df_log["类别_清理"] == "赋分"))
            df_log = df_log[mask].copy()
            df_log.drop(columns=["类别_清理"], inplace=True)

        new_records = []
        success_list = []
        fail_list = []
        for _, row in df_import.iterrows():
            stu_id = row["学号"]
            if stu_id not in stu_df["学号"].values:
                fail_list.append(stu_id)
                continue
            name = stu_df[stu_df["学号"]==stu_id]["姓名"].values[0]
            new_records.append({
                "学号": stu_id,
                "姓名": name,
                "日期": datetime.now().strftime("%Y-%m-%d"),
                "变动类型": "+",
                "类别": "赋分",
                "分数": 0,
                "赋分总分": row["分数"],
                "备注": f"批量导入赋分，折扣{discount}%"
            })
            success_list.append(stu_id)

        if new_records:
            df_new = pd.DataFrame(new_records)
            df_log = pd.concat([df_log, df_new], ignore_index=True)
            df_log.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
            cfg = load_config()
            cfg["import_count"] = cfg.get("import_count", 0) + 1
            save_config(cfg)
            check_and_unlock_achievements()

            # ---- 新增：累计计数增加 ----
            cfg = load_config()
            cfg["total_records_count"] = cfg.get("total_records_count", 0) + len(new_records)
            save_config(cfg)

            # ---- 新增：批量导入经验（每条+2，受每日上限和子上限限制） ----
            today = get_today_date()
            cfg = load_config()
            today_record_exp = cfg.get("today_record_exp", 0)
            # 计算可添加的条数（受子上限限制）
            max_records = (DAILY_LIMIT_RECORD - today_record_exp) // RECORD_EXP_PER_ENTRY
            if max_records > 0:
                exp_count = min(len(new_records), max_records)
                exp_gain = exp_count * RECORD_EXP_PER_ENTRY
                if exp_gain > 0:
                    success, actual, is_level_up = add_experience(exp_gain, "批量导入")
                    if success:
                        cfg = load_config()
                        cfg["today_record_exp"] = cfg.get("today_record_exp", 0) + actual
                        save_config(cfg)
                        main_win = getattr(QApplication.instance(), 'main_window', None)
                        if main_win and hasattr(main_win, 'update_title_with_level'):
                            main_win.update_title_with_level()
            # 刷新主窗口标题
            main_win = getattr(QApplication.instance(), 'main_window', None)
            if main_win and hasattr(main_win, 'update_title_with_level'):
                main_win.update_title_with_level()

        return len(success_list), fail_list
    except Exception as e:
        return 0, [str(e)]

def export_excel_with_template(stu_df, template_path, month=None, output_name=None):
    if not os.path.exists(template_path):
        return False, f"模板文件不存在：{template_path}"
    try:
        wb = load_workbook(template_path)
        ws = wb.active
        header_row = 1
        col_map = {}
        for col in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=header_row, column=col).value
            if cell_value:
                col_map[str(cell_value).strip()] = col

        required = ['班级', '学号', '姓名']
        for key in required:
            if key not in col_map:
                return False, f"模板中找不到列标题：'{key}'"

        total_col = None
        if '总分' in col_map:
            total_col = '总分'
        elif '总计' in col_map:
            total_col = '总计'

        category_mapping = {
            '日常': '日常',
            '听写': '听写',
            '背诵': '背诵',
            '班干加分': '班干加分',
            '考试加分': '考试加分'
        }
        template_categories = {}
        for sys_cat, col_name in category_mapping.items():
            if col_name in col_map:
                template_categories[col_name] = sys_cat

        try:
            df_log = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        except:
            return False, "读取积分记录失败"
        if df_log.empty:
            return False, "无数据可导出"

        if month:
            df_log["日期"] = pd.to_datetime(df_log["日期"])
            df_log = df_log[df_log["日期"].dt.strftime("%Y-%m") == month]
            if df_log.empty:
                return False, f"{month} 无任何积分记录"

        if '类别' not in df_log.columns:
            df_log['类别'] = '日常'

        grouped = df_log.groupby(["学号", "姓名", "类别"])["分数"].sum().reset_index()
        pivot = grouped.pivot_table(index=["学号", "姓名"], columns="类别", values="分数", fill_value=0).reset_index()
        pivot["学号"] = pivot["学号"].astype(str)

        for cat in category_mapping.keys():
            if cat not in pivot.columns:
                pivot[cat] = 0

        assign_df = df_log.groupby(["学号", "姓名"])["赋分总分"].apply(
            lambda x: pd.to_numeric(x, errors='coerce').sum()
        ).reset_index()
        assign_df.columns = ["学号", "姓名", "赋分总分"]
        assign_df["学号"] = assign_df["学号"].astype(str)
        assign_df["赋分总分"] = assign_df["赋分总分"].fillna(0)

        all_stu = pd.merge(stu_df[["学号", "姓名"]], pivot, how="left", on=["学号", "姓名"])
        for cat in category_mapping.keys():
            if cat not in all_stu.columns:
                all_stu[cat] = 0
            else:
                all_stu[cat] = all_stu[cat].fillna(0)

        all_stu = pd.merge(all_stu, assign_df, how="left", on=["学号", "姓名"])
        if '赋分总分_y' in all_stu.columns:
            all_stu['赋分总分'] = all_stu['赋分总分_y'].fillna(0)
            all_stu.drop(columns=['赋分总分_y'], inplace=True)
        elif '赋分总分' not in all_stu.columns:
            all_stu['赋分总分'] = 0
        else:
            all_stu['赋分总分'] = all_stu['赋分总分'].fillna(0)

        all_stu['总积分'] = all_stu[list(category_mapping.keys())].sum(axis=1) + all_stu['赋分总分']
        all_stu = all_stu.sort_values(['总积分', '学号'], ascending=[False, True])

        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

        row_idx = 2
        for _, row_data in all_stu.iterrows():
            if '班级' in col_map:
                ws.cell(row=row_idx, column=col_map['班级'], value=695)
            ws.cell(row=row_idx, column=col_map['学号'], value=row_data['学号'])
            ws.cell(row=row_idx, column=col_map['姓名'], value=row_data['姓名'])
            for col_name, sys_cat in template_categories.items():
                col_idx = col_map[col_name]
                value = row_data[sys_cat]
                if isinstance(value, float) and value.is_integer():
                    value = int(value)
                ws.cell(row=row_idx, column=col_idx, value=value)
            if total_col:
                col_idx = col_map[total_col]
                total_value = row_data['总积分']
                if isinstance(total_value, float) and total_value.is_integer():
                    total_value = int(total_value)
                ws.cell(row=row_idx, column=col_idx, value=total_value)
            if '赋分总分' in col_map:
                col_idx = col_map['赋分总分']
                assign_val = row_data['赋分总分']
                if isinstance(assign_val, float) and assign_val.is_integer():
                    assign_val = int(assign_val)
                ws.cell(row=row_idx, column=col_idx, value=assign_val)
            row_idx += 1

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center_alignment = Alignment(horizontal='center', vertical='center')

        if row_idx > 2:
            for row in range(2, row_idx):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    if cell.coordinate not in ws.merged_cells:
                        cell.border = thin_border
                        cell.alignment = center_alignment

        if output_name:
            output_file = output_name
        elif month:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(EXPORT_DIR, f"班级积分月度汇总_{month}_{ts}.xlsx")
        else:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(EXPORT_DIR, f"班级积分月度汇总_{ts}.xlsx")
        wb.save(output_file)
        return True, f" 已导出（使用模板，含分类数据，带边框居中）：{output_file}"

    except Exception as e:
        return False, f" 导出失败：{str(e)}"

def export_excel_fallback(stu_df, month=None, output_name=None):
    ranking_df, err = calc_total_score(stu_df, month)
    if ranking_df is None or ranking_df.empty:
        return False, err if err else "无排行数据"

    df_export = ranking_df.copy()
    df_export["总计"] = df_export["日常"] + df_export["听写"] + df_export["背诵"]
    df_export["总分"] = df_export["累计总积分"]
    df_export.insert(0, "班级", 695)

    cols_order = ["班级", "学号", "姓名", "听写", "背诵", "日常", "总计", "班干加分", "赋分总分", "考试加分", "总分"]
    df_export = df_export[cols_order]
    df_export = df_export.sort_values("总分", ascending=False)

    if output_name:
        output_file = output_name
    elif month:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(EXPORT_DIR, f"班级积分月度汇总_{month}_{ts}.xlsx")
    else:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(EXPORT_DIR, f"班级积分月度汇总_{ts}.xlsx")

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_export.to_excel(writer, sheet_name="积分汇总", index=False)

            wb = writer.book
            ws = writer.sheets["积分汇总"]

            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_alignment = Alignment(horizontal='center', vertical='center')

            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    if cell.value is not None:
                        cell.border = thin_border
                        cell.alignment = center_alignment

            for col in ws.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        val = str(cell.value)
                        length = len(val) + 2
                        if length > max_length:
                            max_length = length
                    except:
                        pass
                ws.column_dimensions[column_letter].width = max_length

        return True, f" 已导出（汇总表，列宽自适应，带边框居中）：{output_file}"
    except Exception as e:
        return False, f" 导出失败：{e}"

def export_score_overview(output_name=None):
    """导出成绩总览对比表（与成绩界面总览显示完全一致）"""
    if not os.path.exists(CURRENT_SCORES_FILE) or not os.path.exists(PREVIOUS_SCORES_FILE):
        return False, "缺少本次或上次考试数据，无法导出总览"

    try:
        current = pd.read_csv(CURRENT_SCORES_FILE, encoding="utf-8-sig")
        previous = pd.read_csv(PREVIOUS_SCORES_FILE, encoding="utf-8-sig")
    except Exception as e:
        return False, f"读取成绩数据失败：{e}"

    if current.empty or previous.empty:
        return False, "缺少本次或上次考试数据，无法导出总览"

    current["学号"] = current["学号"].astype(str)
    previous["学号"] = previous["学号"].astype(str)

    # 合并数据（与 display_overview 完全一致）
    merged = pd.merge(current, previous, on="学号", how="outer", suffixes=("_本次", "_上次"))
    merged["姓名"] = merged["姓名_本次"].fillna(merged["姓名_上次"])

    # 确保列存在
    for col in ["赋分总分_本次", "赋分校名_本次", "赋分班名_本次"]:
        if col not in merged.columns:
            merged[col] = ""
    for col in ["赋分总分_上次", "赋分校名_上次", "赋分班名_上次"]:
        if col not in merged.columns:
            merged[col] = ""

    # 计算差值
    merged["_diff_total"] = (
        pd.to_numeric(merged["赋分总分_本次"], errors='coerce') -
        pd.to_numeric(merged["赋分总分_上次"], errors='coerce')
    )
    merged["_diff_school"] = -(
        pd.to_numeric(merged["赋分校名_本次"], errors='coerce') -
        pd.to_numeric(merged["赋分校名_上次"], errors='coerce')
    )
    merged["_diff_class"] = -(
        pd.to_numeric(merged["赋分班名_本次"], errors='coerce') -
        pd.to_numeric(merged["赋分班名_上次"], errors='coerce')
    )

    def format_diff(val):
        if pd.isna(val):
            return ""
        if val > 0:
            return f"+{val:.1f}" if isinstance(val, float) and val % 1 != 0 else f"+{int(val)}"
        else:
            return f"{val:.1f}" if isinstance(val, float) and val % 1 != 0 else f"{int(val)}"

    merged["较上次赋分总分"] = merged["_diff_total"].apply(format_diff)
    merged["较上次赋分校名"] = merged["_diff_school"].apply(lambda x: format_diff(x) if pd.notna(x) else "")
    merged["较上次赋分班名"] = merged["_diff_class"].apply(lambda x: format_diff(x) if pd.notna(x) else "")

    # 按学号排序
    merged = merged.dropna(subset=["学号"])
    try:
        merged["学号_int"] = merged["学号"].astype(int)
        merged = merged.sort_values("学号_int").reset_index(drop=True)
    except (ValueError, TypeError):
        merged = merged.sort_values("学号").reset_index(drop=True)

    display_cols = [
        "学号", "姓名",
        "赋分总分_上次", "赋分校名_上次", "赋分班名_上次",
        "赋分总分_本次", "赋分校名_本次", "赋分班名_本次",
        "较上次赋分总分", "较上次赋分校名", "较上次赋分班名"
    ]
    header_names = [
        "学号", "姓名",
        "上次赋分总分", "上次赋分校名", "上次赋分班名",
        "本次赋分总分", "本次赋分校名", "本次赋分班名",
        "较上次赋分总分", "较上次校名", "较上次班名"
    ]
    df_display = merged[display_cols].copy()
    df_display.columns = header_names

    if output_name is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = os.path.join(EXPORT_DIR, f"成绩总览对比_{ts}.xlsx")

    try:
        with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
            df_display.to_excel(writer, sheet_name="成绩总览", index=False)

            wb = writer.book
            ws = writer.sheets["成绩总览"]

            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            center_alignment = Alignment(horizontal='center', vertical='center')

            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    if cell.value is not None:
                        cell.border = thin_border
                        cell.alignment = center_alignment

            # 自适应列宽
            for col in ws.columns:
                max_length = 0
                column_letter = col[0].column_letter
                for cell in col:
                    try:
                        val = str(cell.value) if cell.value is not None else ""
                        length = len(val) + 2
                        if length > max_length:
                            max_length = length
                    except:
                        pass
                ws.column_dimensions[column_letter].width = max(max_length, 12)

        return True, f" 成绩总览已导出：{output_name}"
    except Exception as e:
        return False, f" 导出失败：{e}"

# ======================== Toast 通知 ========================
def show_toast(parent, msg, type_="info", duration=2000):
    titles = {
        "success": "成功",
        "error": "错误",
        "warning": "⚠️ 警告",
        "info": "提示"
    }
    title = titles.get(type_, "提示")
    if type_ == "success":
        InfoBar.success(title=title, content=msg, parent=parent,
                        position=InfoBarPosition.TOP_RIGHT, duration=duration)
    elif type_ == "error":
        InfoBar.error(title=title, content=msg, parent=parent,
                      position=InfoBarPosition.TOP_RIGHT, duration=duration)
    elif type_ == "warning":
        InfoBar.warning(title=title, content=msg, parent=parent,
                        position=InfoBarPosition.TOP_RIGHT, duration=duration)
    else:
        InfoBar.info(title=title, content=msg, parent=parent,
                     position=InfoBarPosition.TOP_RIGHT, duration=duration)

# ======================== 弹窗关闭粒子效果（HarmonyOS 风格） ========================
class ParticleOverlay(QWidget):
    """全屏透明覆盖层：从弹窗区域向外散开的粒子动画"""

    _active = []          # 当前存活的叠加层实例
    _MAX_ACTIVE = 6       # 同时最多保留的粒子层数

    def __init__(self, center_x, center_y, width, height):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        try:
            screen = QApplication.primaryScreen().geometry()
        except Exception:
            screen = QRect(0, 0, 1280, 720)
        self.setGeometry(screen)

        # 粒子数量取决于弹窗面积
        # ★ 粒子数量：大幅提升
        area = max(1, int(width) * int(height))
        if area < 6000:
            # 很小的窗口（如 48×48 的悬浮窗）少放一些，看起来才自然
            num_particles = max(60, min(140, area // 25))
        else:
            num_particles = max(220, min(650, area // 200))

        self.particles = []
        for _ in range(num_particles):
            px = center_x + random.uniform(-width / 2.0, width / 2.0)
            py = center_y + random.uniform(-height / 2.0, height / 2.0)

            # 以弹窗中心为基准的离心方向（略微随机扰动）
            dx = px - center_x
            dy = py - center_y
            base_angle = math.atan2(dy, dx) if (dx or dy) else random.uniform(0, 2 * math.pi)
            angle = base_angle + random.uniform(-0.55, 0.55)

            speed = random.uniform(3.0, 10.0)
            vx = math.cos(angle) * speed
            vy = math.sin(angle) * speed - random.uniform(0.5, 2.5)

            size = random.uniform(1.8, 5.5)
            # 寿命略短一些，避免屏幕长时间残留，同时视觉上更"炸开"
            max_life = random.randint(34, 68)

            # 蓝紫渐变，呼应主题色
            r = random.randint(90, 180)
            g = random.randint(150, 220)
            b = random.randint(220, 255)
            color = QColor(r, g, b)

            self.particles.append({
                "x": px, "y": py,
                "vx": vx, "vy": vy,
                "size": size,
                "life": max_life, "max_life": max_life,
                "color": color,
            })

        self._timer = QTimer(self)
        self._timer.setInterval(20)
        self._timer.timeout.connect(self._step)
        self._timer.start()

        # 限制同时存活数量，防止叠加层堆积
        ParticleOverlay._active.append(self)
        while len(ParticleOverlay._active) > ParticleOverlay._MAX_ACTIVE:
            old = ParticleOverlay._active.pop(0)
            try:
                old._timer.stop()
                old.close()
            except Exception:
                pass

    def _step(self):
        any_alive = False
        for p in self.particles:
            if p["life"] > 0:
                p["x"] += p["vx"]
                p["y"] += p["vy"]
                p["vy"] += 0.20        # 重力
                p["vx"] *= 0.978       # 空气阻力
                p["vy"] *= 0.985
                p["life"] -= 1
                if p["life"] > 0:
                    any_alive = True
        self.update()
        if not any_alive:
            self._finish()

    def _finish(self):
        try:
            self._timer.stop()
        except Exception:
            pass
        try:
            if self in ParticleOverlay._active:
                ParticleOverlay._active.remove(self)
        except Exception:
            pass
        self.close()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        for p in self.particles:
            if p["life"] <= 0:
                continue
            ratio = p["life"] / p["max_life"]
            c = QColor(p["color"])
            c.setAlpha(int(230 * ratio))
            painter.setBrush(c)
            s = p["size"] * (0.35 + 0.65 * ratio)
            painter.drawEllipse(QPointF(p["x"], p["y"]), s, s)

class ClickParticleOverlay(QWidget):
    """轻量点击粒子：在鼠标点击位置溅射一圈小点"""

    _active = []
    _MAX_ACTIVE = 5          # 同时最多 5 个（点击频繁时自动清旧的）
    _MIN_INTERVAL = 0.045    # 最小间隔秒数，避免连点堆爆

    _last_spawn_ts = 0.0

    def __init__(self, center_x, center_y):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        # 覆盖整个屏幕即可
        try:
            screen = QApplication.primaryScreen().geometry()
        except Exception:
            screen = QRect(0, 0, 1280, 720)
        self.setGeometry(screen)

        # 点击粒子的数量少、寿命短、尺寸小，看起来轻快
        num = random.randint(10, 18)
        self.particles = []
        for _ in range(num):
            angle = random.uniform(0, 2 * math.pi)
            speed = random.uniform(1.6, 5.2)
            vx = math.cos(angle) * speed
            vy = math.sin(angle) * speed - random.uniform(0.4, 1.6)   # 整体略微向上

            size = random.uniform(1.5, 3.6)
            max_life = random.randint(20, 36)

            r = random.randint(110, 190)
            g = random.randint(170, 225)
            b = random.randint(230, 255)
            color = QColor(r, g, b)

            self.particles.append({
                "x": float(center_x), "y": float(center_y),
                "vx": vx, "vy": vy,
                "size": size,
                "life": max_life, "max_life": max_life,
                "color": color,
            })

        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._step)
        self._timer.start()

        ClickParticleOverlay._active.append(self)
        while len(ClickParticleOverlay._active) > ClickParticleOverlay._MAX_ACTIVE:
            old = ClickParticleOverlay._active.pop(0)
            try:
                old._timer.stop()
                old.close()
            except Exception:
                pass

    def _step(self):
        any_alive = False
        for p in self.particles:
            if p["life"] > 0:
                p["x"] += p["vx"]
                p["y"] += p["vy"]
                p["vy"] += 0.16
                p["vx"] *= 0.97
                p["vy"] *= 0.98
                p["life"] -= 1
                if p["life"] > 0:
                    any_alive = True
        self.update()
        if not any_alive:
            self._finish()

    def _finish(self):
        try:
            self._timer.stop()
        except Exception:
            pass
        try:
            if self in ClickParticleOverlay._active:
                ClickParticleOverlay._active.remove(self)
        except Exception:
            pass
        self.close()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        for p in self.particles:
            if p["life"] <= 0:
                continue
            ratio = p["life"] / p["max_life"]
            c = QColor(p["color"])
            c.setAlpha(int(220 * ratio))
            painter.setBrush(c)
            s = p["size"] * (0.4 + 0.6 * ratio)
            painter.drawEllipse(QPointF(p["x"], p["y"]), s, s)

class DragParticleOverlay(QWidget):
    """拖动拖尾粒子层：单例，持续接收坐标，在鼠标路径上追加小粒子"""

    _instance = None
    _MAX_PARTICLES = 1800       # 池子上限，超出时裁剪最旧的
    _MAX_IDLE_FRAMES = 40      # 无粒子存活后再空闲这么多帧就关闭自己

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        try:
            screen = QApplication.primaryScreen().geometry()
        except Exception:
            screen = QRect(0, 0, 1280, 720)
        self.setGeometry(screen)

        self.particles = []
        self._idle_frames = 0

        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._step)
        self._timer.start()

    # ---------- 单例入口 ----------
    @classmethod
    def get_instance(cls):
        if cls._instance is None or not cls._instance.isVisible():
            try:
                if cls._instance is not None:
                    cls._instance.close()
            except Exception:
                pass
            cls._instance = DragParticleOverlay()
            cls._instance.show()
            cls._instance.raise_()
        return cls._instance

    # ---------- 外部接口：追加一批粒子 ----------
        # ---------- 外部接口：追加一批粒子 ----------
    def add_point(self, x, y):
        self._idle_frames = 0
        # ★ 每次采样追加 10~16 颗，尾迹更浓密
        num = random.randint(10, 16)
        for _ in range(num):
            # 让粒子稍微围绕路径散开一点，形成"云带"而不是一条细线
            offset_x = random.uniform(-4.0, 4.0)
            offset_y = random.uniform(-4.0, 4.0)

            angle = random.uniform(0, 2 * math.pi)
            speed = random.uniform(0.5, 2.4)   # ★ 扩散略微加大
            vx = math.cos(angle) * speed
            vy = math.sin(angle) * speed + random.uniform(0.15, 0.9)

            size = random.uniform(1.8, 4.2)    # ★ 粒子更大
            max_life = random.randint(22, 42)  # ★ 存活更久

            # 拖尾色调偏亮青蓝，和点击粒子/关闭粒子形成统一语言
            # 额外加入亮白色高光粒子，视觉更闪耀
            if random.random() < 0.18:
                r, g, b = 230, 240, 255
            else:
                r = random.randint(140, 210)
                g = random.randint(180, 230)
                b = random.randint(240, 255)
            color = QColor(r, g, b)

            self.particles.append({
                "x": float(x) + offset_x,
                "y": float(y) + offset_y,
                "vx": vx, "vy": vy,
                "size": size,
                "life": max_life, "max_life": max_life,
                "color": color,
            })

        if len(self.particles) > self._MAX_PARTICLES:
            self.particles = self.particles[-self._MAX_PARTICLES:]

    # ---------- 逐帧推进 ----------
    def _step(self):
        any_alive = False
        for p in self.particles:
            if p["life"] > 0:
                p["x"] += p["vx"]
                p["y"] += p["vy"]
                p["vy"] += 0.07
                p["vx"] *= 0.96
                p["vy"] *= 0.975
                p["life"] -= 1
                if p["life"] > 0:
                    any_alive = True

        self.update()

        if any_alive:
            self._idle_frames = 0
        else:
            self._idle_frames += 1
            if self._idle_frames > self._MAX_IDLE_FRAMES:
                self._finish()

    def _finish(self):
        try:
            self._timer.stop()
        except Exception:
            pass
        DragParticleOverlay._instance = None
        self.close()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        for p in self.particles:
            if p["life"] <= 0:
                continue
            ratio = p["life"] / p["max_life"]
            c = QColor(p["color"])
            c.setAlpha(int(200 * ratio))
            painter.setBrush(c)
            s = p["size"] * (0.3 + 0.7 * ratio)
            painter.drawEllipse(QPointF(p["x"], p["y"]), s, s)


def _spawn_drag_particles(global_x, global_y):
    """在拖动经过的坐标处追加拖尾粒子"""
    try:
        if not load_config().get("enable_drag_particle_effect", True):
            return
        overlay = DragParticleOverlay.get_instance()
        overlay.add_point(global_x, global_y)
    except Exception as e:
        print(f"⚠️ 拖动粒子异常: {e}")


def _spawn_click_particles(global_x, global_y):
    """在全局坐标处溅射一次点击粒子（带间隔保护）"""
    try:
        if not load_config().get("enable_click_particle_effect", True):
            return
        now = time.time()
        if now - ClickParticleOverlay._last_spawn_ts < ClickParticleOverlay._MIN_INTERVAL:
            return
        ClickParticleOverlay._last_spawn_ts = now

        overlay = ClickParticleOverlay(global_x, global_y)
        overlay.show()
    except Exception as e:
        print(f"⚠️ 点击粒子异常: {e}")

class ScrollParticleOverlay(QWidget):
    """滚动粒子：鼠标滚轮滚动时，在光标位置溅射一圈偏向上的小粒子"""

    _active = []
    _MAX_ACTIVE = 6
    _MIN_INTERVAL = 0.035      # 滚动间隔保护
    _last_spawn_ts = 0.0

    def __init__(self, center_x, center_y):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        try:
            screen = QApplication.primaryScreen().geometry()
        except Exception:
            screen = QRect(0, 0, 1280, 720)
        self.setGeometry(screen)

        # 数量少、寿命短、向上飘
        num = random.randint(6, 12)
        self.particles = []
        for _ in range(num):
            # 方向偏上：-90° 左右散布
            base_angle = -math.pi / 2.0 + random.uniform(-0.9, 0.9)
            speed = random.uniform(1.2, 3.5)
            vx = math.cos(base_angle) * speed
            vy = math.sin(base_angle) * speed

            size = random.uniform(1.4, 3.0)
            max_life = random.randint(18, 32)

            # 色调偏亮青蓝
            r = random.randint(120, 190)
            g = random.randint(180, 225)
            b = random.randint(235, 255)
            color = QColor(r, g, b)

            self.particles.append({
                "x": float(center_x), "y": float(center_y),
                "vx": vx, "vy": vy,
                "size": size,
                "life": max_life, "max_life": max_life,
                "color": color,
            })

        self._timer = QTimer(self)
        self._timer.setInterval(16)
        self._timer.timeout.connect(self._step)
        self._timer.start()

        ScrollParticleOverlay._active.append(self)
        while len(ScrollParticleOverlay._active) > ScrollParticleOverlay._MAX_ACTIVE:
            old = ScrollParticleOverlay._active.pop(0)
            try:
                old._timer.stop()
                old.close()
            except Exception:
                pass

    def _step(self):
        any_alive = False
        for p in self.particles:
            if p["life"] > 0:
                p["x"] += p["vx"]
                p["y"] += p["vy"]
                p["vy"] += 0.10        # 轻微重力
                p["vx"] *= 0.97
                p["vy"] *= 0.98
                p["life"] -= 1
                if p["life"] > 0:
                    any_alive = True
        self.update()
        if not any_alive:
            self._finish()

    def _finish(self):
        try:
            self._timer.stop()
        except Exception:
            pass
        try:
            if self in ScrollParticleOverlay._active:
                ScrollParticleOverlay._active.remove(self)
        except Exception:
            pass
        self.close()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        for p in self.particles:
            if p["life"] <= 0:
                continue
            ratio = p["life"] / p["max_life"]
            c = QColor(p["color"])
            c.setAlpha(int(210 * ratio))
            painter.setBrush(c)
            s = p["size"] * (0.35 + 0.65 * ratio)
            painter.drawEllipse(QPointF(p["x"], p["y"]), s, s)


def _spawn_scroll_particles(global_x, global_y):
    """在鼠标滚轮滚动位置溅射一次轻量粒子（带间隔保护）"""
    try:
        if not load_config().get("enable_scroll_particle_effect", True):
            return
        now = time.time()
        if now - ScrollParticleOverlay._last_spawn_ts < ScrollParticleOverlay._MIN_INTERVAL:
            return
        ScrollParticleOverlay._last_spawn_ts = now

        overlay = ScrollParticleOverlay(global_x, global_y)
        overlay.show()
    except Exception as e:
        print(f"⚠️ 滚动粒子异常: {e}")


def _spawn_close_particles(widget):
    """在 widget 当前屏幕上位置播放一次粒子散开效果"""
    try:
        if not load_config().get("enable_particle_effect", True):
            return
        if widget is None:
            return
        rect = widget.geometry()
        if rect.width() <= 0 or rect.height() <= 0:
            return
        try:
            global_tl = widget.mapToGlobal(QPoint(0, 0))
        except Exception:
            return
        cx = global_tl.x() + rect.width() / 2.0
        cy = global_tl.y() + rect.height() / 2.0
        overlay = ParticleOverlay(cx, cy, rect.width(), rect.height())
        overlay.show()
    except Exception as e:
        print(f"⚠️ 粒子效果异常: {e}")


class ParticleEventFilter(QObject):
    """全局事件过滤器：
       - QDialog 隐藏 → 关闭粒子
       - 鼠标左键按下 → 点击粒子
       - 鼠标左键按住拖动 → 拖尾粒子
    只在读取事件坐标，从不 consume 事件（返回 False），
    因此对悬浮窗拖动、滚动条、窗口拖动等一切正常交互零影响。
    """

    # ★ 新增：应用就绪标志。主窗口完全构造完成后才置 True，
    #   开机动画阶段 / 主窗口构造阶段所有粒子触发都被屏蔽。
    _app_ready = False

    # ---- 拖动状态（类级别，跨事件共享）----
    _mouse_pressed = False
    _last_drag_pos = None
    _last_drag_ts = 0.0
    _DRAG_MIN_DIST_SQ = 9.0       # 3px 平方
    _DRAG_MIN_INTERVAL = 0.011     # ~90Hz 采样上限

    def eventFilter(self, obj, event):
        # ★ 应用还没就绪：原样放行所有事件，不触发任何粒子
        if not ParticleEventFilter._app_ready:
            return False
        
        try:
            et = event.type()

            # ---- 1. 弹窗隐藏：关闭粒子 ----
            if isinstance(obj, QDialog):
                if et == QEvent.Type.Hide:
                    if not getattr(obj, "_particle_hide_emitted", False):
                        obj._particle_hide_emitted = True
                        _spawn_close_particles(obj)
                elif et == QEvent.Type.Show:
                    obj._particle_hide_emitted = False

            # ---- 2. 鼠标事件 ----
            if et == QEvent.Type.MouseButtonPress:
                try:
                    if hasattr(event, "button") and event.button() == Qt.MouseButton.LeftButton:
                        gp = event.globalPosition().toPoint()
                        ParticleEventFilter._mouse_pressed = True
                        ParticleEventFilter._last_drag_pos = gp
                        ParticleEventFilter._last_drag_ts = time.time()
                        # 点击粒子
                        _spawn_click_particles(gp.x(), gp.y())
                except Exception:
                    pass

            elif et == QEvent.Type.MouseButtonRelease:
                try:
                    if hasattr(event, "button") and event.button() == Qt.MouseButton.LeftButton:
                        ParticleEventFilter._mouse_pressed = False
                        ParticleEventFilter._last_drag_pos = None
                except Exception:
                    pass

            elif et == QEvent.Type.MouseMove:
                if ParticleEventFilter._mouse_pressed:
                    try:
                        # buttons() 更可靠，某些环境事件序列不完整
                        btns = getattr(event, "buttons", None)
                        if btns is not None and callable(btns):
                            if not (btns() & Qt.MouseButton.LeftButton):
                                ParticleEventFilter._mouse_pressed = False
                                ParticleEventFilter._last_drag_pos = None
                                return False

                        gp = event.globalPosition().toPoint()
                        last = ParticleEventFilter._last_drag_pos
                        now = time.time()

                        if last is None:
                            ParticleEventFilter._last_drag_pos = gp
                            ParticleEventFilter._last_drag_ts = now
                        else:
                            dx = gp.x() - last.x()
                            dy = gp.y() - last.y()
                            dist_sq = dx * dx + dy * dy
                            if (dist_sq >= ParticleEventFilter._DRAG_MIN_DIST_SQ
                                    and now - ParticleEventFilter._last_drag_ts >= ParticleEventFilter._DRAG_MIN_INTERVAL):
                                _spawn_drag_particles(gp.x(), gp.y())
                                ParticleEventFilter._last_drag_pos = gp
                                ParticleEventFilter._last_drag_ts = now
                    except Exception:
                        pass

            # ★ 新增：鼠标滚轮滚动 → 滚动粒子
            elif et == QEvent.Type.Wheel:
                try:
                    gp = event.globalPosition().toPoint()
                    _spawn_scroll_particles(gp.x(), gp.y())
                except Exception:
                    pass

        except Exception:
            pass

        # ★ 关键：始终返回 False，绝不拦截事件，保证拖动功能不受影响
        return False

# ======================== 进度条动画通用工具 ========================
def animate_progress_bar(progress_bar, target_value, duration=1800):
    """
    为进度条添加从 0 到目标值的缓动动画。
    是否动画由全局配置 enable_progress_animation 决定。
    :param progress_bar: 目标 QProgressBar
    :param target_value: 目标值（会转为 int）
    :param duration: 动画时长（毫秒），默认 1800
    """
    if progress_bar is None:
        return

    target_value = int(target_value)

    # 读取全局开关
    if not load_config().get("enable_progress_animation", True):
        progress_bar.setValue(target_value)
        return

    # 若有旧动画在跑，先停掉（避免重复启动）
    old = getattr(progress_bar, '_current_anim', None)
    if old is not None:
        try:
            old.stop()
        except Exception:
            pass

    anim = QPropertyAnimation(progress_bar, b"value")
    anim.setDuration(duration)
    anim.setStartValue(0)
    anim.setEndValue(target_value)
    anim.setEasingCurve(QEasingCurve.Type.OutQuad)

    progress_bar._current_anim = anim  # 引用挂在 bar 自身上，防止被 GC

    def on_finished(a=anim, bar=progress_bar):
        if getattr(bar, '_current_anim', None) is a:
            bar._current_anim = None

    anim.finished.connect(on_finished)
    anim.start()

# ======================== 开机动画（全新实现，复刻旧版效果） ========================
import sys
import math
import random
import time
from PyQt6.QtCore import Qt, QTimer, QPointF
from PyQt6.QtGui import QColor, QFont, QPainter, QBrush, QPen
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QProgressBar
)

# 版本号，可根据需要修改
SYSTEM_VERSION = "v6.4.0.110"


# ======================== 支持字体大小属性动画的 QLabel ========================
class AnimatedLabel(QLabel):
    """支持通过 pyqtProperty 动画化字体大小的 QLabel。"""

    # 微软雅黑优先；找不到时回退到 Segoe UI / Inter
    FONT_STACK = (
        "'Microsoft YaHei UI', 'Microsoft YaHei', "
        "'Segoe UI', 'Inter', sans-serif"
    )

    def __init__(self, text, color, font_size, parent=None):
        super().__init__(text, parent)
        self._color = color
        self._font_size = font_size
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(
            f"color: {self._color}; "
            f"font-family: {self.FONT_STACK}; "
            f"font-size: {self._font_size}px; "
            f"font-weight: 700; "
            f"background: transparent;"
        )
        self.adjustSize()

    def get_font_size(self):
        return self._font_size

    def set_font_size(self, size):
        """改变字体大小时保持控件中心位置不变。"""
        x, y = self.x(), self.y()
        w, h = self.width(), self.height()
        self._font_size = int(size)
        self._apply_style()
        w2, h2 = self.width(), self.height()
        self.move(x + (w - w2) // 2, y + (h - h2) // 2)

    font_size = pyqtProperty(int, get_font_size, set_font_size)


# ======================== 开机动画 ========================
class BootSplash(QWidget):
    """开机动画：
        1. 版权信息显示 2 秒（仅版权文字）
        2. 白色"P"从屏幕下方升起
        3. "P"缩小，同时"ick"出现，组成"Pick"
        4. "Pick"向左滑动，蓝色"OS"从右侧滑入，组成"PickOS"
        5. "OPenPick"从屏幕下方滑入（PickOS 下方）
        6. "Powered by"一行从中间向两边展开（OPenPick 上方）
    """

    FINAL_SIZE = 160        # PickOS 最终字体大小（px）
    BIG_SIZE = 330          # P 起始字体大小（px）
    COLOR_PICK = "#FFFFFF"  # Pick 白色
    COLOR_OS = "#2196F3"    

    def __init__(self, callback, enable_animation=True):
        super().__init__()
        self.callback = callback

        # 窗口设置
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint |
                            Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(1280, 720)
        self.center()

        self.bg_color = QColor(10, 14, 23)

        # 自定义属性所用的内部状态
        self._powered_width_value = 0

        # ---------- 阶段 1：版权页（仅版权文字） ----------
        self._build_credit_page()

        # ---------- 阶段 2：PickOS 组装标签 ----------
        self.p_label = AnimatedLabel("P", self.COLOR_PICK, self.BIG_SIZE, self)
        self.p_label.hide()

        self.ick_label = AnimatedLabel("ick", self.COLOR_PICK, self.FINAL_SIZE, self)
        self.ick_label.hide()

        self.os_label = AnimatedLabel("OS", self.COLOR_OS, self.FINAL_SIZE, self)
        self.os_label.hide()

        # P 初始定位到屏幕下方之外
        self.p_label.move(
            (self.width() - self.p_label.width()) // 2,
            self.height() + 30,
        )

        # ---------- 阶段 3：底部信息 ----------
        self._build_bottom_labels()

        self._animations = []
        self._ick_effect = None

        self.show()

        if not enable_animation:
            QTimer.singleShot(100, self.skip)
        else:
            QTimer.singleShot(2000, self._start_animation)

    # ------------------------------------------------------------------
    # 版权页（只显示版权文字）
    # ------------------------------------------------------------------
    def _build_credit_page(self):
        self.credit_page = QWidget(self)
        self.credit_page.setGeometry(0, 0, self.width(), self.height())
        self.credit_page.setStyleSheet("background: transparent;")

        layout = QVBoxLayout(self.credit_page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addStretch(1)

        credit_label = QLabel("Copyright © 2026 wuweijia-star  ·  GPLv3")
        credit_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_label.setStyleSheet(
            "color: white; "
            "font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', "
            "'Segoe UI', sans-serif; "
            "font-size: 44px; font-weight: bold; "
            "background: transparent;"
        )
        layout.addWidget(credit_label)

        layout.addStretch(1)

    # ------------------------------------------------------------------
    # 底部信息：OPenPick + Powered by（容器 + 子 QLabel）
    # ------------------------------------------------------------------
    def _build_bottom_labels(self):
        # ---- OPenPick（位置更低） ----
        self.openpick_label = QLabel("OPenPick", self)
        self.openpick_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.openpick_label.setStyleSheet(
            "color: white; "
            "font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', "
            "'Segoe UI', sans-serif; "
            "font-size: 28px; font-weight: bold; "
            "background: transparent;"
        )
        self.openpick_label.adjustSize()
        self._openpick_target_x = (self.width() - self.openpick_label.width()) // 2
        self._openpick_target_y = self.height() - 60   # 距屏幕底部约 120px
        self.openpick_label.move(self._openpick_target_x, self.height() + 40)
        self.openpick_label.hide()

        # ---- Powered by：容器裁剪 + 子 QLabel 居中 ----
        # 思路：容器宽度从 0 缓动到完整宽度，QLabel 始终在容器里居中。
        # 因为容器设置了 WA_ClipChildren，超出容器的部分会被裁剪，
        # 视觉上 Powered by 的文字就会从中间向两边"生长"出来。
        self.powered_container = QWidget(self)
        self.powered_container.setStyleSheet("background: transparent;")
        # Qt6 中 QWidget 的子控件默认会被父控件裁剪，无需 WA_ClipChildren

        self.powered_label = QLabel("—— Powered by ——", self.powered_container)
        self.powered_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.powered_label.setWordWrap(False)
        self.powered_label.setStyleSheet(
            "color: rgba(255,255,255,0.75); "
            "font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', "
            "'Segoe UI', sans-serif; "
            "font-size: 18px; "
            "background: transparent;"
        )
        self.powered_label.adjustSize()
        self._powered_full_width = self.powered_label.width()
        self._powered_height = self.powered_label.height()
        self.powered_label.setFixedSize(
            self._powered_full_width, self._powered_height
        )

        # Powered by 的目标 y：在 OPenPick 上方留 12px 间距
        self._powered_target_y = (
            self._openpick_target_y - self._powered_height - 12
        )

        # 容器初始宽度为 0，位置在屏幕中央
        self.powered_container.setFixedSize(0, self._powered_height)
        self.powered_container.move(self.width() // 2, self._powered_target_y)
        self.powered_label.move(-self._powered_full_width // 2, 0)
        self.powered_container.hide()

    # ------------------------------------------------------------------
    # 自定义属性：powered_width（供 QPropertyAnimation 使用）
    # ------------------------------------------------------------------
    def _get_powered_width(self):
        return self._powered_width_value

    def _set_powered_width(self, v):
        self._powered_width_value = int(v)
        w = self._powered_width_value
        # 容器宽度变化，且始终保持在屏幕水平居中
        self.powered_container.setFixedWidth(max(0, w))
        self.powered_container.move(
            (self.width() - w) // 2, self._powered_target_y
        )
        # 子 QLabel 相对容器居中
        self.powered_label.move((w - self._powered_full_width) // 2, 0)

    powered_width = pyqtProperty(int, fget=_get_powered_width, fset=_set_powered_width)

    # ------------------------------------------------------------------
    # 工具：让两个 QLabel 的文字基线对齐
    # ------------------------------------------------------------------
    @staticmethod
    def _align_baseline(label_a, label_b, base_y):
        fm_a = label_a.fontMetrics()
        fm_b = label_b.fontMetrics()
        baseline_a = base_y + (label_a.height() - fm_a.height()) // 2 + fm_a.ascent()
        b_top = baseline_a - fm_b.ascent() - (label_b.height() - fm_b.height()) // 2
        return b_top

    # ------------------------------------------------------------------
    # 第一步：隐藏版权页，P 从底部升起
    # ------------------------------------------------------------------
    def _start_animation(self):
        self.credit_page.hide()
        self.p_label.show()

        target_x = (self.width() - self.p_label.width()) // 2
        target_y = (self.height() - self.p_label.height()) // 2

        anim = QPropertyAnimation(self.p_label, b"pos", self)
        anim.setDuration(700)
        anim.setStartValue(QPoint(target_x, self.height() + 30))
        anim.setEndValue(QPoint(target_x, target_y))
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.finished.connect(self._step_shrink)
        anim.start()
        self._animations.append(anim)

    # ------------------------------------------------------------------
    # 第二步：P 缩小，同时 ick 出现
    # ------------------------------------------------------------------
    def _step_shrink(self):
        tmp_ratio = self.FINAL_SIZE / self.BIG_SIZE
        p_w_small = int(self.p_label.width() * tmp_ratio)
        p_h_small = int(self.p_label.height() * tmp_ratio)

        ick_w = self.ick_label.width()
        ick_h = self.ick_label.height()

        pick_w = p_w_small + ick_w
        pick_start_x = (self.width() - pick_w) // 2
        final_y = (self.height() - p_h_small) // 2

        anim_shrink = QPropertyAnimation(self.p_label, b"font_size", self)
        anim_shrink.setDuration(400)
        anim_shrink.setStartValue(self.BIG_SIZE)
        anim_shrink.setEndValue(self.FINAL_SIZE)
        anim_shrink.setEasingCurve(QEasingCurve.Type.OutCubic)

        anim_move = QPropertyAnimation(self.p_label, b"pos", self)
        anim_move.setDuration(400)
        anim_move.setStartValue(self.p_label.pos())
        anim_move.setEndValue(QPoint(pick_start_x, final_y))
        anim_move.setEasingCurve(QEasingCurve.Type.OutCubic)

        ick_x = pick_start_x + p_w_small
        ick_y = final_y + (p_h_small - ick_h) // 2
        self.ick_label.move(ick_x, ick_y)
        self.ick_label.show()

        effect = QGraphicsOpacityEffect(self.ick_label)
        effect.setOpacity(0.0)
        self.ick_label.setGraphicsEffect(effect)
        self._ick_effect = effect

        anim_fade = QPropertyAnimation(effect, b"opacity", self)
        anim_fade.setDuration(400)
        anim_fade.setStartValue(0.0)
        anim_fade.setEndValue(1.0)

        def _after_shrink():
            p_w = self.p_label.width()
            ick_w_now = self.ick_label.width()
            pick_w_now = p_w + ick_w_now
            new_start_x = (self.width() - pick_w_now) // 2
            new_y = (self.height() - self.p_label.height()) // 2
            self.p_label.move(new_start_x, new_y)

            ick_top = self._align_baseline(self.p_label, self.ick_label, new_y)
            self.ick_label.move(new_start_x + p_w, ick_top)

            QTimer.singleShot(250, self._step_add_os)

        anim_shrink.finished.connect(_after_shrink)

        anim_shrink.start()
        anim_move.start()
        anim_fade.start()

        self._animations.extend([anim_shrink, anim_move, anim_fade])

    # ------------------------------------------------------------------
    # 第三步：Pick 左移，OS 从右侧滑入
    # ------------------------------------------------------------------
    def _step_add_os(self):
        if self._ick_effect is not None:
            self.ick_label.setGraphicsEffect(None)
            self._ick_effect = None

        p_w = self.p_label.width()
        ick_w = self.ick_label.width()
        os_w = self.os_label.width()

        pick_w = p_w + ick_w
        total_w = pick_w + os_w

        total_start_x = (self.width() - total_w) // 2
        final_y = self.p_label.y()

        pick_new_x = total_start_x
        os_new_x = total_start_x + pick_w

        os_new_y = self._align_baseline(self.p_label, self.os_label, final_y)

        os_start_x = self.width() + 60
        self.os_label.move(os_start_x, os_new_y)
        self.os_label.show()

        anim_p_slide = QPropertyAnimation(self.p_label, b"pos", self)
        anim_p_slide.setDuration(1000)
        anim_p_slide.setStartValue(self.p_label.pos())
        anim_p_slide.setEndValue(QPoint(pick_new_x, final_y))
        anim_p_slide.setEasingCurve(QEasingCurve.Type.OutCubic)

        ick_new_x = pick_new_x + p_w

        anim_ick_slide = QPropertyAnimation(self.ick_label, b"pos", self)
        anim_ick_slide.setDuration(600)
        anim_ick_slide.setStartValue(self.ick_label.pos())
        anim_ick_slide.setEndValue(QPoint(ick_new_x, self.ick_label.y()))
        anim_ick_slide.setEasingCurve(QEasingCurve.Type.OutCubic)

        anim_os_slide = QPropertyAnimation(self.os_label, b"pos", self)
        anim_os_slide.setDuration(800)
        anim_os_slide.setStartValue(QPoint(os_start_x, os_new_y))
        anim_os_slide.setEndValue(QPoint(os_new_x, os_new_y))
        anim_os_slide.setEasingCurve(QEasingCurve.Type.OutCubic)

        # OS 滑入完成后，开始 OPenPick 上滑
        anim_os_slide.finished.connect(self._step_show_openpick)

        anim_p_slide.start()
        anim_ick_slide.start()
        anim_os_slide.start()

        self._animations.extend([anim_p_slide, anim_ick_slide, anim_os_slide])

    # ------------------------------------------------------------------
    # 第四步：OPenPick 从下方滑入
    # ------------------------------------------------------------------
    def _step_show_openpick(self):
        self.openpick_label.show()

        anim = QPropertyAnimation(self.openpick_label, b"pos", self)
        anim.setDuration(500)
        anim.setStartValue(QPoint(self._openpick_target_x, self.height() + 40))
        anim.setEndValue(QPoint(self._openpick_target_x, self._openpick_target_y))
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.finished.connect(self._step_expand_powered)
        anim.start()
        self._animations.append(anim)

    # ------------------------------------------------------------------
    # 第五步：Powered by 从中间向两边展开
    # ------------------------------------------------------------------
    def _step_expand_powered(self):
        # 重置到初始状态（宽度 0）
        self._powered_width_value = 0
        self.powered_container.setFixedWidth(0)
        self.powered_container.move(self.width() // 2, self._powered_target_y)
        self.powered_label.move(-self._powered_full_width // 2, 0)
        self.powered_container.show()

        # 动画 powered_width 属性：0 → 完整宽度
        anim = QPropertyAnimation(self, b"powered_width", self)
        anim.setDuration(500)
        anim.setStartValue(0)
        anim.setEndValue(self._powered_full_width)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()
        self._animations.append(anim)

        # 停留后结束
        QTimer.singleShot(1000, self.skip)

    # ------------------------------------------------------------------
    def center(self):
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), self.bg_color)

    def skip(self):
        """开机动画结束：立即关闭，并启动主窗口（只触发一次）"""
        if getattr(self, "_skip_done", False):
            return
        self._skip_done = True

        app = QApplication.instance()
        if app is not None:
            app.setQuitOnLastWindowClosed(False)

        self.close()
        if self.callback:
            cb = self.callback
            self.callback = None   # 防止重复触发
            QTimer.singleShot(0, cb)

    def _fade_out_and_launch(self):
        """立即关闭开机动画并启动主窗口（无渐暗）"""
        app = QApplication.instance()
        if app is not None:
            app.setQuitOnLastWindowClosed(False)
        self.close()
        if self.callback:
            QTimer.singleShot(0, self.callback)

# ======================== 成绩总览界面（支持多考试切换、持久化、单科王、总览对比，统计两行） ========================
class ScoreInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("scoreInterface")
        self.current_data = None
        self.previous_data = None
        self.current_exam = "本次"
        self.build_ui()
        self.load_saved_data()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        top = QHBoxLayout()
        title = TitleLabel("📊 成绩总览", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        top.addWidget(title)
        top.addStretch()

        top.addWidget(QLabel("考试："))
        self.exam_combo = QComboBox()
        self.exam_combo.addItems(["本次考试", "上次考试", "总览"])
        self.exam_combo.currentTextChanged.connect(self.on_exam_changed)
        top.addWidget(self.exam_combo)

        import_btn = QPushButton("📂 导入成绩")
        import_btn.clicked.connect(self.import_score_file)
        import_btn.setFixedWidth(180)
        top.addWidget(import_btn)

        bonus_btn = QPushButton("🎯 考试加分")
        bonus_btn.clicked.connect(self.open_bonus_dialog)
        bonus_btn.setFixedWidth(180)
        top.addWidget(bonus_btn)

        clear_btn = QPushButton("🗑️ 清空当前")
        clear_btn.clicked.connect(self.clear_current)
        clear_btn.setFixedWidth(180)
        top.addWidget(clear_btn)

        layout.addLayout(top)

        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # 统计信息标签（支持多行，自动换行）
        self.stats_label = QLabel("暂无数据")
        self.stats_label.setStyleSheet("color: #A0A0B0; font-size: 15px;")
        self.stats_label.setWordWrap(True)
        self.stats_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.stats_label)

        # 单科王标签
        champion_layout = QHBoxLayout()
        champion_layout.setContentsMargins(0, 0, 0, 0)
        self.subject_champions_label = QLabel("")
        self.subject_champions_label.setStyleSheet("color: #A0A0B0; font-size: 14px;")
        self.subject_champions_label.setWordWrap(True)
        self.subject_champions_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        champion_layout.addWidget(self.subject_champions_label)
        champion_layout.setStretch(0, 1)
        layout.addLayout(champion_layout)

        back_btn = QPushButton("🔙 返回主菜单")
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)

    def load_saved_data(self):
        if os.path.exists(CURRENT_SCORES_FILE):
            try:
                df = pd.read_csv(CURRENT_SCORES_FILE, encoding="utf-8-sig")
                if not df.empty:
                    self.current_data = df
                else:
                    self.current_data = None
            except:
                self.current_data = None
        else:
            self.current_data = None

        if os.path.exists(PREVIOUS_SCORES_FILE):
            try:
                df = pd.read_csv(PREVIOUS_SCORES_FILE, encoding="utf-8-sig")
                if not df.empty:
                    self.previous_data = df
                else:
                    self.previous_data = None
            except:
                self.previous_data = None
        else:
            self.previous_data = None

        if self.current_data is not None:
            self.exam_combo.setCurrentText("本次考试")
        elif self.previous_data is not None:
            self.exam_combo.setCurrentText("上次考试")
        else:
            self.exam_combo.setCurrentText("本次考试")

        self.on_exam_changed(self.exam_combo.currentText())

    def on_exam_changed(self, text):
        if text == "本次考试":
            self.current_exam = "本次"
            if self.current_data is not None:
                self.display_data(self.current_data)
            else:
                self.clear_table("暂无本次考试数据，请导入")
        elif text == "上次考试":
            self.current_exam = "上次"
            if self.previous_data is not None:
                self.display_data(self.previous_data)
            else:
                self.clear_table("暂无上次考试数据")
        else:  # 总览
            self.current_exam = "总览"
            self.display_overview()

    def clear_table(self, message="暂无数据"):
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.stats_label.setText(message)
        self.subject_champions_label.setText("")

    def import_score_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择成绩Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            show_toast(self, f"读取文件失败：{e}", "error")
            return

        required = ["姓名", "学号"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            show_toast(self, f"文件缺少必要列：{', '.join(missing)}", "error")
            return

        priority = ["学号", "姓名", "语文", "数学", "英语", "物理", "化学", "生物", "总分", "原分校名", "化学赋分", "生物赋分", "赋分总分", "赋分校名", "赋分班名"]
        all_cols = df.columns.tolist()
        available_cols = [col for col in priority if col in all_cols]
        other_cols = [col for col in all_cols if col not in available_cols and col != "班次"]
        available_cols.extend(other_cols)

        df_clean = df[available_cols].copy()
        df_clean["学号"] = df_clean["学号"].astype(str)

        if self.current_data is not None:
            try:
                self.current_data.to_csv(PREVIOUS_SCORES_FILE, index=False, encoding="utf-8-sig")
                self.previous_data = self.current_data.copy()
            except Exception as e:
                show_toast(self, f"备份上次考试失败：{e}", "error")
                return

        try:
            df_clean.to_csv(CURRENT_SCORES_FILE, index=False, encoding="utf-8-sig")
            self.current_data = df_clean
            # ----- 新增：增加导入次数成就计数 -----
            cfg = load_config()
            cfg["import_count"] = cfg.get("import_count", 0) + 1
            save_config(cfg)
            check_and_unlock_achievements()  # 立即检查成就
            # ----- 新增结束 -----
            show_toast(self, f"成功导入并保存 {len(df_clean)} 条记录作为本次考试", "success")
        except Exception as e:
            show_toast(self, f"保存本次考试失败：{e}", "error")
            return

        self.exam_combo.setCurrentText("本次考试")

    def display_data(self, df):
        if df is None or df.empty:
            self.clear_table("无数据")
            return

        columns = df.columns.tolist()
        self.table.setUpdatesEnabled(False)          # ← 新增
        try:
            self.table.setColumnCount(len(columns))
            self.table.setHorizontalHeaderLabels(columns)

            self.table.setRowCount(len(df))
            for i, row in df.iterrows():
                for j, col in enumerate(columns):
                    val = row[col]
                    if isinstance(val, float):
                        if val.is_integer():
                            val = int(val)
                        else:
                            val = round(val, 2)
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(i, j, item)

            for i in range(len(df)):
                self.table.setRowHeight(i, 30)
        finally:
            self.table.setUpdatesEnabled(True)       # ← 新增

        # 总分统计
        score_col = None
        if "赋分总分" in df.columns:
            score_col = "赋分总分"
        elif "总分" in df.columns:
            score_col = "总分"

        if score_col:
            scores = pd.to_numeric(df[score_col], errors='coerce')
            if not scores.empty and scores.notna().any():
                valid_scores = scores.dropna()
                if not valid_scores.empty:
                    max_score = valid_scores.max()
                    min_score = valid_scores.min()
                    avg_score = valid_scores.mean()
                    max_row = df[df[score_col] == max_score].iloc[0]
                    min_row = df[df[score_col] == min_score].iloc[0]
                    stats_text = (f"最高：{max_row['姓名']}（{max_row['学号']}）{max_score}分 | "
                                  f"最低：{min_row['姓名']}（{min_row['学号']}）{min_score}分 | "
                                  f"平均：{avg_score:.2f}分")
                    self.stats_label.setText(stats_text)
                else:
                    self.stats_label.setText("无有效成绩")
            else:
                self.stats_label.setText("无有效成绩")
        else:
            self.stats_label.setText(f"共 {len(df)} 名学生，无有效成绩列")

        # ---- 单科王（并列显示） ----
        subject_map = {
            "语文": "语文",
            "数学": "数学",
            "英语": "英语",
            "物理": "物理",
            "化学": "化学赋分",
            "生物": "生物赋分"
        }
        champion_texts = []
        for display_name, col_name in subject_map.items():
            if col_name in df.columns:
                scores = pd.to_numeric(df[col_name], errors='coerce')
                valid_scores = scores.dropna()
                if not valid_scores.empty:
                    max_score = valid_scores.max()
                    max_rows = df[df[col_name] == max_score]
                    students = []
                    for _, row in max_rows.iterrows():
                        stu_id = row.get('学号', '')
                        stu_name = row.get('姓名', '')
                        if stu_name and stu_id:
                            students.append(f"{stu_name}({stu_id})")
                    if students:
                        champion_texts.append(f"{display_name}:{', '.join(students)}{max_score}分")
        if champion_texts:
            self.subject_champions_label.setText(" | ".join(champion_texts))
        else:
            self.subject_champions_label.setText("")

    # ========== 总览对比视图（进步/退步分两行） ==========
    def display_overview(self):
        """显示本次与上次考试的对比数据，支持并列显示，可选排名依据"""
        if self.current_data is None or self.previous_data is None:
            self.clear_table("缺少本次或上次考试数据，无法对比")
            return

        # 合并数据
        current = self.current_data.copy()
        previous = self.previous_data.copy()
        current["学号"] = current["学号"].astype(str)
        previous["学号"] = previous["学号"].astype(str)

        merged = pd.merge(current, previous, on="学号", how="outer", suffixes=("_本次", "_上次"))
        merged["姓名"] = merged["姓名_本次"].fillna(merged["姓名_上次"])

        # 确保列存在
        for col in ["赋分总分_本次", "赋分校名_本次", "赋分班名_本次"]:
            if col not in merged.columns:
                merged[col] = ""
        for col in ["赋分总分_上次", "赋分校名_上次", "赋分班名_上次"]:
            if col not in merged.columns:
                merged[col] = ""

        # ---- 直接计算数值差值（不使用字符串解析） ----
        # 总分差值（浮点数）
        merged["_diff_total"] = (
            pd.to_numeric(merged["赋分总分_本次"], errors='coerce') -
            pd.to_numeric(merged["赋分总分_上次"], errors='coerce')
        )
        merged["_diff_school"] = -(
            pd.to_numeric(merged["赋分校名_本次"], errors='coerce') -
            pd.to_numeric(merged["赋分校名_上次"], errors='coerce')
        )
        merged["_diff_class"] = -(
            pd.to_numeric(merged["赋分班名_本次"], errors='coerce') -
            pd.to_numeric(merged["赋分班名_上次"], errors='coerce')
        )

        # 读取配置，选择排序依据
        cfg = load_config()
        rank_by = cfg.get("overview_rank_by", "赋分班名")   # 默认赋分班名
        if rank_by == "赋分总分":
            diff_col = "_diff_total"
        elif rank_by == "赋分校名":
            diff_col = "_diff_school"
        else:  # 赋分班名
            diff_col = "_diff_class"

        # ---- 学号数字排序 ----
        merged = merged.dropna(subset=["学号"])
        merged["学号_int"] = merged["学号"].astype(int)
        merged = merged.sort_values("学号_int").reset_index(drop=True)

        # ---- 填充表格（显示文本差值） ----
        # 为了展示，我们仍用字符串格式化显示差值（保留一位小数）
        def format_diff(val):
            if pd.isna(val):
                return ""
            if val > 0:
                return f"+{val:.1f}" if isinstance(val, float) and val % 1 != 0 else f"+{int(val)}"
            else:
                return f"{val:.1f}" if isinstance(val, float) and val % 1 != 0 else f"{int(val)}"

        merged["较上次赋分总分"] = merged["_diff_total"].apply(format_diff)
        merged["较上次赋分校名"] = merged["_diff_school"].apply(lambda x: format_diff(x) if pd.notna(x) else "")
        merged["较上次赋分班名"] = merged["_diff_class"].apply(lambda x: format_diff(x) if pd.notna(x) else "")

        display_cols = [
            "学号", "姓名",
            "赋分总分_上次", "赋分校名_上次", "赋分班名_上次",
            "赋分总分_本次", "赋分校名_本次", "赋分班名_本次",
            "较上次赋分总分", "较上次赋分校名", "较上次赋分班名"
        ]
        header_names = [
            "学号", "姓名",
            "上次赋分总分", "上次赋分校名", "上次赋分班名",
            "本次赋分总分", "本次赋分校名", "本次赋分班名",
            "较上次赋分总分", "较上次校名", "较上次班名"
        ]
        df_display = merged[display_cols].copy()
        df_display.columns = header_names

        self.table.setUpdatesEnabled(False)          # ← 新增
        try:
            self.table.setColumnCount(len(header_names))
            self.table.setHorizontalHeaderLabels(header_names)
            self.table.setRowCount(len(df_display))
            for i, row in df_display.iterrows():
                for j, col in enumerate(header_names):
                    val = row[col]
                    if isinstance(val, float):
                        if val.is_integer():
                            val = int(val)
                        else:
                            val = round(val, 2)
                    item = QTableWidgetItem(str(val) if pd.notna(val) else "")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(i, j, item)
            for i in range(len(df_display)):
                self.table.setRowHeight(i, 30)
        finally:
            self.table.setUpdatesEnabled(True)       # ← 新增

        # ---- 进步/退步排行榜（含并列处理，排除差值为0） ----
        top_n = cfg.get("overview_top_n", 5)
        valid = merged[merged[diff_col].notna()].copy()
        if valid.empty:
            self.stats_label.setText("无有效差值数据")
            self.subject_champions_label.setText("")
            return

        # 进步榜（差值 > 0）
        prog_valid = valid[valid[diff_col] > 0].copy()
        if not prog_valid.empty:
            prog_sorted = prog_valid.sort_values(diff_col, ascending=False).reset_index(drop=True)
            if len(prog_sorted) > top_n:
                threshold_prog = prog_sorted.iloc[top_n-1][diff_col]
                prog_df = prog_sorted[prog_sorted[diff_col] >= threshold_prog]
            else:
                prog_df = prog_sorted
        else:
            prog_df = pd.DataFrame()

        # 退步榜（差值 < 0）
        reg_valid = valid[valid[diff_col] < 0].copy()
        if not reg_valid.empty:
            reg_sorted = reg_valid.sort_values(diff_col, ascending=True).reset_index(drop=True)
            if len(reg_sorted) > top_n:
                threshold_reg = reg_sorted.iloc[top_n-1][diff_col]
                reg_df = reg_sorted[reg_sorted[diff_col] <= threshold_reg]
            else:
                reg_df = reg_sorted
        else:
            reg_df = pd.DataFrame()

        stats_lines = []
        if not prog_df.empty:
            prog_list = []
            for _, row in prog_df.iterrows():
                name = row["姓名"]
                stu_id = row["学号"]
                diff_val = row[diff_col]
                # 显示差值时保留整数或一位小数
                if isinstance(diff_val, float) and diff_val.is_integer():
                    diff_str = f"{int(diff_val):+d}"
                else:
                    diff_str = f"{diff_val:+.1f}"
                prog_list.append(f"{name}({stu_id}){diff_str}")
            stats_lines.append("🏆 进步前{}：".format(top_n) + " ".join(prog_list))
        else:
            stats_lines.append("🏆 进步前{}：无".format(top_n))

        if not reg_df.empty:
            reg_list = []
            for _, row in reg_df.iterrows():
                name = row["姓名"]
                stu_id = row["学号"]
                diff_val = row[diff_col]
                if isinstance(diff_val, float) and diff_val.is_integer():
                    diff_str = f"{int(diff_val):+d}"
                else:
                    diff_str = f"{diff_val:+.1f}"
                reg_list.append(f"{name}({stu_id}){diff_str}")
            stats_lines.append("📉 退步前{}：".format(top_n) + " ".join(reg_list))
        else:
            stats_lines.append("📉 退步前{}：无".format(top_n))

        self.stats_label.setText("\n".join(stats_lines))
        self.subject_champions_label.setText("")

    def clear_current(self):
        if self.current_exam == "本次":
            if not os.path.exists(CURRENT_SCORES_FILE):
                show_toast(self, "本次考试数据已为空", "info")
                return
            # ---- 自定义确认对话框 ----
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("确认清空")
            msg_box.setText("确定要清空本次考试的所有数据吗？\n此操作不可恢复！")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.No)
            msg_box.setStyleSheet("""
                QMessageBox {
                    background: #16161E;
                }
                QLabel {
                    color: #E0E0E0;
                    font-size: 15px;
                }
                QPushButton {
                    border-radius: 10px;
                    padding: 8px 24px;
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                                stop:0 #64B5F6, stop:1 #42A5F5);
                    color: white;
                    font-weight: 600;
                    border: 1px solid rgba(100,181,246,0.3);
                    min-width: 80px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                                stop:0 #42A5F5, stop:1 #1E88E5);
                    border-color: #64B5F6;
                }
                QPushButton:pressed {
                    background: #1E88E5;
                }
            """)
            reply = msg_box.exec()
            if reply != QMessageBox.StandardButton.Yes:
                return
            os.remove(CURRENT_SCORES_FILE)
            self.current_data = None
            self.clear_table("本次考试已清空")
            show_toast(self, "本次考试数据已清空", "info")

        elif self.current_exam == "上次":
            if not os.path.exists(PREVIOUS_SCORES_FILE):
                show_toast(self, "上次考试数据已为空", "info")
                return
            # ---- 同样自定义 ----
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("确认清空")
            msg_box.setText("确定要清空上次考试的所有数据吗？\n此操作不可恢复！")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setDefaultButton(QMessageBox.StandardButton.No)
            msg_box.setStyleSheet("""
                QMessageBox {
                    background: #16161E;
                }
                QLabel {
                    color: #E0E0E0;
                    font-size: 15px;
                }
                QPushButton {
                    border-radius: 10px;
                    padding: 8px 24px;
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                                stop:0 #64B5F6, stop:1 #42A5F5);
                    color: white;
                    font-weight: 600;
                    border: 1px solid rgba(100,181,246,0.3);
                    min-width: 80px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                                stop:0 #42A5F5, stop:1 #1E88E5);
                    border-color: #64B5F6;
                }
                QPushButton:pressed {
                    background: #1E88E5;
                }
            """)
            reply = msg_box.exec()
            if reply != QMessageBox.StandardButton.Yes:
                return
            os.remove(PREVIOUS_SCORES_FILE)
            self.previous_data = None
            self.clear_table("上次考试已清空")
            show_toast(self, "上次考试数据已清空", "info")

        else:
            show_toast(self, "总览模式下不能清空，请切换到具体考试", "warning")

    def open_bonus_dialog(self):
        if self.current_data is None or self.current_data.empty:
            show_toast(self, "请先导入本次考试成绩", "warning")
            return
        # 检查是否有 stu_df 和 id_to_name
        if self.app.stu_df is None:
            show_toast(self, "请先加载学生名单", "warning")
            return
        dialog = ExamBonusDialog(self, self.current_data, self.previous_data,
                                self.app.stu_df, self.app.id_to_name)
        dialog.exec()

# ======================== 考试加分弹窗（优化样式 + 提交确认） ========================
class ExamBonusDialog(QDialog):
    def __init__(self, parent, current_data, previous_data, stu_df, id_to_name):
        super().__init__(parent)
        _c = get_theme_colors()
        self.parent = parent
        self.current_data = current_data
        self.previous_data = previous_data
        self.stu_df = stu_df
        self.id_to_name = id_to_name
        self.setWindowTitle("🎯 考试加分规则")
        self.resize(750, 600)
        self.setStyleSheet(DIALOG_STYLE + f"""
            QListWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 8px;
                padding: 4px;
                outline: none;
                color: {_c['fg']};
            }}
            QListWidget::item {{
                padding: 6px 10px;
                border-radius: 4px;
            }}
            QListWidget::item:selected {{
                background: {_c['bg_card']};
                color: {_c['fg_bright']};
            }}
            QListWidget::item:hover {{
                background: {_c['bg_card']};
            }}
            QSpinBox {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                color: {_c['fg']};
                padding: 4px 8px;
            }}
            QSpinBox:focus {{
                border: 2px solid #64B5F6;
            }}
            QComboBox {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                padding: 6px 10px;
                color: {_c['fg']};
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {_c['fg_dim']};
                margin-right: 6px;
            }}
            QComboBox QAbstractItemView {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 14px;
            }}
            QPushButton {{
                border-radius: 10px;
                padding: 8px 20px;
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #64B5F6,stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(100,181,246,0.3);
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #42A5F5,stop:1 #1E88E5);
            }}
            QPushButton:pressed {{
                background: #1E88E5;
            }}
            QPushButton:disabled {{
                background: #555555;
                color: #888888;
            }}
        """)

        self.rules = load_config().get("exam_bonus_rules", [])
        # 确定可用科目列（与成绩界面保持一致）
        self.subject_columns = [col for col in ["语文", "数学", "英语", "物理", "化学赋分", "生物赋分"] 
                                if col in self.current_data.columns]
        if not self.subject_columns:
            # 若上述列都不存在，则尝试常见列名
            self.subject_columns = [col for col in ["语文", "数学", "英语", "物理", "化学", "生物"] 
                                    if col in self.current_data.columns]
            if not self.subject_columns:
                self.subject_columns = ["语文", "数学", "英语", "物理", "化学", "生物"]  # 占位
        self.selected_rule_index = -1

        self.build_ui()
        self.refresh_list()

    def build_ui(self):
        _c = get_theme_colors()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        info = QLabel("定义加分规则：根据本次考试的成绩排名、校排名、班排名、进步/退步名次或单科王对学生加分。")
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 14px;")
        layout.addWidget(info)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        add_btn = QPushButton("➕ 新增")
        add_btn.clicked.connect(self.add_rule)
        edit_btn = QPushButton("✏️ 编辑")
        edit_btn.clicked.connect(self.edit_rule)
        delete_btn = QPushButton("🗑️ 删除")
        delete_btn.clicked.connect(self.delete_rule)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        action_layout = QHBoxLayout()
        action_layout.setSpacing(12)
        submit_btn = QPushButton("📤 提交加分")
        submit_btn.clicked.connect(self.submit_bonus)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        action_layout.addWidget(submit_btn)
        action_layout.addStretch()
        action_layout.addWidget(close_btn)
        layout.addLayout(action_layout)

    def refresh_list(self):
        self.list_widget.clear()
        for i, rule in enumerate(self.rules):
            desc = self.rule_description(rule)
            item = QListWidgetItem(desc)
            item.setData(Qt.ItemDataRole.UserRole, i)
            self.list_widget.addItem(item)
        if 0 <= self.selected_rule_index < len(self.rules):
            self.list_widget.setCurrentRow(self.selected_rule_index)

    def rule_description(self, rule):
        type_map = {
            "赋分总分": "总分",
            "赋分校名": "校名",
            "赋分班名": "班名",
            "进步排名": "进步名次",
            "退步排名": "退步名次",
            "单科王": "单科王"
        }
        type_str = type_map.get(rule["type"], rule["type"])
        if rule["type"] == "单科王":
            subject = rule.get("subject", "所有科目")
            if subject == "all":
                subject = "所有科目"
            start = rule["range_start"]
            end = rule["range_end"]
            score = rule["score"]
            sign = "+" if score >= 0 else ""
            return f"单科王({subject}) {start}~{end} 名 → {sign}{score}分"
        else:
            start = rule["range_start"]
            end = rule["range_end"]
            score = rule["score"]
            sign = "+" if score >= 0 else ""
            return f"{type_str} {start}~{end} 名（分） → {sign}{score}分"

    def on_item_clicked(self, item):
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is not None:
            self.selected_rule_index = idx

    def get_selected_index(self):
        cur = self.list_widget.currentRow()
        if cur < 0:
            QMessageBox.warning(self, "提示", "请先选择一条规则")
            return None
        return cur

    def add_rule(self):
        self.edit_rule_impl(-1)

    def edit_rule(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        self.edit_rule_impl(idx)

    def delete_rule(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if QMessageBox.question(self, "确认删除", f"确定删除规则“{self.rule_description(self.rules[idx])}”吗？",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            del self.rules[idx]
            self.selected_rule_index = -1
            self.save_rules()
            self.refresh_list()

    def edit_rule_impl(self, idx):
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑规则" if idx >= 0 else "新增规则")
        dialog.resize(450, 350)
        dialog.setStyleSheet(DIALOG_STYLE + f"""
            QSpinBox {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                color: {_c['fg']};
                padding: 4px 8px;
            }}
            QSpinBox:focus {{
                border: 2px solid #64B5F6;
            }}
            QComboBox {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                padding: 6px 10px;
                color: {_c['fg']};
            }}
            QComboBox::drop-down {{
                border: none;
                width: 24px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {_c['fg_dim']};
                margin-right: 6px;
            }}
            QComboBox QAbstractItemView {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 14px;
            }}
            QDialogButtonBox QPushButton {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #64B5F6,stop:1 #42A5F5);
                border-radius: 8px;
                padding: 6px 18px;
                color: white;
                font-weight: 600;
                border: none;
            }}
            QDialogButtonBox QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #42A5F5,stop:1 #1E88E5);
            }}
            QDialogButtonBox QPushButton:pressed {{
                background: #1E88E5;
            }}
        """)

        layout = QFormLayout(dialog)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(14)
        layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        # 类型
        type_combo = QComboBox()
        type_combo.addItems(["赋分总分", "赋分校名", "赋分班名", "进步排名", "退步排名", "单科王"])
        type_combo.setCurrentText("赋分校名")
        layout.addRow("类型：", type_combo)

        # 科目（仅当类型为单科王时显示）
        subject_container = QWidget()
        subject_layout = QHBoxLayout(subject_container)
        subject_layout.setContentsMargins(0, 0, 0, 0)
        subject_combo = QComboBox()
        # 从本次数据中获取科目列（语文、数学、英语、物理、化学、生物）
        subject_cols = ["语文", "数学", "英语", "物理", "化学", "生物"]
        available_subjects = [col for col in subject_cols if col in self.current_data.columns]
        if not available_subjects:
            available_subjects = ["语文", "数学", "英语", "物理", "化学", "生物"]  # 默认
        subject_options = ["所有科目"] + self.subject_columns
        subject_combo.addItems(subject_options)
        subject_layout.addWidget(subject_combo)
        subject_container.setVisible(False)  # 初始隐藏
        layout.addRow("科目：", subject_container)

        # 范围
        range_start = QSpinBox()
        range_start.setRange(-9999, 9999)
        range_start.setValue(1)
        range_end = QSpinBox()
        range_end.setRange(-9999, 9999)
        range_end.setValue(10)

        range_layout = QHBoxLayout()
        range_layout.addWidget(QLabel("从"))
        range_layout.addWidget(range_start)
        range_layout.addWidget(QLabel("到"))
        range_layout.addWidget(range_end)
        layout.addRow("范围（名次/分数）：", range_layout)

        # 分数
        score_spin = QSpinBox()
        score_spin.setRange(-999, 999)
        score_spin.setValue(5)
        layout.addRow("加减分数：", score_spin)

        # 填充数据
        if idx >= 0:
            rule = self.rules[idx]
            type_combo.setCurrentText(rule["type"])
            if rule["type"] == "单科王":
                subj = rule.get("subject", "所有科目")
                if subj == "all":
                    subj = "所有科目"
                subject_combo.setCurrentText(subj)
                subject_container.setVisible(True)
                range_start.setValue(1)
                range_end.setValue(1)
                range_start.setEnabled(False)
                range_end.setEnabled(False)
            else:
                range_start.setValue(rule["range_start"])
                range_end.setValue(rule["range_end"])
                range_start.setEnabled(True)
                range_end.setEnabled(True)
            score_spin.setValue(rule["score"])

        # 类型切换逻辑
        def on_type_changed(text):
            if text == "单科王":
                subject_container.setVisible(True)
                range_start.setValue(1)
                range_end.setValue(1)
                range_start.setEnabled(False)
                range_end.setEnabled(False)
            else:
                subject_container.setVisible(False)
                range_start.setEnabled(True)
                range_end.setEnabled(True)
                # 恢复默认值（如果之前是单科王，可能值被改了，但这里不自动恢复）
        type_combo.currentTextChanged.connect(on_type_changed)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addRow(btn_box)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            rule_type = type_combo.currentText()
            new_rule = {
                "type": rule_type,
                "range_start": range_start.value(),
                "range_end": range_end.value(),
                "score": score_spin.value()
            }
            if rule_type == "单科王":
                subj = subject_combo.currentText()
                new_rule["subject"] = "all" if subj == "所有科目" else subj
                # 固定范围1~1
                new_rule["range_start"] = 1
                new_rule["range_end"] = 1
            else:
                if new_rule["range_start"] > new_rule["range_end"]:
                    new_rule["range_start"], new_rule["range_end"] = new_rule["range_end"], new_rule["range_start"]

            if idx >= 0:
                self.rules[idx] = new_rule
            else:
                self.rules.append(new_rule)
            self.save_rules()
            self.refresh_list()

    def save_rules(self):
        cfg = load_config()
        cfg["exam_bonus_rules"] = self.rules
        save_config(cfg)

    def submit_bonus(self):
        # （业务计算逻辑保持不变，请直接保留你原来文件里的这一整段逻辑）
        if self.current_data is None or self.current_data.empty:
            QMessageBox.warning(self, "错误", "请先导入本次考试成绩")
            return
        if not self.rules:
            QMessageBox.warning(self, "提示", "没有定义任何规则，请先新增")
            return
        # 准备数据
        current = self.current_data.copy()
        current["学号"] = current["学号"].astype(str)
        previous = None
        if self.previous_data is not None and not self.previous_data.empty:
            previous = self.previous_data.copy()
            previous["学号"] = previous["学号"].astype(str)
        # 合并数据
        if previous is not None:
            merged = pd.merge(current, previous, on="学号", how="outer", suffixes=("_本次", "_上次"))
            merged["姓名"] = merged["姓名_本次"].fillna(merged["姓名_上次"])
            for col in ["赋分总分", "赋分校名", "赋分班名"]:
                if col + "_本次" in merged.columns:
                    merged[col] = merged[col + "_本次"]
                if col + "_上次" in merged.columns:
                    merged[col] = merged[col].fillna(merged[col + "_上次"])
        else:
            merged = current.copy()
            merged["姓名"] = merged["姓名"]
            for col in ["赋分总分", "赋分校名", "赋分班名"]:
                if col not in merged.columns:
                    merged[col] = None
        # 检查进步/退步规则是否需要上次数据
        has_progress_rule = any(r["type"] in ["进步排名", "退步排名"] for r in self.rules)
        if has_progress_rule and previous is None:
            if QMessageBox.question(self, "提示", "缺少上次考试数据，进步/退步规则将无效。是否继续？",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No:
                return
        # 读取配置，获取排名依据（与总览一致）
        cfg = load_config()
        rank_by = cfg.get("overview_rank_by", "赋分班名")
        if rank_by == "赋分总分":
            diff_col = "_diff_total"
        elif rank_by == "赋分校名":
            diff_col = "_diff_school"
        else:
            diff_col = "_diff_class"
        # 预计算差值（用于进步/退步规则）
        if has_progress_rule and previous is not None:
            merged["_diff_total"] = (
                pd.to_numeric(merged["赋分总分_本次"], errors='coerce') -
                pd.to_numeric(merged["赋分总分_上次"], errors='coerce')
            )
            merged["_diff_school"] = -(
                pd.to_numeric(merged["赋分校名_本次"], errors='coerce') -
                pd.to_numeric(merged["赋分校名_上次"], errors='coerce')
            )
            merged["_diff_class"] = -(
                pd.to_numeric(merged["赋分班名_本次"], errors='coerce') -
                pd.to_numeric(merged["赋分班名_上次"], errors='coerce')
            )
        # 预计算单科王（最高分及其学生）
        subject_columns = ["语文", "数学", "英语", "物理", "化学赋分", "生物赋分"]
        subject_top_students = {}
        for subj in self.subject_columns:
            # 尝试查找原始列名或带 _本次 的列名
            col_candidates = [subj, subj + "_本次"]
            col = None
            for c in col_candidates:
                if c in merged.columns:
                    col = c
                    break
            if col is None:
                continue
            scores = pd.to_numeric(merged[col], errors='coerce')
            if not scores.isna().all():
                max_score = scores.max()
                top_students = merged[scores == max_score]["学号"].tolist()
                subject_top_students[subj] = top_students
        # ================== 新增：预计算进步/退步 Top N 并列名单 ==================
        cfg = load_config()
        rank_by = cfg.get("overview_rank_by", "赋分班名")  # 默认使用赋分班名
        if rank_by == "赋分总分":
            diff_col = "_diff_total"
        elif rank_by == "赋分校名":
            diff_col = "_diff_school"
        else:
            diff_col = "_diff_class"
        has_progress_rule = any(r["type"] in ["进步排名", "退步排名"] for r in self.rules)
        prog_top_students = set()   # 存放进步前 N 名的学号
        reg_top_students = set()    # 存放退步前 N 名的学号
        if has_progress_rule and previous is not None:
            for rule in self.rules:
                rule_type = rule["type"]
                if rule_type not in ["进步排名", "退步排名"]:
                    continue
                n = rule["range_end"]
                if n <= 0:
                    n = 5
                diff_list = []
                for _, row in merged.iterrows():
                    stu_id = str(row["学号"])
                    if pd.isna(stu_id):
                        continue
                    diff_val = row.get(diff_col, None)
                    if pd.isna(diff_val):
                        continue
                    if rule_type == "进步排名":
                        if diff_val > 0:
                            diff_list.append((stu_id, diff_val))
                    else:
                        if diff_val < 0:
                            diff_list.append((stu_id, -diff_val))
                if not diff_list:
                    continue
                diff_list.sort(key=lambda x: x[1], reverse=True)
                if len(diff_list) > n:
                    threshold = diff_list[n-1][1]
                    top_ids = [sid for sid, val in diff_list if val >= threshold]
                else:
                    top_ids = [sid for sid, _ in diff_list]
                if rule_type == "进步排名":
                    prog_top_students.update(top_ids)
                else:
                    reg_top_students.update(top_ids)

        valid_current_ids = set(current["学号"].astype(str))
        bonus_list = []
        # 遍历每条规则，运行时防护 start>end
        for rule in self.rules:
            if rule["range_start"] > rule["range_end"]:
                rule["range_start"], rule["range_end"] = rule["range_end"], rule["range_start"]

        def get_numeric_value(row, candidates):
            for col in candidates:
                if col not in row:
                    continue
                val = row[col]
                if val is None:
                    continue
                if isinstance(val, (int, float)):
                    return float(val)
                try:
                    return float(val)
                except (ValueError, TypeError):
                    continue
            return None

        for _, row in merged.iterrows():
            stu_id = str(row["学号"])
            # 过滤：只处理本次成绩存在的学生，跳过outer合并带入的旧学生
            if stu_id not in valid_current_ids:
                continue
            name = row.get("姓名", self.id_to_name.get(stu_id, ""))
            if not name:
                continue
            total = 0
            rule_descs = []

            # 动态组装候选字段，区分是否merge过
            if previous is not None:
                cand_total = ["赋分总分_本次", "赋分总分", "总分", "总成绩", "总分（赋分）"]
                cand_school = ["赋分校名_本次", "赋分校名", "校名次", "校排名", "校名", "学校排名"]
                cand_class = ["赋分班名_本次", "赋分班名", "班名次", "班排名", "班级排名", "班名"]
            else:
                cand_total = ["赋分总分", "总分", "总成绩", "总分（赋分）"]
                cand_school = ["赋分校名", "校名次", "校排名", "校名", "学校排名"]
                cand_class = ["赋分班名", "班名次", "班排名", "班级排名", "班名"]

            for rule in self.rules:
                rule_type = rule["type"]
                r_start = rule["range_start"]
                r_end = rule["range_end"]
                score = rule["score"]

                value = None
                if rule_type == "赋分总分":
                    value = get_numeric_value(row, cand_total)
                elif rule_type == "赋分校名":
                    value = get_numeric_value(row, cand_school)
                elif rule_type == "赋分班名":
                    value = get_numeric_value(row, cand_class)
                elif rule_type == "进步排名":
                    if stu_id in prog_top_students:
                        total += score
                        rule_descs.append(self.rule_description(rule))
                    continue
                elif rule_type == "退步排名":
                    if stu_id in reg_top_students:
                        total += score
                        rule_descs.append(self.rule_description(rule))
                    continue
                elif rule_type == "单科王":
                    subject = rule.get("subject", "")
                    if subject == "all":
                        for subj in self.subject_columns:
                            if subj in subject_top_students and stu_id in subject_top_students[subj]:
                                total += score
                                rule_descs.append(f"单科王({subj})")
                        continue
                    else:
                        if subject in subject_top_students:
                            if stu_id in subject_top_students[subject]:
                                value = 1
                            else:
                                continue
                        else:
                            continue
                if value is None:
                    continue
                if r_start <= value <= r_end:
                    total += score
                    rule_descs.append(self.rule_description(rule))

            if total != 0:
                bonus_list.append((stu_id, name, total, rule_descs))

        if not bonus_list:
            QMessageBox.information(self, "提示", "没有学生符合任何规则，无需加分")
            return
        self.show_confirm_dialog(bonus_list)

    def show_confirm_dialog(self, bonus_list):
        """显示加分确认表格"""
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("确认加分明细")
        dialog.resize(750, 450)
        dialog.setStyleSheet(DIALOG_STYLE + f"""
            QTableWidget {{
                background: {_c['bg']};
                alternate-background-color: {_c['bg_alt']};
                gridline-color: {_c['border']};
                border-radius: 8px;
                border: 1px solid {_c['border']};
                selection-background-color: #64B5F6;
                selection-color: white;
            }}
            QTableWidget::item {{
                padding: 6px;
                color: {_c['fg']};
            }}
            QTableWidget::item:selected {{
                background: #64B5F6;
                color: white;
            }}
            QHeaderView::section {{
                background: {_c['bg_card']};
                padding: 6px;
                border: none;
                font-weight: 600;
                color: {_c['fg_dim']};
                border-bottom: 2px solid {_c['border']};
            }}
            QPushButton {{
                border-radius: 8px;
                padding: 6px 18px;
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #64B5F6,stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: none;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #42A5F5,stop:1 #1E88E5);
            }}
            QPushButton:pressed {{
                background: #1E88E5;
            }}
            QPushButton#cancelBtn {{
                background: #555555;
            }}
            QPushButton#cancelBtn:hover {{
                background: #6A6A7A;
            }}
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        label = QLabel(f"共 {len(bonus_list)} 名学生将获得加分，请确认：")
        label.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
        layout.addWidget(label)

        table = QTableWidget()
        install_scrolling_header(table)
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["学号", "姓名", "加分", "规则"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setRowCount(len(bonus_list))
        for i, (sid, name, score, descs) in enumerate(bonus_list):
            table.setItem(i, 0, QTableWidgetItem(sid))
            table.setItem(i, 1, QTableWidgetItem(name))
            table.setItem(i, 2, QTableWidgetItem(f"{score:+d}分"))
            table.setItem(i, 3, QTableWidgetItem("；".join(descs)))
        layout.addWidget(table)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setObjectName("cancelBtn")
        cancel_btn.clicked.connect(dialog.reject)
        ok_btn = QPushButton("✅ 确认提交")
        ok_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.commit_bonus(bonus_list)

    def commit_bonus(self, bonus_list):
        """实际写入积分"""
        date_str = datetime.now().strftime("%Y-%m-%d")
        success = 0
        for sid, name, score, descs in bonus_list:
            note = "考试加分：" + "；".join(descs)
            ok, _ = save_score_record(sid, name, date_str, "+", "考试加分", score, note, assign_score=None)
            if ok:
                success += 1
        show_toast(self.parent.app, f"成功为 {success} 名学生添加考试加分", "success")
        if self.parent and hasattr(self.parent, 'app'):
            self.parent.app.refresh_activity()
        self.accept()


class HeartPopup(QDialog):
    def __init__(self, text, x=None, y=None):
        super().__init__()
        self.setWindowTitle("💖 爱心便签")
        self.setFixedSize(200, 120)

        # 随机背景色（与原版完全一致）
        colors = ['#FFB6C1', '#ADD8E6', '#90EE90', '#FFFACD', '#DDA0DD',
                  '#F0E68C', '#FFDAB9', '#B0E0E6', '#E6E6FA', '#F5DEB3']
        bg = random.choice(colors)
        self.setStyleSheet(f"""
            QDialog {{
                background: {bg};
                border-radius: 10px;
            }}
            QLabel {{
                background: transparent;
                color: #222;
                font-family: 'Microsoft YaHei';
                font-size: 14px;
                font-weight: bold;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        label = QLabel(text, self)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setWordWrap(True)
        layout.addWidget(label)

        # 定位（与原版一致）
        screen = QApplication.primaryScreen().availableGeometry()
        if x is not None and y is not None:
            x = max(0, min(x, screen.width() - self.width()))
            y = max(0, min(y, screen.height() - self.height()))
            self.move(x, y)
        else:
            x = random.randint(0, screen.width() - self.width())
            y = random.randint(0, screen.height() - self.height())
            self.move(x, y)

    def showEvent(self, event):
        super().showEvent(event)
        self.raise_()
        self.repaint()

# ======================== 导航图标临时浮层（用于弹动/旋转） ========================
class _NavIconOverlay(QLabel):
    """覆盖在导航图标上的临时浮层，用来显示放大/旋转后的图标。"""

    def __init__(self, pixmap, global_center, base_size):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)

        self._base_pixmap = pixmap
        self._center = QPoint(global_center)
        self._base_size = max(8, int(base_size))
        self._scale = 1.0
        self._angle = 0.0

        self._render()

    def set_scale(self, s):
        self._scale = float(s)
        self._render()

    def set_angle(self, a):
        self._angle = float(a)
        self._render()

    def _render(self):
        try:
            size = max(1, int(self._base_size * self._scale))
            scaled = self._base_pixmap.scaled(
                size, size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            if abs(self._angle) > 0.01:
                rotated = QPixmap(scaled.size())
                rotated.fill(Qt.GlobalColor.transparent)
                painter = QPainter(rotated)
                painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                cx = scaled.width() / 2.0
                cy = scaled.height() / 2.0
                painter.translate(cx, cy)
                painter.rotate(self._angle)
                painter.translate(-cx, -cy)
                painter.drawPixmap(0, 0, scaled)
                painter.end()
                scaled = rotated

            self.setPixmap(scaled)
            self.resize(scaled.size())
            self.move(
                self._center.x() - scaled.width() // 2,
                self._center.y() - scaled.height() // 2,
            )
        except Exception:
            pass

# ======================== 主窗口（应用全局样式 + 顶部时钟） ========================
class MainWindow(FluentWindow):
    # 静默更新：子线程 → 主线程
    _silent_update_signal = pyqtSignal(str, str, str)

    def __init__(self):
        super().__init__()
        # ----- 根据用户主题配置设置 qfluentwidgets 主题 -----
        setTheme(get_qfluent_theme(), save=True)
        setThemeColor("#64B5F6")
        # ---------------------------------
        self.setWindowTitle("PickOS 班级管理系统")
        self.resize(1280, 720)

        

        # 数据成员
        self.stu_df = None
        self.id_to_name = {}
        if os.path.exists(CLASS_LIST_FILE):
            self.current_file = CLASS_LIST_FILE
        else:
            self.current_file = read_config()
        self.template_path = read_template_config()
        self.start_time = time.time()
        self.backup_mgr = BackupManager(backup_dir=BACKUP_DIR)
        self.input_frame = None
        

        init_score_file()
        self.load_student_data(self.current_file)

        # 创建页面
        # 先创建其他所有页面（因为首页需要引用它们）
        self.inputInterface = InputInterface(self)
        self.rankingInterface = RankingInterface(self)
        self.exportInterface = ExportInterface(self)
        self.settingsInterface = SettingsInterface(self)
        self.settingsInterface.setObjectName("settingsInterface")
        self.aboutInterface = AboutInterface(self)
        self.listManagerInterface = ListManagerInterface(self)
        self.scoreInterface = ScoreInterface(self)
        # 最后创建首页（此时 app.inputInterface 等已存在）
        self.homeInterface = HomeInterface(self)

        self.addSubInterface(self.homeInterface, FluentIcon.HOME, "首页", NavigationItemPosition.TOP)
        self.addSubInterface(self.inputInterface, FluentIcon.EDIT, "录入", NavigationItemPosition.TOP)
        self.addSubInterface(self.rankingInterface, FluentIcon.ROBOT, "排行", NavigationItemPosition.TOP)
        self.addSubInterface(self.scoreInterface, FluentIcon.LIBRARY, "成绩", NavigationItemPosition.TOP)
        self.addSubInterface(self.listManagerInterface, FluentIcon.PEOPLE, "名单", NavigationItemPosition.TOP)
        self.addSubInterface(self.exportInterface, FluentIcon.SHARE, "导出", NavigationItemPosition.TOP)
        self.addSubInterface(self.settingsInterface, FluentIcon.SETTING, "设置", NavigationItemPosition.BOTTOM)
        self.addSubInterface(self.aboutInterface, FluentIcon.INFO, "关于", NavigationItemPosition.BOTTOM)

        

        # 应用现代化样式表（增强版）
        self.setStyleSheet(_replace_theme_colors("""
            QWidget {
                font-family: "Inter", "Segoe UI", "Microsoft YaHei", sans-serif;
                font-size: 17px;
            }
            QPushButton, PrimaryPushButton {
                border-radius: 12px;
                padding: 12px 32px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(100, 181, 246, 0.3);
                font-size: 19px;
            }
            QPushButton:hover, PrimaryPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
                border-color: #64B5F6;
            }
            QPushButton:pressed, PrimaryPushButton:pressed {
                background: #1E88E5;
                border-color: #1E88E5;
            }
            QPushButton:disabled, PrimaryPushButton:disabled {
                background: #555555;
                color: #888888;
                border-color: #555555;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                border-radius: 10px;
                padding: 10px 14px;
                border: 1px solid #3A3A4A;
                background: #1A1A24;
                color: #F0F0F0;
                font-size: 17px;
                selection-background-color: #64B5F6;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {
                border: 2px solid #64B5F6;
                background: #1E1E2A;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #A0A0B0;
                margin-right: 8px;
            }
            QComboBox QAbstractItemView {
                border-radius: 10px;
                background: #1E1E2A;
                border: 1px solid #3A3A4A;
                padding: 4px;
                outline: none;
            }
            QComboBox QAbstractItemView::item {
                padding: 8px 12px;
                border-radius: 6px;
                color: #E0E0E0;
                font-size: 16px;
            }
            QComboBox QAbstractItemView::item:selected {
                background: #64B5F6;
                color: white;
            }
            QComboBox QAbstractItemView::item:hover {
                background: #3A3A4A;
            }
            QCheckBox {
                spacing: 10px;
                font-size: 17px;
                color: #E0E0E0;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
                border-radius: 6px;
                border: 2px solid #5A5A6A;
                background: #1A1A24;
            }
            QCheckBox::indicator:checked {
                background: #64B5F6;
                border-color: #64B5F6;
                image: url(:/qfluentwidgets/images/check_white.svg);
            }
            QCheckBox::indicator:hover {
                border-color: #64B5F6;
            }
            QRadioButton {
                spacing: 10px;
                font-size: 17px;
                color: #E0E0E0;
            }
            QRadioButton::indicator {
                width: 20px;
                height: 20px;
                border-radius: 10px;
                border: 2px solid #5A5A6A;
                background: #1A1A24;
            }
            QRadioButton::indicator:checked {
                background: #64B5F6;
                border-color: #64B5F6;
            }
            QRadioButton::indicator:hover {
                border-color: #64B5F6;
            }
            QTableWidget {
                gridline-color: #2A2A3A;
                background: #16161E;
                alternate-background-color: #1E1E2A;
                border-radius: 14px;
                border: 1px solid #2A2A3A;
                selection-background-color: #64B5F6;
                selection-color: white;
                font-size: 16px;
            }
            QTableWidget::item {
                padding: 10px;
                color: #E0E0E0;
                border: none;
            }
            QTableWidget::item:selected {
                background: #64B5F6;
                color: white;
            }
            QTableWidget::item:hover {
                background: #2A2A3A;
            }

           
            QHeaderView::section {
                background: #1E1E2A;
                padding: 10px;
                border: none;
                font-weight: 600;
                color: #C0C0D0;
                border-bottom: 2px solid #2A2A3A;
                font-size: 16px;
            }
            QHeaderView::section:checked {
                background: #64B5F6;
                color: white;
            }
            QScrollBar:vertical {
                width: 6px;
                background: transparent;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #4A4A5A;
                border-radius: 4px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background: #6A6A7A;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar:horizontal {
                height: 6px;
                background: transparent;
                margin: 0px;
            }
            QScrollBar::handle:horizontal {
                background: #4A4A5A;
                border-radius: 4px;
                min-width: 30px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #6A6A7A;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            QGroupBox {
                border: 1px solid #3A3A4A;
                border-radius: 14px;
                margin-top: 18px;
                padding-top: 14px;
                font-weight: 600;
                color: #D0D0E0;
                font-size: 17px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 10px;
                color: #C0C0D0;
                background: transparent;
            }
            QLabel {
                color: #E0E0E0;
            }
            QTabWidget::pane {
                border: 1px solid #3A3A4A;
                border-radius: 14px;
                background: #16161E;
            }
            QTabBar::tab {
                background: #1E1E2A;
                padding: 10px 20px;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                margin-right: 4px;
                color: #A0A0B0;
                font-weight: 500;
                font-size: 16px;
            }
            QTabBar::tab:selected {
                background: #2A2A3A;
                color: white;
                border-bottom: 3px solid #64B5F6;
            }
            QTabBar::tab:hover {
                background: #2A2A3A;
            }
            CardWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #2A2A3A, stop:1 #1E1E2A);
                border-radius: 18px;
                border: 1px solid #3A3A4A;
            }
            CardWidget:hover {
                border: 2px solid #64B5F6;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #353545, stop:1 #252535);
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            QFrame {
                border: none;
            }
        """))

        # 窗口居中
        screen_geo = QApplication.primaryScreen().availableGeometry()
        self.move((screen_geo.width() - self.width()) // 2,
                (screen_geo.height() - self.height()) // 2)

        # ---------- 添加顶部实时时钟 ----------
        self._init_time_label()
        self._init_level_label()

        self._set_window_icon()
        reset_daily_experience_if_needed()
        self._init_online_timer()

        self.update_title_with_level()
        QApplication.instance().main_window = self
        # 设备运行经验 / 首次成就检查 / 在线检查，统一移到
        # start_post_boot_tasks() 里，由主入口在"主窗口真正显示后"调用，
        # 避免新手引导 / 密码验证阶段弹成就、升级通知。
        # 在 __init__ 末尾添加
        self.create_tray_icon()
        self.floating_window = None   # 悬浮窗实例

        # ---------- 表格选中高亮：点击表格外自动清除 ----------
        self.table_selection_cleaner = TableSelectionCleaner(self)
        app_instance = QApplication.instance()
        if app_instance is not None:
            app_instance.installEventFilter(self.table_selection_cleaner)
        self.table_selection_cleaner.register(getattr(self.rankingInterface, "table", None))
        self.table_selection_cleaner.register(getattr(self.scoreInterface, "table", None))
        self.table_selection_cleaner.register(getattr(self.listManagerInterface, "table", None))

        # ★ 主窗口构造完毕，正式启用粒子效果
        ParticleEventFilter._app_ready = True

        # 连接静默更新信号到主线程槽
        self._silent_update_signal.connect(self._on_silent_update_found)

        # ======================== 导航栏切换动画 ========================
    def switchTo(self, interface, *args, **kwargs):
        """重写 switchTo：切换后播放导航图标弹动，设置额外齿轮旋转。"""
        super().switchTo(interface, *args, **kwargs)
        try:
            route_key = interface.objectName()
        except Exception:
            return
        # 等布局稳定后再取图标的全局位置
        QTimer.singleShot(30, lambda k=route_key: self._animate_nav_icon(k))

    def _animate_nav_icon(self, routeKey):
        if not routeKey:
            return
        try:
            if not load_config().get("enable_nav_animation", True):
                return
        except Exception:
            pass

        btn = self._find_nav_button(routeKey)
        if btn is None:
            return

        icon_label = getattr(btn, "iconLabel", None)
        if icon_label is None or icon_label.width() <= 0:
            return

        is_settings = (
            hasattr(self, "settingsInterface")
            and routeKey == self.settingsInterface.objectName()
        )
        self._play_nav_icon_animation(icon_label, spin=is_settings)

    def _find_nav_button(self, routeKey):
        """在 qfluentwidgets 的导航面板中按 routeKey 查找按钮，兼容多种版本。"""
        try:
            panel = getattr(self.navigationInterface, "panel", None)
            if panel is None:
                return None

            # 1) 标准方式：NavigationPanel.items 字典
            items = getattr(panel, "items", None)
            if isinstance(items, dict):
                btn = items.get(routeKey)
                if btn is not None:
                    return btn

            # 2) 兜底：QScrollArea.widget() 的 widget 里 findChildren
            inner = None
            try:
                inner = panel.widget()
            except Exception:
                inner = None
            if inner is not None:
                from qfluentwidgets import NavigationWidget  # 延迟导入，兼容不同版本
                for w in inner.findChildren(NavigationWidget):
                    if getattr(w, "objectName", lambda: "")() == routeKey:
                        return w
                # 有些版本按钮的 routeKey 存在 property 里
                for w in inner.findChildren(NavigationWidget):
                    if getattr(w, "routeKey", None) == routeKey:
                        return w

            # 3) 再兜底：panel 直接 findChildren
            try:
                from qfluentwidgets import NavigationWidget
                for w in panel.findChildren(NavigationWidget):
                    if getattr(w, "routeKey", None) == routeKey:
                        return w
            except Exception:
                pass
        except Exception as e:
            print(f"查找导航按钮异常: {e}")
        return None

    def _get_icon_pixmap(self, icon_label):
        """尽量从 IconWidget 中拿到高质量 pixmap。"""
        try:
            icon = getattr(icon_label, "icon", None)
            if icon is None:
                return None

            # FluentIcon / FluentIconBase
            if hasattr(icon, "icon") and callable(icon.icon):
                qicon = icon.icon()
                if qicon is not None and not qicon.isNull():
                    return qicon.pixmap(128, 128)

            # 普通 QIcon
            if hasattr(icon, "pixmap"):
                pm = icon.pixmap(128, 128)
                if pm is not None and not pm.isNull():
                    return pm
        except Exception:
            pass
        return None

    def _play_nav_icon_animation(self, icon_label, spin=False):
        """在图标上覆盖一层临时浮层，播放弹动（和齿轮旋转）。"""
        try:
            pixmap = self._get_icon_pixmap(icon_label)
            if pixmap is None or pixmap.isNull():
                return

            global_tl = icon_label.mapToGlobal(QPoint(0, 0))
            center = QPoint(
                global_tl.x() + icon_label.width() // 2,
                global_tl.y() + icon_label.height() // 2,
            )
            # 图标控件很小，浮层以它为基础尺寸做缩放
            base_size = max(icon_label.width(), icon_label.height(), 20)

            overlay = _NavIconOverlay(pixmap, center, base_size)
            overlay.show()
            overlay.raise_()

            # ---- 弹动：1.0 -> 1.40 -> 0.86 -> 1.0 ----
            scale_anim = QVariantAnimation(overlay)
            scale_anim.setDuration(420)
            scale_anim.setStartValue(1.0)
            scale_anim.setKeyValueAt(0.35, 1.40)
            scale_anim.setKeyValueAt(0.70, 0.86)
            scale_anim.setKeyValueAt(1.00, 1.0)
            scale_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            scale_anim.valueChanged.connect(
                lambda v: overlay.set_scale(float(v))
            )

            # ---- 齿轮旋转：0° -> 720° ----
            angle_anim = None
            if spin:
                angle_anim = QVariantAnimation(overlay)
                angle_anim.setDuration(700)
                angle_anim.setStartValue(0.0)
                angle_anim.setEndValue(720.0)
                angle_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
                angle_anim.valueChanged.connect(
                    lambda v: overlay.set_angle(float(v))
                )

            def _cleanup():
                try:
                    overlay.close()
                except Exception:
                    pass

            scale_anim.finished.connect(_cleanup)

            # 挂在 overlay 上，防止动画被 GC
            overlay._scale_anim = scale_anim
            overlay._angle_anim = angle_anim

            scale_anim.start()
            if angle_anim is not None:
                angle_anim.start()

        except Exception as e:
            print(f"图标动画异常: {e}")

    def start_post_boot_tasks(self):
        """主窗口真正显示后调用：设备运行经验、首次成就检查、在线检查。
        只执行一次，供新手引导完成 / 常规启动两条路径共用。
        """
        if getattr(self, "_post_boot_started", False):
            return
        self._post_boot_started = True

        # 从现在起允许成就/升级通知弹出
        POST_BOOT_READY["value"] = True

        # 1) 设备运行经验（可能升级，此时主窗口已可见，InfoBar 能正常挂上去）
        try:
            self._check_device_runtime()
        except Exception as e:
            print(f"⚠️ 设备运行经验异常：{e}")

        # 2) 首次成就检查（含 first_launch、walk_* 等）
        QTimer.singleShot(600, check_and_unlock_achievements)

        # 3) 在线经验检查一次
        QTimer.singleShot(1000, self._check_online_experience)

        # 4) 启动后 5 秒静默检查更新
        QTimer.singleShot(5000, self._silent_check_update)

    def _silent_check_update(self):
        """启动 5 秒后静默检查更新（后台线程，不阻塞 UI）"""
        is_frozen = getattr(sys, "frozen", False)
        # 源码运行时也允许检查（方便调试）；如果不想在开发环境被弹窗打扰，
        # 可以在 config.json 里加 "dev_skip_silent_update": true 来控制。
        if not is_frozen:
            try:
                if load_config().get("dev_skip_silent_update", False):
                    print("ℹ️ [静默更新] 源码模式，dev_skip_silent_update=True，跳过")
                    return
            except Exception:
                pass

        def check_thread():
            try:
                print("🔍 [静默更新] 开始检查...")
                has_update, latest_ver, changelog, url = UpdateManager.check_update()
                print(f"🔍 [静默更新] 检测结果 has_update={has_update}, latest={latest_ver}, url={url}")
                if has_update and url:
                    # ★ 用信号回主线程，PyQt 会自动用 QueuedConnection 排队
                    self._silent_update_signal.emit(
                        latest_ver or "", changelog or "", url or ""
                    )
                elif not has_update:
                    print("🔍 [静默更新] 已是最新版本")
            except Exception as e:
                print(f"⚠️ [静默更新] 检查失败：{e}")

        threading.Thread(target=check_thread, daemon=True).start()


    @pyqtSlot(str, str, str)
    def _on_silent_update_found(self, latest_ver, changelog, url):
        """主线程收到静默更新结果 → 弹窗询问"""
        print(f"🔍 [静默更新] 主线程收到弹窗请求 v={latest_ver}")
        try:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("发现新版本")
            msg_box.setText(f"检测到新版本 {latest_ver}，是否立即查看并更新？")
            msg_box.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)
            msg_box.setStyleSheet(DIALOG_STYLE)
            # ★ 保证弹出窗口位于主窗口之上
            msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
            msg_box.raise_()
            msg_box.activateWindow()

            if msg_box.exec() == QMessageBox.StandardButton.Yes:
                self._update_window = _NeuUpdateWindow(latest_ver, changelog, url)
                self._update_window.show()
                self._update_window.raise_()
                self._update_window.activateWindow()
        except Exception as e:
            print(f"⚠️ [静默更新] 弹窗失败：{e}")
            import traceback
            traceback.print_exc()

    def _check_device_runtime(self):
        """检查设备运行时长经验：根据安装以来的总天数，每过一天加50经验（不占用每日上限）"""
        cfg = load_config()
        install_time = cfg.get("install_time", time.time())
        current_days = int((time.time() - install_time) // 86400)
        last_recorded_days = cfg.get("device_total_days_recorded", 0)

        if current_days > last_recorded_days:
            delta_days = current_days - last_recorded_days
            exp_gain = delta_days * DEVICE_RUNTIME_DAILY_EXP
            if exp_gain > 0:
                success, actual, is_level_up = add_experience(
                    exp_gain,
                    "设备运行",
                    check_limit=False,      # 不检查每日上限
                    update_daily=False      # 不计入今日经验
                )
                if success:
                    cfg = load_config()
                    cfg["device_total_days_recorded"] = current_days
                    save_config(cfg)
                    self.update_title_with_level()
                    show_toast(
                        self,
                        f"📱 设备运行 +{actual} 经验（累计运行{current_days}天）",
                        "info",
                        duration=2000
                    )

    def _init_level_label(self):
        """在标题标签后面插入等级标签（大字号）"""
        layout = self.titleBar.layout()
        if layout is None:
            return
        # 查找标题标签（文本为 "PickOS 班级管理系统"）
        title_label = None
        for i in range(layout.count()):
            widget = layout.itemAt(i).widget()
            if isinstance(widget, QLabel) and widget.text() == "PickOS 班级管理系统":
                title_label = widget
                break
        if title_label is None:
            return
        # 创建等级标签
        self.level_label = QLabel("Lv.1", self.titleBar)
        self.level_label.setStyleSheet("""
            color: #64B5F6;
            font-size: 20px;
            font-weight: bold;
            background: transparent;
            padding-left: 6px;
        """)
        self.level_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.level_label.setFixedHeight(30)
        self.level_label.adjustSize()
        # 插入到标题标签后面（索引+1）
        layout.insertWidget(layout.indexOf(title_label) + 1, self.level_label)
        # 更新等级
        self.update_title_with_level()

    

    def update_title_with_level(self):
        # 不传参，让 calculate_level 内部读取 total_experience
        level = calculate_level()
        if hasattr(self, 'level_label'):
            self.level_label.setText(f"Lv.{level}")
            self.level_label.adjustSize()
        self.setWindowTitle("PickOS 班级管理系统")

    def _init_online_timer(self):
        """初始化在线计时器（每30秒检查一次）"""
        self.online_timer = QTimer(self)
        self.online_timer.timeout.connect(self._check_online_experience)
        self.online_timer.start(30000)  # 30秒检查一次
        

    def _check_online_experience(self):
        """检查在线计时经验"""
        reset_daily_experience_if_needed()
        cfg = load_config()
        today = get_today_date()
        last_date = cfg.get("last_experience_date", "")
        
        # 日期变更时重置在线秒数
        if last_date != today:
            cfg["online_seconds_today"] = 0
            cfg["last_experience_date"] = today
            save_config(cfg)
        
        # 获取上次更新时间
        last_update = cfg.get("last_online_update", time.time())
        now = time.time()
        elapsed = now - last_update
        
        # 累加在线秒数
        online_seconds = cfg.get("online_seconds_today", 0) + elapsed
        cfg["online_seconds_today"] = online_seconds
        cfg["last_online_update"] = now
        save_config(cfg)
        
        # 检查是否达到2分钟（120秒）的整数倍
        intervals = int(online_seconds // ONLINE_INTERVAL_SECONDS)
        if intervals > 0:
            # 计算应得经验
            exp_gain = intervals * ONLINE_EXP_PER_INTERVAL
            # 检查在线每日上限
            today_online_exp = cfg.get("today_online_exp", 0)
            max_gain = min(exp_gain, DAILY_LIMIT_ONLINE - today_online_exp)
            
            if max_gain > 0:
                # 扣除已使用的秒数
                used_seconds = intervals * ONLINE_INTERVAL_SECONDS
                cfg["online_seconds_today"] = online_seconds - used_seconds
                cfg["today_online_exp"] = today_online_exp + max_gain
                save_config(cfg)
                
                # 添加经验（受每日上限限制）
                success, actual, is_level_up = add_experience(max_gain, "在线计时")
                if success and actual > 0:
                    # 更新窗口标题（等级可能变化）
                    self.update_title_with_level()


    def _set_window_icon(self):
        """从 Base64 字符串加载图标并设为窗口图标（任务栏也会显示）"""
        # 请将您的 Base64 图标字符串赋值给下面的变量
        base64_str = "iVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAACRWElEQVR4nO29CbhsRXUovHZ3n3PufJmRQQUcAAHFAVRERdQYRcVEUAERxdmo6Es+h/ei+aMm+fKMGofgiCiIAoJG0UTjPEXAATSRmRcZBEEB4XK595zT3fv/VlWtWsOuvXt39+6+ffEu6Hu691BVu3ateagsz/McaoO/1N+SQ5a1ACBLXxmu8Z8sfPrhrP+dQw9aWQfyvB/bzrJ29Qjitdg3Dd33gePx/fYHtjMe5KEf6pufST8zPk94Vjdu3UYRaC6pDXmsDvi+XJfhPhqnH8eApzLX0lzq9qvG7b+n+sK2iu36a4tzI/uitvz1fnztxNhkH7J/Xiv2XC7GVJwHfk5+h+VzSW3x+vO44X/3xFjs+OSzyrnUa1vPRRnQ3NdfMzjKAQudJ5lJhR+MpB34Ha+lj385dL4fJkEiiRuuasO2lV5s/LD8lxYWTjy98PQzjAJ6jPpFFcfP49KLwH7k9WWIn404XtF6lkbIFBSvtc9W9j70s1kEorb1d5obzRT4e3GOJMNhIifbTS1neia7ltPIr5/H9kkEq3xdaoJP/dG4yu6j+WiF/iQTsb/rMqZ60KpuhBHLP5y8LTWY9EB9c0zF8CHppXjCgH3StXWQVUoTPE48xu14Skzt1WkXb/HEp3wSmTh5AsWLiJ69XxNZyn4PHGWCSDAiDIPw9fsZDur0X7ymDPlpHLTmcP57Yk16JGFEpbXl54EQyq85iYjlDMZfm5IgUxISPw8xP+rLn5YErp3g8DQOnm/fN61xuk4+Y3NEoFO4PVBvQlqaQJx0Ghw+f7ET+buPdzgOT5PD1xCSymMktnooimpyEul6K+J6CsrjYmLgJ3KAsJN4uTxu2VYcpWhbjo2fISWmplSXonpUBTxvXhS2KlEdTlEHRm2LCKR+p9ymVhf0fNC8yXsI0VsCOWiRS4mS+2OiL5kMEQfuNw/IWvaMlkhZqcGe179TUox8fvm7+Cx+Xnrh4wkC4mCG4x5ImGm9Z/UIgOaOxUEzMdAPX861EfnpQfrm4eRilQNkBJVITPexOE1EST+wHoM9P1icpkWrCQ8tnOQdou006MUsXzbdlxb3q8VT+UxMNIq2hhQC1gePPOMQkjJJkJ+NkBGPE1Px3+VasAzE38NI3S6RJrAfv9by3HPz4vpt1ZxrumcQE7HtsMTL68sSVk0Q6HlpfLpNZKx8rb5fMxBeByQRpaGT1rUZYemlFV8mcXNL2exEyIWkdfO0/osfEr+obSlOWc7viQK3ISdDT4B92Wygsc/WF8hrEdfOhz1HbeDYinoscS8WXWlskrPXWZD0bvQL5sVGhLY+sPTn/7KaRuMfBIR4WeI9sQrJnFf3nRa7rQ2FoIrD6bVVNp+ZIzrZyDaguoSD3ge98zzvmrZIGpESt147vM6LjIquKT6fn+MqItfRk1Auikn9iQdHx4r36wFZrpe6LkUVWVROi2opEdoepzEnn9/0KYkDE5s0pDi6/yv1QQ9tI83wfVpCSPUhEVlSeyKU9hp7b91FKnXQ4pjqccrUM5A4Xn6XJtDuiDrHKgURCElEky2qMZV5B/KhuD8RIu1lqb6nbVQS6o/WA6/pVqs4Pgm+b2nD8OMhJkNSFAPbTKqIeNbvd/HWRKepm7W+VpyclB5bdCP54Vdxpyzp3rPcW/+WXK9v9CiaoLQ4xAYlS+T8y04ZEHnCNfeV4/NU3V9DFL76ea04pxFYv2RaBJAwctUDllIIyQgpuG3tWmVCxMclklqVRqqBCHr+025HmvciMbbXpZ93EHPxv1P6f54gQMNJQDDAPZiStIvHND5pzwNzdfy3WtLza7faKBsIgB5MkZrQefk9pXPXs+DzRPMDS2Dkqz/pWmSS3FtPcErM5MUgEV8ujJSakOZWWvSVL9gSxwqWOBCKkkfacFo9h0WiSeDvZX1cEyY9h6mYCGvopL801vIYDclI6lm9U880eO2k1ZB+rXm01znjXOlY+T0VY2PoPL/PsrmRfWhiQqqTlXJr2gA0pZaLquifzZELmGP+ujqIb6m7FNP1AIdFfOv60ro7UUzmwPJFsJhGCJ2i+EXVIy0mM9Ut+oDlS672JVc9Z1EqSy/2amnAunrtWMjiTs8ndVAyVKX87JAI2ilTJVPPSO9IivqMbEVJrQhl9p7iddXSEkt9UlVjpNOMwhotJef386FdkNyGdFkWpRF6FunKLq5Lz6AkPklpsEwK8tDSE0GDkTf6B3CL3r0YT4WY4tVFVsthuL/JAMcb4KfVIlcKcT2pU4JZwHLiLXG0II+njTds3KFr5PNLzmH/ipbiApHzTog3bAAII6cmesXrytdH+lp9PS3y1DOWjctKSHysnppTThRdq5VzlYVYEOmbl2CZHQeseZsPISK9LyklSiS3SErX2OcI3+L8yahCOmeZKKsO8jHLAuuyfr+f48lWiyOzivqqFxOtjpv2dac4qBV5EIbXWeuCFtukCEaivNSnJBWlcUlJoSpM1T4bnZNiuabgxXmV9w2ClNiYcova+eDFbK3MRavz6KHE9pl0m8MQekKsavG1acgLkYKQsD1Yxkh2Ji0JSNtRSmUa5p2xBGZtH3nFtTQ+/p2ay05RJLWiE4sWEknKrcyW0vcBnFEptB99m5N5sVI00wgsI8X0xDEVpeslx5IxC1KHs88bR2AIpeUKEqzqI8Nj81rShvaRp+eUjXRsz7Dch8RWO65hEVAiSR0CP4wlfpKQRySiyEFt6LXvxD8b5SRo1ZYt9ik8kQTFxrmoEZUwFw7Kq34WbkcHTWlo6UWd0ss1ZWPoJxaM7SS055rGv1ZMaRaknlzk3v3EuCFh7NIcXBKHajFbEkpLIKoIHomHEvnL5tzcKWwCaSuzvVYSOsuphyfKKXFaEiTyhAzK6yjOqwwXHw3qq0S5isvQxE9ycryu+GzEMMgrQnPJtik6rpG2aDiVtgFtwGPiXDWPEresREBSnsY5bDPYACwC298pZKrSiWlRht/qPTaP+AR6wRQfNjU5RdDc2784qVJo7mntBsVxDALto2abgeynuj1tUygDHUuuDXOW4I0DRCxlRGJ1uG26z+ENpcNDbtSjrLC29Xk7Hnkd/ouWeikJ9MT9lvGwCiZVDH8Jvw9vk6hDPC1OSmKVxjk3VkoHLrooLCcfRHlkqCNC0epZvGew/joKDA7h9ZOuXU7tgm7N3ME+j5UyiDBIqy95F6rFtaKYJxOlZGivFglllJfV22vM0MRUMG5fjyfldqvXDiMHi+PD2iXq95cbfTr9nUX7ohomXaDyeMrFp59Pu69ln9LVV+ZO1moGrxPZVmkugDToQQ091z5E2YCq7p2MAdC2be0Z9IKtS0XGJuggkZQeBgbhmWhqYiFDc/X9MrLMf4gQpTmoJTZyQQ6H/MNeOwpUEd+sth2AjbkUp2DjTZqu+ZCZsep1wUS5DLE0DhSvswRN+vQt/khbgWzDRxHqcSFYKYul8Eq3ab/fy9MLUwaJFPXGIhcSnNHFC0gpQhID5mCpwItJgebEclwshTBnKV4zZu8JXblVEexhx1kVpJKbTLctC0XD6uj3ygxUOX+DuNq4kIsoTssc2VYwqiorGUeZdJ0ZSbUs4m/QGCgCtdzj1inj1JxqKsWKYkAN6TBKV3KCRb9CRSBL6ySlAAs6A1GLbVUupyYWmQ1UGqTb2j7TCDV4/iYt6hdhHIKuo+uYuGk9lzji+AQvL5E8pOE3vSaImVQNQKoMqteEkdiOSY4jOfLKfj1zld4eK2WYOADmhCnDS3lYrY4ak6Jr1fVSdJ1k2a7Zdz1NGjgLbcvM86hQ9OBkJTaQZt9hXvCkCMJDzqycVrzMORl2HGwn0OqdO5IIKbeh0VVeH61SqjNuHfjaFsQ8WtLaXHRNEHCiSLGzcIezVlY9tDQybln4Y0D+SSHJ5ECuDVZrytN+GUGa9BZkSQkmuLCdQ8uK4nX6thZ6jQd6OfLz+L82zBePp+ZDSkn9GjEyIeDPSwBS92fjAUYHaj1TT0paj7VZYfbBtYFsG2wDuehJfE0RL59HT6HokyNsueHI/lNmQ5K/k60ZQ67EB6/GSE+ODiwqSge63VRflkBJHGS1iYzXHS3GSDFEivMWkbWIIl+GtnqnDGk26mwbbIMEx00iFycpjeZWHAx59BKRjYoTkjicVq7v9DouBhelrtWl5my1q2LyT3lfVnWQfTCzloZVPzYXCFSkplLkKBPbZSSUHlx1yKu2vG+DbUBrRgYkkbvNLvJJQ5algrNIOpHeMVrLaaRMu5+rQLclA4gGj7mq/FtW8ddJAJZLWz3Mfqff4le0OGpLt/af8vc/FiPcNhgVikZkXi/Wi9Ms5LFLNjZyrAd7voZocchrLR4Oun6Q+gGJ8uvAz4KRgNrHWRY6qMV8rbcVs410YI0cjA5m2UYItsEsQp7UvaVIToFJdSWTKs9Y1T1l15el1FePx9pWOtW6hQyK4QfwHF8HFxQJslYjmGhsE/23waRBZsCNJi1ksSovt1lE4hSylSEg2hE8upUnR6WyWFMwmOhwAFM6UpQYdEgHpowl3amct2KMABowqjO2KKebv+uBbOP+26B50Ig1zhrLksFb/m8IHUvdVdJaipikDevsAs2HliYI4cnI5+0WTDBsrkOLEb9c39IGP9KF/DRQxpP2XcqQySqf5zZvwDZoDihzrsxgiMf7fbuHhAZdzZmOFRG1VajUrO4QSCzjbKhtu8uP1dXt+AcRMYmvZDxlNduDx1+2YXiGbnYGIhuAdcOQ646vI8riG9RVSkjnL+YKUBuSGm2DbdAM2ESsIpRb7Rmke5t+W4LBqkV6DWvubrmurOOgvW3So+ZL1sk8AFllyINNjmIG612XHge9KkPRhBrvRFFQyhySAyHrPtEJ/SA0CTquXrbFD6zTjeVEbINt0DykPE26xkIZEMLLAiCEI/Tbi9bauC3v1+pCWT/F0F7ZD2fycUES2Qdozq4UEr3BCCcfSddmJrMBrV8zTFLWD1V8aCODsgkrm4hUfXemnJyo0BwhwHZ7vUE5+MODew4bmyJeWl5bjxzFGjw7oH314q//AVsr5MHwpoumSC8XMzBZhltWDKrrDSjbsUdmQ0opumg0JPXCRt2yh04eS/UVW+r3l/Nivf4y/39xDzMdLKSLgBRr44G5Rlpqt3kHtnbw+jUjTTMVhqYFeWRKZYjJRjXL1IZ17VVnMhb719mHEu90xSJZVo6N+9YLICN8RSCQHKD8K/UU+i0fgPQVuTMvHU+1S+fKtmwaffFh7sI111wDp512GszNzbnfvBDFovSzw08jzuOBfjAWYZoEGY7kR85Ku9WK/UBCJ9Sx3RYh8J5QMzXohW50rqkWtNs03gwy7Mf14e/zY+mFMfJ82/7xWXCMimNn/plp/p3QGypDU+ALzUe73XbbVs115mDNmrWw3Xbbwbr162HdunWw/fbbw7p1a2HVqlWwatXqUFk6LZH59vxczSZkBakmnjEFQgnYMJ5mkqk+wp0F7lzk8LJ/6yGw/chdj3Xb9lls7Q6xM9CgIgfyvEVmf97aAvRatBshNLsQcJHhYv3mN78JT33qUxttexsUod1pOaRfvWoVrFu7FnbaeWe4/157w4Me9ADY98EPhgc9aF/Ya6/7w84776LuI4K6pSSEujsHSUjbr2yiTgoYQUVriR2JWPTn+4og8x/KJQ9ZOSlVmk+PI6gA1UUD9MNAaakqnfGUoowpYpE1SgB+8IMfwJFHHllhoW0AmjVbzDQ464aYRudmw/98EmkpoNSw6667wr777QsHH3wwPP7wx8Ohhz4a9txzT3Vdt9t1UsG0JAONzC1T3MZyS4nc2qpPVXYGqwDZwJ2NtOtcgnXbWQKgx8W/y2IICEfZXCMkgKobi+6GIgEoK1ckqaC0FUir5/g2ACIAP77wx3DYYw+bLAHYBlFVQUiJzilD7A477ggHH/wwOPJJT4I//dOnO8KA7wyBVKvJEgK2XVk9GYSaVrhL1WPkNUWlyngT2CqQ97KobndS0rY4mXvAgXpshLSqg+2L1Txr/Y9XegmALsjMbrqSS1v3hsxlLtoBJGXVe8UVN0CQFGlU1ko2gIt/cjE8+tBHw1ShTOX7YwVHG/x7JPuItJ8gIOI/4hGPgGOPPQaOPfZ5sNdeeymbAZ5vVj1IVdQlkEjNQTKyf7S3sP8+Xi2+F20/9cZvo2k5Db9MVSjGJpQZ69M2ATm2IAFYy3wKpG9RU9Ey6sKGqSoCoB9uVEmACMDPfvYzOOSQQ4T/cxtmzgrg+yAOLyWE7XfYHp5z9HPgJS95CRx++OFxHS0vL0On02ncTlBnz8dMuQNtCvC44ynfrqtMHdBjlp433x6PTY+/KN2gDSaP76FlrfKiS/GXPzaqSXbOGyHIQZPIUr0XYL0gjcHgOM7MWpr/uIG4u/cKeGKAnP6O2++A008/HZ7whCfAE5/4BDjzzDPg7rvvdp4cvA7tBM2X/UIoM5B5xMH1i5y/GApsn2vYEUgiIr032tgow5K118165srUFvbmSIKC9pn4XTaqiY7tSNbAZxGnaMn1g/YD4AmelrG3jRxDPOA2mE0gbwARA7IF/OAHP4QXvegkOOTQQ+AjH/0wbNiwwUkBeD0SgmYgK7F3ZQU7AK9vGbsvVeZh4p98O8ToivfZHZaZ2+viKDovJ70rENX9S7kcCwQgBalijJzIUPSV4vW0fbjrytgIpgekf26DrQNk9CYSAvxccfkV8OpXvQYOffSh8MlPnu4TVzod6HaXG5EGSBzOMrQ1yHWrOSsjpLRpWbfgUD0nCIlr2UgDUJqIV9/gaG13UGi/ggDIwcn0wnRDNDBb1slP3hRF8i1AcBC2noi32QZSEUg9QELw0peeDE9+8pHwg+9/HzodrxY0He6dibVbBLTyW1F8FNBVsihzkYgKZTMW9xGEEpG/fLx+rGX1N2sRABuBVGZUK6seNNxMNaXjNW7009Je6THnL8+GbGdLQIr5zCCQekCE4Hvf+z4c+eQj4XWvey3cdttt7hgaCaVnYRigbdtzZ7diUZkZGLveGElTsS0SBp8jvZ6NepK7y2epE2JsbXaS+dn6HnqbcCI6rXJjgtxXDP2d0uofri4R8YlQ1EXGJpHWGW0GBKnUhjLPSvwZijFaI2wVck1KOCnrLzWWAcSsVrtTJgSI8Pj3Qx/6F3jc4w6Dr3/9a85IiOMbhQiQZJpFA3XxQYtG6UGTMejl2oC7MlVAtpcyzBdGaqoP+2MeD6XqotvFY5T2l2hQRAuFfdL5RqKK/loWXUQLlWGe+WQTUppqP6Xuie8uIi4+YgiMKSEaIaK/vK9BUl2dsVY9w6htz4gX1Yv8mbMDXHnlVXDUUc+Et7/tbdBd9pGE+Hc4KFYhlg+b8ojVZ1SSqlodn1Ri5vhsMLeZfNJ4V/0c6QAmqdrQ85FkESWA4k65vlGkjJJ6FKkUJXiQejAYpM4jemsyJly2P6qNpvRU4Pgtg8yUTGTbKNpf0og+QNIYGar6SM1NE/1OUGIgT0ALpYG8D+9817vg6U//U7jyiiugM9dxRGIcaTIXO+6wusuTUn+N5saAKA3qspS+ZLTyuHRPSuRN9V90v6eeIT12rGyUqBumKSFvBuofqK0CgupHPPkOCdmryjKNBS4GgESXZpHHR8ELIoMUNk+4V+irkBoohr6UIJS933HGbjn/ICnBepJmFProOszBSQPf/vZ3XPzA588916kJw6ie5QzKf09vs2XF7XIpl/dmpMVQbozzq4tsEak8f8tN5EuT57ThvZB9Gj8hEMhfIN0ggyPzeH5H3d572N1y60Or1UAI6SBxPAxXZT+G7ePi+fBeTJRnsZ+UeN6Eh8n2kaLvZf2kVNO6/cu/MHlpAJH+1t/9Dp73/OfD3/zN21Xo8bCQidTxqiXJQWtVE0hqQ+3eHT9mZkztSGJAeKo3TUm2lZWP239CZKbn4OSeYGpCXLoMOUeN3KP2Um03QQdcbv84K7COyN5iihyZhbyZiCIVEaqy9UgEte9+VJ297Bnkbzumqv6aUkkmANJT8I53vNOFEy8tLbljo3oIisiiv6eBJ1Hm0LChvI4ORu3I9uxxYszlL4Xdi+UqEUU6tsjnWNzUEwG5qbyJ3RfcSd1JJrtBSjEmVWJ8CtCILYGwNqpghNBCaiGCleKmaXWrORilrSp7hISqV2AJR+rvKOMa5/5g+MUPqgSf+tSn4JhjjoE777wzEIFx4gUyJS7T72IhEFrbNtxdRvDVmXyEdLXgosSRotrSziCL8lqXIp1ztY3TlEQmHzAV8QOQwQt1QecI+AeVVlg2WDQgBoxDSOJ0BL++mjfy9YcTqbmP95I0YNuuaZ+omophjqf0+6o+B7WfUmOtFFMFVYzQSiPZ8CoBEoELLrgAnv3sZ8Ott97qVMImgobyyPD8oIvrOaVrl09omjHzQ0tcYEY7aF9Cy8RTC1TX5nQkgmuVk4GPRBk9AdIgUUf3L1ZS4QGndZRU+PFwMJ4RqAyJxDMQ64/SgbkvgSDqiWYxO9HaCOxxu9YG2Qiq1JhUO4NgCGKARGBurgPf//734eijnw2/+93vnHowjjoAcdDYRm+AVEtSMi+MYlp9yjtA7aciDrMhxEt6WX7PjrTxh6UJRwBIt0nn8kvjCBMJSSh04xJkFBXFD/jiCUTVdGRUAzCuDQAquJnkTKnQcfotiEHpWIbhzNRuE1CF0AhVNNg+67DPkBpL1fikdDEELC97SeDCCy+C5z3vebBhw10uA270qMF+jUHQGic80a50Zp6yTmU5ly62Lc/b+4iYyFiClNReVOGDCiCr92pdoUxf95SNjsjQRgkpIkFljrRRw4pVI8OwbVQhVhVHJBsBHU/dl4l1TOqOsyUMN8RkH02AVSkHIXRamBvcRwrqiPqW2AwBTh2Y68B3v/tdOPHEF8HS0nJJXHw16KzWOsCBcroNsh/odT7ImEfXcLBd6iVITk/4bCc3pbtFN6BuUCJoOWcmNcF2SNzd6l16IPo9ELUcl6WE9odBsDIub4/Jc27KEjYAS9Pc8eAiDBdNTQEYlsjYNVU10EEPYVXgum1YqWTMyeoGSeBLX/oSnHzyybHK8zBEIFNBahLBWpWTru1bzPXpwQivyqsUaTezjujzx4oU0sYtFHU7bsMfc25AyQa0QS79oGWWdmlcSMdRkx9TnpOERhpaRoN2u+MMP43qw4N+2/uLwg1/EQbKiSYQ5qNd697tIM7cVN9W2HQDGLGtEiDD4FlnnQVvOOUU5xkYLWKwZdyCZdS/mqOn7GpF70G8ooRzp6R0LfJrKFNhci8BcIECqftbMb8KpBivM6vkZLFeJI0dKZ2lPP5gECCVH3nVWLG0LhLYgB/ZnvxrzhkhqDk9fxTIxuD040BeQ/KCZojAh/7lX+Dd7/6/oa7AcLkDZclC8kPXlTFIHVHLwThV7nQfnpwyvPT9zl1OwqQY1bJSZ3o87G4PXgDujA7XMUzQPTrNkAbpgxA41VJPChMIKnWsKeDouQGYItrtNpArnlaZPLj5DSpQlGACg3c/w1zgNUninHi2cTWfUSAh1TQWlTmiBFJqTxkTKKPwLW95K/zbv33VZRJ2x64wVKRaVfNn7QIcrNNzYcCp9r3dgs4l1AW0DfTDhjJAG4FKqUATBm1LcIFAvhUlgpdsdlCeyCOpoNz00FOiIvJLL0JK9hv9rfdEOvDYazkOLUvox0HiicSM0qP5nLIT5GUydw1pY1qSgarOXO/6CkYz9liabJfWLu6UhPaAK6+83EkC47gHLT4MIp7FUmOC6SVvJSlcqxcyackjNF4lNzyRdjktYvEQ/DUtXZwAAdNpi5FDxUCE9GDpWvmA2t0ns5QgUK1ictCwxpoYYuxeaB1LVjUocmQ5VAFZcaMMcw01YlW0FKLba6YFCSQbKRarKVWgTN1NzvnwgMiO26Tdcsutru7gPfdsjMdHAc4b0HsiVofAy3WtmUcx8Yi4d+wxUGmUGALmkw2h4P7je1ySUVStpRSRkReABh308jjbjLSDkIoezPo3rQW0WEuQsMRaOEcL620qCEhK6nmw5BcXofghKas1pFXdl4KGDWFjqzzD3DsJaaUZ4TCqAsj5L774J/DWt741GAXHUQUYd6qWq649AAmOQNzdTiZzEYqgdRK0K9IjbXekihYLhLoko5JJNDYATZ10WmS8KmmwSAcH6Qe0IpAGIgLjRwM2Bm7ejR4foTihnmhEmaxC18+3PMKnYFwum9LmxhmHlHxtH/K6Ee0BH/zgh+DLX/4izM3Nj0UE5NZdgxmQlAS8jYwfwiYbeXGf7Wj+GJvuUrYym0vgVe70NX1yA9KD1DG+eYo0iXx+6zWYCRAW0yJiSpEhnI/TIq6vEvm3FLKnYJzxWAGxqVco1aiG2mUDXA6nnPJG+P3vbnVIPF64cL3itzLNnjwBWgLW7fhMXYkTss6Au0JNShF1SPSXx3rh45KBWBe3UI2HoxEB9g6wtJFC+i1CBBTxFRPrvoZ5spy9FGFIR6uB8PT+t2awGtyohCSrMT9aMh7dHtBpw69/fR28613vaiB9uB4D1R6v4jZ5HjG1VJDnul6HLvghK3XbvtPStwSxMUi5FaueaBPuHJgiLO0N1sAxDox5f4FLh6Ady30olsca+uiAEl2NjSAl0sb+YOuGOsiYeu6kjaTkvpSUUdZmDej3+k4VOPXDH4Hvf+87zjbQ602oUlUAHY8vawISDhYnsjzwiJhpbD1hQ0gxdtoLQZ2NBa+M5ZDCEJOPk3AL+uN+SyXc0slu62R1/IZWfmhSRwGOQRSceTdt0nD6mOPwdMxw+riGDAERh/8ooUz6qXN9ytBoCfGQrzt3Lm8fO/KWt/4fV0gktaabBSqlVxxLygbmPXLaek/eB/+PXmDayEjZi3QdBeJx62prMNVY/Fo1s5H9qWOewtB5MmwQx6fEBX++qblmpmu47qiQGph4R4Locl/YdyrLk677Y0X8QWDnxur7khDLZWrVDnnPEPUF27it/I9/DGd95sxYfrxJSMXO6NLkZRweGY1cTGZhFaIEU8Z4iaOCsIRUYVEWPPjsRYiuThyw8mtVDDOLGRwLoCMGOWjI99oUyP3la6+FKvqW+mu/kwEwVYikifTkPxaQazVFFMqWyZjTmwf3NVYZxk1HcA01KQVo5NZh9xwvkwrD7wgcSlUE6gg8srU82ol2uUYA4aXCXo+LTW3jZSeQWSMTDW+d5PDELQSVFnlZ7TP8jl9Nam9BGCIj4jbWX/s91Dkmz1Vx/ppqQT9sLf8///Nr+Jd/+ZcQG7AMk4B0SbwUEXDfKhcp0xXJXO0x2z4F3XkiYeIAbKGC1Azy4NiPWQQbkBD1FvWAfKwpqz8Wfhh6e3DxmLref4V8SoY/+ps0VPl7kzUhtwkFzYGVGspsBwOIQJZlcOqpp8JNN90ErdZ4YcJVkFrrRY+YLuBBXFvr8ISzgzIJrS6FEo5/tngnheh6ClFn1rKhKv/aNF/5fYv7/dXjCmt+Xk9SiESjiuNE7We6z0oLCrkafdptX0XXfuQ1VF57i7+bQZCyExAMIXjlWB3HhQnfAh/5yEfGdgvqtlPbd6eAq24xasotvov3xeCgWgFBvj05lmgDKIYoVsnFMhKpjPoQWBciWSfJFQKNgi/6MMKLi49d8nJSXDwYYVK3OM1fGqtCG7EkyAS0AiqP7T7BGxKTYELlXL/Xnt9vz37kNW6LNSobFdrttDsuln6miEJdO8EQtoDTTjsNfv/738eNRsYfYJUbnXNn0h6IrGJ/DpISivUNNNGxxIcTjDq6yVaJb75oxZScnY7xeRIx7NthHYe9BPgAzS0qVwaacGyE8mDRFUrDTVig4/NL3d59RZsAuwdTqlyTBI84OwIhsIVVq1bC2rVrYd26tbBy1SqYn5+Duc4cdNptpyqRKwznDSvoYOmspeUluGfjJrjrrjthw4a7nZvMAkkIRFzuDdBHW0C75VSA8847D171qlfFWgLDghTdEbSxW1rjC3cmturz7RWJgG1f9m1FWMmJ+OKORsiUzGtBZu6RZTJxVSQCfB/rMnIglk2OBoS4I1HslJlDIr9EZCISvpyqeYOiWvCEjJoW6aXLaq+97g+PeOQj4SEP2R/2228/uP/97w+77rIrrFu/DlatWgUrVyyExVyWb+Elg+VuFxY3L8Ldd9/trOI33HgjXHvtNXD55VfApZf+Aq688kq44/bbzW5MLeghISjUrdvKIPd/Pv7xj8PJJ7/E5QkU61kMA7SNXj4gNZdwhqz0lPaL9xDip8rx+0hAHmMZ15cPyO1l/X43p5u9HYBz2tMPXaaLpCfIFgAtEgV2h4wDWAcAxd6LL74YHv3oRzdDACwRLdMv7fGy99AA4kuE32677dyzPvGIJ7r98Q488EBYv279wLZ8IJMP9ir6kQe/hxtuuAEuvOgi+MY3vgHf/va34dprrvEnWiuh3cmgt4xSGCbX1KmmOwSkBdLGIWtlrqYEFg55+tOfMbIUwCpuarGwgThVLIek41S5fF1vk431Mq+gSHBSulILCUDPyLF1ZllzcOnr9wMut4RZAuCp3qD76md4XXTRRfCYxzwGxgYmkgx2aqSUoEBIBnmziL+wsABHPvnJcMyxz4UnP+nJjstLIN2dhuF3NJbqWfUcF3VJrZfiHEu486474Tvf/S6cdtqn4Wtf/xZ0F++ChdVrATWx5cXFQAgmKBGkjLWpa6D+ENohGOg5z3kOfPGLX3Tq0Vh1JkskXI/IdnCkDhOueAnCvzbGM64U1A/xAJZAWP3VVhMK5/r9fu4/vbzXW85RIhj86YV7+nkdoPb50x3q/jrQ6/Xc34suukjKoMN/svBphU9W8rHn5W8Y/5NlWd5ut+Pv7bffPn/d61+XX3LJz81zd/PlpeW82/VzOmnAPnCul7vdfHl5WZ37/o8vzJ959J+58c7Pz+er1+2QQ2dVDllnjHmQ7wZT48Sc13mP9ljNuc+yLF+1alV++RWXq/XV8GxGPGC8wH78e6T3Sef5uMSjZXFOtmNx1f/uhWMe13s26qcuBx5OJyozgEzCmozUe5jYgtJrlLqWKPIhr6HvDTE6sj4jF0Ix/y//6q/gZz/7KXzg/R+Agw9+eDznLfRtV//eP/fkrfMkkaAREcViGgt+Hv+YR8MX//V8OP2MM2DPPfaAjXfdDqtWroB2Z8XYdR6cS1K4XCn7Ln0xjAV58Hzcc889cPZnz47HmgRdQVsuLJ0lK/+aFgo2BrmdX9rGxlEusV+mJkjZixSj/DMct+HrqW382xwtJQr905/+1FFvouRDcX75vWU4u/zI85LbZM1x/U6nk7/2da/Nr77m6viM3a7n9LMI+D6dVOBtSvlV/3NtfswLnpejhje/sDrvLKzNATrjSQE+Od68H8+tk9x/jPfRarXc34MOPDDfvGlzwxKrxa9e4VyeM8fn44yrNOeMq1IKQO6+FDg+/l2Kfchr8d6WTM5hK2K6Qin/Te8/XgVc/ph+NxsTQ1QTDTbWRVk9sKR3hIlm5OpmsHKfULp+DCZB8eeOkz7+8fDd734HPviBD8IDH/DA6OLDPQ+sDj4rgO/XSQVZ23kRHrjXPvC5z50N737Pe2F+3nOpzsKqIXfZ8cCRlGLxyMKYhRvGfx/9EBn4q8sug4t/crH7Ppq70+68jYBjt7k28hyXCrfhvHKLcpYe/HEuy4drJGzhF/fikNWFCQ9DVWBKOSzOWJnoIcWMmtOgxBO9RVKTewPSS6otDtvFUrD8G3lfnItfx5QOafPK+fl5+Id/+AdnWX/c4w53xMwjvo/S21pgDlUDzKvv9+Ev3/BGOPOcz8FOu2zvWHlrfuVIMnryDserTG5tg0ylHd4L7iw0jhqg627KYxqhqQ9dadufq+paB/F5Y7FODGqVuLdcNiBVFI3NlXQz3irnh+EByrTIsqzCqYHi9gHcb4H4ZdLAGOCLUPRgr732gn//t3+Ht7zlLdEKjee2JsSXgAE1nawFi8vL8JynHwXnnHsu7LDDOugvL0GrPT9CixhZ6b0rPv5K68kR5BpviJl861vfcsFQo0lfwwxGSt9aEqizWY7n/J2Ena2sf4X58sJBNfuJKg1bPKHKmIGQDmscCUZpwkpkWusRx0ShjxGHii8GFxRy+ac85SluO+sjn3xk3KyiaVF/i0TrZRksdDqwuLQET3jMY+Dscz4La9etgX5veWjC5iMWi6pdcq2MKf7btq+44gq49tprx1ADhgNZ4UdK5qX2auHzL5d8We2QRCG+BVu2m0WQdiHHuM42SGVgub3Wj/QYhmvY/2mFv1zdqN59pb8bXFCxizBvyOVf9cpXwle+8lW4733vGwNOmrTme7uCT3fdEoDThqoNVts58vDD4TNnfQbmF+bcuWnlFIzaTZ7njhBv3rwZfvSjH8VjkwOtCnAZ8JStQNwVEdov1NQY2eYmcSyjqsD6oA7VTexi0tBD+rb1Huqj9kMOjla7M/yLt8a/8uGOTQwI+ZGT/O3f/i18+CMfcfH5+Hu0aLNyIAkNs//+679/CZs3b4JpA7nvkAgggXv2Uc+A977vvTEHf2QY4h00gbPf+ua3pkK0uJoWc3Rd+DMNvPFO2X6eKQKi3IDaxbC1AbkBf/7zn+et1ghuQOv+8xaSYsDPmG4+ci+9733/7MaLwTSTCDLBNuldvvV/vzk/4ogn5N3uln2/2PPi4qL7fvJLX+rdbSLYaeRPcA8OfN8jvL9WeF/77LNPfvfdd/vnmNIcMj6yC7A4o9K9Tq7K+gF6pibgvQVGVCGs2YOmpQG1T3L+U0/9MLzhDadEw1LT4jlZknGzi1e+8hXwD3//j7Awv9JJAluy8hJOKW7KiXPwvve+F/bbfz9Xk2/U55cVccpE31IDYV0nUe7bvO666+Cyyy6fsj0lK7jviuMjD5oO+knNRcpmJ8tbGqe23dNPDmJ2iQUuqJFfkH2shh6VoudwXB/4wAfg1a9+VUT+pkVKygXA9k844Xj42Mc+7vpesWIUy7uG4j6SwwMRwXXr1sE/veefxnr+QpmJWjcN917zYAdAe83Pf/7zeGwaIJN7ykR4rzKwqm7tCIzLVN1behZiUdB4i7iRihSkDHezSwB6IlOqtiGwBJpCTVpAb3vb2+F1r3udQ85JuPgoNx/f28tf/jI499zzYGHFQjg2fl826WtUoF15j3r6UfDkpzw5xjqMOKixxlIHCA8uucQTgOnVQ9EW+6rr9PU0ZmLePthP1xokLZcO5fqErOJbFB2stTIfcnOQ5iG6iMLW4OFgPShzrzQwLnL1HX/8cfCOd/yt+z4Jzk9RhIhcf/M3fwNnnHGm+45FPpoVW5txstOa+au//EtVYGaEhqoNt6mhDjn8PIztsssui/kX0wFpCGTX+8C7ogeBg43kOQ/BC0AdaVcCXUAKMO1gQkVAJYEoylNbsmTUSH3nNdf7kGuU0ngPedQh8NGPftR9p3p7TQMhP0atvfOd74wBRvF8QwSgqTqBFGX3lKc8FR73uMPGkwLi4MLfUVSDCiCkuf766+GuuzaMR7DGILY2alCMUDFc/474es3cdUSuDwVWZbnlLEqqw79lx2UuQt/x9KSAsWvva8I4NpDOv+uuu8JnP/dZWLNmbTzeNJCe+pvf/AZe85rXRD1bLtL5hfFtAE0DIf1rX/ta93tYpEoiQ513NyTu5mFcWCcQKyTJY9OGgUFQ6bviX6sGtIpU3YsNUhqQpYaYmtgCAzwoEkGqqgRNijhI8WxoTmW1HImrQ7xvmk9E9k9+8pPwwAc+MIr+kwDS+0855ZRQ0rpY0Rb98A5myHxD83HUUc+EffbZe+jYgAICSGG0QSE0D5WyMD345pt/m+57SpAu7jIofsbaCJjJt2gnoKKIYFmiJwIyg6lMFNJRR4MGNQEYp+nUHqkwmuj/tr9+GzzjGeOUlKpfCencc8+F888/v3RrK+5/digARUOuWbMGjjvuuHisEWj0MfOYbfebG2+ErRN0tBvhLW6jMZAVauqiiYMlAukURzOUCdabd5wvDGfkAqFS48lGQ/4jjjgC/s9f/3VE0EkA+ftRLH3zm99UrZtS5mVzvTfSCo35Wc96dvQONA4NqXRUD3HWPWESuGyYtdPhvy0Ucq3Ob0UJ3UCxKum9AFKhwKPEEgWitv3227uqsnNznYkTO1yY//iP/+j2ua/azAL3SxgXWL1rLn2b5ueggw5y9Q2pGs/4DTd0jYFbbr1lixu664E10Ke4masHkAo0kH+Jq7DYb4nELEzG2BSZHn+MylWEgO94xzuc3k9W/0kAtX3VVVfCqaf6/eyquGevS2rBeO+KJDxvOB6fEJAagGXLH3XIIaKPMcE6s0oHAEOtr5tvvtnf1qCmMil2Wo4SjO+BAOSFXUs1Jlikb8hUPiuQigB0kNV+Q6R7P+pRj4JXvOIVE0X+OLosg7//h7+HjRvvGeiaWu75eIBxXpuuRNPcs9G4Dz3kUGgSKgmJjRXIEnwwAb//3e8TDSRk5rjpLYcqKJ5MdUzcJ3du2n4DaoU0wuuHkwuZR9LRSG51eSnyEweh77NZmmpoKHo2w7H6yE/Ih0TgPe95j7O4IwGYlGRExOXSX1wCZ3/ubLch6iDdOW9Yt64XoTZMWwAHHXSg+9vcnnwDWH8WL6zd5j2bfEZllYerRQVkq8SQTH9tB4Lap/tHBFkdSNcVSEMgAD5wQFINHTyQGnkqLmDLw9iLZ4RFQYa/448/3m3QMUnDn9SbkdgsLi6FoJqi5V9Cn6rQwuwBLdp99t4bVq5cCZs2bZpcsE2W8O8NAd2wTRoF2NDtcjOd/7zzHjjzziXY2JvH5JTI6rO8D1iIu9/LYCnvw1Ivh8UcoL3cg/07XXjjA7eDPVcvOGkA92AcDwiny+bQD1z4pqyRrxzBR5i32Qf7uBT7MGANUtDN2rVr4O1vf7vZCWlyhr9f/vIXcN7nz6tdpWYcrjJpiMbTHbaH9dutdwRgYiC5YlZC6FOHwnXL3WWx/kWwnCtQDPDjOzbCq6/bAK37rYVd55BIdCDv+52G8qwHvTzDLBtoE4fOu7Cin8MPN/bgv37y/+D0R+4Ne6xZ0TB7lVv4SSnfeQHC8N3mg3IG5CykXAjDlgObziJSnLfOok+pRvJczWq+r3zlK2HfffeduO5Pc/7uf3o3bN68GPsfBHG3YJg8jLouVq9eAztsv4P7PjYRtbdb+3YK8sET5PdjSBFcTMTqwfuvvB3m1q2BE3ZowQNW9GDNihxWrwJYWJ1Be3UHFtbMwYo1HVhY3YLV61qw43Zt2HnHFvzZ3rvAzruvg/deegP0sU5H2LR1fEir+IS/HeZYtg5gGQWZXU5STBEdcQrd+x18L3HfnXbaCf7XG/9Xcy6sAaGzl19+OXzh/C8OV6NuihLAqMiL256hN2Cyht2SY1kNj0FJW86EnmWwYakPv12eg106HbhtKYdzbl8Jt3f9htHI9bVE2Q4CRObQa4d1OTxrj/Xw819ugmsWe7DHXAtQDsB2RwtzlzIERebqc/ieOvyy0KWTVSB/udFjJtyAMfpnDBtA3QVgdP8Xv/jFsNvuu09c9yfO+tGPfdSFpZZF/VVa8GH2gPR9/LsQQpYntqbsUs5LvldIUritehEyL963AeayHBZ7LdgU1H8fixPEi4BjUYPAC/oZ3J0D3L5dG5byNty4KQcsnr6+04eVuAtTKKs21GOGlH4dwYvjlhvC9qUNQN5sGQa5CfNG88IbhTCUsVwpQ9xK3Be37sJklkkTQvIyYNTfOeecE48NM153D8w2TGV8g6SBivnzRVtTVwB0c4BeBtDNMPsS68gFZ0PYbRj66HX3UkDszsn7ANlSDpvQFdgC2LwZ4K52DvOOQ6OsMMraYtG/DF+dPUnfEKyVsS9OBU4l/0wyym10mM54SPd+0Yte5CLYJq37E6fHeP/f3vzbgYE/Fvq4UceMgiSeWCyFjs0izM37isYpgt+DHNpYlCUHWMqEDi+fpS8+3UA1lnLoLWewYbkNeasDS/0+bOqiFAGw3MdLhre3cTFRja8eWrFOgFixgzrIZhThyzddnBQw91/vavtNQw2iwiKf+MQn4hiGAr+bxkzLAFg6fNM998Asw8LCgvubMtFlGNCTobbfh26OgT3+uPKok2ROf3vOQAD9bgZ39zLo97CwCyK+VymcVOEV9EbG76UBtucNwbK2jui/YSqnjM/9T4K9996nmWIWFUBBRbh5yE9/+tMYPjsckO45/DucFje+++4NcNvtt0+1z2EB4xQcJIaHc9t2yJ3Dcp5Br5WILUEbgEP88HEqQAZ5N3dxAUg8MG/D5bT1bCZOU7UceQ0YAuBFgzRw3fHhdwTaQm7ACRAs4v7r1yP3f0Pk/pOcD2r79NNPH9nTMN74JvuuZcGNO+7wBGBWYbv1693fZNXdoFcjU++SjBA167ChKd1HRCASgwyWe4hXGGeQQw+DhvBGjB9oZP5lMVCGVlo5KQMrw8wO0AvRWW/NjxEJDPb1ghe8APbee++o+08y4w8NT7++7jr48gUXxGPTJACT3reRS27d6GIbpldyqz7Q+91p550rLsqhj9JZjvsY9kWYIF3gxf0oBRBhQCRHXb+bQdbvOXrQdVntfXCF+8dcW3KjUA+8fsLOQO5rye0jRMdsQYghrxNASBK9V6xY4ar7TkP3J2Q/++zPwV133hkJ0LDQdC5Ak0DPc/kVvu7+LG+IunMgAMlCOIA6e8t5AjKcb5cR5M94dx9xfC/2kweAiIHbFhI9AeQdpDB92op+LKDyfViQlovStiDrD4nszQYpNg0q1mFCuv+znvVMOOCAA6JrblIg96b7zGc+E49tNZuDDgmXhrLbswh5YCz3u+/93N8U4XfCerflcRw3FJKGPkR6X53bE4ZIDML3HkDm2D7VcswZ+R0BGFcKKN7vytb5hsui/Gx+JAUTzLYUMCmg2ntYeBNh0mIq9ff1r30NfvXfvxra9ScBNwidVUAihx6An/zkp+73rIn/CLS3wv332gvKAEe9ud+DZUzwoWq8DtGNBBA1aU0kltFzgFImfqDlbnNxBLHE3uhQNqetougvkd1a/on7N5sPPgloKpKagETvww47DB7/+Mln/Emq/ZGPfVT9HgWmvU/DsETu6quvhquuumomCUDm6gGCq+y85x67x2OF65yhDbP92o4IeMO7N+RFQiBNbXng/FEd8ATBB7MJx59L6xv3GfTaiTUBaehpsCqBTzHkc7MIkxkXTRhG/Y2qh4+S9XfJJZe4nWlHc/0xzOzbCvP4ne98BxYXF6cyt6Miz2677Q477+JtAElanAEsZ33YBD3YiD8IsaOxT4j90cFPx5Hr55BjigC6AcOep/0YvzGZZ6pRqtbq/ORn78+UJBBj3SdgQCLRe//994ejjz564ro/AiHBpz/96bCPIG70Ear6jACzalijcX3pS/86G2soi7/iWsct5xEl93nAA2HFqlWw3EPizC49whDP6OdgqZ/DZunmkx86BsYb0McgIi9FLLvvmBeQOUkCVQPMB3D7FFOHpQ/hDHu1VYYKApCyC8jNPmZrQRHCjL1BSAVxed3rXusCQaaR9OPj/n/vyn37Y+PX3ps1wPRZZCCXXf4r+OGPfuTe3pYxVmaAdMi5kKXLLoCP68/hYfs/zEfKtNNBcSsX0Pi3jKE8cI8L9e1DlmMeACUKJAgBqQnoOswANi62odXpwNqVLVi3ogXrOwBri8MdOK+45UedGhCCAFhjoBT/rQRArrZxiECZ4XE8aHoBkei98867wDHHHDvxlF+5zdcXvnC+K0I5jvFvlgHXPXJSDHDavGnzUNmNzUHm1iJO74odHgxrt9sFsmwOoL0AedaBboZW/RWwqd2DTQc8Ei648Xb43VLXudIQyVx+D7Sg3c/g15t7cH13Bew9l8EdrtlWEPfRA2DcgI4AUFyAzwvIewvQ2yWDsy+/E/bddR7WdJZhddaHuRYWEEHWnrvEoLa7t+8IC2+NmMPaVhsessM8HH6/7YI9YrCbusMpgngxcjUZ5Vdm8Buf+1Gl4aYZE8UB6RJnowOl/B577DHOBzwN4x/2iWL/xz9+WszBGBcmPeZhgQgpErhPnf4pd2z6RC4L9rU2vO4tfwcveM3JcNu6ObgdE3ncmm+5pJ7N/S5sxJLgi304555F2JBlsKnbdaW78n4LepgKjAicteGwB8/B3NrNcNEd6wCWgo6P3L/LEX8Or/BZgw0Adf6sn8NdN3dgtwN3haVNd8Ev79wM2TKK/T3IXWJR7lKN3a0Ya9D3+QJLfYDFft+pDYvdLiz9ehMcfflt8MEj94JVaE9xNQXKGVZHijq+KhDn/+vEgYanfkK2g2xCkXgnn3wyTAOIwHz/+9+Dn/70J2Mb/2xFoFkBknI+/OFT4fe/v22LcP922xP3E1/8Enjb378Z/m/eh28td2EDIilKJ31f5quNqbnQclV9WutbMO+4ex96rR60+m2nc7f6XWh3u/D7TQC/vn0BNt7ZcVzfG/SDBOACgPy90QAYcgIclt2dwZWXrYa79pmHXfZowcqsBR3sI8tc0VBULTA4yBUUw2OYnejG2oMVvT6s7Pdh3UIGF/z8Bjjo0lvgfx2yJ2zCwLU21iOU9o1SFYCIgE6oQWvkLAf/SGi1Kfd5/LZoUR555JPgEY94xMSTfhBo7k877TT3DLRIxwVsZ1YAYxJwHq+/7tdw6odxT4Mto/vnrssMjj7+BfCxpRw+dPtm2LxxwXNscsu51+GS8kOhjCC2O5G87ZHaXdMBWMZPQGoQbZDI7x6ejksVgOqGZwB3teDmS1fAzYiZaPWLYbx4PeobLR4f3odWwi6Oo4U5yDC/pge7PnQn+PFt98DdvS4s5y2XojwXDIxJAuB1BS4E4isDkRSwdSB+DG1sUD8nAvgXf/FanwSEiDjhkl84/t/85kb4yle+Eo81AeQdmQUXm2cyHXjb298Gt/3+jlpVjRsHh5t9gGwOeuvWwv8sZTAHbVheBOgveqR2XJNCeV0Frwyy5VjUJzxMwBVp0HP6N7v3PNKbWIBQCcgfD1mCpL7iF1QfYjZRIBwIuMFLlCSchdITF6ditGBpUwa/uWkebsmW4XeLfS9FYCudHOYSUkAoCeaLA7BxD/+dHY4xDBDBHldzobDfhz70ofC0pz3NW+YntMEnASUWnXPOufCHP9zZqFiMVWlmSfQ/++zPwhlnfMbt5jx9w5+Adgda8z71dhnTb50xztfp4zz+gLyRORKqCB8/lfmx3J2MfrYOgLMBBK7rVAJ0Q4RmqA88LqMFbRgxXocEIJYiciQK+pv7sLmVwyYkIq0cOlnuVIh2C82Veh10iCGwUW42Fsqo4Kh63lza76tf/WpXBGKSO/wSIMIvLy/BmWee2Ti3noX3ilWJ8BmvvfZaOOWUN4asvy3t3ViAXqvjKu/E+XacVSCeO+ZEAkZ0h7hiYy25kUwB2WOusNg2SPaDCCzUgVg6LBgHIkEQyO/aEhGEzsYQxtTLoLfch8VuDnMd9EaiHSmHbpbDnGP2mTUCTsYiv0XSgRuIeSfL/+677w7Pf/7zphL4Q8a/733vR/CLX/xiuIq/NWAS8RHDgHuWDOvdbXZl1G699dYt5PaTkAFgwE2WwaIbIxXrCIQgBu3IXaKCGO0C9bnIp4OoAhDCB/FfcvJoANRGwFgrwCE5iiKhQ+onFhVFBA9MziE+ZRf2wzlUEfrQQ3Wgl0G3lUMXkb+VQaefQauNpcBZFQgsbSvGfANsuBw9p5wo5Ete8hLYfvsdpsL9ifieccYZgeCMF/lXbH/LvWMiZEhYX/f618J//ud/zgDyB3DVe3Isywc9RCgUp9GQl0nkNCHxkvs7Q6HM+xcqAun2ZANwhMRFHAWxPTx/8OtHx75KGjIbDEabAakbFGDkHiD+7eXLzk3ZQ2mg7Ymb29DPDcmrCgidWRIRZyEJiDjv6tWr4aSTTppK4I/vow233PJbuOCCL7tjTRvFtlQgEUk2+Iyv+YtXwyc+fhq0O22xW/GWhMx9FpEAIG7iFOEXRwBM5B7l9BPXjwgZbrRhxBSQ4hBbiu4k4gcDHuUK0G6isWCo3V1UVBSK14d2cC5du+IajAbEP668GFYY8kQA3YfSFig2B72XQN5M8U00/D3oQQ+aSuAPGcbOO+98uO222yfCHQnhpknoSXL6wx/ugFe9+lVwztnn+mebCeRHQATsuUAclLVcQM5SQD6cJlegwyCoiscnohCucVMr9Hu6WBruJEFBQuAQnTg9GQZD8FD0DtCHmgzEPBoT6b7QDhZWxgCiQPj9LkMUvkzfPYiNQe4dMO7jEKfEzT6mBShhdHtdOPPMMybWxzSNbeTNQOT/+c9/Bi992Uvh0kt+MTtiv4QMkb8HSzgsRH50Abpc3kAAIlcOBAAlA0JIBMeFA8FwH3mOrnUpftrvTxWDYokwROLQR8wbQJ0eGyKjILkiSD0IO4+QMZBsA0sYLuTPOROCe0zfVx4qgBBMVrHdAkCPN4oq4P3Rfef6+5Pg+mtNmPuT7/8nF/0EfvrTnzUW+WdhGv5/fBYymOL3D3zwA/D2t70d7rrrLkcMUCKYLWg5BEbdH20APr42B2cRBCIA5KtnDsqBO2JOKVgO/7gAWuESiK5EEvPJsEcGQUEMKFHISR/BvedsA4GASykkSid0reftOW4oMB+KiWL8MHkTMJyYCvq6sP97IQEgxHfJECMQAZycl7/iFW6LqmkY/wgxP3/uuVEVmASiTCoQCNujoh5kK/nhD38Af/M3/x98+9vfZgln5pAfwSMCJvxscrt5BAPg5oDQXnEWujc56vl2h1wq+ARjCNAmYNUF56ISonzg+KI0eIwYJKMhGfXyEhejVD9cG9gvEjHMZeqwm1W4KZn7+3vvfQQgH8/1t+eee8Lxxx8/Fdcfcf877/wDfP68z8djk4Amn4WQntqltjF34Z//+Z/h7HPOcXo+SQIzncnYB7h7MXP7+DkDIO5Mjhk2Tl7GC4Qxj0KDo74fxH3i7Mj6s/BbSgEYZBeRNxAUrP9H7j8iAC6aj3IGgpEvIjsZFqWxL4yTwoLxfgxP7rdgDt8LJhFRKAFKZ5k3+TlDYJCVaxOAWQghrYJx9ysgW8iJJ54IO2y//VSMf5Ro9PnPfx6uv/6GierImFAy3ByFRBYxt66dVst9aG7uvvtu+MEPvg8f/8Qn4IILLoDusuf0M6nvS3BInKO8DP99Uw5/wHL/S22ATt9LAMQryTVHxjkSwaPrj+ICRHQe+tpwhxCX70yIStWBZByA2RvAxfSTaC9UB1IvkLsjUSCnFI2NpIB2DvlKf/lOKwA6SHdCgFOO94Z2HA0LqcK1CcCsGwtJVO/M8d5tw7r+VqxYgBee+MJ4bNJAojEm/kwa0PXmXnhtlaY8MOz222+DSy+9FL761X+Dr/7bV+DKK67ifgLXn2nkJ8gQQzbAj8/+Edxz7KEACz2A+/Ugw/r8bR8228K8hbwFreWwxTbe5mOaHEfOlsNmORh/g5t7LGfQvwurgxKBkeHEYmcg+yGdf90ytNdg7H7bITSq6piN2G7hNl5IBHruGJYowSQ9t/FoEEnarT5sXu7A3T/fAIc+PHNpw23o+g3AcJ/AIElI+lVrNeADbtq0ySNVLFrIEU9lujbr4byYPDfxE0L3oZ+SSyGHe3HjQqxj7h5QlioNiRo+ZSJUTM0cIs3NzcE9GzcOXQ+AxP8nP/mp8JD9HzI11x/2cdFFF8HFFzeX9lsGuOccbimO1Xc77SAKCs4upYNud9nVI8BNOjZs2AB33HEH3HTTTXDV1VfBJZdeApf8/BK44fobYtuYzYfva6tBfIK+D+K5+QvvhhUbHgitRz0d+is7kK0C6C9hEbAlgIWOlwq4preI0AviPNCpHrQW+wB74rE2wD1zPqMvBgPhhWRUpO/B7Yiuu+0XIdshh94fnA81qhLLuJhdXrKnWfjBn1nWhxZuH4ynkBjM5XDPDb+DP73fZnjcg+8Lmzb2Yd18IBQBd3y4QiIQKAUkJiByHXPsMXDVlVd5P3kPfaeeDMaFQ41mJRwkZFdhrH7USQLik/WYPmRQoo+q+ef2Sud+aMNSDAF2hTS6y0NHAdK1r3zly9XvacBnP3tWVAUmYSgjhPzyBV+Gix92sevD7WQE9C6ICOP78O8Bkd99ul1H+Emsl0DviO0BWxHiR+gHnnMzbP6Pl0J2ydMh2+VgyOe3h8Oe9TjYcb/7wMaN3p+eI4vHsp0U1ehy6JBJ4I6+GF/fhw0bNsIvbswh230OYO089O6a9xgWi4OSy46Cg2S0IUB7NUB2xQbYcbEP69dlMI+JPBjCi9wfy5CRGN/PXPlwt4WYIyxd6PeWYd1CDofstwBHPvx+0F9E3Om6+3CsnmAw8hMByPKK1S4JwH777wfXXnMt3NuAym2h6+/iiy+G+fn5iasANOUYIHPQQQfBb35z00yX/XJE1vmPuU7ErI51NMgcYlEB3vWrd4RzL/ohbLfXfrB5U4830nK2Pk8yZZQuVgRCRnz74l3w8k/eBHfuvh3MrV0NyzcteHtARHraHCT476PVHwOC2tDe5x7Y47Zb4KQn7wY7zK+CVXMAK+YyVxKslfniJJSu76sEhShSPJf3YUULXMrv8uYurGjljiCsWchgzYo2rJlvw4q5jrMLzImt7GqpACjerVhYETnyUC9fyu+J37H0QIIOaYmiZndDGgKl8Q+z/qYZ+ffFf/3i1JC/7tbuyU0vnY4720bg8SB3CxGLcfa6y3DsSSfAfffZD66/bRk6VEiF7G20yaZI6vPG/Azme/Mwv9EXBUGViykFxeiLIKLo5qNCH4i8LVi3CmAdtKG92HeaBybwINI65A/bewVvvxsURb1g0Q/8r5flsBIJRzuHhbkMVrTbsAKDsgLSO8lFZwPWmyBJ9YderHbt1FxLbponuO5I78aEn+OOe4HrcRrlsynt9+Mf+0Qcx0hAt9WYo1nc0XmWIEOc7HVh3br18LJXvdKV/dppVSjIE1T/GGMS/PKuvJ877UX0e5Z60JpbBOiscWXEo+GP6gHGTD7pv+coQ0Rt7G6hlcHaOYD1K8BJAfgb/6OSDk7ldTeiYQ/VOa/nY2MoDcxlGSy0kRDksKLTgvk21gLAsmLBXiOeu6YEMHu7tTYBFPf/nOc8B/bYY0+XfYdZeJMEkjB++MMfwoUXXuheyFCGs9l2xmy10AqG4Oc971h45EEPgZvvWYZdV7VD5Jxf/74eH6cAyzog7U4Gd/YyaJPjPUbpydr/Qv8n/38MBgpRlHkOK1s5rJ4D2G6hBavnWrCig5YGVMMCDlIRH2/fD/TExyGgdwBFfCz+Md/G7xm0Oy1nrnAEwBYE+WNedxTBduKJL4wFUSYNREjPOstv9ukrMQ1BAGiY+b34xUwZMuSouc8APeUNp7i4+fUoAjjs8sU4vd7tub3byivo4Mi9nZe/ncMSctc+JhRgdl7Q71sCyZ1TnmoEiByD4Pvvu3JEXWf8QxF+VSd3kgDaAVArxf0GvRE8jNuhs9czcvG7FWpAtoN3oOWMiel9AioJQLzcxTNsnSXCysAHqvTh4IMfBocffng8NkmgoqI33ngjnH/+F+OxoUEKY7MmmDVQjm1Lcf+XnPxSOPCAA2GxuwyrcTeg6Er2aIb+ACeMt7y+7U5h7A7aD1oZFvX15bycqC8Ci6haDwUDkXQQtwn3EX+4ryCSG+Tacy2A+TYiP8Aq9ER20NDoa3c6cIQgSBIhTcGTBu+e9F5DfzESDkc8yoqCDmI23vBwb4QcXnTSSS5+oNftTrzmH7n7zj//PPjDH/4wmWg5+47zKSO9FKS2gkXTCgbY+9znPvC/3/oWJ6F1Wu3ILb2Y75ENt+fyyYBUQZeM+SHBNsO8e78ZSI41xtx237IgqIznp5BfvM4TBJeOjP20EGlbTpSfb2WwgMjf8uW9HBsmW4AboDcMOjsSheaERB8v7nP0XwrqBQI5I+C9x+1Dxr+ddtoZjj/uuKlk/UXjX3cZzjzTi/8Tg+loM8U+tkJ1JAtRoO9617tgt912K3iBYp4PIRE5pmJcSpCQQ8qtq0iH3L5DXJ/Ef1FnkMR/WRkYUb/vo/WQ07daPsnIuV8xgrPt4wHUXqHE/EksSM5/9UupLddvBcS8NtALfuELT4Bdd93VvfRJh/5SHxddeKHb8beRyD9y6ZDkB4YTT+qlicxY1Y/8PUzfW4hwtIME9sxnPtNt/DKqC5jjI9CW4GKCmcO7XYFC0Q5VB4COezcg7gxEwVgucCcgp7f8h0g+CkV2HJ8i+/TkDfvK712KfQ0gxMOAH9rtZzp5Dlnc7XesDUZi4Qn6LV45W4cmA6pfGH/usi3P+XfccSf4wAfe30ibzpuHLj/aCmxZJPvQcVnDz6gHiPwoRXSwhLf7D+MDyNMvibyXJLIGprS2G3CLQ0N6JRl8jjzySBeFN43dfnwfLbjppt/A+V/4gjs2sZj5lIegKb1ctitc2SpF1nopBkFL1M+bItA6eN/73gN7771PIwFgjoMvdX1RkTaVFguGOqouFJOCqLJwUA2whgA6GjvzPsI4w/RdzOLzeRZRyh8RF8s2Cp39egCDyFw+mhuOuP+0KuXggvvsZz8Ld9zud8JpjACQuD0tGl3oZ8hgLWsw3AJA83/iiSfCiSe+qLHCL46Zu7z+figu2gdAXd6F/JLFMET+BeNiDAYKhKINnRB2TRZIUgBkJb/moHY24BYDu7jHcDORxReLfR511FFTq/iLC27jxo0x7Xes+cwTwVkpY5zVw1NxA0281mHbLUSF5luE8++7777w/n9+f6NrAJOCWmiqx8o/Xb9Bh5MGSA2IAUJiEbuAoCDq530XtYcuRRKtfLgPXTsKoQ9egBLJYfZtAFUPPORk0CRgwc9Vq1Y5yj8t49/Xv/51uOKKK5uJ+3frIx/OVjAJiMgvK9bOLpCrbNWq1fCpT30Ktt9he5eK3hQBcBb8fAVAdzXA0oIvMtqdC9V/0SYgqv5Im4BLS8ZYgDasnOu4/H8ML0IVIAIVI6gJnF5ffdPsqwApGEEKIOMfRnsdd9xx7tikdX/qF+H00z+pfo/e4IDfFqoiBkeJGZBrqioiccaIAaUv4xo49dQPwWMe85jGEr/o0dcuAKzbcRFuX25Dfpfbkyts/iFyAKh0mLwZ4wBWoxSwCNutAJjDKL4QE+BcfVT/QuxBUud5G7MB+ESEGRAWRkwqQqCX/4xnPAP23nvvqJdPEsjAeNlll8E3v/mt5rb7sqJ9nUUxTOjwMAQ2uKkq+8pnJ+/jzW9+M5x00oubLfiaIa73Ydc1K+Gw+7Xg/93ZgvZ95yHfgKoAPT+qBkQE2KeHaT2t3hJk2+eQ33IXHHzoWpce3MJiH1L8D4VvBgHH67Dh0N1ecm/1DAhKPxO7y46h/5O+/PKXT6/oB/WBm33innhjG/+GdfPVIQ6Wo9u+TOBJwehoz88gULGVY485Fv7u7/7OFTtputpz5iJ72/DWP90LLj33JvjvuzFsbwXAiiWAfMnbBBxHD8U+MbEHawW0cmjPd6F1553wnPvn8IgH7gTdxR6sWoE4R4FDIVJwwBR7Lj8cU6uVC4B0qD2FNNkk5OMTAdK7H/7wh8MRRxwxFe5Pxj+sif+5z30uHmu2kwoDXGqllK2gMtddHbVBuv7Kxti04XEE5H/qU58Kp3/q9KAKNL8LdoYbjOY92Oe+O8B5JyzAVy6/B363vAxLPaystAQtF+7rxX+XrhPy913Zu6wNO++/IzzigeshX8pgztUAxC29wW/pTdGHph5plWhfFwaQQQ4xnIa+XOg6tVhGWEA0Qa94xSt83H/YuWYaRT+++tWvwHXXXTe+8a9u1EeKS6cgRShS0YT2nmFUiFR+wBZA/ic96Ynw+c+fC6tWr3bBNpNI+c4ycFtvb+72YacdVsIxj10Nt9/Th3tQC3ARgrxDsg+tD8hPlcZ7OSwvLsNCC2BFx+fxYx0AT6xCvj/GA1RMZBUhkJvmjpALMEUCkHJhZaNLAxz3vxM897nPdcemUfSDbA4f+chH4zhmwoVaNs/jDs9KEfbdbCHO//jHHw5fOP8LsH79dhOp9pQLxMKWF3AjHizK0e3BaleuL4fNvRyWsP6lswHIaGl+QW3IYeU8lusCWNnOXAbgQrsFc0FmCNkGXHszgF5XVMpdv4QqtWAgASARY/369ZNdyCnEryO61jT+PPvZz4add955qhV/v/vd78APfvCD0eP+x7B5xPubar8OQlcR6HGfZQTkP+ywx8D5558P222/w1TeewvVixxcPj8W8kCknp/LYLGbwzJuQOrSALwq4GuCeEs/FfbA6mMu+6+dwcIcuOw/TABy9UDd/PkJ5Gkk8Z9iQlISZprz1yYAZMneaeedYaIwaHGMuHjoxaP4Py2gyf7oRz8WbQEjG/8s4lh9ugyxqnTysuvqznHKMDgjur8U+8877zzYYYedJrrFW2YQyyFy7iUB9xtF+nbmvIBdLC4SgoGcGuBuxQzAYO3HKj5ZqAUw55HfV/QNdQhEym/o3Uxy3DFE/E35Zz3BwOZqSQAIu+++e/KBpwIjcg9CPDT8HXrooVPb7Qfn6Oqrr4Ivf/nLzbj+htHlB+nbZXp+Sj1I9a+sUOIie2iKHD+K4EHae9azngmf+cxnXH0/ssVMGnJhjPMu89zl8SPyYtFfLPbhCuLTLr04QS3edNTt3AMtDAR2uf+uAlAoBoplQrwp3lsD2ZjsKxX5XIG0MQfdgnJs3k1I+0QOEQi0//77xwfdWoDG+vrXv35qdQ2p6McnTz/dbcTRWNx/yvUmLfApxKvLhQcF8VA78TpjjqYdcCg5XrY7heVC+xMg8p9wwgku5HpaFZ7LtmDzYj3WmchC0d/g0Q87A5F3j8r80X59WMWH6vhzCbAyio6I7JkLcXQfLEQc3hMHXe7OpxJHW0E+ALrdrvv7k59cnGdZ5j5CW5/8x8/L0Pe1MZ4SID/44IPzzZs35/1+P580YB/4ufV3t+a77bab67+FBdnGeXb7262qoDy2xIfmSc5Xat6q5hIZkd9GJvwN75r6SrVt+6zbV0Mfes/tVit/1zvf6d5Dr9eL63ZLQD+sA/z0wqfb7+XL/V6+5D7dfLHnP/43nuu6a7p5P+/lvbwb7i/pIXywr57ok58Zj6fux2voHP4daA4ni/n++z8E9tprr6kk0DQJb3rTmxw3mMZGFiT+n3vOuXDzzTfHffJGAhucI7k/STJy51irDtK1QwILDwHH7KsO4ahZYUxinFOIACSRH7k8rsuvfPWr8H/++q/VrsWjQj6kpGhLrlOJLl+H33+8ROA9BcjlMecfP85m4I7jd3ILeqOg6ME+fUG88lvkyWfWO3P5KGTZDokeNYCo6Utf+lLHGTqdznQlgBG4Ao6TuD9yhElLAETxN23alB900EG+NOQo3D/FWeljuT0wp3Z/STJISQ5VnFpJF8XrSfILUqr/O633X8H18fPc5/55fuONN7p3sLy8PBVJLwV1+6V14r6737mTENx/JW14rt1PSBhLeb+/XNFXT3y8hMAfnKvlwRKABPSjT31bqFG4WKDIf/VXfxW5/7Sy/r70pX+F//qv/4Js1MCfvMb5KBX4LaJ8iWoTK2p1f/kpa7fkOqqB6w4lNl2dlmGYtiXHud5uu+3gwx8+Fc4773zYY489orGvubHkxSOJupgyDoAlATK+FVslyQDPOz4epClKEaA2pVSBnF0+l+f2+A05fmtANiC1442BHmLBsXoSAFGcjRs35g/ed1/HDcbSbSf4wXHh+Pbff383XklxJwkoZSwtLeWHHnpogUsNzf3LJAApCVRIBY1/rFQy5XeK71PO55/8yVPzX/7yF1E6xbkfDfpRl5bHnO7eI11Zf/x5yVX7Bb0/3a7pJejg/Fu3qduS49VSRHn7NFY5XuL+3G4tCYACWTCHHnPpx40/niQQJcasLxzvtAp+Inf6ylcucBuMUhTg0EAcvCoaMnVPGVdvCvIp9pXQ8/0uuD2XxXnaJ0+Dr33t63DQQQ91Vn88P4pNynPYfvzI4640V/Ro6I+/lv5qyZK/pwwyxWeTEXr+O9sO9F6OJBHQWMn4I5/Drjc5BrqHQDxzbVoZqM5vf/vbfMcdd9wyHoEBH+ISyIVR90fuMA3d31udl/PHPvYxo3N/y3HLrOyDdPl7wcdy/HXr1uVvetOb8ltvvdXNOc436vvjv7s8aRnn392E3owShz8vWpqApEljkVzfr+ei5EHSibExxHFbCYAlgdoEQBoD3/yWNzez0CdEAP7jP77uxtnEIhkEKPYjnHvuuWEMDahGZQh/L0d8VN/kmlqxYkX+spe9NL/88svVGmwCz1Lifd8Yy3rBUOaNbV3xt5sUsXX79VSB5kE+y7JRK8YkAFEKuOW3+S677DJTtgBaOC884QQ31mn5gZEbLS4u5g87+ODRCMAgRK9C/nsBQSBuj358OrZ+/fr8la98RX7ppZfEeV5aatLC3y9wxl5/Oe/1kGEQQi/nvf5i/JDV3EsA/F3r1aIHQxCaBC+B4PrGPplwUb9amknbAUiKGYoASMR6/wc/4F7WVF2CFZwDF9Luu+/uXEKIlKMbhoafi7POOiuOY2jkt8a81oDAm3vBhxiHlSD32Wef/G1v++v86muuUXPc3LssutJyh/CIzIjoS5oo9JAILDEh6EkCoLmpDs6hoBzqr66LsK67mpHaGgl1v1J1IPVBSzpZuK62oYzcgPj3qU99Cnz3u9+DdqcNve6E6twPkQBy1mfOhONPeKGr+DKHqVQTTEInN83i4iI86lGPhMsvv2L0uP+qOHzVKWyVYI1a0kCKhtqnPOUpcPzxx8HTnvanzr2HQAVbmwvl1W45GToLzpQVdu6V8dWuaEeozOPu8UE6/n824Gn8kVFQPkbfG9xSRUh8VqA/TNfJcl5+HNX4SW34cXNVIGncDM/jPqEyUTjeYQJcD1noReKL+fBHPgqHPe4wuOO225updjsG8v/5n/8ZHH/CCdDt9lzRj0kD+Z1PPfVUuOyyy5sp9+XeSzZSFdhZR3gZhbawYgEOOeQQOProZ8FRRz0T9t/vIfEceW2aTuCx0X280QY4pMDtwT0OhaIdtJRjyKOItcetwgmJZDpuQMKQdhOTdYrJGuFXjBmQSEmx+n4A2hugnsh4LzghyHsEdH+E57GSfHiuzBkCBhQNSAElWvz7174GRz/naOgu4T5ImHgwPUmAkG6//faD733vey7ffxqhyhRYdMMN18PBBz/c7fSLMFKyUTqJK02Tt6AEkFqEcnGS+5WkQwv3v//94ZBDD4EnHXEEPPGII+CAhxwQz+E7xGWIuydN6t2l8uXz8DvLUd/CzThxg05foZuJMRFiCqihQN1Q38/RDM+xOckmYJgLv/VrxZ/TyIz9+0SfluLeMnGN26SxY1CQfCYmZlwMRFIvzgL07TLy+2f3+kLMGhrGZ45iP4r/nzrzDDj5xS/z+U85UnsshTod5N9ll13gW9/8Bhx40EOnUusPgfLLTzrpJDjjjDOG5/6SIVBCV5rI8/X22JBQzkkS3RoEHpaw7bDjjrD33nvBQQcdCI965KPg4IMPhgMOODCK93Ie8X1N451pAuDr7iNQFJ7/4gkApuj53HtBkTHakhBOhFsGIT1gqubk8mVR4Q7/6Xk5wSGlfXa+DyMAKe7A/yYcDX1HhE/leJOkQRWEpLnF7T0mCYDofsigmeVuF+Y6HXj/hz4Cb3jdayDL5v0eB72libEsQrg1a9bAV7/6VXjCE54w0aIPKcnn61/7GjzjqKPcsUZr/dkU3xnR+0ntQ4Lf6czB2jVrYLvtt4ft1q+HPffcE/Z5wD7wgH32gQc9+MFu56Xdd9u9gNikCrgkmXZ1fbumgBe3/ksEAFwiVZBagrgedX7auy+Gu1ATOPaIlu68Uwpy5M7++ZyNyHH4VoEAuL6JZoQCoVmoByD78RIDSwoULMREQXIPGfTDacF+W6IwzggkdTgCEOjIiNF9FKWFyPehUz8Kb3zDKdBdXnSLpNtFtWAyOv+6devg3HPPhac97WlTy/smEXfzps3wqEMeBVdcccV4to/UdFsC4Doefcw0viOOeCK86U1vhuXlxch1pCgYo9DcvnR+E41WiLJDAr+wYgUszM+7vytWrIDVq1bBmrVrYdXKlaV9k1RUjGybPNB6pkg5Lp2FY8LvXtf3BEAgOOnssfCGuM5f4Dm9++kNh5yM6aMWOYrQgzQQ+pwN8WIdV2dZgmwN9N1fSmpWbNEYCAfb8Jgg6Os65TuJ1NuFhAw2vV4XXvuaV8KDH/wAePnLXg7XX/draLWweUSaXmM7uyDyP+AB+8CZZ34GHvvYx06N8yNgX2hg/Lu/f5dD/pEMf2ZOlQgqoSjNjUQI6L3usdvu8PSnPx2aBqv3S2SfeiVpATZ5pqBnOa4NQRTPigU3vcPSE0tHRIjTt3w6JBGVsGsP3YTI6Mp9h+u4R1I5wrZfOao+QUSPRT+k2M7jLSI4qzItTAGOxEpa++kh/BjddoORCLC9oIX6us1yKi8wWA6I7N1eD/7kKU+BH/zwe/C8Fzwf+v2uQ/52qzOynkcEhiSNY577XPjhD3/okH9a5Z4k8v/HN74B7/6nfxov1x9BUnOr43O6Fx8bUxXYvLjo5mtpacn9HfaDzyo/5AYlRMf3gB+Kzd/SuSJFu4Vf03FcyLlz5tS+ig+NO4jvqBq4Ql5B/6d7QsUfRyBCVp2v8ddzkrXT1d1rTVFxf94TCiYeciyR4JBxNfocqPwXZj3iuu8YLm3Lg1EdQfYYcJYh/u45BUSdkEaHYYiAKzAYarLdb8/7wTmfOxvO/8J5cOBDD4KeIwTeQFdngUR9MySCkMj//ve/Dz5/3nlwn/vsNtVyTyRl/M///D84+eSXQHe5WygCURsMMY8tRBWRNtokf00jj+CaI0SluR3mQ8Y6+kxbpK8C/y7I1cjvRTMy4a6L4nPmryPuTK4xLNQRECmqSyT1u2aQlQaEioTF2waKhDqguWvPl//gvkLf0ULPSM8LIteygWLWlMzExI0Lf8i5kISIPqja4PZjIqBB6y6sPw2z2HGBIbKjSvDnf/ZcuOjCC+ETp50GD3v4w8Jxz00k95Af0tuI86xduxZe/OKT4D//80fw+te/Id4/LeSn3WM3bNgALzj+ePjNjb8Zv9KP5egFPJIGJ2gEmlDDZg8sQtDi75VMsjf4cQxAFniwVy+pTQ8BWdGw5yv1xYAhL6IHJkZEwxkG8Xqs8kuvFK/xFncdDBTy+NH9GMqDe5WCx+oJBuID5e6LZ3DPLX39ek74r//4dqgtmp9gzMSwQKY+xP21EcMaiuq+nF6PEXXz4mb45je/AV+54Ctw4YUXwbXXXgN3370xeSfuQXDwwQ+Dpz/jGfDnf/Zn8KAHPVil3U6L88gCD897/vPg/PO+MH6Jb9WBmlo95Q0hPo332c96Fnzpy1/eIoUyRwVbZouO6fdvuRwucjL0BQSP0XgIlM4rGV+hZ+Eu8/d4qz7hO0qm/roMxX7XlsQRTyLcaAJXLvrm2W6SI7KTcTKI7PyMFCOAa9+rwukgHzl2God3I1q9X46lw1TBigopM/QwiMe53Mh9FuYX4JlHPct9UA+97vrr4Kqrr4Ybr78BNmy4yz332rXrYK+99oZ9993XBY4QICHB89NcuMThkeD8xV+8xiE/eSAaA2kWmbi7T4qXsw5FoxeVvy5W4ZFEweu1tIuOLZXtfyISYbRhVd/0MtiFFpl9tPTTGEk3IIMuGfu8BQAlSH9tvEIEBwUrfoz6DEbG8Mzs5vPqA+IRSRSaqHF1YPm8XLuA8VmGDcdQYG1BLGtQDqr+IvJiPhvxEObn5+FBD3yQ+5S+hnC9txtMtwgpFZrAsb/hDW+AU0/9cKw5PxZYji+PC9eO9zo1rAbMON7ryDfJAUXwjrJPMXen8wzE8ciS7v9yGCwzklzoy7Jdr90zXnjERGQNfUUvmT/O4xH7/qn3KHT6RE6A9iT4drVFX9qGrR6ZktaFcVN5RPQi7LCUJW9sPqZfuoVszTNJ0ekvIv60LPypmIbNmzfBy172cjjrrLOaq+1fsPbTQfECXY39YAi0asG9AMrjTchKTcEuKdeHnBT2QmjGpe8h9ZbjAMrCLtsJ/dk5ykK3njv7uB+OE2AXY1uI77gFmLQ1+LbL7GhSRNdTI6+nBB9prGwLW4BUkXgeKDFI9KbaFBgmpYDyVddEObBZsiATkHsLkf/qq6+Gl7/8ZfC9732/WeRPMSs6JyPAHHtpFvlH8lg0DoQwEhlo4XpOzeIyndOqgNV3PUJICdGKxsV2itl5uYgTYDsYc2U2/JEh0bUaEM0ZCV1AD9seOLLP98fEgsYV8K1QUqy4OJjrMwMtquv2uzboy3N+rP45W1ZPGgyzsJCaA3IzksRx7rnnwOMf/3iH/J1OQ8jvOhKl2RSjEdvF0rHU2r9XEADPjb19hR6aautp3ZTBLnTmdCzm+pp4tr4fG9os4bDtpiTf0H5AXNbvYx5gtAFoAkTRe/o4Von2WUZCuYibfQZbgvprkTotEVmjoPciUwQuESPZJtczBOh6EsUX8E3l4MMpUyWStyYglyQFGl1/w/Vw4okvhOc//wVwyy23BJ2/YdcZG6bT59y7NoFBTcFMEAAE6e6SkEJQ6Z+nY4E7Bx998ZwkLjKsFkykojR+2zGQuuCPxaMtb2R06cAh088TGcIHCrFmFUVJdrkwD5JU4E7Yop0SaeX9nsDx+Omv8UJEG54NEy7WBxChwCQWFKxTamC6ga0HyOZAqbxkj/j9738Hnz7jDHjPe94DN990c9xnrjHOXwZ2muNUm/r+MwrsUioOtlq9K7NI073eN84iOeuyHtHQMyMlAbL6pwzFZW7AMvWgL55PZN25d0LuOEr7lzvvcNEOOi8CjcM7DcG+MaMvhA0LFYbVI62nW0MlqzAy/Z7O01zQXhFGtWRFwP3b0TnFKTHDvjy+mYsWzNZWYeTCk2IvWfUp4ONXv/oVnHPO2fDpT38arr/+hnjNxBHfQqVkOltWQKk7c566tRsNolzaqq310zLxiBCE8+b5N7v8OPHH32PjBez6zkxyDBfmkHYuJnISMWkOmJBIP39Af1PURbYpbQJM1Pz9jItsp/Dcn34TIdBqERPLQIhMxSG/MZlWqUJFoKCKKgPfIDaUeuF1F8FkIZV3sLS8BNdccw1859vfgQsuuMAVENm8ebM7R5F9YyN/Fb6WTVNqmuOx5pAfM/uaBo5qY0gnpfDairH2gTm5eHv3V08K+9q1gc8Did7sNoz9invYKEciMkkVUnUAwyyoXSF9+AB/ebVQEWTwnFZX/NiCaoARiGIcbMln/Cu+b7tvgDX4+Tb0MYyU9UZIr1FKoobueB3sh9kEgtrYggZlIB9Yv2TfUUhxNPpXceDNw8aNG+HKK690f2+66Sa46qqr4KqrrnTbdV1x5ZWwuHkxXis3nWgEBiE/TZtUP8vsAd6X1My4xHsaxxhoJT0pJmtkL6oB2tXldVSXDaes4+6KcD1xK9mu1O8JaDIpAMcSDJYONLEqU1vomqp5ojbl+K1R0X+XMf/+WtrOyyIwfaRBLytdRKwqsa2Dyo/57MKOaE1KP3TUz1OUALQII4sMpCZCnrMiID88U2i+VoYmyvJG47oZKcT10l9cCoc/7vDS60j3b4Tj14FsWEJBsnWzrsDm3a70Pjn4RVua0+9R5cZHNUcG+9ADk2hPfXG7rRa5w7RaIteZZmrUj7U/gDhH/ci/KUMh3a8TiOKsKNsASQrkESAjntDTTQxMUSykvorqkVTDfDu+HypG4q8J2YoxSYja8W12SBzShgZJKYlaaB1H6yEakeSArQFDLwoSt1JulOHrFErx3xahIDFxokhfhuzy3cpjltbKkw1b7YsLbVxgCZAXluupzmgSHNCuM1pfkhNyQI8Hsp57VYR1aslwqCiJl/bUKHLfhnaTUd9lAUPmSYy0Q0gpQ4WlJZ6l5Kq29bWWMHFItPbE+b5TRkHt++d3lmMuQFFES3FzbogGJkUvm0lF0K9ZqpmotX9xLN4MD+0WZxROzPdtETkFWhJM44c8rqrQCu7fkBSQj1G7QKp29Jt1eUlYpLXdqoYS0dxRc1/qQW3kHhl3OT6eXXFim13Thr9HIobuKw9cmaPq5K67KVzgtpnosCqgcVvOBzNbtgFIGwZJU9JAKQllesGx9EPXpMqLsfokE6E6ZVRdPrCuQQbiQXhgtoBhvZWrxSUdGkz9yppqMxj0YtdT1TCTwwo3WwmwQeRvdk78AkN6Ij1Hvh4e6abEmWVSDhvA/GJli3Q6pkRKcKnQYCtBdEVdf4s8liilCE0u1nDKLS7vKRpxtB2A2y0jGOlrrLgvXZZEgEU1InEvh9FrvCzOrR6AIgCyakgauK54kcJY0bqc41j9kJMn5eDkxM4A8st1UMUUxO8oalkVQF7HyePFthp8lP5Yxj8pHpNezRzFI460unPyjdazdeIZi6IEmsOxj5/UTJbu9L2cAsw6OLVFXJ2t6KQng0IqOscpwCmOy94JqQbrXYI1I5XVe4uhyxIXYqCQiIiUeQRyPjnmn+scyufSaoNUsWgu/b2KAOiAiRT7qUJ4C1LBNWfMZHHZYu5Hi4uDQSYSNcY5q5C+rA+xZmTQaCXYdTaMOj0kNNMkcXSpZ8r0VWBXG/qvgx5Moq50BWpOLl1dcgHTMSIqmgN7h4mvMlz0KEiJIByVwTG5dE9CYlySOxeL5rh/VQgzGSjl2rcuPD1GJj5WRbIEEnGuk7CNSW5Pc0T4advkFGEctgsEKrps2BWSzsqSkzvKUrP32RVvuUd9cDHXTcrOseEhHq2uSpBqc4RnrgOYnjpu86Qj+5p0nF6b58sCgakXu3ON+zcRnKN6KOFe1Cb7/7naDp21T1bMDNR9SK4rr0urI+7frIpie+JHTFSL6FoNsG3rgCJtY3GqjePTNF6SfihBid33bNSUc8bEllUxNlImkoGyEt0/9cByQvVE0ktKA7Ur/8rJYXFpWHB52KGZsSzeKY5MQ5N/00Q9fcx+5PGqexsATxhHA5u5xyKpr2cnrhR//fKkkNgylVAuSq2v+rXI3I6DePj61GQNIv40Jtm/PW+Th7QKZD/+eua+7HKz3F8vpGLWoh4Tj83W9g/qgUdhcUzaD6gEGM0jtU4uSX+gkHDPeoRWBYiC2ACEInZkJXnIepL1X7pXtzOSBDCumyupp5vzljHZ61NtWakg9V31GU5MiCCMAuxTzgv2GZmb70R+Mk5FRchvpmHfs+TAfn3JYh1YfBUXP8e+63tDHr4CuXZsbEB4jiTTygRiSWO3VIuLOQr+Ti3BMvf2dgsZlkz5DLpf+czSf28ZLT0TEYR2LFIiCpYVNgVV5c6j+O8rCzkCUFQDJMUKk+LSNcvEKgL9gkYVxbUfddh7c8hHjaBjE2o5jZLTUhXRV9RqqolF8UGgSRinnJoOD/d/cS89UirIeBW5XihyGdNfVcSgBa2na0+S5KLyWmpLSg56rWlPhH1Zpv+MaK2UDUKMSqUhyD4DfpMRiSzusWVezkF5rIGWuqpTllEZyJS3QHMe5llW7IxGQOt+SHPm9HbD+prhQk/TBGKc5CK312kc7xA3uqGwwBpHphauyNTDT5mXKAWDpiF17ZjCjC206poeg7BITkjILyU1z4Vw37tgUHa+QuaIeoce6TqjB7UqgDYK+vvbA6Q9SZ1TBu2Ebpdbyk4SR997FqPf3NqX/F+dcEPty6w/bQRMR03ymGQCUJGRynZ6alMRlnrsO0bpS6vzNDZHANigY6mcpLb6+CCqnvbtlsnT8uGbzCzMh76UEr6YFIg3XOCEzXRdek8DQgAtpl133dU32YA7kBcwS4RyUZPYT3YwXHyemPK60QYpK2ZThzwHMpKteBGvU25LMispSRDySFGd7qdNO70UQ/TJp9VSJSBfAYie2SM4bbYpqxH7hB9dfUhuz62TiWRAnPRM6NdlOLv76dUg2m+QdX8y2tqwfiY6+FekA7OYJQdWeBNxMZSJVDzQYsiiJRajR6eVgnjWop95iDacJSf+mCyHblba102H5997730aaE2K4hSJF/TRUDNfygbhRLzFIUSQFByZSOYLCP3ZDZ1i6K1+a9qPRIJE7rRvnouMpJJ+CDlCXIHjrpixFPYTiNcRzgiJOS4TGRlrKxlzzQDerlt6ThiYCBDhs/OjmSlvLcZuRbsdubRpEIHpFPUQkm3VcNSgyiUDfY+Nk06DVTnGgwal5tieJ84Jg9wEEbcpoCSpRzziEWMbSWWgmF9IPROUQro/77QrNesY2IMVdeJYdH1AuU190YDHHF4Hvcj2ZbBNic6v1m0rcP4+ZLk3qnnujwRC1vIrW6eBkCWlW+qP+2J+ZMt165qBWpKx+iY/hysXjvNBhUkNQdRjkM/g31X0AtQJACoag6rAD1pXN+GBayIyBrcudDsmCYhrjal63L9ta8B4AVTdCDdWOeigA0Mm3egqFrunOJMtIistUBUUQzonNYA75/hKNVlyHfBa4GP+r0YQq2LKBU7JN+IjxxQ5LiEQcWFvDPTvOrg3aRdgb2aPzyg38vRcncYoMwSDROS65jBoiaC8fTiN2br8JBOV6dH8vJ74kDrFuCsrK2mVgtSUsEaKkUJiWiMVHwYs1U1915SsEb+9igNogAjQejElzLcmIKL6/Oc/323pTTUQRwW2ZMs+3JmAFsT5BeKQ9BiQJvLJ4CmQUXA2c1OqoTrBi3X2xFObMfotuIIYV4hByUkFVogVHowcBG4d0GaghMhU1ERmr3IYMgXqeMFREyttcCelSLZFxU1l2XFGeE6KkjYCXQEpVTuhuISjChAmKmbiUWxxGcfPVCZTUY8qc29YEUbqYil35PDQGJpOE+GL0zJ+k5mvvoubqr7kJS92x5rYWakYa+56U6K41+/9dtpenA471LoNYgmhwioy3DOw2nCCkIAi4ZibFdcXifzdOJxiXgBJJFLNzZW04lUY/L8FeYuIQ1BkQm2LaDx04bZBVYhPTvMfiJ6TImSgEPfH5tKgSkUXnpXSJPLLeacMS98GlzTTxCwVvkzHY09kcNAiOg08RQCqoOq8fjjJWZuQANyj9mfUoDAI8ubF/+OOO85tszY+96cyVmXeHQTDMIIEGVWnxO67XiEga7W/N6oT9DuoDcwRmQsyBw1cWXHx9ITqXIB+4Lyc7k6Si0diHjsbxwMhNeoEtY06OeMqzRvHAvjnod2MqQk2OGqJk8V9KiRK0ZD6fZi5FpIKt0MElQuHdGShBO3nLw9S0NSsML0lor68t+zc+EClohth4NMSAhruh8Tl7bbbDt70pjc1Ilkx2Mg0zXEy5Joit594DOm72jDM1hWp8/ojfqGzTk4LWha2lHEJQSolq73rj9oO7YdxSakgUzvmhjZyQnApkOAWYMSpse1gI8K9/+LmQZnbCiyYDIL5wdsdbMSjI1jOsxDGHKZEZ/zJgDjN+QtqjiTMYcNSHoOfoaiOiUQl93a0qJ9ajanFIxeAPV612KSE0TwRwG3Jt0Z9vUmgIqdvfvNbYJ999nHfhzP+8fbTsuCHrChjjXduC60o8vo2cIH7WnjeM0Dc1evGtBhpzaHbzW/a6Yi4l+FDH4RARCQkg0HLfdCzUdQOvnplsI3LLKgl8WA/cGMJcm1qNx5+9T/FhiPuGaUe78crVWhvBwiSjkP6NmQ51uyzKUw4R6TfE+JiW6S6ycIf0oZAyB8IUyA4VISUJSrplveuwBAK7G/WUI6cOlOvNSQxkMQm9X080OId/NEB7WD8pCc9Cd74xlPiBqvDgQzIkgRbgn1nhNDBGMhsMSCdX4g+QjDcQemrFKsekZyGQcggreQ6ai6MNpySkgVJCtIlKa6Jrsrc+O71s1GMvUNmaschgKiy7ySsgGRk+xBDtKnzXpKga+k5MhFeLec2PJ3Y6jvtXheSDtjxEVEO8xUDsdzmoFUKcxoxdVqkvE4ulNQCsZKG3VZpdNAcxY7vjwv599prL/j0pz8FCwsr4kYo9UHHb+goPRvYJew3MXw2ZNuRtZ58/v6igCzB4EUI5r7IMUp3W0oMlkYuznjTBi4yZkvkLq6HTLUtg3IonsE/t48T4Gh6/1ChLFLMIwhr0KXwilLihJNxfco4BhKY/PXkT5GZgD7Wn7IguSBJYZ8AMc9uA9PYuG+HpQDGQ0ca0iGWZWCuC1S6CBb5LWL6iWh6i7E/VvGfkH/33XaH888/H+573/uNxP0HT1/PhLwS8MLyEkBYqCGCzses0yImQlGWyw8mbl32kUZ4ey9mEur6fvZeaecqmaMoRsenCpZ/YVOI5ggZeITxDmI78jhOMt4xceVnpeRp+1xk1yDCIqItHe5x3ECUHsR2k/p1epXMz49XY8hBWgIp8V0PjoMcErMnJQR3ifRNet0mVWt+HEALrBtZY0av2QZ8Tr+PYRcOOOAh8PWvf81F/eHvUdx+XB1HcPdgNfbA7y29pyTr7bQoydpv143TheMSlMhA97P9ISJCoQ+6167DMntWOBfrEmbhmFAZxPokYqWMjIHN4vh9fIDo20kN3I606PugHX8vH0dJQO4HSZ43GcfPT6nVs+Qb9B6NELPAcQt+jogo03wZI+AgkNRyGIlB6DtClyMJgANCxkdc2hbsjwXxcd6Q0z/nOUfDt7/9LTjwoINgeXl5ZJ+/9AgVRX1W9WhDTB2gw4jJgTXCvRbaLwTGqD7UaKJKgn31HQfEj2tRXIPt0TlZd5+ukVl1dC8hv0R6yamtlMHFTWUwkKu05FQBs2swGUpDAJFXhZADS47tCaC2acgx2jmhuS8G+cg5cZ6YqNoYpu1u43khS42Z1DKQWVaDrk2IaK5z+3LyUmIwCvSa3tF3hsAFbrRaCvF32mkn+Of3vxe++MV/hV12uY87Njc3NyYhlfqwjMwjDwBzKJlYohdlCIP1I4+HvYU/FegSDIXKdBRUiEz2iyClj8wgtNxwxCbakA/ccPuMiIiRWt2HyqCR6C6kDWErcYSAbARid+Bo43Dlt0QqLs2RWutUEEVKUyn8lLhljbVhS/NApHTkpJ4LBFEQRLyOxK6tdJ1/iXyxtkjaHVf4Rfoh1dFHxxXdc6cPN5ZbINr1UKV3ikui/hdezUiGOLk+EOGZOK5cuRJOeOHx8Na3/O/o6iOpYHTwiIA1FWghajcgG640ByVI6eXW9mOBqgeZBRglR2oX/+1EpOREISEiZ50gAXopgRHd9xObVuNtiYMCicy4bZ0/5+twXN8TuhjSTFmo5K1AouBshVy2wx2Llnh+PiIaeiz+PjaIptRkEuvlsOl6yr2QxJLsAMENKB+Mf1MjUhcsLmJdE8BuikDnrJknJd7Q6EkvGi14ZXm56/TfeyPMzc/BgQccCM9+9jPh2GOfBwcccKDK+BsVdHitF0l1QRd6X6ksUclROMLMY0W/BnHQiFYIhXU6Ndl15kSGHEsMtJmGvN+qGbofkkI1ovnHlIRFqwgUM+CDgPBfGXwTKX3w+1PkXoKZkxocgnWsW9MG+BR1f1nxN9wnNvvQ48d/7XsjtSwUBOGb6CXqpIPiU1pgY4VMDbWgkbqKCEi3Uz0iQO3e9373hZe+7GToIEJQckZ4C+2WfzYfiuprzNO4XAhnv+fDOIMYFcdL24yZIBh/D+4x2I8771ACSJAPXZ/41YeH+mw8b5mnjS+0VZgMTngNSjJz8/OwZs1quN/97wcHHXgQHHjAAe44AhK68bg+I3A6FkRcWdgcJjxt3OuOf2uiUCbO2vPFMbGxzs8lpwHLIhrUn1w3JG1S/Xvbrs6Ii2C4py6mIfvQabfxLiREAbnJlRe5tbjcSw/UnxbJGaTRlVRkWSczNUZqh8ZJNQGIofJ5KnKS9ftd0bPdrCFtZEgjrh18cTENz9XtHu/bAIHi+sdJ7Q1vJL4XudCI28lELeY68ph1o0l3siTwliCUrSv5O8UISHyV1xUz6/zHJ+m441E81pIuKJ3actVBoMfIWX9MJL1koiXrWF2I9r6IOQipMVVJW2UMVI6Nkpck8WVG5TwxmgCkOisDEsHMLj/JWm9lhEOKW3LTiKA2jID8yB2QIxMhS9knygqVWN1qsB1B7w3HepeNllS9hOtNS4Xx6juyxIan4wCLzonUVNeh3K2J3UZFnd5yccuRDCAykoisokglAnO1Hpv5J3fBYSkgCOaxBJcct//LwrsVlVsiT14SQjn+MqLFjJKTkGLZIzZh0FzHgySZ0rkU0UnNre/X2uyKhMG2w4SUiVSsCJQCGdSQfpG61LAYSsxWsoMbLC2oJJERFzlOZqfTTHThvRvkltKs0tC5xJuteHcpYk/cl/ziIsEmWSJcEgAq/inHpdUJHV5LzIjakOMKcQvmubXtCqFqvWlk53FxurG7giLznCrgJQ4KEeYn5eKdkTooSWqwdCRF+XRyFoE0AMa7wzc/RxWYUqTiOk2R9VZ9vQ3OSNUMsN/p2vo6/zaoDzLSk/37zAX8exWZdUIK00DuP7wXE1u68d1JTqwkCsNDdIUob0Vng52UMsoQQY+N1ZeUHUT3JZ9ZPlNe8Kv79mUCkl6nvr2C/97hdGBmdF2i3DgZAWOogBMDJKeWz0nfreu8OBdFkNZ/4dblgYk8xFJ/v6BdQvwscn9p0NIIXqS23hebykJsSsTdBho04vNW3xGZC1ZzBImYVt/XC8+LuJyNp7GexHG55ZUUyfk8MQKqtiPXBH9E2LH7gxKDDACSce/hOWJRDjn2THDgKr1aSiY2f4WR0QX1WISMUYQUGSkqC0miGbvl8XBoEUsJpF5KQsr4YjftKXoGOEvQj8fsDqx15jrcWOv89EdOJBk5SHSzCSZ2IreJ701DatdYtqSn3rEV5/n9kNjrFxJl8vnztLj5eokMRTXB3ysDg6rsLXIMHiF07rxFbkIYv3NxLE6ixP+WITyBGwdvg+ZDVCRUehYsRxYjjYFT/JviAvxtlD0oZidZwKQo5jMBIjxjAoOGYfRmecJDwJ4SutfjYLu4O7B+MPu9CDyZpIPJF1HMHbfMvdl9ALZBEeQWWoyEus5+aiFLndSKnMLAqfhEmTGLRHQp5kYqkrg+XFGaJ+LHoSVL6YHAvnwyEIvx1GZmpJ5c9BuImBqTnBMfgclrOOU2NWMz0YDUprszEFEtNYs5Ftdrj4zcUlzOV/gdGbIluiwBUl0BkwxkF0Ad0ItCimlazDR3xSql/dLCm3/Mqb1NARWoYFcQHU8tIhKvra+fDWvOqCVoOiODjAwkcZv0Xa3iUbt6ofr+094G6suvqbJgM8+EJEeX7Rhx2l3ep1GZdlLr1s+X97yWSa+yLZpnEQYdogexVkJUFwQD1QFK1CepRFoVl/1p+403sMptwPR+V/r5OkXqUIX4kpuXWS7Tg0tbXXmQMj9aTgAtVN/OtriAUYDnz+bGU0AS/ZahvyQhCKlNZbqBQTIi6Fmosx9y2AvbyTEC+3T7FBela8rWI4n21nAmOXyZ6oHiMZUHz6Iyotc1XUvnwjcTt19sn8GXDfPhyJEIuCKcxHPR/CYJlZRGzNNW2MSsm5TGpmo6iOfgQqR+DCIOQFLhZFel4rre0XR4sDkGsgAFPztPtN5uaXqga7RtOUg/P+nHZdVirLgo/egp4k/tSx85n6M2NBIyUnBEKC1QNkyxXsxj033qtchFNGg8UryV95dFxZWpMxKsyGzPMecsjqO+dFp0QVoVxrrQaa7L1xy9Q0+AqS36W5x/2W9JHAAkX3b1NaOL6KmFTHqatnbaa4ZBxGHvsQu+aoHUba9qcdUHftF1+pMuviyxCKXF3KTpxuAm6Q2gtqV7V+vRXOfBIqlvi+8rZgWSL1+uN7Z823dgd9ShPux40iK9FpFDPb3S9cLtyrkvD/hKQx7UjlZ0W6YQvz7yy7ES0eXy4HrcNizaPXm/3xuwCb3W8SfF/Yq55+lw4lHVAAoc0caUqug4Gd6qOWaRGJZvg8bRhzKCzY5ruGeSeq6OdJSiuHQbEXfgvRzoXTIHt+qZFd0lEeTn1gjiK+dy1l7ZuqKYOArXlZCei2LVKhqTDT+WY+d2NNHj6/OY7yFKeccdgugZeWy+LYkLmsjFApym+nF6LuS9FlF5vnXhFZ4PXpO0P4ElfPb55TP5e1upxouD9PqjzNVvPtU2TxgUtc95lBgBaznmY4To9J194eyyZBGKjZXkCqJ58IUxvLTST6g11ePT8RDDgbyNuSP5++WzssRABEz62dPiKP02vmQVV+6v8UckB01xH2qXdVGdyFNOMKSKyRGC8h4btCMJFs1VeM/xXx6Pmwf3aCbYJhCF2HRh3mkNaCLkeleqTJFIxVoBERlprotzWIZqfu1Z9Sw9j0XVKryHPKyCokFFc0ANUheqB6k4dy1CyixEKebKPOxm7Asyg027w+Q4qG9fjbUYXCKlIqmXy+jHcKWytuuXwOfsdYPyFegaa7Hn76rwpCBqdjcm7zJLBbfIBCH6rd28DOEduVh/QlQozK0W2W0bhLxIFOUzy2q4qchSC1IykGKvEPeFsTB3r473LXCPEXzypIKW17KgLEXqVz4Lc9ri+uVU5rRqRsDifMr156EsQapsbvge91RFfzA9QOp7aGYkMbxosZX530zNWLcb19BXtC/ojSGKoqXcl07u1d4R3FpSaZJOeM943a+MwHNn4nFCSBe4ISQRsthLFym3bceqX75OzJLSm9x3rkyMT4nSof1YS47n0WMKLXIfgcdjlZF/3LZCbFUAg/rl8TKQt4K9QmXrgrmqfS6WPCQy50IKcvUkQx/+nJQWLQ5IaQsScyiuFGXJNWOgMdJcyTHRfKUIHUt7TNj8Dsf6HM8dj1F/VD0AVcxBAS8IfJhWa3ikJENO4owYsM3yovN1gHV3jQS0hbVsRyNh8Zg2lmhdil5WKiecFkta52PupZ9fExWuga+9Ipr4yBRe6xUoGgmJ6+H7pYg2LnHl996TKo+QMtxPKRYjxxRjjgYnP1eMXEgM/PPoucX+gz0kqhQ0Rs/9LTPi15mKAeB5l2+R20x9N4Q0CzsBE02TInIMLU5ni1pbAHN1P16d4VqsJ5Gyq0kVg9+zBMnh5Tj5t+zPEnyPY55hZX2s/ySsrlI8Lm4Zzg/WjAvO7vZKi7Q+5+eQSxa9NeKkKLekiKnzCDLenMVnfU+ZO8yqPBxLr0tLFeeDF3lW6t4qcnAeT1pMJC6T2nde9mVFfv+XNvqUmWTe2McLm9shw14/cQyJgv/up4A5rW9LWsZJvckGEHt2g2mRn0CqbqYYBgSJgTbzCIk7LNmE3kXcAG8oQnMv3gNJrHFj0lSKt5yzKjVGQooA0HEp7Wjk93ihHX12PJEA2AaKlKOejjqOdb482sldKYxB+lpN6ewkVPZepKQqIKkoGmubAQzwWEibgDCAOenBL+CiK03rv7r/VNyFVWMICQjp9VwUbRqaO/t5DAZRJb6LnXGUfiuBuKWUnPQcKwJNNQeS213p9qXdwO60w4iQ2q/ASrTeDUe1CPpRfxcxTm4ItKsvbWSCtxrGZ/fEINqhpCQ9/74fWfostW4sWBGepVDtPbP94Xz4+pilLaMRkLmS7IwHm34hTUgARSgjLrIEtTYW+klM62h2AUr/aFp3Q4pZNIiSeKej5lhdsRyVfeEpg5UU0bUR0Y/R6vHaZRci2eKQ6Rk194zqgchVZxHUgq2IwyoZn5eSUTHgxnsBpFtPGK+cxBCuibeEeniCuHJikVUb5LtkI6x0d2pOKKFqXeThrEBOHEPcASi0ie/dqb1StRJFP2ncrsKPlJJT/bJ0OpwRnV3RKQmU2tRzROu0vJ9QFJRvkGK/N8KxK0nr6YMGbCsFSa5VDunBMgGytgRd0YbGW2UFDW1FsY769dQ0fS+1nZIoysV5G9tAVF8SVXZfkXpQXMTaM0K730oiZaUdbxRyCOmak/1bqcdyjvD8ke3RfMvSUsX5DMtaVfOJ7TrECMUt1f59KSA7AX23qgyCraUvrih4Yaz0xC5BCKH4TKODiE+xE/RuhefBl/EKl4fIaN+ldhtTpWPLAGjc9ZBftBcvL3J6G75NOKIZULq/6AbkCZLU34q5flDajVT0dad1ZU256oYNs7hHImpaBC8GxJSJ/1bElGK5FtPEKBL3kzow7D4ELPZJoytLN1YKkiKvnrPw+sWitsQr8LFo8Ct7t5bTy8CmLEE0JcckJApISZuCRh2U1BQpSco7dcwBB0aVq1VpCcwSNVqj0sUZJIbg8gOZc+/+6GPRGBiZTug3Po8nDtwzuUCJYch3Gu4XQUaMltLlXnx/RVXUq4m6sAtLTTq4jNze6aDfQj0AKXJpX64MLuFFpMVZnhg2nNE5flnp0N+iyK+PyZedDwx31ROq2/HuJK2rFowjpfUAue/y87ZPS0zYim771GBFPGPUCsUmnAso5V+nPyp4R6pDNB6Zmcd/fW17MaYoMRF3i+xMqDJdUeSDGQSvEyJKcuHKuUlxd77HBuFoYyUbUO36k+tVcs7MIVIo4+X9fiwAxW2Mg+QSK/7orb+DLiDckyXcNiKpJIbyWnufJQx0L75vqYpZRizXVXWgn5IAdPaXRIByXzPfSy/Z6s78MOWGPU571Mclp6IHtiJsmQsvNXEBcRQnZO4o3YjWtSYlo6I7znK34nPyOKhfL3FoQsNekLSVHhKSFEkIxFV8Fh4Hp5T1b/V3gdzS0u2mJmT1qXLvsk25mEmyItBIzm5Z6RoeVC1a1uynPqyaIOdeMi8Zgi3zEJjIMU9uQd9tKqoN3vR+sJS81/nRDkN9Yb1B0X/Y9NQXRxGZrCq82L5P+Vw6NsEPg9a1ltyKNqZqApSCVtqvGZ+lIOLaLZk4WCJeYR6MB0bEQQZa8MvSlFDuLWBDgjUhKHtYuSBJgiHrt3RzUcCN5LS0HVRoKSl18JzpZ0+NxxIhakeO1fchXz73KyUwImJW8iGxlcZFvzxh5Tm2c4n/JoJF6H63yy1H0Ok95zQhTT+/Ffvl+Gir6yoIbYa1UuSMkrCLXl1AleWw8jkpuEhs6UUVitxuPlq39q7Pvt8NiDYEVfVMwhiiS1PnAvhvMiBLglZpfcCZlISkV0euD56T4mat9cBIAKlwXc399YPJ66y0IDmrFKm1DUHem5IQWBKRJaAs8sgFYW0L9GqFUSc+iyVSMpaAn7f82eicdD/Zt2u5Js9lsbS1//jnlDYJEjGlKqARvwjyXaWJRRrEqlaXiYPhMbWuLKUwXb2mGFglG6fn1rYFLU1YRLbvHxL327UgGQK124+SRaxLQCJ9qFUQVRdcE8HIT2oDrWu/55/fmJbUAL+G9LsveBDiY9q1Q+uCVWz/Hn0RVtwhiZlFSpWi/gYb7UtKgoUmFOLSpELhBQ8yujG193raIChmzhGB4DFpsZOoI42xyIm8i8ouPk0MKAjKzoP2hKTDi8ue3V9fjAXgPuXCTMUBkFrCi6IKgaX1Wy8eNhRxHACCdC2F+S7EQWCYKW2HbWdWJgjRYi6MKoHA0l1YBUT47dqUKcM8r9rHXmyrKKHmnCIUp47HG+czioM0M0VVuDhGv4ELI2nAJTdfwWUqwoRFS/pPPC3xzrpBLQOnOR7gdbNegDLQ7sGUlGA39hBUttitWhRaR5QLgzirHWIKgcsCIuwxOSmybS1K6lgDG8Mtx0sSgOyPr2OOIRFXzouUBGihSE4jxWwaPz+XF84DYiZ3l5FBXfJZpegeEEhFudHq6wU1gCrYUJ07Tkf1V0puZyUtghD8FKUx+17s2OMsqzUQtuYsWdx2nq3vXI6xr+8k6hb2AAxvR6iqhGh+rq2+zXEMtMtxsG/FGAe5K3EZYso1K7l70RuUJopibSh8TEM9X1zoLMUd02JIUYZMu9W07kNUvSg6mZcWKaiwNsf2ZB8WIeR3OTEsSejnIxtBmY5FolxR9+JrbR1+m/5LKcZ2rijtWI6F+mJ92y8qqfOXqWf2Y95F2f1RP6ZrKXLTj4n7FaJBdElKA11gEiptWHLZcL8rLy4Ldsp3KOe5nZD87PU2WcpCRm9AzoTi9vFMQGZf4FMQTCF5uOujahTeqSOeaCjUVns9/4LcFLY9t9enNnIRrcVYlqpQ6hEkADngYuqrFtll0EQ6Hl8+bEkvKoBBioGWI0ouLlUD+p1Sa4wLbEBcf2rMkhqzJb7qJdFLsfYPKT2UZfMVCSpzfyIIVOvOPrPlpBbkMcuNNd/1/YTNQCh+nniIyhWwXVhiGyPui9b8EJNvPQkeRP/qviLXTD+XfapWsNAHY2SwBfi7mQmxFuDbdowqPjdv3W3HwDarVP92raaqNpMEWo7ERbtdPc4/hASgqag0ZFiubjlNWTBNUY+VabDyeouIzKmL7cqXX+yTEUwccd9Z3EuPrUwUZXVBj0tLSRwUwjny6fHZ9suuYxGSxWg9R4xeLP5qW4OYM2mFl5tnuIg4Y6WPlZIEd6L5Iu5H79wdlvkDYvzOyWALe9Bpy+nt/JNXo+w9W8IhiYF5dvBE00v93F+eSBumPQ/84/Lzp4meJOZlyC9XHvXDlbQ59qaaP6diR4bxBgyoCSgrAAmdTwwSd8Um9cAfs3u22e+yBoA1auWlSMj+UInoVo9LcQC7kNg3GymsIhySiMkFqsVVjgLkGutlEo2NJSiGaEpRUr7UlBRCSGClAkJGuSj5/qL3xNwrVYFYxVbOXChpTfOeoYGLd6KJrkSbIFPaH4nE0mhbXMyJ2QxzTpJA8Zzuz0gflhZlwesSJAp0/2UtsQ5DUBC15YWe8B6T25FBhUdDPxMTGB6fNryn11MdyaAuDCAAnJtOoo/mJjiIYNhJhsRWi2GMOIwgOhoxVTOOuaz286YnmcdRNllZzcWqkUoHa8jIt/KgFoou00pXESmrcxms7hsCXGgsiuulrMDSck7dh/mJ34MEQJwu7OhLST0ynBiTZNiNSU0ZAu3uFxKTu1QGetn3UyTaadE+FeilZjw5g9GTkYnvoR8VBEfBO+ECfm9UWpuDesvFbptZy89Q3GCXJYIyl3hTiD+yDaA6PNZ+T10nESklFlM/coEQQWCrrg4trUpBhaTltDx0mCCxvVTsxyOhdksW56kqsclGcOkxZCWcXrtG+/1uWKTy2a3V23Ij/WyFsREHt/XxClKS5PZEPMo4n38WtglRG3RNEQlkXYiiNyY3EoQcq7apyP7VvFgBCuS6kWW+yE4gKx7yHOl3nLIH6eeT8TA8dslMKMXX2oPGK43XoBEwBVLfL6NudE6+mBRSMqWPGW3EbcxmEFLg9b3aCEFdXCNlEdZBP1w9iA2OHHXnzxMRohfSKnlhnAQiuYqKMS+AR0zmPvRctJ02XUNBLN2ok6YNmylVQV4j34W8Rr83NjamiFbfuLjonO3XSiLpBCf5zrhQKyGlzWiUkmJqbPxcHFwlbBxAUgtxeVIHZORdMb+l2E9KYrHfrbQjpRe5fqQRWie5pd7NFicAOkrPiuFF7moz4DTIdNyiSFi0dpL3wS+ANGIZkTe0R+OQ8ejFghJWxZCiN7tZykp++/mRG1pYzkB9yGMsIchQ0uIzUewFW6eLHBuhrEQ3RcGl9EyKwZB9p8TvlHRhx8LcnM7ruASppvgKNnpJdkOL9Mx2vBR5KN+dblMfEyoOhPMq9qFsLsukS5YEPamkvAobhZhSYZjJWG9BiqlMAv5/YeioH8Zzod0AAAAASUVORK5CYII="   # ← 在这里粘贴您的 Base64 字符串
        
        if not base64_str:
            return  # 如果没有提供，则跳过
        
        try:
            # 去除可能的多余空白字符
            base64_str = ''.join(base64_str.split())
            icon_data = base64.b64decode(base64_str)
            pixmap = QPixmap()
            pixmap.loadFromData(icon_data)
            if not pixmap.isNull():
                self.setWindowIcon(QIcon(pixmap))
            else:
                print("⚠️ 图标加载失败：QPixmap 无法解析数据")
        except Exception as e:
            print(f"⚠️ 图标加载异常：{e}")

        

        
        
    
    def _init_time_label(self):
        """字符拆分：Wed与Jun间距调大，Jun与日期间隙调小"""
        _c = get_theme_colors()
        self.time_container = QWidget(self.titleBar)
        self.time_container.setStyleSheet("background: transparent;")
        layout = QHBoxLayout(self.time_container)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        font = QFont("Microsoft YaHei", 18, QFont.Weight.Bold)

        self.char_labels = []
        for i in range(19):
            lbl = QLabel()
            lbl.setFont(font)
            lbl.setStyleSheet(f"""
                color: {_c['fg']};
                background: transparent;
            """)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            if i <= 10:
                if i == 3:
                    lbl.setFixedWidth(18)
                elif i == 7:
                    lbl.setFixedWidth(1)
                else:
                    lbl.setFixedWidth(14)
            else:
                lbl.setFixedWidth(14)
            layout.addWidget(lbl)
            self.char_labels.append(lbl)

        self._animations = []
        self._last_sec = -1
        self.update_time_text()
        self.update_time_label_position()

        self.time_timer = QTimer(self)
        self.time_timer.timeout.connect(self.update_time_text)
        self.time_timer.start(1000)
        self.time_container.raise_()

    def update_time_text(self):
        now = datetime.now()
        time_str = now.strftime("%a %b %d %H:%M:%S")
        for i, ch in enumerate(time_str):
            if i < len(self.char_labels):
                self.char_labels[i].setText(ch)

        self.time_container.adjustSize()
        self.update_time_label_position()

        current_sec = now.second
        if self._last_sec != -1 and self._last_sec != current_sec:
            self._start_character_wave()
        self._last_sec = current_sec

    def _start_character_wave(self):
        """仅时间部分（索引11~18）从右向左波浪缓动"""
        for anim in self._animations:
            anim.stop()
        self._animations.clear()

        start_idx = 11
        end_idx = 18
        step = 40
        total = end_idx - start_idx + 1

        for offset in range(total):
            i = end_idx - offset
            lbl = self.char_labels[i]
            orig_pos = lbl.pos()
            anim = QPropertyAnimation(lbl, b"pos")
            anim.setDuration(150)
            anim.setStartValue(QPoint(orig_pos.x(), orig_pos.y() - 4))
            anim.setEndValue(QPoint(orig_pos.x(), orig_pos.y()))
            anim.setEasingCurve(QEasingCurve.Type.OutQuad)

            delay = offset * step
            QTimer.singleShot(delay, lambda a=anim: self._start_anim(a))
            self._animations.append(anim)

    def _start_anim(self, anim):
        def on_finished():
            if anim in self._animations:
                self._animations.remove(anim)
        anim.finished.connect(on_finished)
        anim.start()

    def update_time_label_position(self):
        """将时钟容器置于标题栏水平居中"""
        if not hasattr(self, 'time_container'):
            return
        titleBar = self.titleBar
        x = (titleBar.width() - self.time_container.width()) // 2
        y = (titleBar.height() - self.time_container.height()) // 2
        self.time_container.move(x, y)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_time_label_position()

    # ---------- 原有方法（保持不变） ----------
    def after_boot(self):
        self.show()

    def load_student_data(self, file_path):
        df, id_map, msg = load_student_list(file_path)
        if df is not None:
            self.stu_df = df
            self.id_to_name = id_map
            self.current_file = file_path
            save_config_path(file_path)
            self.backup_mgr.save_student_backup(df)
            show_toast(self, msg, "success")
            return True

        backups = self.backup_mgr.get_student_backups()
        if not backups:
            show_toast(self, f" 无法加载名单且无备份：{msg}", "error")
            return False

        latest = os.path.basename(backups[-1])
        reply = QMessageBox.question(self, "恢复备份",
                                     f"名单文件加载失败，是否从最新备份恢复？\n{latest}",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            backup_df = self.backup_mgr.load_student_backup(backups[-1])
            if backup_df is not None:
                id_map = dict(zip(backup_df["学号"], backup_df["姓名"]))
                self.stu_df = backup_df
                self.id_to_name = id_map
                self.current_file = file_path
                show_toast(self, f" 已从备份恢复 {len(backup_df)} 名学生", "success")
                return True
            else:
                show_toast(self, f" 备份文件损坏：{os.path.basename(backups[-1])}", "error")
                return False
        else:
            show_toast(self, " 恢复取消，请手动处理名单文件", "warning")
            return False

    def show_toast(self, msg, type_="info", duration=2000):
        show_toast(self, msg, type_, duration)

    def set_input_frame(self, frame):
        self.input_frame = frame

    def refresh_input_reasons(self):
        if self.input_frame:
            self.input_frame.refresh_reasons()

    def refresh_input_shortcut(self):
        if self.input_frame:
            self.input_frame.update_shortcut_config()

    def refresh_ranking(self):
        if hasattr(self, 'rankingInterface'):
            self.rankingInterface.refresh()

    def show_change_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择名单 Excel 文件", "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if not file_path:
            return
        # 1) 先把用户选择的 Excel 转存为标准「班级名单.xlsx」（覆盖旧文件）
        if not self._save_as_class_list(file_path):
            return
        # 2) 再统一从标准文件加载
        if self.load_student_data(CLASS_LIST_FILE):
            self.switchTo(self.homeInterface)

    def _save_as_class_list(self, source_path):
        """读取用户选择的名单 Excel，保存为标准班级名单文件
        （路径：user/data/班级名单.xlsx，存在则覆盖）
        返回 True 表示成功；False 表示失败（已弹出提示）"""
        try:
            df = pd.read_excel(source_path)   # 不强制 engine，兼容 .xls / .xlsx
        except Exception as e:
            show_toast(self, f"读取名单失败：{e}", "error")
            return False
        if "学号" not in df.columns or "姓名" not in df.columns:
            show_toast(self, "文件必须包含'学号'和'姓名'两列", "error")
            return False
        try:
            df_copy = df.copy()
            df_copy["学号"] = df_copy["学号"].astype(str)          # 学号统一转文本，避免前导零/科学计数
            df_copy = df_copy.drop_duplicates(subset=["学号"]).reset_index(drop=True)
            os.makedirs(os.path.dirname(CLASS_LIST_FILE), exist_ok=True)
            df_copy.to_excel(CLASS_LIST_FILE, index=False, engine="openpyxl")
        except Exception as e:
            show_toast(self, f"保存标准名单失败：{e}", "error")
            return False
        return True

    def reset_scores(self):
        # 自定义 QMessageBox
        # ✅ 第一个 QMessageBox —— 用 DIALOG_STYLE
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("⚠️ 确认重置")
        msg_box.setText("确定要重置所有积分记录吗？\n重置前会自动备份当前数据。")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)
        
        # ★ 核心修复：获取返回值并判断
        reply = msg_box.exec()
        if reply != QMessageBox.StandardButton.Yes:
            return  # 如果点了 No 或者直接关掉，直接结束函数，不再往下走

        # ✅ QInputDialog —— 用 DIALOG_STYLE
        input_dialog = QInputDialog(self)
        input_dialog.setWindowTitle("确认操作")
        input_dialog.setLabelText("请输入“确认重置”以继续：")
        input_dialog.setTextEchoMode(QLineEdit.EchoMode.Normal)
        input_dialog.setStyleSheet(DIALOG_STYLE)
        ok = input_dialog.exec()
        confirm = input_dialog.textValue()
        if not ok or confirm != "确认重置":
            show_toast(self, " 操作取消", "warning")
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_csv = os.path.join(BACKUP_DIR, f"备份_积分记录_{ts}.csv")
        if os.path.exists(DATA_CSV):
            shutil.copy2(DATA_CSV, backup_csv)
            show_toast(self, f" 数据已备份至 {backup_csv}", "success")
            cfg = load_config()
            cfg["backup_count"] = cfg.get("backup_count", 0) + 1
            save_config(cfg)
            check_and_unlock_achievements()

        # 备份 Excel 文件（可选）
        if self.template_path and os.path.exists(self.template_path):
            ok, msg = export_excel_with_template(self.stu_df, self.template_path, month=None, output_name=os.path.join(BACKUP_DIR, f"备份_积分记录_{ts}.xlsx"))
            if not ok:
                show_toast(self, f"⚠️ Excel备份导出失败（不影响重置）：{msg}", "warning")
        else:
            ok, msg = export_excel_fallback(self.stu_df, month=None, output_name=os.path.join(BACKUP_DIR, f"备份_积分记录_{ts}.xlsx"))
            if not ok:
                show_toast(self, f"⚠️ Excel备份导出失败（不影响重置）：{msg}", "warning")

        # ========== 关键：清空积分记录 ==========
        try:
            cols = ["学号", "姓名", "日期", "变动类型", "类别", "分数", "赋分总分", "备注"]
            pd.DataFrame(columns=cols).to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
            show_toast(self, f"✅ 积分已重置，备份文件：{backup_csv}", "success")
            # 刷新排行榜页面（如果可见）
            self.refresh_ranking()
        except Exception as e:
            show_toast(self, f"❌ 清空记录失败：{e}", "error")

    def generate_heart_points(self, n, screen_w, screen_h):
        """生成爱心形状的坐标点（屏幕居中）"""
        points = []
        for i in range(n):
            t = i / n * 2 * math.pi
            x = 16 * (math.sin(t) ** 3)
            y = 13 * math.cos(t) - 5 * math.cos(2*t) - 2 * math.cos(3*t) - math.cos(4*t)
            sx = int(screen_w / 2 + x * 20 - 50)
            sy = int(screen_h / 2 - y * 20 - 80)
            sx = max(0, min(sx, screen_w - 150))
            sy = max(0, min(sy, screen_h - 80))
            points.append((sx, sy))
        return points

    def _show_hearts(self):
        if hasattr(self, '_heart_running') and self._heart_running:
            return
        self._heart_running = True

        tip_list = ["多喝水哦~", "好好爱自己", "好好吃饭", "保持好心情", "好好睡觉", "我想你了", "别熬夜",
                    "天冷了多穿衣服", "顺顺利利", "爱你", "好好休息", "你是最棒的", "未来可期", "坚持就是胜利",
                    "别放弃", "加油", "相信你自己", "每天进步一点点", "向前走，别回头", "阳光总在风雨后"]

        all_windows = []
        screen = QApplication.primaryScreen().availableGeometry()
        sw, sh = screen.width(), screen.height()

        # 阶段1：爱心形状（120个）
        heart_points = self.generate_heart_points(120, sw, sh)
        for px, py in heart_points:
            tip = random.choice(tip_list)
            popup = HeartPopup(tip, px, py)
            popup.show()
            popup.raise_()
            popup.activateWindow()
            all_windows.append(popup)
            QApplication.processEvents()
            loop = QEventLoop()
            QTimer.singleShot(10, loop.quit)
            loop.exec()

        loop = QEventLoop()
        QTimer.singleShot(1500, loop.quit)
        loop.exec()

        for w in all_windows:
            w.close()
        all_windows.clear()
        QApplication.processEvents()

        # 阶段2：随机散落（减少至80个，更流畅）
        for _ in range(120):
            tip = random.choice(tip_list)
            popup = HeartPopup(tip)
            popup.show()
            popup.raise_()
            popup.activateWindow()
            all_windows.append(popup)
            QApplication.processEvents()
            loop = QEventLoop()
            QTimer.singleShot(10, loop.quit)
            loop.exec()

        def close_all_and_thank():
            try:
                for w in all_windows:
                    w.close()
                all_windows.clear()
                self._show_thank_you()
            except Exception as e:
                print("Error in close_all_and_thank:", e)
            finally:
                self._heart_running = False

        QTimer.singleShot(3000, close_all_and_thank)


    def _show_thank_you(self):
        self.raise_()
        self.activateWindow()

        thank_win = QWidget()
        thank_win.setWindowFlags(
            Qt.WindowType.Window |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint
        )
        thank_win.setFixedSize(700, 200)
        thank_win.setStyleSheet("background: #f0f0f0; border: 2px solid #2E86C1; border-radius: 16px;")
        layout = QVBoxLayout(thank_win)
        label = QLabel("感谢您的使用\nPickOS\n有你更出色", thank_win)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-family: 'Microsoft YaHei'; font-size: 28px; font-weight: bold; color: #2E86C1;")
        layout.addWidget(label)
        screen = QApplication.primaryScreen().availableGeometry()
        thank_win.move(screen.center() - thank_win.rect().center())
        thank_win.show()
        thank_win.raise_()
        QApplication.processEvents()

        loop = QEventLoop()
        QTimer.singleShot(3000, loop.quit)
        loop.exec()

        thank_win.close()
        

    def show_heart_animation(self):
        splash = QWidget()
        splash.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        splash.setFixedSize(300, 150)
        splash.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2E86C1, stop:1 #1A5276); border-radius: 16px;"
        )
        layout = QVBoxLayout(splash)
        label = QLabel("PickOS", splash)
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("color: white; font-size: 40px; font-weight: bold;")
        layout.addWidget(label)
        splash.show()

        screen = QApplication.primaryScreen().availableGeometry()
        splash.move(screen.center() - splash.rect().center())

        click_count = 0
        last_click_time = time.time()

        def shake_window():
            x = splash.x()
            y = splash.y()
            offsets = [5, -5, 3, -3, 0]
            def do_shift(step=0):
                if step < len(offsets):
                    splash.move(x + offsets[step], y)
                    QTimer.singleShot(20, lambda: do_shift(step + 1))
            do_shift()

        def on_click(event):
            nonlocal click_count, last_click_time
            now = time.time()
            if now - last_click_time > 1.0:
                click_count = 0
            click_count += 1
            last_click_time = now
            shake_window()
            if click_count >= 5:
                splash.close()
                self._show_hearts()

        splash.mousePressEvent = on_click

    def create_tray_icon(self):
        icon = self.windowIcon()
        if icon.isNull():
            icon = QIcon.fromTheme("application-exit")
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("PickOS 班级管理系统")
        tray_menu = QMenu()
        show_action = tray_menu.addAction("显示主窗口")
        show_action.triggered.connect(self.show_window)
        tray_menu.addSeparator()
        quit_action = tray_menu.addAction("退出")
        quit_action.triggered.connect(self.quit_app)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()
        # 不再弹出通知

    def on_tray_activated(self, reason):
        """托盘图标被激活（双击）时显示主窗口"""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window()

    def show_window(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()
        self.hide_floating_window()    # 激活窗口获得焦点

    def changeEvent(self, event):
        if event.type() == event.Type.WindowStateChange:
            if self.isMinimized():
                # ---- 最小化：先播放粒子效果 ----
                # 防止重复触发（例如系统连续发多次 WindowStateChange）
                if not getattr(self, '_particle_min_emitted', False):
                    self._particle_min_emitted = True
                    try:
                        # 用 normalGeometry 保证拿到的不是最小化后的异常尺寸
                        _spawn_close_particles(self)
                    except Exception as e:
                        print(f"⚠️ 最小化粒子异常: {e}")

                # 最小化 → 显示悬浮窗，隐藏主窗口
                self.show_floating_window()
                self.hide()
            else:
                # 从最小化恢复时重置标记，下次最小化还能再播放
                self._particle_min_emitted = False
                if hasattr(self, 'floating_window') and self.floating_window is not None:
                    self.floating_window.hide()
        super().changeEvent(event)

    def closeEvent(self, event):
        # 已经播放过粒子（第二次进入），直接退出
        if getattr(self, '_particle_close_done', False):
            if hasattr(self, 'floating_window') and self.floating_window is not None:
                self.floating_window.close()
                self.floating_window.deleteLater()
            if hasattr(self, 'tray_icon'):
                self.tray_icon.hide()
            event.accept()
            QApplication.quit()
            return

        # 第一次：先忽略关闭，播放粒子效果，再延迟退出
        event.ignore()
        self._particle_close_done = True

        # 在窗口当前所在屏幕位置播放粒子
        _spawn_close_particles(self)

        # 隐藏主窗口，让粒子层单独显示
        self.hide()
        if hasattr(self, 'floating_window') and self.floating_window is not None:
            self.floating_window.close()
            self.floating_window.deleteLater()
            self.floating_window = None
        if hasattr(self, 'tray_icon'):
            self.tray_icon.hide()

        # 延迟退出，让粒子播完
        QTimer.singleShot(900, QApplication.quit)

    def quit_app(self):
        if hasattr(self, 'floating_window') and self.floating_window is not None:
            self.floating_window.close()
            self.floating_window.deleteLater()
        if hasattr(self, 'tray_icon'):
            self.tray_icon.hide()
        QApplication.quit()

    def show_floating_window(self):
        if self.floating_window is None:
            self.floating_window = FloatingWindow(self)
        self.floating_window.show()
        self.floating_window.raise_()

    def hide_floating_window(self):
        if self.floating_window is not None:
            self.floating_window.hide()

class TableSelectionCleaner(QObject):
    """点击表格外区域时，自动清除表格的选中高亮。
    只调用 clearSelection()，不会清除 item 的背景色 ——
    因此搜索结果设置的黄色高亮背景会被保留。
    ★ 新增：点击按钮（含按钮内部子控件）时不清空，避免删除了
       「先选行再点表格外按钮」的交互。
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._tables = []

    def register(self, table):
        if table is not None and table not in self._tables:
            self._tables.append(table)

    # ---------- 工具：判断点击是否发生在按钮上 ----------
    @staticmethod
    def _is_button_click(obj, gp):
        try:
            from PyQt6.QtWidgets import QAbstractButton
        except Exception:
            return False

        # 1) 从事件接收者向上找
        try:
            probe = obj
            depth = 0
            while probe is not None and depth < 25:
                if isinstance(probe, QAbstractButton):
                    return True
                parent = probe.parent() if hasattr(probe, "parent") else None
                probe = parent
                depth += 1
        except Exception:
            pass

        # 2) 从鼠标位置向上找（更保险，覆盖事件被父控件拦截的情况）
        try:
            w = QApplication.widgetAt(gp)
            depth = 0
            while w is not None and depth < 25:
                if isinstance(w, QAbstractButton):
                    return True
                w = w.parentWidget()
                depth += 1
        except Exception:
            pass

        return False

    def eventFilter(self, obj, event):
        if event.type() != QEvent.Type.MouseButtonPress:
            return False
        try:
            if event.button() != Qt.MouseButton.LeftButton:
                return False
            gp = event.globalPosition().toPoint()
        except Exception:
            return False

        # ★ 关键：点击按钮 → 保留表格的选中，不清空
        if self._is_button_click(obj, gp):
            return False

        for table in list(self._tables):
            try:
                if table is None or not table.isVisible():
                    continue
                tl = table.mapToGlobal(QPoint(0, 0))
                if not QRect(tl, table.size()).contains(gp):
                    table.clearSelection()
            except Exception:
                pass
        return False

class FloatingWindow(QWidget):
    """可拖动的图标悬浮窗，双击恢复主窗口，右键菜单含排行榜"""
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(48, 48)

        # 获取主窗口图标并缩放
        icon = main_window.windowIcon()
        if not icon.isNull():
            pixmap = icon.pixmap(48, 48)
        else:
            pixmap = QPixmap(48, 48)
            pixmap.fill(Qt.GlobalColor.transparent)

        self.icon_label = QLabel(self)
        self.icon_label.setPixmap(pixmap)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_label.setStyleSheet("background: transparent;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.icon_label)

        self.setToolTip("PickOS 运行中 (双击恢复)")

        # 默认放在屏幕右下角
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.width() - self.width() - 30,
                  screen.height() - self.height() - 50)

        self.drag_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self.drag_pos is not None:
            delta = event.globalPosition().toPoint() - self.drag_pos
            self.move(self.pos() + delta)
            self.drag_pos = event.globalPosition().toPoint()

    def mouseReleaseEvent(self, event):
        self.drag_pos = None

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            # ---- 先播放粒子效果，再隐藏悬浮窗、恢复主窗口 ----
            try:
                _spawn_close_particles(self)
            except Exception as e:
                print(f"⚠️ 悬浮窗粒子异常: {e}")

            # 立刻隐藏悬浮窗，让粒子层单独显示（避免图标残留在粒子中间）
            self.hide()
            self.main_window.show_window()

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        menu.setPalette(QApplication.style().standardPalette())
        show_action = menu.addAction("显示主窗口")
        show_action.triggered.connect(self.main_window.show_window)

        # 新增：查看积分排行榜
        rank_action = menu.addAction("📊 查看积分排行榜")
        rank_action.triggered.connect(self.show_ranking_dialog)

        menu.addSeparator()
        quit_action = menu.addAction("退出")
        quit_action.triggered.connect(self.main_window.quit_app)
        menu.exec(event.globalPos())

    def show_ranking_dialog(self):
        """弹出积分排行榜对话框（排名、学号、姓名三列，自适应）"""
        stu_df = self.main_window.stu_df
        if stu_df is None or stu_df.empty:
            QMessageBox.information(self, "提示", "请先加载学生名单")
            return

        result, err = calc_total_score(stu_df, month=None)
        if result is None or result.empty:
            QMessageBox.information(self, "提示", "暂无积分数据" if not err else err)
            return

        # 提取排名、学号、姓名三列
        df_display = result[["排名", "学号", "姓名"]].copy()

        dialog = QDialog(self)
        dialog.setWindowTitle("📊 积分排行榜")
        dialog.setStyleSheet(DIALOG_STYLE)   # 使用全局暗色样式
        dialog.resize(215, 420)

        table = QTableWidget()
        install_scrolling_header(table)
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["排名", "学号", "姓名"])
        table.setRowCount(len(df_display))
        for i, row in df_display.iterrows():
            table.setItem(i, 0, QTableWidgetItem(str(row["排名"])))
            table.setItem(i, 1, QTableWidgetItem(str(row["学号"])))
            table.setItem(i, 2, QTableWidgetItem(str(row["姓名"])))
            # 居中对齐
            for j in range(3):
                table.item(i, j).setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        # 列宽自适应内容，但限制最小宽度
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(False)   # 不拉伸最后一列，允许各列按内容宽度
        # 设置最小宽度避免过窄
        table.setColumnWidth(0, 60)
        table.setColumnWidth(1, 80)
        table.setColumnWidth(2, 120)

        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)

        # 布局
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.addWidget(table)

        close_btn = QPushButton("关闭")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        dialog.exec()

    def closeEvent(self, event):
        # 关闭悬浮窗只是隐藏，不销毁
        self.hide()
        event.ignore()

# ======================== 首页（便签替换天气卡片，双击操作，无分隔线，右侧对齐优化，自定义弹窗） ========================
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QTextEdit, QStackedWidget,QSplitter, QMessageBox,
    QInputDialog, QDialog, QDialogButtonBox, QLineEdit, QProgressBar, QDateEdit
)
from qfluentwidgets import (
    CardWidget, BodyLabel, TitleLabel, ComboBox, SwitchButton,
    PushButton, InfoBar, InfoBarPosition
)
import json
import time
from datetime import datetime

class HomeInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("homeInterface")
        
        self.today_date = datetime.now()
        self.start_time = time.time()
        self.version = SYSTEM_VERSION
        
        # 便签数据
        self.notes = []          # 每个元素 {"title": "标题", "content": "内容"}
        self.current_index = -1  # 当前选中的便签索引
        self._load_notes()
        
        # UI 引用
        self.cpu_label = None
        self.mem_label = None
        self.disk_label = None
        self.battery_label = None
        self.net_down_label = None
        self.net_up_label = None
        self.cpu_progress = None
        self.mem_progress = None
        self.disk_progress = None
        self.battery_progress = None
        self.todo_count_label = None
        self.total_count_label = None
        self.top_student_label = None
        self.life_time_label = None
        self.version_label = None
        
        # 便签卡片 UI 引用
        self.notes_list = None
        self.notes_content = None
        self.notes_card = None
        
        self._last_net_recv = 0
        self._last_net_sent = 0
        self._last_net_time = time.time()
        
        # 倒计时数据
        self.countdown_events = []
        self.selected_index = -1
        self._load_countdown_data()
        
        # 日历状态：默认选中今天
        today = datetime.now().date()
        self.calendar_year = today.year
        self.calendar_month = today.month
        self.selected_date = today   # 选中的日期（初始为今天）

        # 进度条动画相关
        self.progress_animation_enabled = load_config().get("enable_progress_animation", True)
        self._progress_needs_animation = True
        self._animating_bars = set()
        self._progress_anims_keepalive = []
        
        self.build_dashboard_ui()
        
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_dashboard_data)
        self.refresh_timer.start(1000)
        QTimer.singleShot(100, self.refresh_dashboard_data)

    # ==================== 便签数据持久化 ====================
    def _notes_file_path(self):
        return os.path.join(DATA_DIR, "notes.json")

    def _load_notes(self):
        path = self._notes_file_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    self.notes = json.load(f)
                if not isinstance(self.notes, list):
                    self.notes = []
                for note in self.notes:
                    if "title" not in note:
                        note["title"] = "未命名"
                    if "content" not in note:
                        note["content"] = ""
                if not self.notes:
                    self.notes = [{"title": "欢迎", "content": "此间方寸，\n可容纳万千思绪；\n一纸浅墨，\n以铭记朝暮日常。"}]
                return
            except Exception as e:
                print("加载便签失败:", e)
        self.notes = [{"title": "欢迎", "content": "此间方寸，\n可容纳万千思绪；\n一纸浅墨，\n以铭记朝暮日常。"}]
        self._save_notes()

    def _save_notes(self):
        try:
            with open(self._notes_file_path(), "w", encoding="utf-8") as f:
                json.dump(self.notes, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("保存便签失败:", e)

    # ==================== 便签卡片 ====================
    def build_notes_card(self, card):
        _c = get_theme_colors()
        layout = card.main_layout
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(0)
        splitter.setStyleSheet("QSplitter::handle { background: transparent; }")
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(6)
        self.notes_title_label = BodyLabel("请选择一条便签")
        self.notes_title_label.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 13px; font-weight: 600;")
        left_layout.addWidget(self.notes_title_label)
        self.notes_content = QTextEdit()
        self.notes_content.setPlaceholderText("在此编辑便签内容...")
        self.notes_content.setStyleSheet(f"""
            QTextEdit {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border']};
                border-radius: 8px;
                color: {_c['fg']};
                padding: 8px;
                font-size: 14px;
            }}
            QTextEdit:focus {{
                border: 2px solid #64B5F6;
            }}
        """)
        self.notes_content.textChanged.connect(self._on_content_changed)
        left_layout.addWidget(self.notes_content)
        self.save_status = BodyLabel("已保存")
        self.save_status.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 10px;")
        left_layout.addWidget(self.save_status, alignment=Qt.AlignmentFlag.AlignRight)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)
        header_layout = QHBoxLayout()
        header_layout.addWidget(BodyLabel("便签列表"))
        header_layout.addStretch()
        add_btn = QPushButton("+ 添加")
        add_btn.setFixedSize(70, 28)
        add_btn.setStyleSheet("""
            QPushButton {
                background: rgba(100,181,246,0.2);
                color: #64B5F6;
                border: 1px solid rgba(100,181,246,0.3);
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: rgba(100,181,246,0.3);
            }
        """)
        add_btn.clicked.connect(self._add_note)
        header_layout.addWidget(add_btn)
        right_layout.addLayout(header_layout)
        self.notes_list = QListWidget()
        self.notes_list.setStyleSheet(f"""
            QListWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border']};
                border-radius: 8px;
                padding: 4px;
                outline: none;
            }}
            QListWidget::item {{
                padding: 8px 12px;
                border-radius: 4px;
                color: {_c['fg']};
            }}
            QListWidget::item:selected {{
                background: rgba(100,181,246,0.25);
                border-left: 3px solid #64B5F6;
                color: {_c['fg_bright']};
            }}
            QListWidget::item:hover {{
                background: {_c['bg_card']};
            }}
        """)
        self.notes_list.itemClicked.connect(self._on_note_selected)
        self.notes_list.itemDoubleClicked.connect(self._on_note_double_click)
        right_layout.addWidget(self.notes_list)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 200])
        layout.addWidget(splitter)
        self._refresh_note_list()
        if self.notes:
            self.notes_list.setCurrentRow(0)
            self._on_note_selected(self.notes_list.item(0))

    # ==================== 便签内容编辑 / 新建 ====================
    def _on_content_changed(self):
        """便签内容编辑后自动保存到 notes.json"""
        if self.current_index < 0 or self.current_index >= len(self.notes):
            return
        new_content = self.notes_content.toPlainText()
        self.notes[self.current_index]["content"] = new_content
        self._save_notes()
        self.save_status.setText("已保存")

    def _add_note(self):
        """新建一条便签"""
        title, ok = QInputDialog.getText(
            self, "新建便签", "请输入标题：",
            QLineEdit.EchoMode.Normal, "新便签"
        )
        if ok and title.strip():
            title = title.strip()
            self.notes.append({"title": title, "content": ""})
            self._save_notes()
            self._refresh_note_list()
            new_idx = len(self.notes) - 1
            self.notes_list.setCurrentRow(new_idx)
            self._on_note_selected(self.notes_list.item(new_idx))
            # 自动聚焦内容编辑框，并将光标移到最后
            self.notes_content.setFocus()
            self.notes_content.moveCursor(QTextCursor.MoveOperation.End)
            show_toast(self, f"已添加便签：{title}", "success")

    def _refresh_note_list(self):
        self.notes_list.clear()
        for i, note in enumerate(self.notes):
            item = QListWidgetItem(note["title"])
            item.setData(Qt.ItemDataRole.UserRole, i)
            self.notes_list.addItem(item)
        if 0 <= self.current_index < self.notes_list.count():
            self.notes_list.setCurrentRow(self.current_index)
        elif self.notes_list.count() > 0:
            self.notes_list.setCurrentRow(0)

    def _on_note_selected(self, item):
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is None or idx < 0 or idx >= len(self.notes):
            return
        self.current_index = idx
        note = self.notes[idx]
        self.notes_title_label.setText(note["title"])
        self.notes_content.blockSignals(True)
        self.notes_content.setText(note["content"])
        self.notes_content.blockSignals(False)
        self.save_status.setText("已加载")
        self.notes_list.setCurrentRow(idx)

    def _on_note_double_click(self, item):
        idx = item.data(Qt.ItemDataRole.UserRole)
        if idx is None or idx < 0 or idx >= len(self.notes):
            return
        _c = get_theme_colors()
        note = self.notes[idx]
        title = note["title"]
        dialog = QDialog(self)
        dialog.setWindowTitle("操作便签")
        dialog.setFixedSize(380, 180)
        dialog.setStyleSheet(f"""
            QDialog {{
                background: {_c['bg']};
                border-radius: 14px;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 15px;
            }}
            QPushButton {{
                border-radius: 8px;
                padding: 8px 18px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: none;
                min-width: 80px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
            }}
            QPushButton#deleteBtn {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #EF5350, stop:1 #D32F2F);
            }}
            QPushButton#deleteBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #D32F2F, stop:1 #B71C1C);
            }}
        """)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(18)
        title_label = QLabel(f"便签：{title}")
        # ★ 修复：原字符串少了 f 前缀
        title_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title_label)
        info_label = QLabel("请选择要执行的操作：")
        info_label.setStyleSheet(f"color: {_c['fg_dim']};")
        layout.addWidget(info_label)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        rename_btn = QPushButton("重命名")
        rename_btn.clicked.connect(lambda: self._rename_note(idx) or dialog.accept())
        delete_btn = QPushButton("删除")
        delete_btn.setObjectName("deleteBtn")
        delete_btn.clicked.connect(lambda: self._delete_note(idx) or dialog.accept())
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(rename_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        dialog.exec()

    def _rename_note(self, idx):
        if idx < 0 or idx >= len(self.notes):
            return
        _c = get_theme_colors()
        old_title = self.notes[idx]["title"]
        dialog = QDialog(self)
        dialog.setWindowTitle("重命名便签")
        dialog.setFixedSize(400, 160)
        dialog.setStyleSheet(f"""
            QDialog {{
                background: {_c['bg']};
                border-radius: 14px;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 14px;
            }}
            QLineEdit {{
                border-radius: 8px;
                padding: 8px 12px;
                border: 1px solid {_c['border_light']};
                background: {_c['bg_alt']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
            }}
            QLineEdit:focus {{
                border: 2px solid #64B5F6;
                background: {_c['bg_card']};
            }}
            QPushButton {{
                border-radius: 8px;
                padding: 6px 18px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: none;
                min-width: 70px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
            }}
        """)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(12)
        label = QLabel("请输入新标题：")
        layout.addWidget(label)
        line_edit = QLineEdit(old_title)
        line_edit.selectAll()
        layout.addWidget(line_edit)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_title = line_edit.text().strip()
            if new_title:
                self.notes[idx]["title"] = new_title
                self._save_notes()
                self._refresh_note_list()
                self.notes_list.setCurrentRow(idx)
                self._on_note_selected(self.notes_list.item(idx))
                show_toast(self, f"已重命名为：{new_title}", "success")
            else:
                show_toast(self, "标题不能为空", "warning")

    def _delete_note(self, idx):
        if idx < 0 or idx >= len(self.notes):
            return
        _c = get_theme_colors()
        title = self.notes[idx]["title"]
        dialog = QDialog(self)
        dialog.setWindowTitle("确认删除")
        dialog.setFixedSize(380, 140)
        dialog.setStyleSheet(f"""
            QDialog {{
                background: {_c['bg']};
                border-radius: 14px;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 14px;
            }}
            QPushButton {{
                border-radius: 8px;
                padding: 6px 18px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: none;
                min-width: 70px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
            }}
            QPushButton#deleteBtn {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #EF5350, stop:1 #D32F2F);
            }}
            QPushButton#deleteBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #D32F2F, stop:1 #B71C1C);
            }}
        """)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(12)
        label = QLabel(f"确定要删除便签“{title}”吗？")
        layout.addWidget(label)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        delete_btn = QPushButton("删除")
        delete_btn.setObjectName("deleteBtn")
        delete_btn.clicked.connect(dialog.accept)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.notes.pop(idx)
            self._save_notes()
            self._refresh_note_list()
            if self.notes_list.count() > 0:
                self.notes_list.setCurrentRow(0)
                self._on_note_selected(self.notes_list.item(0))
            else:
                self.current_index = -1
                self.notes_title_label.setText("请添加便签")
                self.notes_content.blockSignals(True)
                self.notes_content.clear()
                self.notes_content.blockSignals(False)
                self.save_status.setText("")
            show_toast(self, f"已删除便签：{title}", "info")

    # ==================== 构建整体界面 ====================
    def build_dashboard_ui(self):
        _c = get_theme_colors()
        _is_dark = _is_dark_color(_c["bg"])
        _bg_grad = ("qradialgradient(cx:0.1, cy:0.1, radius:2, "
                    f"stop:0 {_c['bg']}, stop:1 {_c['bg_alt']})")
        self.setStyleSheet(f"#homeInterface {{ background: {_bg_grad}; }}")
        
        # ★ 修改点 1：外层使用 QVBoxLayout，里面装 QScrollArea
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setStyleSheet("""
            QScrollArea { background: transparent; border: none; }
            QScrollBar:vertical {
                width: 6px; background: transparent; margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: rgba(150, 150, 150, 0.5); border-radius: 3px; min-height: 30px;
            }
            QScrollBar::handle:vertical:hover { background: rgba(150, 150, 150, 0.8); }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
        """)

        content_widget = QWidget()
        content_widget.setStyleSheet("background: transparent;")

        # 原来的 main_layout 现在放在 content_widget 里
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(24, 16, 24, 16)
        main_layout.setSpacing(18)

        top_header = QHBoxLayout()
        top_header.setSpacing(20)
        left_time_layout = QVBoxLayout()
        left_time_layout.setSpacing(4)
        self.time_large = TitleLabel("18:07")
        self.time_large.setStyleSheet(f"font-size: 46px; font-weight: 900; color: {_c['fg']}; letter-spacing: 1px;")
        self.date_sub = BodyLabel("8月3日 星期一\n农历六月廿一 | 四时有序，今日安好")
        self.date_sub.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 14px; line-height: 1.5;")
        left_time_layout.addWidget(self.time_large)
        left_time_layout.addWidget(self.date_sub)
        top_header.addLayout(left_time_layout)
        top_header.addStretch()
        self.greet_label = BodyLabel("")
        self.greet_label.setStyleSheet(f"font-size: 20px; color: {_c['fg']}; font-weight: 500;")
        self.greet_label.setWordWrap(True)
        top_header.addWidget(self.greet_label)
        top_header.addStretch()
        right_action_layout = QHBoxLayout()
        right_action_layout.setSpacing(12)
        search_edit = QLineEdit()
        search_edit.setPlaceholderText("搜索内容、文件与工具")
        search_edit.setFixedWidth(360)
        
        if _is_dark:
            search_bg = "rgba(255,255,255,0.05)"
            search_border = "rgba(255,255,255,0.1)"
            btn_bg = "rgba(255,255,255,0.2)"
            btn_border = "rgba(255,255,255,0.3)"
            sign_bg = "rgba(120, 160, 190, 0.18)"
            sign_color = "#A8C6E0"
            sign_border = "rgba(120, 160, 190, 0.35)"
        else:
            search_bg = "rgba(0,0,0,0.05)"
            search_border = "rgba(0,0,0,0.1)"
            btn_bg = "rgba(0,0,0,0.05)"
            btn_border = "rgba(0,0,0,0.1)"
            sign_bg = "rgba(120, 160, 190, 0.1)"
            sign_color = "#2E86C1"
            sign_border = "rgba(120, 160, 190, 0.4)"

        search_edit.setStyleSheet(f"background: {search_bg}; color: {_c['fg']}; border-radius: 18px; padding: 8px 16px; border: 1px solid {search_border};")
        right_action_layout.addWidget(search_edit)
        page_map = {
            "录入": self.app.inputInterface,
            "排行": self.app.rankingInterface,
            "成绩": self.app.scoreInterface,
            "名单": self.app.listManagerInterface
        }
        for btn_text in ["录入", "排行", "成绩", "名单"]:
            btn = QPushButton(btn_text)
            btn.setFixedSize(80, 34)
            btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: {btn_bg};
                    color: {_c['fg']};
                    border: 1px solid {btn_border};
                    border-radius: 18px;
                    padding: 4px 10px;
                    font-size: 14px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background: {_c['bg_card']};
                    color: {_c['fg_bright']};
                }}
                QPushButton:pressed {{
                    background: rgba(100,181,246,0.5);
                }}
                QPushButton:focus {{
                    outline: none;
                    background: rgba(100,181,246,0.3);
                }}
            """)
            btn.clicked.connect(lambda checked, page=page_map[btn_text]: self.app.switchTo(page))
            right_action_layout.addWidget(btn)
        top_header.addLayout(right_action_layout)
        main_layout.addLayout(top_header)

        sign_btn = QPushButton("📅 签到")
        sign_btn.setFixedSize(80, 34)
        sign_btn.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        sign_btn.setStyleSheet(f"""
            QPushButton {{
                background: {sign_bg};
                color: {sign_color};
                border: 1px solid {sign_border};
                border-radius: 18px;
                padding: 4px 10px;
                font-size: 14px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background: rgba(120, 160, 190, 0.3);
                color: #C8DDF0;
            }}
            QPushButton:pressed {{
                background: rgba(120, 160, 190, 0.45);
            }}
        """)
        sign_btn.clicked.connect(self.do_sign_in)
        right_action_layout.addWidget(sign_btn)

        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(18)
        self.card_notes = self.create_dashboard_card("📝 便签")
        self.build_notes_card(self.card_notes)
        row1_layout.addWidget(self.card_notes, stretch=1)

        self.card_calendar = self.create_dashboard_card("📅 日历与倒计时")
        self.build_calendar_card(self.card_calendar)
        row1_layout.addWidget(self.card_calendar, stretch=1)
        main_layout.addLayout(row1_layout)

        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(18)
        self.card_device = self.create_dashboard_card("事情正在稳定发展")
        self.build_device_card(self.card_device)
        row2_layout.addWidget(self.card_device, stretch=1)
        self.card_todo = self.create_dashboard_card("星空正在点点汇聚")
        self.build_todo_card(self.card_todo)
        row2_layout.addWidget(self.card_todo, stretch=1)
        self.card_life = self.create_dashboard_card("生活也在缓慢向前")
        self.build_life_card(self.card_life)
        row2_layout.addWidget(self.card_life, stretch=1)
        main_layout.addLayout(row2_layout)
        main_layout.addStretch()

        # 把内容塞进滚动区
        scroll_area.setWidget(content_widget)
        outer_layout.addWidget(scroll_area)

    def create_dashboard_card(self, title_text):
        _c = get_theme_colors()
        _is_dark = _is_dark_color(_c["bg"])
        if _is_dark:
            _card_bg     = "rgba(30, 34, 48, 0.6)"
            _card_border = "rgba(255, 255, 255, 0.05)"
            _title_color = "#b0b8c8"
            _dot_color   = "rgba(255, 255, 255, 0.2)"
        else:
            _card_bg     = "rgba(255, 255, 255, 0.9)"
            _card_border = "rgba(0, 0, 0, 0.08)"
            _title_color = "#506070"
            _dot_color   = "rgba(0, 0, 0, 0.2)"

        card = CardWidget()
        card.setStyleSheet(f"""
            CardWidget {{
                background: {_card_bg};
                border: 1px solid {_card_border};
                border-radius: 16px;
            }}
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 14, 16, 14)
        card_layout.setSpacing(10)
        title_row = QHBoxLayout()
        title_label = BodyLabel(title_text)
        title_label.setStyleSheet(
            f"font-weight: 600; font-size: 15px; color: {_title_color};"
        )
        title_row.addWidget(title_label)
        title_row.addStretch()
        for _ in range(3):
            dot_btn = QPushButton("·")
            dot_btn.setFixedSize(20, 20)
            dot_btn.setStyleSheet(
                f"background: transparent; color: {_dot_color}; "
                f"border: none; font-weight: bold;"
            )
            title_row.addWidget(dot_btn)
        card_layout.addLayout(title_row)
        content_layout = QVBoxLayout()
        content_layout.setSpacing(12)
        card_layout.addLayout(content_layout)
        card.main_layout = content_layout
        card.title_label = title_label
        return card

    # ==================== 日历卡片（含年月切换、点击高亮、年月与倒计时齐平） ====================
    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())

    def build_calendar_card(self, card):
        _c = get_theme_colors()
        layout = card.main_layout
        self._clear_layout(layout)

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 4, 0, 4)

        btn_style = f"""
            QPushButton {{
                background: transparent;
                color: {_c['fg_dim']};
                border: 1px solid {_c['border_light']};
                border-radius: 4px;
                font-size: 18px;
            }}
            QPushButton:hover {{
                background: {_c['bg_card']};
                color: {_c['fg']};
            }}
        """

        left_btn = QPushButton("◀")
        left_btn.setFixedSize(32, 32)
        left_btn.setStyleSheet(btn_style)
        left_btn.clicked.connect(lambda: self._change_month(-1))
        top_row.addWidget(left_btn)

        self.month_label = BodyLabel(f"{self.calendar_year}年 {self.calendar_month}月")
        self.month_label.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {_c['fg']}; padding: 0 10px;")
        self.month_label.mousePressEvent = lambda e: self._edit_year_month()
        self.month_label.setCursor(Qt.CursorShape.PointingHandCursor)
        top_row.addWidget(self.month_label, alignment=Qt.AlignmentFlag.AlignCenter)

        right_btn = QPushButton("▶")
        right_btn.setFixedSize(32, 32)
        right_btn.setStyleSheet(btn_style)
        right_btn.clicked.connect(lambda: self._change_month(1))
        top_row.addWidget(right_btn)

        top_row.addStretch()

        right_widget = QWidget()
        right_layout = QHBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        countdown_widget = QWidget()
        countdown_widget.setStyleSheet(f"background: {_c['bg_alt']}; border-radius: 8px; padding: 4px 12px;")
        countdown_widget.mouseDoubleClickEvent = lambda e: self._show_countdown_manager()
        countdown_layout = QHBoxLayout(countdown_widget)
        countdown_layout.setContentsMargins(8, 4, 8, 4)
        countdown_layout.setSpacing(6)
        self.cd_prefix_label = BodyLabel("距离")
        self.cd_name_label = BodyLabel("")
        self.cd_days_label = BodyLabel("")
        self.cd_days_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #4A90E2;")
        self.cd_suffix_label = BodyLabel("天")
        countdown_layout.addWidget(self.cd_prefix_label)
        countdown_layout.addWidget(self.cd_name_label)
        countdown_layout.addWidget(self.cd_days_label)
        countdown_layout.addWidget(self.cd_suffix_label)
        countdown_layout.addStretch()
        right_layout.addWidget(countdown_widget)

        refresh_btn = QPushButton("⟳")
        refresh_btn.setFixedSize(32, 32)
        refresh_btn.setToolTip("回到今天")
        refresh_btn.setStyleSheet(btn_style)
        refresh_btn.clicked.connect(self._reset_to_today)
        right_layout.addWidget(refresh_btn)

        top_row.addWidget(right_widget)
        layout.addLayout(top_row)

        grid_widget = QWidget()
        grid_widget.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        grid_widget.wheelEvent = self._on_calendar_wheel
        grid_layout = QGridLayout(grid_widget)
        grid_layout.setSpacing(6)
        grid_layout.setHorizontalSpacing(12)
        week_list = ["日", "一", "二", "三", "四", "五", "六"]
        for i, w in enumerate(week_list):
            lab = BodyLabel(w)
            lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lab.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 13px;")
            grid_layout.addWidget(lab, 0, i)

        import calendar
        cal = calendar.monthcalendar(self.calendar_year, self.calendar_month)
        for r, week in enumerate(cal):
            for c, day in enumerate(week):
                if day == 0:
                    lab = BodyLabel("")
                    lab.setFixedSize(34, 34)
                    grid_layout.addWidget(lab, r + 1, c)
                    continue
                lab = BodyLabel(str(day))
                lab.setAlignment(Qt.AlignmentFlag.AlignCenter)
                lab.setFixedSize(34, 34)
                is_selected = (self.selected_date.year == self.calendar_year and 
                               self.selected_date.month == self.calendar_month and 
                               self.selected_date.day == day)
                if is_selected:
                    lab.setStyleSheet("background: #7c6df0; border-radius: 8px; color: #fff; font-weight: bold;")
                else:
                    lab.setStyleSheet(f"color: {_c['fg']};")
                lab.mousePressEvent = lambda event, y=day: self._on_date_clicked(y)
                lab.setCursor(Qt.CursorShape.PointingHandCursor)
                grid_layout.addWidget(lab, r + 1, c)
        layout.addWidget(grid_widget)

        self._update_countdown_display()

    def _change_month(self, delta):
        month = self.calendar_month + delta
        year = self.calendar_year
        if month > 12:
            month = 1
            year += 1
        elif month < 1:
            month = 12
            year -= 1
        self.calendar_year = year
        self.calendar_month = month
        # 若选中日期不在新月份，调整为1号
        try:
            day = self.selected_date.day
            max_day = calendar.monthrange(year, month)[1]
            if day > max_day:
                day = max_day
            self.selected_date = datetime(year, month, day).date()
        except:
            self.selected_date = datetime(year, month, 1).date()
        self.refresh_calendar_card()

    def _on_date_clicked(self, day):
        try:
            new_date = datetime(self.calendar_year, self.calendar_month, day).date()
            self.selected_date = new_date
            self.refresh_calendar_card()
        except ValueError:
            pass

    def _edit_year_month(self):
        """单击年月，分别输入年份和月份（更直观）"""
        from PyQt6.QtWidgets import QInputDialog

        year, ok1 = QInputDialog.getInt(self, "修改年份", "请输入年份 (1900-2100):",
                                        self.calendar_year, 1900, 2100, 1)
        if not ok1:
            return

        month, ok2 = QInputDialog.getInt(self, "修改月份", "请输入月份 (1-12):",
                                         self.calendar_month, 1, 12, 1)
        if not ok2:
            return

        try:
            import calendar
            max_day = calendar.monthrange(year, month)[1]
            day = min(self.selected_date.day, max_day)
            self.calendar_year = year
            self.calendar_month = month
            self.selected_date = datetime(year, month, day).date()
            self.refresh_calendar_card()
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无效的日期：{e}")

    def _reset_to_today(self):
        """重置日历到今天的日期并高亮"""
        today = datetime.now().date()
        self.calendar_year = today.year
        self.calendar_month = today.month
        self.selected_date = today
        self.refresh_calendar_card()

    def _on_calendar_wheel(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            # 按住了 Ctrl 键：切换月份
            delta = event.angleDelta().y()
            if delta > 0:
                self._change_month(-1)
            elif delta < 0:
                self._change_month(1)
            event.accept()  # 拦截事件
        else:
            # 没有按 Ctrl 键：忽略事件，把滚轮传递给外层的 QScrollArea
            event.ignore()

    # ==================== 倒计时管理 ====================
    def _load_countdown_data(self):
        cfg = load_config()
        events = cfg.get("countdown_events", [])
        if not events:
            today = datetime.now().date()
            year = today.year
            national_day = datetime(year, 10, 1).date()
            if national_day < today:
                national_day = datetime(year + 1, 10, 1).date()
            events = [{"name": "国庆假期", "date": national_day.strftime("%Y-%m-%d")}]
            cfg["countdown_events"] = events
            cfg["selected_countdown_index"] = 0
            save_config(cfg)
        self.countdown_events = events
        self.selected_index = cfg.get("selected_countdown_index", 0)
        if self.selected_index >= len(events):
            self.selected_index = 0 if events else -1

    def _save_countdown_data(self):
        cfg = load_config()
        cfg["countdown_events"] = self.countdown_events
        cfg["selected_countdown_index"] = self.selected_index
        save_config(cfg)

    def _update_countdown_display(self):
        if self.selected_index < 0 or self.selected_index >= len(self.countdown_events):
            self.cd_name_label.setText("未选择")
            self.cd_days_label.setText("--")
            return
        event = self.countdown_events[self.selected_index]
        name = event["name"]
        date_str = event["date"]
        try:
            target = datetime.strptime(date_str, "%Y-%m-%d").date()
            delta = target - datetime.now().date()
            days = delta.days
            if days < 0:
                days = 0
            self.cd_name_label.setText(name)
            self.cd_days_label.setText(str(days))
        except:
            self.cd_name_label.setText("无效日期")
            self.cd_days_label.setText("--")

    def refresh_calendar_card(self):
        if hasattr(self, 'card_calendar'):
            self.build_calendar_card(self.card_calendar)

    def _show_countdown_manager(self):
        dialog = CountdownManagerDialog(self, self.countdown_events, self.selected_index)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.countdown_events = dialog.events
            self.selected_index = dialog.selected_index
            self._save_countdown_data()
            self.refresh_calendar_card()

    # ==================== 设备监控卡片 ====================
    def build_device_card(self, card):
        _c = get_theme_colors()
        layout = card.main_layout
        row_top = QHBoxLayout()
        left_dev = QVBoxLayout()
        left_dev.setSpacing(6)
        # CPU
        cpu_row = QHBoxLayout()
        cpu_label = BodyLabel("CPU")
        cpu_label.setFixedWidth(50)
        cpu_label.setStyleSheet(f"color: {_c['fg_dim']};")
        cpu_progress = QProgressBar()
        cpu_progress.setRange(0, 100)
        cpu_progress.setValue(0)
        cpu_progress.setFixedWidth(100)
        cpu_progress.setFixedHeight(6)
        cpu_progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(150, 150, 150, 0.2);
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                border-radius: 3px;
            }
        """)
        self.cpu_progress = cpu_progress
        cpu_val = BodyLabel("0%")
        cpu_val.setStyleSheet(f"color: {_c['fg']}; font-weight: 500;")
        self.cpu_label = cpu_val
        cpu_row.addWidget(cpu_label)
        cpu_row.addWidget(cpu_progress)
        cpu_row.addWidget(cpu_val)
        cpu_row.addStretch()
        left_dev.addLayout(cpu_row)
        # 内存
        mem_row = QHBoxLayout()
        mem_label = BodyLabel("内存")
        mem_label.setFixedWidth(50)
        mem_label.setStyleSheet(f"color: {_c['fg_dim']};")
        mem_progress = QProgressBar()
        mem_progress.setRange(0, 100)
        mem_progress.setValue(0)
        mem_progress.setFixedWidth(100)
        mem_progress.setFixedHeight(6)
        mem_progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(150, 150, 150, 0.2);
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #81C784, stop:1 #4CAF50);
                border-radius: 3px;
            }
        """)
        self.mem_progress = mem_progress
        mem_val = BodyLabel("0 GB / 0 GB")
        mem_val.setStyleSheet(f"color: {_c['fg']}; font-weight: 500;")
        self.mem_label = mem_val
        mem_row.addWidget(mem_label)
        mem_row.addWidget(mem_progress)
        mem_row.addWidget(mem_val)
        mem_row.addStretch()
        left_dev.addLayout(mem_row)
        # 磁盘
        disk_row = QHBoxLayout()
        disk_label = BodyLabel("磁盘")
        disk_label.setFixedWidth(50)
        disk_label.setStyleSheet(f"color: {_c['fg_dim']};")
        disk_progress = QProgressBar()
        disk_progress.setRange(0, 100)
        disk_progress.setValue(0)
        disk_progress.setFixedWidth(100)
        disk_progress.setFixedHeight(6)
        disk_progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(150, 150, 150, 0.2);
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #FFB74D, stop:1 #FF9800);
                border-radius: 3px;
            }
        """)
        self.disk_progress = disk_progress
        disk_val = BodyLabel("0 GB 可用")
        disk_val.setStyleSheet(f"color: {_c['fg']}; font-weight: 500;")
        self.disk_label = disk_val
        disk_row.addWidget(disk_label)
        disk_row.addWidget(disk_progress)
        disk_row.addWidget(disk_val)
        disk_row.addStretch()
        left_dev.addLayout(disk_row)
        # 电量
        battery_row = QHBoxLayout()
        battery_label = BodyLabel("电量")
        battery_label.setFixedWidth(50)
        battery_label.setStyleSheet(f"color: {_c['fg_dim']};")
        battery_progress = QProgressBar()
        battery_progress.setRange(0, 100)
        battery_progress.setValue(0)
        battery_progress.setFixedWidth(100)
        battery_progress.setFixedHeight(6)
        battery_progress.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(150, 150, 150, 0.2);
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #E57373, stop:1 #F44336);
                border-radius: 3px;
            }
        """)
        self.battery_progress = battery_progress
        battery_val = BodyLabel("----")
        battery_val.setStyleSheet(f"color: {_c['fg']}; font-weight: 500;")
        self.battery_label = battery_val
        battery_row.addWidget(battery_label)
        battery_row.addWidget(battery_progress)
        battery_row.addWidget(battery_val)
        battery_row.addStretch()
        left_dev.addLayout(battery_row)
        row_top.addLayout(left_dev, stretch=1)
        # 网络速率
        right_dev = QVBoxLayout()
        right_dev.setSpacing(8)
        right_dev.addStretch()
        self.net_down_label = BodyLabel("↓ 0 KB/s")
        self.net_down_label.setFixedWidth(100)
        self.net_down_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.net_down_label.setStyleSheet(f"color: {_c['fg']}; font-size: 16px; font-weight: 500;")
        right_dev.addWidget(self.net_down_label, alignment=Qt.AlignmentFlag.AlignCenter)
        self.net_up_label = BodyLabel("↑ 0 KB/s")
        self.net_up_label.setFixedWidth(100)
        self.net_up_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.net_up_label.setStyleSheet(f"color: {_c['fg']}; font-size: 16px; font-weight: 500;")
        right_dev.addWidget(self.net_up_label, alignment=Qt.AlignmentFlag.AlignCenter)
        right_dev.addStretch()
        row_top.addLayout(right_dev, stretch=0)
        layout.addLayout(row_top)

    # ==================== 待办卡片 ====================
    def build_todo_card(self, card):
        _c = get_theme_colors()
        layout = card.main_layout
        todo_top = QHBoxLayout()
        todo_top.addWidget(BodyLabel("今日积分统计"))
        todo_top.addStretch()
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self.refresh_dashboard_data)
        refresh_btn.setStyleSheet(f"background: {_c['bg_alt']}; border-radius: 6px; padding: 4px 12px; color: {_c['fg']}; border: none;")
        todo_top.addWidget(refresh_btn)
        layout.addLayout(todo_top)
        content = QVBoxLayout()
        self.todo_count_label = BodyLabel("今日已录入：0 条")
        self.todo_count_label.setStyleSheet(f"color: {_c['fg']}; font-size: 16px;")
        self.total_count_label = BodyLabel("累计录入总条数：0 条")
        self.total_count_label.setStyleSheet(f"color: {_c['fg']}; font-size: 16px;")
        self.top_student_label = BodyLabel("积分冠军：暂无")
        self.top_student_label.setStyleSheet(f"color: {_c['fg']}; font-size: 16px;")
        content.addWidget(self.todo_count_label)
        content.addWidget(self.total_count_label)
        content.addWidget(self.top_student_label)
        layout.addLayout(content)

    # ==================== 生活卡片 ====================
    def build_life_card(self, card):
        _c = get_theme_colors()
        layout = card.main_layout
        row_top = QHBoxLayout()
        row_top.addWidget(BodyLabel("系统状态"))
        row_top.addStretch()
        layout.addLayout(row_top)
        content = QVBoxLayout()
        self.life_time_label = BodyLabel("本次运行时间：0 天 0 小时 0 分 0 秒")
        self.life_time_label.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
        content.addWidget(self.life_time_label)
        self.device_total_time_label = BodyLabel("设备总运行：-- 天 -- 小时 -- 分 -- 秒")
        self.device_total_time_label.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
        content.addWidget(self.device_total_time_label)
        self.version_label = BodyLabel(f"版本：{SYSTEM_VERSION}")
        self.version_label.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
        content.addWidget(self.version_label)
        copyright_label = BodyLabel("Copyright © 2026 wuweijia-star  ·  GPLv3")
        copyright_label.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 16px;")
        content.addWidget(copyright_label)
        layout.addLayout(content)

    # ==================== 刷新数据 ====================
    def refresh_dashboard_data(self):
        now = datetime.now()
        self.time_large.setText(now.strftime("%H:%M"))
        week_days = ["一","二","三","四","五","六","日"]
        self.date_sub.setText(
            f"{now.month}月{now.day}日 星期{week_days[now.weekday()]}\n"
            f"四时有序，今日安好"
        )

        # 读取动画开关，消耗动画标记
        self.progress_animation_enabled = load_config().get("enable_progress_animation", True)
        animate = getattr(self, '_progress_needs_animation', False)
        self._progress_needs_animation = False

        if HAS_PSUTIL:
            try:
                cpu = psutil.cpu_percent()
                self.cpu_label.setText(f"{cpu}%")
                self._update_progress(self.cpu_progress, int(cpu), animate=animate)

                mem = psutil.virtual_memory()
                mem_used = mem.used / (1024**3)
                mem_total = mem.total / (1024**3)
                self.mem_label.setText(f"{mem_used:.1f} GB / {mem_total:.1f} GB")
                mem_percent = (mem.used / mem.total) * 100
                self._update_progress(self.mem_progress, int(mem_percent), animate=animate)

                disk = psutil.disk_usage('/')
                disk_free = disk.free / (1024**3)
                disk_total = disk.total / (1024**3)
                self.disk_label.setText(f"{disk_free:.0f} GB 可用")
                disk_percent = (disk.used / disk.total) * 100
                self._update_progress(self.disk_progress, int(disk_percent), animate=animate)

                net = psutil.net_io_counters()
                current_time = time.time()
                time_delta = current_time - self._last_net_time
                if time_delta > 0:
                    down_rate = (net.bytes_recv - self._last_net_recv) / time_delta
                    up_rate = (net.bytes_sent - self._last_net_sent) / time_delta
                    self.net_down_label.setText(f"↓ {down_rate / 1024:.1f} KB/s")
                    self.net_up_label.setText(f"↑ {up_rate / 1024:.1f} KB/s")
                else:
                    self.net_down_label.setText("↓ -- KB/s")
                    self.net_up_label.setText("↑ -- KB/s")
                self._last_net_recv = net.bytes_recv
                self._last_net_sent = net.bytes_sent
                self._last_net_time = current_time

                battery = psutil.sensors_battery()
                if battery:
                    percent = battery.percent
                    plug = "接通" if battery.power_plugged else "电池"
                    self.battery_label.setText(f"{percent}% {plug}")
                    self._update_progress(self.battery_progress, percent, animate=animate)
                else:
                    self.battery_label.setText("电源直供")
                    self._update_progress(self.battery_progress, 0, animate=animate)
            except Exception:
                pass
        else:
            self.cpu_label.setText("N/A")
            self.mem_label.setText("N/A")
            self.disk_label.setText("N/A")
            self.net_up_label.setText("↑ N/A")
            self.net_down_label.setText("↓ N/A")
            self.battery_label.setText("N/A")
            self._update_progress(self.cpu_progress, 0)
            self._update_progress(self.mem_progress, 0)
            self._update_progress(self.disk_progress, 0)
            self._update_progress(self.battery_progress, 0)

        self._update_todo_stats()

        elapsed = time.time() - self.start_time
        days = int(elapsed // 86400)
        hours = int((elapsed % 86400) // 3600)
        minutes = int((elapsed % 3600) // 60)
        seconds = int(elapsed % 60)
        self.life_time_label.setText(f"本次运行：{days} 天 {hours} 小时 {minutes} 分 {seconds} 秒")

        cfg = load_config()
        device_install_ts = cfg.get("install_time", time.time())
        total_device = time.time() - device_install_ts
        d_dev = int(total_device // 86400)
        h_dev = int((total_device % 86400) // 3600)
        m_dev = int((total_device % 3600) // 60)
        s_dev = int(total_device % 60)
        self.device_total_time_label.setText(
            f"设备总运行时间：{d_dev} 天 {h_dev} 小时 {m_dev} 分 {s_dev} 秒"
        )

        hour = now.hour
        minute = now.minute
        time_val = hour + minute / 60.0
        if 0 <= time_val < 5.5:
            greeting = "Late Night\n夜深了，好好休憩，静待朝光。"
        elif 5.5 <= time_val < 8.0:
            greeting = "Early Morning\n早安，迎着晨光，稳步启程。"
        elif 8.0 <= time_val < 11.5:
            greeting = "Morning\n上午好，沉心专注，慢慢来。"
        elif 11.5 <= time_val < 14.0:
            greeting = "Noon\n午安，短暂休整，蓄好力气。"
        elif 14.0 <= time_val < 18.0:
            greeting = "Afternoon\n下午好，稳住节奏，不必焦虑。"
        elif 18.0 <= time_val < 19.5:
            greeting = "Dusk\n傍晚好，落日缓缓，轻卸疲惫。"
        else:
            greeting = "Night\n晚上好，忙碌今日，应当从容。"
        self.greet_label.setText(greeting)

    def showEvent(self, event):
        """切换到首页：恢复刷新计时器并触发一次进度条动画"""
        super().showEvent(event)
        self.progress_animation_enabled = load_config().get("enable_progress_animation", True)
        self._progress_needs_animation = True
        if hasattr(self, '_animating_bars'):
            self._animating_bars.clear()
        # ★ 恢复定时器
        if not self.refresh_timer.isActive():
            self.refresh_timer.start(1000)
        # 延迟触发刷新（进度条动画）
        QTimer.singleShot(100, self.refresh_dashboard_data)

    def hideEvent(self, event):
        """离开首页：暂停刷新计时器，避免在后台拖慢其他页面的切换"""
        super().hideEvent(event)
        self.refresh_timer.stop()


    def _update_progress(self, progress_bar, value, animate=False):
        """更新进度条数值，animate=True 时从 0 缓动到目标值"""
        if progress_bar is None:
            return
        value = int(value)

        # 若该进度条正在动画中，跳过本次普通更新（避免动画被打断）
        cur = getattr(progress_bar, '_current_anim', None)
        if cur is not None and cur.state() == QPropertyAnimation.State.Running:
            return

        if animate:
            animate_progress_bar(progress_bar, value)
        else:
            # 普通更新：若开关开启但用户刚进首页需要动画，这里不动；
            # 若开关关闭或普通刷新，直接赋值
            progress_bar.setValue(value)

        

    def _get_lunar_str(self):
        try:
            import lunardate
            today = datetime.now()
            lunar = lunardate.LunarDate.from_solar_date(today.year, today.month, today.day)
            month_names = ["正", "二", "三", "四", "五", "六",
                           "七", "八", "九", "十", "冬", "腊"]
            if lunar.leap:
                month_str = f"闰{month_names[lunar.month-1]}"
            else:
                month_str = month_names[lunar.month-1]
            day_strs = ["初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", "初九", "初十",
                        "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
                        "廿一", "廿二", "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"]
            day_str = day_strs[lunar.day-1] if 1 <= lunar.day <= 30 else str(lunar.day)
            return f"农历{month_str}月{day_str}"
        except Exception as e:
            print("农历转换失败:", e)
            return "农历日期获取失败"

    def _update_todo_stats(self):
        today_str = datetime.now().strftime("%Y-%m-%d")
        count_today = 0
        total_count = 0

        # ---- 读取配置中的累计总数 ----
        cfg = load_config()
        total_count = cfg.get("total_records_count", 0)

        # ---- 首次迁移：若配置计数为0但CSV有数据，自动初始化 ----
        if total_count == 0 and os.path.exists(DATA_CSV):
            try:
                df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
                if not df.empty:
                    total_count = len(df)
                    cfg["total_records_count"] = total_count
                    save_config(cfg)
            except Exception:
                pass

        # ---- 统计今日录入（仍从CSV实时计算） ----
        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
            if not df.empty:
                df['日期'] = pd.to_datetime(df['日期']).dt.strftime("%Y-%m-%d")
                today_df = df[df['日期'] == today_str]
                count_today = len(today_df)
        except Exception:
            pass

        # ---- 更新界面标签 ----
        self.todo_count_label.setText(f"今日已录入：{count_today} 条")
        self.total_count_label.setText(f"累计录入总条数：{total_count} 条")

        
        # ---- 积分冠军（5 秒缓存，避免每秒读整张 CSV） ----
        now = time.time()
        cache = getattr(self, "_top_student_cache", None)
        if cache and (now - cache["ts"] < 5.0):
            top_student = cache["text"]
        else:
            top_student = "暂无"
            if self.app.stu_df is not None:
                result, _ = calc_total_score(self.app.stu_df, month=None)
                if result is not None and not result.empty:
                    top_row = result.iloc[0]
                    top_student = f"{top_row['姓名']} ({top_row['累计总积分']}分)"
            self._top_student_cache = {"ts": now, "text": top_student}

        self.top_student_label.setText(f"积分冠军：{top_student}")

    def destroyEvent(self, event):
        self.refresh_timer.stop()
        super().destroyEvent(event)

    def do_sign_in(self):
        """执行签到，即使经验已满也可签到，但可能不获得经验"""
        cfg = load_config()
        today = get_today_date()
        last_sign_date = cfg.get("last_sign_in_date", "")
        streak = cfg.get("sign_in_streak", 0)

        # 检查今天是否已签到
        if last_sign_date == today:
            show_toast(self, "今天已签到，明天再来吧！", "info")
            return

        # 计算连续签到天数
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        if last_sign_date == yesterday:
            streak += 1
        else:
            streak = 1  # 断签重置

        # 计算应得经验（但可能因上限而无法获得）
        if streak >= 7:
            exp_gain = SIGN_IN_DAILY_BONUS
        else:
            exp_gain = SIGN_IN_REWARDS.get(streak, 2)

        # 检查今日剩余经验
        remaining = get_remaining_today_exp()
        actual_gain = min(exp_gain, remaining) if remaining > 0 else 0

        # 更新签到状态（无论如何都保存连续天数）
        cfg["sign_in_streak"] = streak
        cfg["last_sign_in_date"] = today
        save_config(cfg)

        # 尝试添加经验（如果 actual_gain > 0）
        if actual_gain > 0:
            success, actual, is_level_up = add_experience(actual_gain, "签到")
            if success:
                streak_text = f"连续签到 {streak} 天" if streak < 7 else "连续签到 7+ 天（满勤）"
                show_toast(
                    self,
                    f"签到成功！+{actual} 经验\n{streak_text}",
                    "success"
                )
                self.app.update_title_with_level()
                if hasattr(self.app, 'settingsInterface'):
                    self.app.settingsInterface.refresh_level_display()
                check_and_unlock_achievements()
                return
            else:
                show_toast(self, "签到失败，请重试", "error")
                return
        else:
            # 经验已满，但签到成功（无经验）
            streak_text = f"连续签到 {streak} 天" if streak < 7 else "连续签到 7+ 天（满勤）"
            show_toast(
                self,
                f"签到成功！但今日经验已满，无法获得经验\n{streak_text}",
                "warning"   # 用警告色提示
            )
            # 仍触发成就检查（可能解锁连续签到成就）
            check_and_unlock_achievements()
            return


# ======================== 倒计时管理对话框 ========================
class CountdownManagerDialog(QDialog):
    def __init__(self, parent, events, selected_index):
        super().__init__(parent)
        _c = get_theme_colors()
        self.parent = parent
        self.events = events.copy()
        self.selected_index = selected_index
        self.setWindowTitle("📌 管理倒计时")
        self.resize(500, 400)
        
        self.setStyleSheet(DIALOG_STYLE + f"""
            QListWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 8px;
                padding: 4px;
                outline: none;
            }}
            QListWidget::item {{
                background: transparent;
                color: {_c['fg']};
                padding: 6px;
            }}
            QListWidget::item:selected {{
                background: {_c['bg_card']};
                color: {_c['fg_bright']};
            }}
            QListWidget::item:hover {{
                background: {_c['bg_card']};
            }}
            QLineEdit, QDateEdit {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                padding: 6px 10px;
                color: {_c['fg']};
            }}
            QLineEdit:focus, QDateEdit:focus {{
                border: 2px solid #64B5F6;
            }}
            QDateEdit::drop-down {{
                border: none;
                width: 20px;
            }}
            QDateEdit QCalendarWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QToolButton {{
                background: {_c['bg_card']};
                color: {_c['fg']};
                border: none;
                padding: 4px;
            }}
            QDateEdit QCalendarWidget QToolButton:hover {{
                background: {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QMenu {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
            }}
            QDateEdit QCalendarWidget QAbstractItemView:enabled {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
                selection-color: white;
            }}
            QDateEdit QCalendarWidget QAbstractItemView:disabled {{
                color: {_c['fg_dim']};
            }}
            QDialogButtonBox QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                border-radius: 6px;
                padding: 6px 16px;
                font-weight: 600;
                border: none;
            }}
            QDialogButtonBox QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
            }}
            QLabel {{
                color: {_c['fg']};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.list_widget.itemDoubleClicked.connect(self.on_item_double_clicked)
        layout.addWidget(self.list_widget)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        add_btn = QPushButton("➕ 添加")
        add_btn.clicked.connect(self.add_event)
        edit_btn = QPushButton("✏️ 编辑")
        edit_btn.clicked.connect(self.edit_event)
        delete_btn = QPushButton("🗑️ 删除")
        delete_btn.clicked.connect(self.delete_event)
        set_btn = QPushButton("✅ 设为当前")
        set_btn.clicked.connect(self.select_item)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)

        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(set_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.refresh_list()

    def on_item_double_clicked(self):
        idx = self.get_selected_index()
        if idx is not None:
            self.selected_index = idx
            self.accept()

    def refresh_list(self):
        _c = get_theme_colors()
        self.list_widget.clear()
        for i, ev in enumerate(self.events):
            name = ev["name"]
            date = ev["date"]
            try:
                target = datetime.strptime(date, "%Y-%m-%d").date()
                delta = target - datetime.now().date()
                days = delta.days
                if days < 0:
                    days = 0
                text = f"{name}  ({days}天)"
            except:
                text = f"{name}  (日期无效)"
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, i)
            if i == self.selected_index:
                item.setBackground(QColor(_c['bg_card']))
            self.list_widget.addItem(item)

    def get_selected_index(self):
        cur = self.list_widget.currentRow()
        if cur < 0:
            QMessageBox.warning(self, "提示", "请先选择一项")
            return None
        return cur

    def select_item(self):
        idx = self.get_selected_index()
        if idx is not None:
            self.selected_index = idx
            self.refresh_list()

    def add_event(self):
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("添加倒计时")
        dialog.resize(300, 200)
        dialog.setStyleSheet(DIALOG_STYLE + f"""
            QLineEdit, QDateEdit {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                padding: 6px 10px;
                color: {_c['fg']};
            }}
            QLineEdit:focus, QDateEdit:focus {{
                border: 2px solid #64B5F6;
            }}
            QDateEdit::drop-down {{
                border: none;
                width: 20px;
            }}
            QDateEdit QCalendarWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QToolButton {{
                background: {_c['bg_card']};
                color: {_c['fg']};
                border: none;
                padding: 4px;
            }}
            QDateEdit QCalendarWidget QToolButton:hover {{
                background: {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QMenu {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
            }}
            QDateEdit QCalendarWidget QAbstractItemView:enabled {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
                selection-color: white;
            }}
            QDateEdit QCalendarWidget QAbstractItemView:disabled {{
                color: {_c['fg_dim']};
            }}
            QDialogButtonBox QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                border-radius: 6px;
                padding: 6px 16px;
                font-weight: 600;
                border: none;
            }}
            QDialogButtonBox QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
            }}
            QLabel {{
                color: {_c['fg']};
            }}
        """)
        layout = QFormLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        name_edit = QLineEdit()
        layout.addRow("事件名称：", name_edit)
        date_edit = QDateEdit(datetime.now().date())
        date_edit.setCalendarPopup(True)
        layout.addRow("目标日期：", date_edit)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addRow(btn_box)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_edit.text().strip()
            if not name:
                QMessageBox.warning(self, "错误", "名称不能为空")
                return
            date = date_edit.date().toString("yyyy-MM-dd")
            self.events.append({"name": name, "date": date})
            self.refresh_list()

    def edit_event(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        ev = self.events[idx]
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑倒计时")
        dialog.resize(300, 200)
        dialog.setStyleSheet(DIALOG_STYLE + f"""
            QLineEdit, QDateEdit {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 6px;
                padding: 6px 10px;
                color: {_c['fg']};
            }}
            QLineEdit:focus, QDateEdit:focus {{
                border: 2px solid #64B5F6;
            }}
            QDateEdit::drop-down {{
                border: none;
                width: 20px;
            }}
            QDateEdit QCalendarWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QToolButton {{
                background: {_c['bg_card']};
                color: {_c['fg']};
                border: none;
                padding: 4px;
            }}
            QDateEdit QCalendarWidget QToolButton:hover {{
                background: {_c['border_light']};
            }}
            QDateEdit QCalendarWidget QMenu {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
            }}
            QDateEdit QCalendarWidget QAbstractItemView:enabled {{
                background: {_c['bg_alt']};
                color: {_c['fg']};
                selection-background-color: #64B5F6;
                selection-color: white;
            }}
            QDateEdit QCalendarWidget QAbstractItemView:disabled {{
                color: {_c['fg_dim']};
            }}
            QDialogButtonBox QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                border-radius: 6px;
                padding: 6px 16px;
                font-weight: 600;
                border: none;
            }}
            QDialogButtonBox QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
            }}
            QLabel {{
                color: {_c['fg']};
            }}
        """)
        layout = QFormLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        name_edit = QLineEdit(ev["name"])
        layout.addRow("事件名称：", name_edit)
        date_edit = QDateEdit(datetime.strptime(ev["date"], "%Y-%m-%d").date())
        date_edit.setCalendarPopup(True)
        layout.addRow("目标日期：", date_edit)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addRow(btn_box)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_edit.text().strip()
            if not name:
                QMessageBox.warning(self, "错误", "名称不能为空")
                return
            date = date_edit.date().toString("yyyy-MM-dd")
            self.events[idx] = {"name": name, "date": date}
            self.refresh_list()

    def delete_event(self):
        idx = self.get_selected_index()
        if idx is None:
            return
        if QMessageBox.question(self, "确认删除", f"确定删除“{self.events[idx]['name']}”吗？",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            del self.events[idx]
            if self.selected_index >= len(self.events):
                self.selected_index = -1
            self.refresh_list()

class RecentRecordsDialog(QDialog):
    """最近录入记录窗口：支持点击勾选删除、双击修改/删除"""

    def __init__(self, parent, app, limit=None):
        super().__init__(parent)
        self.app = app
        if limit is None:
            limit = load_config().get("recent_records_limit", 100)
        self.limit = limit
        self.full_df = pd.DataFrame()
        self.setWindowTitle("📋 最近录入记录")
        self.resize(1000, 600)
        self.setStyleSheet(DIALOG_STYLE + """
            QTableWidget::item { padding: 8px; }
            QTableView::indicator {
                width: 22px;
                height: 22px;
                border-radius: 5px;
                border: 2px solid #5A5A6A;
                background: #1A1A24;
            }
            QTableView::indicator:hover {
                border-color: #64B5F6;
                background: #1E1E2A;
            }
            QTableView::indicator:checked {
                background: #64B5F6;
                border-color: #64B5F6;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # ---- 顶部工具栏 ----
        top = QHBoxLayout()
        info_label = QLabel(
            f"显示最近 {self.limit} 条记录（点击左侧方块勾选，双击行可修改/删除）"
        )
        info_label.setStyleSheet("color: #A0A0B0; font-size: 14px;")
        top.addWidget(info_label)
        top.addStretch()

        select_all_btn = QPushButton("✅ 全选")
        select_all_btn.clicked.connect(self.select_all)
        top.addWidget(select_all_btn)

        unselect_all_btn = QPushButton("◻️ 全不选")
        unselect_all_btn.clicked.connect(self.unselect_all)
        top.addWidget(unselect_all_btn)

        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.clicked.connect(self.load_data)
        top.addWidget(refresh_btn)

        delete_btn = QPushButton("🗑️ 删除选中")
        delete_btn.clicked.connect(self.delete_checked)
        top.addWidget(delete_btn)
        layout.addLayout(top)

        # ---- 表格 ----
        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(
            ["选择", "序号", "日期", "学号", "姓名", "类型", "类别", "分数", "备注"]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        self.table.setColumnWidth(0, 70)
        self.table.setColumnWidth(1, 55)
        self.table.setColumnWidth(2, 105)
        self.table.setColumnWidth(3, 105)
        self.table.setColumnWidth(4, 85)
        self.table.setColumnWidth(5, 65)
        self.table.setColumnWidth(6, 95)
        self.table.setColumnWidth(7, 75)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)   # 行高加大，更好按
        self.table.doubleClicked.connect(self.on_double_click)
        layout.addWidget(self.table)

        # ---- 底部 ----
        bottom = QHBoxLayout()
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #A0A0B0;")
        bottom.addWidget(self.status_label)
        bottom.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        bottom.addWidget(close_btn)
        layout.addLayout(bottom)

        self.load_data()

    # ---------- 点击第一列切换勾选（整列都可点，无需精准点击方块） ----------
    

    def select_all(self):
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(Qt.CheckState.Checked)

    def unselect_all(self):
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(Qt.CheckState.Unchecked)

    # ---------- 数据加载 ----------
    def load_data(self):
        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        except Exception:
            df = pd.DataFrame()
        self.full_df = df

        if df.empty:
            self.table.setRowCount(0)
            self.status_label.setText("暂无记录")
            return

        df["学号"] = df["学号"].astype(str)
        recent = df.tail(self.limit).iloc[::-1]   # 倒序显示最近

        self.table.setRowCount(len(recent))
        for i, (idx, row) in enumerate(recent.iterrows()):
            # 复选框
            chk = QTableWidgetItem()
            chk.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
            )
            chk.setCheckState(Qt.CheckState.Unchecked)
            chk.setData(Qt.ItemDataRole.UserRole, idx)
            chk.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 0, chk)

            def make_item(text):
                item = QTableWidgetItem(str(text) if pd.notna(text) else "")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                return item

            self.table.setItem(i, 1, make_item(i + 1))
            self.table.setItem(i, 2, make_item(row.get("日期", "")))
            self.table.setItem(i, 3, make_item(row.get("学号", "")))
            self.table.setItem(i, 4, make_item(row.get("姓名", "")))
            self.table.setItem(i, 5, make_item(row.get("变动类型", "")))
            self.table.setItem(i, 6, make_item(row.get("类别", "")))
            self.table.setItem(i, 7, make_item(row.get("分数", "")))
            self.table.setItem(i, 8, make_item(row.get("备注", "")))

        self.status_label.setText(f"共显示 {len(recent)} 条记录（总计 {len(df)} 条）")

    # ---------- 勾选删除 ----------
    def get_checked_indices(self):
        result = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.CheckState.Checked:
                idx = item.data(Qt.ItemDataRole.UserRole)
                if idx is not None:
                    result.append(idx)
        return result

    def delete_checked(self):
        indices = self.get_checked_indices()
        if not indices:
            show_toast(self, "请先勾选要删除的记录", "warning")
            return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("确认删除")
        msg_box.setText(f"确定要删除选中的 {len(indices)} 条记录吗？\n此操作不可撤销！")
        msg_box.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)
        if msg_box.exec() != QMessageBox.StandardButton.Yes:
            return

        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
            valid_indices = [i for i in indices if i in df.index]
            df = df.drop(valid_indices).reset_index(drop=True)
            df.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")

            cfg = load_config()
            cfg["total_records_count"] = max(
                0, cfg.get("total_records_count", 0) - len(valid_indices)
            )
            save_config(cfg)

            self.load_data()
            self.app.refresh_ranking()
            show_toast(self, f"已删除 {len(valid_indices)} 条记录", "success")
        except Exception as e:
            show_toast(self, f"删除失败：{e}", "error")

    # ---------- 双击修改/删除（跳过第一列） ----------
    def on_double_click(self, index):
        row = index.row()
        col = index.column()
        # 第一列只用于勾选，不触发修改/删除
        if col == 0:
            return

        chk = self.table.item(row, 0)
        if chk is None:
            return
        idx = chk.data(Qt.ItemDataRole.UserRole)
        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        except Exception:
            return
        if idx not in df.index:
            show_toast(self, "记录已被删除", "warning")
            self.load_data()
            return

        record = df.loc[idx].to_dict()

        dlg = QDialog(self)
        dlg.setWindowTitle("操作记录")
        dlg.setFixedSize(460, 200)
        dlg.setStyleSheet(DIALOG_STYLE)
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(25, 25, 25, 25)
        lay.setSpacing(15)

        info = QLabel(
            f"学号：{record.get('学号', '')}    姓名：{record.get('姓名', '')}\n"
            f"日期：{record.get('日期', '')}    类别：{record.get('类别', '')}    "
            f"类型：{record.get('变动类型', '')}    分数：{record.get('分数', '')}"
        )
        info.setStyleSheet("color: #E0E0E0; font-size: 14px;")
        info.setWordWrap(True)
        lay.addWidget(info)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        edit_btn = QPushButton("✏️ 修改")
        edit_btn.clicked.connect(lambda: (dlg.accept(), self._do_edit(idx)))
        del_btn = QPushButton("🗑️ 删除")
        del_btn.clicked.connect(lambda: (dlg.accept(), self._do_delete(idx)))
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dlg.reject)
        btn_row.addWidget(edit_btn)
        btn_row.addWidget(del_btn)
        btn_row.addStretch()
        btn_row.addWidget(cancel_btn)
        lay.addLayout(btn_row)

        dlg.exec()

    def _do_edit(self, idx):
        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        except Exception as e:
            show_toast(self, f"读取失败：{e}", "error")
            return
        if idx not in df.index:
            show_toast(self, "记录已被删除", "warning")
            self.load_data()
            return
        record = df.loc[idx].to_dict()
        dlg = ModifyRecordDialog(self, record, idx, self.app)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_data()
            self.app.refresh_ranking()
            show_toast(self, "记录已修改", "success")

    def _do_delete(self, idx):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("确认删除")
        msg_box.setText("确定要删除这条记录吗？此操作不可撤销！")
        msg_box.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)
        if msg_box.exec() != QMessageBox.StandardButton.Yes:
            return
        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
            if idx in df.index:
                df = df.drop(idx).reset_index(drop=True)
                df.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
                cfg = load_config()
                cfg["total_records_count"] = max(
                    0, cfg.get("total_records_count", 0) - 1
                )
                save_config(cfg)
                self.load_data()
                self.app.refresh_ranking()
                show_toast(self, "记录已删除", "success")
        except Exception as e:
            show_toast(self, f"删除失败：{e}", "error")
    

# ======================== 录入页面（焦点跳转逻辑移植自老版 + 优化） ========================
class InputInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("inputInterface")
        self.app.set_input_frame(self)

        self.reasons_data = load_config().get("custom_reasons", [])
        self.last_category = load_config().get("last_input_category", "日常")
        self.last_reason = load_config().get("last_input_reason", "")
        self.last_discount = load_config().get("last_input_discount", "100")
        self.fix_params = load_config().get("fix_params_serial", False)

        # 新增：读取快捷开关配置
        self.enable_shortcut = load_config().get("enable_reason_number_shortcut", True)

        self.serial_mode = False
        self.serial_start = None
        self.serial_end = None
        self.current_serial = None
        self.serial_paused = False
        self.skipped_ids = []
        self.serial_timer = None

        self.build_ui()

    def build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(15)

        title_row = QHBoxLayout()
        title = TitleLabel("📝 录入积分", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        title_row.addWidget(title)
        title_row.addStretch()

        self.records_btn = QPushButton("📋 录入记录")
        self.records_btn.setMinimumWidth(160)
        self.records_btn.setFixedHeight(40)
        self.records_btn.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.records_btn.setStyleSheet("""
            QPushButton {
                border-radius: 10px;
                padding: 6px 18px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(100, 181, 246, 0.3);
                font-size: 16px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
                border-color: #64B5F6;
            }
            QPushButton:pressed {
                background: #1E88E5;
            }
        """)
        self.records_btn.clicked.connect(self.open_recent_records)
        title_row.addWidget(self.records_btn)

        main_layout.addLayout(title_row)

        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form_layout.setSpacing(12)

        self.serial_check = QCheckBox("连续模式（自动跳到下一个学号）")
        self.serial_check.toggled.connect(self.toggle_serial_mode)
        form_layout.addRow(self.serial_check)

        self.stu_id_combo = QComboBox()
        self.stu_id_combo.setEditable(True)
        self.stu_id_combo.lineEdit().setPlaceholderText("输入学号")
        self.stu_id_combo.currentTextChanged.connect(self.on_id_changed)
        self.update_student_list()
        self.name_label = QLabel("")
        self.name_label.setStyleSheet("color: #A0A0A0;")
        id_layout = QHBoxLayout()
        id_layout.addWidget(self.stu_id_combo)
        id_layout.addWidget(self.name_label)
        self.single_widget = QWidget()
        self.single_widget.setLayout(id_layout)
        form_layout.addRow("学号：", self.single_widget)

        self.serial_widget = QWidget()
        serial_layout = QHBoxLayout(self.serial_widget)
        serial_layout.setContentsMargins(0, 0, 0, 0)
        self.start_serial = QLineEdit()
        self.start_serial.setPlaceholderText("起始学号")
        self.end_serial = QLineEdit()
        self.end_serial.setPlaceholderText("结束学号")
        self.serial_status = QLabel("未开始")
        self.serial_status.setStyleSheet("color: #6C63FF; font-weight: bold;")
        # ===== 修改后 =====
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.skip_btn = QPushButton("跳过")        # ★ 新增
        self.skip_btn.clicked.connect(self.skip_serial)   # ★ 新增
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.stop_serial)
        self.fix_check = QCheckBox("固定参数")
        self.fix_check.setChecked(self.fix_params)
        self.fix_check.toggled.connect(self.on_fix_toggled)
        serial_layout.addWidget(QLabel("起始："))
        serial_layout.addWidget(self.start_serial)
        serial_layout.addWidget(QLabel("结束："))
        serial_layout.addWidget(self.end_serial)
        serial_layout.addWidget(self.serial_status)
        serial_layout.addWidget(self.pause_btn)
        serial_layout.addWidget(self.skip_btn)     # ★ 新增
        serial_layout.addWidget(self.stop_btn)
        serial_layout.addWidget(self.fix_check)
        self.serial_widget.setVisible(False)
        form_layout.addRow(self.serial_widget)

        self.date_edit = QLineEdit()
        self.date_edit.setText(datetime.now().strftime("%Y-%m-%d"))
        form_layout.addRow("日期：", self.date_edit)

        self.op_group = QButtonGroup(self)
        self.op_plus = QRadioButton("加分")
        self.op_minus = QRadioButton("扣分")
        self.op_plus.setChecked(True)
        self.op_group.addButton(self.op_plus, 1)
        self.op_group.addButton(self.op_minus, 2)
        op_widget = QWidget()
        op_layout = QHBoxLayout(op_widget)
        op_layout.setContentsMargins(0, 0, 0, 0)
        op_layout.addWidget(self.op_plus)
        op_layout.addWidget(self.op_minus)
        form_layout.addRow("类型：", op_widget)

        self.category_combo = ComboBox()
        self.category_combo.addItems(["日常", "听写", "背诵", "班干加分", "考试加分"])
        self.category_combo.setCurrentText(self.last_category)
        form_layout.addRow("类别：", self.category_combo)

        self.score_edit = QLineEdit()
        self.score_edit.setPlaceholderText("输入分数")
        form_layout.addRow("分数：", self.score_edit)

        self.assign_edit = QLineEdit()
        self.assign_edit.setPlaceholderText("可选")
        form_layout.addRow("赋分总分：", self.assign_edit)

        self.discount_edit = QLineEdit()
        self.discount_edit.setText(self.last_discount)
        form_layout.addRow("折扣比例（%）：", self.discount_edit)

        self.reason_combo = QComboBox()
        self.reason_combo.setEditable(True)
        self.reason_combo.lineEdit().setPlaceholderText("输入或选择事由")
        self.reason_combo.currentTextChanged.connect(self.on_reason_changed)
        self.refresh_reasons()
        form_layout.addRow("事由：", self.reason_combo)

        self.keep_student = QCheckBox("保留学号")
        self.keep_reason = QCheckBox("保留事由")
        keep_widget = QWidget()
        keep_layout = QHBoxLayout(keep_widget)
        keep_layout.setContentsMargins(0, 0, 0, 0)
        keep_layout.addWidget(self.keep_student)
        keep_layout.addWidget(self.keep_reason)
        form_layout.addRow(keep_widget)

        btn_layout = QHBoxLayout()
        self.submit_btn = QPushButton("确认提交 (Ctrl+Enter)")
        self.submit_btn.clicked.connect(self.submit)
        self.clear_btn = QPushButton("清空表单")
        self.clear_btn.clicked.connect(self.clear_form)
        self.back_btn = QPushButton("返回主菜单")
        self.back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        btn_layout.addWidget(self.submit_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.back_btn)
        form_layout.addRow(btn_layout)

        main_layout.addWidget(form_widget)
        main_layout.addStretch()
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        # 初始化 Enter 键绑定
        self._setup_enter_bindings()
        self._setup_tab_order()
        
        self.reason_combo.installEventFilter(self)
        # 【新增】初始重置，保证进入时表单为默认状态
        self.reset_form()

    # ========== 新增：重置表单（核心） ==========
    def reset_form(self):
        """重置录入表单到默认状态（类别=日常，事由=无，清空学号/姓名/分数/赋分，关闭复选框）"""
        self.category_combo.setCurrentText("日常")
        self.reason_combo.setCurrentText("无")
        self.stu_id_combo.setCurrentText("")
        self.name_label.setText("")
        self.score_edit.clear()
        self.assign_edit.clear()

        # 关闭保留选项和固定参数
        self.keep_student.setChecked(False)
        self.keep_reason.setChecked(False)
        self.fix_check.setChecked(False)

        # 如果处于连续模式，同时复位连续状态
        if self.serial_mode:
            self.start_serial.clear()
            self.end_serial.clear()
            self.current_serial = None
            self.serial_status.setText("未开始")
            self.skipped_ids = []
            self.serial_paused = False
            self.pause_btn.setText("暂停")
            if self.serial_timer:
                self.serial_timer.stop()
                self.serial_timer = None

    # ========== 界面显示事件（每次进入时重置） ==========
    def showEvent(self, event):
        super().showEvent(event)
        self.reset_form()
        # 【关键】确保过滤器安装到组合框本身（而不是 lineEdit）
        self.reason_combo.installEventFilter(self)
        if self.serial_mode:
            QTimer.singleShot(50, self.start_serial.setFocus)
        else:
            QTimer.singleShot(50, lambda: self.stu_id_combo.lineEdit().setFocus())

    def open_recent_records(self):
        """打开最近录入记录窗口（条数从配置读取）"""
        dialog = RecentRecordsDialog(self, self.app)
        dialog.exec()
        self.app.refresh_ranking()

    # ---------- 键盘事件 ----------
    def keyPressEvent(self, event):
        # Ctrl+Enter 提交
        if event.key() == Qt.Key.Key_Return and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.submit()
            return
        # Esc 返回首页
        if event.key() == Qt.Key.Key_Escape:
            self.app.switchTo(self.app.homeInterface)
            return
        # 其他按键交给父类处理
        super().keyPressEvent(event)

    # ---------- 学号列表更新 ----------
    def update_student_list(self):
        ids = list(self.app.id_to_name.keys()) if self.app.id_to_name else []
        # 按学号数字大小排序（若学号非纯数字则回退为字符串排序）
        try:
            ids.sort(key=lambda x: int(x))
        except (ValueError, TypeError):
            ids.sort()
        self.stu_id_combo.clear()
        # 构建 姓名 -> [学号...] 映射（支持重名）
        self.name_to_ids = {}
        for sid in ids:
            name = self.app.id_to_name.get(sid, "")
            self.stu_id_combo.addItem(f"{sid}  {name}")
            self.name_to_ids.setdefault(name, []).append(sid)
        # 配置补全器：开启包含匹配，输入姓名/学号片段均可弹出候选
        try:
            from PyQt6.QtWidgets import QCompleter
            completer = self.stu_id_combo.completer()
            if completer is not None:
                completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
                completer.setFilterMode(Qt.MatchFlag.MatchContains)
                completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

                # ----- 统一补全弹出列表的样式（与 QComboBox 下拉一致） -----
                popup = completer.popup()
                if popup is not None:
                    _c = get_theme_colors()
                    popup.setStyleSheet(f"""
                        QListView {{
                            background: {_c['bg_alt']};
                            border: 1px solid {_c['border_light']};
                            border-radius: 10px;
                            padding: 4px;
                            outline: none;
                            color: {_c['fg']};
                            font-size: 16px;
                        }}
                        QListView::item {{
                            padding: 8px 12px;
                            border-radius: 6px;
                            color: {_c['fg']};
                        }}
                        QListView::item:selected {{
                            background: #64B5F6;
                            color: white;
                        }}
                        QListView::item:hover {{
                            background: {_c['bg_card']};
                        }}
                    """)

                # ----- 连接 activated：从补全列表选中后规范化并切换焦点 -----
                try:
                    completer.activated.disconnect()
                except TypeError:
                    pass
                completer.activated.connect(
                    lambda _text: self._on_stu_id_enter()
                )
        except Exception:
            pass

    def _resolve_student_id(self, text):
        """从任意输入文本解析学号，支持：
           1) 纯学号
           2) '学号  姓名'（下拉项格式）
           3) 纯姓名（完整）
           4) '姓名 学号'
           5) 姓名部分匹配（前缀优先，其次子串）
        """
        if not text:
            return ""
        text = text.strip()
        # 1. 整体就是学号
        if text in self.app.id_to_name:
            return text
        parts = text.split()
        # 2. 下拉项格式 '学号  姓名'，取第一段
        if parts and parts[0] in self.app.id_to_name:
            return parts[0]
        # 3. 输入的是完整姓名（含重名取第一个）
        if hasattr(self, 'name_to_ids') and text in self.name_to_ids:
            ids = self.name_to_ids[text]
            if ids:
                return ids[0]
        # 4. '姓名 学号' 形式，最后一段为学号
        if len(parts) >= 2 and parts[-1] in self.app.id_to_name:
            return parts[-1]
        # 5. 部分姓名匹配：优先前缀匹配，其次子串匹配
        if hasattr(self, 'name_to_ids'):
            for name, ids in self.name_to_ids.items():
                if name.startswith(text) and ids:
                    return ids[0]
            for name, ids in self.name_to_ids.items():
                if text in name and ids:
                    return ids[0]
        return ""

    def get_current_student_id(self):
        """从学号组合框获取当前学号（支持学号/姓名/混合输入）"""
        return self._resolve_student_id(self.stu_id_combo.currentText())

    def on_id_changed(self, text):
        if hasattr(self, 'name_label') and self.name_label is not None:
            stu_id = self._resolve_student_id(text)
            if stu_id and stu_id in self.app.id_to_name:
                name = self.app.id_to_name[stu_id]
                # 重名提示
                if hasattr(self, 'name_to_ids') and len(self.name_to_ids.get(name, [])) > 1:
                    self.name_label.setText(f"👤{stu_id} {name}（重名）")
                else:
                    self.name_label.setText(f"👤{stu_id} {name}")
            else:
                self.name_label.setText("")

    # ---------- 模式切换（增加重置） ----------
    def toggle_serial_mode(self, checked):
        self.serial_mode = checked
        self.single_widget.setVisible(not checked)
        self.serial_widget.setVisible(checked)
        # 切换时重置表单
        self.reset_form()
        if checked:
            QTimer.singleShot(50, self.start_serial.setFocus)
        else:
            QTimer.singleShot(50, lambda: self.stu_id_combo.lineEdit().setFocus())
        # 重新绑定 Enter 键
        self._setup_enter_bindings()
        self._setup_tab_order()

    # ---------- 连续模式控制 ----------
    def toggle_pause(self):
        if not self.serial_mode or self.current_serial is None:
            return
        self.serial_paused = not self.serial_paused
        self.pause_btn.setText("继续" if self.serial_paused else "暂停")
        if not self.serial_paused:
            self.do_serial_submit()

    def stop_serial(self):
        if not self.serial_mode or self.current_serial is None:
            return
        if QMessageBox.question(self, "停止连续录入", "确定要停止连续录入吗？已录入的将保存。",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            if self.serial_timer:
                self.serial_timer.stop()
                self.serial_timer = None
            self.current_serial = None
            self.serial_status.setText("已停止")
            self.pause_btn.setText("暂停")
            self.serial_paused = False
            self.skipped_ids = []
            # 清空起始和结束学号
            self.start_serial.clear()
            self.end_serial.clear()
            show_toast(self, "⏹ 连续录入已停止", "info")

    def _finish_serial(self):
        """连续录入全部完成：弹通知、清空表单与起止学号、复位状态"""
        msg = "✅ 连续录入完成"
        if self.skipped_ids:
            msg += f"，跳过学号：{', '.join(self.skipped_ids)}"
        show_toast(self, msg, "success")

        # 清空起始和结束学号
        self.start_serial.clear()
        self.end_serial.clear()

        # 复位内部状态
        self.current_serial = None
        self.serial_status.setText("已完成")
        self.pause_btn.setText("暂停")
        self.serial_paused = False
        self.skipped_ids = []

        # 清空录入内容
        self.score_edit.clear()
        self.assign_edit.clear()
        self.reason_combo.setCurrentText("无")
        self.date_edit.setText(datetime.now().strftime("%Y-%m-%d"))

    def skip_serial(self):
        """跳过当前学号，直接跳到下一个（不写入任何记录）"""
        # 未进入连续模式或尚未开始时提示
        if not self.serial_mode or self.current_serial is None:
            show_toast(self, "连续模式未开始，请先输入起止学号并提交", "info")
            return

        # 已完成所有学号
        if self.current_serial > self.serial_end:
            show_toast(self, "已完成所有学号，无需跳过", "info")
            return

        # 记录被跳过的学号
        stu_id = str(self.current_serial)
        self.skipped_ids.append(stu_id)
        show_toast(self, f"⏭ 已跳过学号 {stu_id}", "info")

        # 学号 +1，刷新状态显示
        self.current_serial += 1
        if self.current_serial > self.serial_end:
            # ★ 跳过后已全部完成 → 立即自动结束
            self._finish_serial()
            return

        self.serial_status.setText(f"当前：{self.current_serial}")

        # 非固定参数模式下清空表单（与原有跳过逻辑保持一致）
        if not self.fix_params:
            if not self.keep_reason.isChecked():
                self.score_edit.clear()
                self.assign_edit.clear()
                self.reason_combo.setCurrentText("无")

        # 处理后续流程（固定参数下会自动提交下一个；否则聚焦到事由框）
        self._after_submit_serial()

    def on_fix_toggled(self, checked):
        self.fix_params = checked
        cfg = load_config()
        cfg["fix_params_serial"] = checked
        save_config(cfg)

    # ---------- 事由管理 ----------
    def refresh_reasons(self):
        cfg = load_config()
        self.reasons_data = cfg.get("custom_reasons", [])
        names = [r["name"] for r in self.reasons_data]
        names.insert(0, "无")
        self.reason_combo.clear()
        self.reason_combo.addItems(names)
        self.reason_combo.setCurrentText("无")
        
        # 【关键】安装过滤器到组合框（而不是其 lineEdit）
        self.reason_combo.installEventFilter(self)
        self.enable_shortcut = cfg.get("enable_reason_number_shortcut", True)

    def on_reason_changed(self, text):
        if text == "无" or not text:
            return
        for r in self.reasons_data:
            if r["name"] == text:
                # 设置类型
                if r["type"] == "+":
                    self.op_plus.setChecked(True)
                else:
                    self.op_minus.setChecked(True)
                # 设置分数
                self.score_edit.setText(str(r["score"]))
                # 设置类别（如果有匹配）
                category = r.get("category", "")
                if category:
                    self.category_combo.setCurrentText(category)
                break

    # ---------- 清空表单（保留原逻辑） ----------
    def clear_form(self):
        if not self.serial_mode:
            if not self.keep_student.isChecked():
                self.stu_id_combo.setCurrentText("")
                self.name_label.setText("")
            self.score_edit.clear()
            self.assign_edit.clear()
            if not self.keep_reason.isChecked():
                self.reason_combo.setCurrentText("无")
        else:
            self.score_edit.clear()
            self.assign_edit.clear()
            if not self.keep_reason.isChecked():
                self.reason_combo.setCurrentText("无")

    # ---------- 提交 ----------
    def submit(self):
        if self.serial_mode:
            self.do_serial_submit()
        else:
            self.do_single_submit()

    def do_single_submit(self):
        date_str = self.date_edit.text().strip()
        if not date_str:
            show_toast(self, " 日期不能为空", "error")
            return
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except:
            show_toast(self, " 日期格式应为 YYYY-MM-DD", "error")
            return

        stu_id = self.get_current_student_id()
        if stu_id not in self.app.id_to_name:
            show_toast(self, " 学号不存在", "error")
            return
        name = self.app.id_to_name[stu_id]

        op = "+" if self.op_plus.isChecked() else "-"
        category = self.category_combo.currentText()
        try:
            score_val = float(self.score_edit.text().strip())
        except:
            show_toast(self, " 请输入有效分数", "error")
            return
        score = score_val if op == "+" else -score_val

        try:
            discount = float(self.discount_edit.text().strip())
        except:
            show_toast(self, " 折扣比例必须为数字", "error")
            return
        if discount < 0:
            show_toast(self, " 折扣不能为负数", "error")
            return
        actual_score = round(score * (discount / 100.0), 2)

        assign_str = self.assign_edit.text().strip()
        assign_score = None
        if assign_str:
            try:
                assign_score = float(assign_str)
                assign_score = round(assign_score * (discount / 100.0), 2)
            except:
                show_toast(self, " 赋分总分必须为数字", "error")
                return

        note = self.reason_combo.currentText().strip()
        if note == "无":
            note = ""
        if discount != 100:
            note = f"{note} (折扣{discount}%)" if note else f"折扣{discount}%"

        ok, msg = save_score_record(stu_id, name, date_str, op, category, actual_score, note, assign_score)
        if ok:
            show_toast(self, msg, "success")
            self.app.refresh_ranking()
            cfg = load_config()
            cfg["last_input_category"] = category
            cfg["last_input_reason"] = note
            cfg["last_input_discount"] = str(discount)
            save_config(cfg)

            if not self.keep_student.isChecked():
                self.stu_id_combo.setCurrentText("")
                self.name_label.setText("")
            if not self.keep_reason.isChecked():
                self.score_edit.clear()
                self.assign_edit.clear()
                self.reason_combo.setCurrentText("无")
            self.date_edit.setText(datetime.now().strftime("%Y-%m-%d"))
            # 提交后聚焦学号框（普通模式）
            if self.keep_student.isChecked() and self.reason_combo.currentText() != "无":
                self.score_edit.setFocus()
                self.score_edit.selectAll()
            else:
                self.stu_id_combo.setFocus()
        else:
            show_toast(self, msg, "error")

    def do_serial_submit(self):
        if self.serial_paused:
            show_toast(self, "⏸ 已暂停，请点击「继续」", "info")
            return

        # 初始化连续模式
        if self.current_serial is None:
            start_str = self.start_serial.text().strip()
            end_str = self.end_serial.text().strip()
            if not start_str or not end_str:
                show_toast(self, " 请输入起始和结束学号", "error")
                return
            try:
                self.serial_start = int(start_str)
                self.serial_end = int(end_str)
            except:
                show_toast(self, " 学号必须为整数", "error")
                return
            if self.serial_start > self.serial_end:
                show_toast(self, " 起始学号不能大于结束学号", "error")
                return
            self.current_serial = self.serial_start
            self.skipped_ids = []
            # ★ 修改：直接显示第一个待录学号
            self.serial_status.setText(f"当前：{self.current_serial}")
            self.pause_btn.setText("暂停")

        # 检查是否完成（兜底，例如暂停后继续时）
        if self.current_serial > self.serial_end:
            self._finish_serial()
            return

        stu_id = str(self.current_serial)
        # 学号不在名单中 → 跳过
        if stu_id not in self.app.id_to_name:
            self.skipped_ids.append(stu_id)
            show_toast(self, f"⚠️ 学号 {stu_id} 不在名单中，已跳过", "warning")
            self.current_serial += 1
            if self.current_serial > self.serial_end:
                # ★ 跳过之后已全部完成 → 立即自动结束
                self._finish_serial()
                return
            self.serial_status.setText(f"当前：{self.current_serial}")
            if not self.fix_params:
                if not self.keep_reason.isChecked():
                    self.score_edit.clear()
                    self.assign_edit.clear()
                    self.reason_combo.setCurrentText("无")
            # 跳过学号后，继续下一个（自动或手动）
            self._after_submit_serial()
            return

        # 准备提交当前学号
        date_str = self.date_edit.text().strip()
        if not date_str:
            date_str = datetime.now().strftime("%Y-%m-%d")
            self.date_edit.setText(date_str)
        op = "+" if self.op_plus.isChecked() else "-"
        category = self.category_combo.currentText()
        try:
            score_val = float(self.score_edit.text().strip())
        except:
            show_toast(self, f" 学号 {stu_id} 分数无效，跳过", "error")
            self.current_serial += 1
            if self.current_serial > self.serial_end:
                # ★ 跳过之后已全部完成 → 立即自动结束
                self._finish_serial()
                return
            self.serial_status.setText(f"当前：{self.current_serial}")
            if not self.fix_params and not self.keep_reason.isChecked():
                self.score_edit.clear()
                self.assign_edit.clear()
                self.reason_combo.setCurrentText("无")
            self._after_submit_serial()
            return
        score = score_val if op == "+" else -score_val
        try:
            discount = float(self.discount_edit.text().strip())
        except:
            discount = 100
        actual_score = round(score * (discount / 100.0), 2)

        assign_str = self.assign_edit.text().strip()
        assign_score = None
        if assign_str:
            try:
                assign_score = float(assign_str)
                assign_score = round(assign_score * (discount / 100.0), 2)
            except:
                assign_score = None

        note = self.reason_combo.currentText().strip()
        if note == "无":
            note = ""
        if discount != 100:
            note = f"{note} (折扣{discount}%)" if note else f"折扣{discount}%"

        name = self.app.id_to_name[stu_id]
        ok, msg = save_score_record(stu_id, name, date_str, op, category, actual_score, note, assign_score)
        if ok:
            show_toast(self, f" 学号 {stu_id} {name} 录入成功", "success")
            self.app.refresh_ranking()
        else:
            show_toast(self, f" 学号 {stu_id} 录入失败：{msg}", "error")

        # 提交完成，递增学号
        self.current_serial += 1
        if self.current_serial > self.serial_end:
            # ★ 最后一个学号已录完 → 立即自动结束
            self._finish_serial()
            return

        self.serial_status.setText(f"当前：{self.current_serial}")

        # 处理后续操作
        self._after_submit_serial()

    def _after_submit_serial(self):
        if self.fix_params:
            if not self.serial_paused:
                QTimer.singleShot(100, self.do_serial_submit)
        else:
            if self.keep_reason.isChecked():
                self.score_edit.setFocus()
                self.score_edit.selectAll()
            else:
                self.score_edit.clear()
                self.assign_edit.clear()
                self.reason_combo.setCurrentText("无")
                self.reason_combo.setFocus()
                if not self.enable_shortcut:      # 仅当快捷模式关闭时自动弹出
                    self._open_reason_dropdown()

    # ---------- 焦点跳转逻辑（移植自老版 + 优化） ----------
    def _setup_enter_bindings(self):
        # 清除旧的 returnPressed 连接
        for widget in [self.stu_id_combo.lineEdit(), self.reason_combo.lineEdit(),
                    self.score_edit, self.start_serial, self.end_serial]:
            try:
                widget.returnPressed.disconnect()
            except TypeError:
                pass

        if self.serial_mode:
            self.start_serial.returnPressed.connect(self._on_start_serial_enter)
            self.end_serial.returnPressed.connect(self._on_end_serial_enter)
            self.reason_combo.lineEdit().returnPressed.connect(self._select_reason_and_focus_score)
            self.score_edit.returnPressed.connect(self.submit)
        else:
            self.stu_id_combo.lineEdit().returnPressed.connect(self._on_stu_id_enter)
            self.reason_combo.lineEdit().returnPressed.connect(self._select_reason_and_focus_score)
            self.score_edit.returnPressed.connect(self.submit)

        # ---------- 新增：下拉列表选择后自动跳转到分数框 ----------
        try:
            self.reason_combo.activated.disconnect()
        except TypeError:
            pass
        self.reason_combo.activated.connect(self._focus_score)

    def _setup_tab_order(self):
        """设置 Tab 键顺序（按添加顺序已合理，无需额外设置）"""
        pass

    # ---------- 辅助焦点方法 ----------
    def _on_stu_id_enter(self):
        if self.serial_mode:
            return
        # 规范化输入：无论输入的是学号、姓名还是混合输入，统一为 "学号  姓名"
        # （输入汉字姓名时，此处会自动填入列表第一个匹配项）
        stu_id = self._resolve_student_id(self.stu_id_combo.currentText())
        if stu_id and stu_id in self.app.id_to_name:
            name = self.app.id_to_name[stu_id]
            normalized = f"{stu_id}  {name}"
            if self.stu_id_combo.currentText() != normalized:
                self.stu_id_combo.setCurrentText(normalized)

        # 切换焦点：若事由已填则跳分数，否则跳事由
        reason = self.reason_combo.currentText().strip()
        if reason and reason != "无":
            self.score_edit.setFocus()
            self.score_edit.selectAll()
        else:
            self.reason_combo.setFocus()
            if not self.enable_shortcut:
                self._open_reason_dropdown()

    def _select_reason_and_focus_score(self):
        """事由按 Enter：选中事由并跳转到分数（优化为一次 Enter 生效）"""
        choice = self.reason_combo.currentText().strip()
        if choice and choice != "无":
            self.on_reason_changed(choice)
        # 延迟聚焦，确保下拉菜单已完全关闭
        QTimer.singleShot(0, self._focus_score)

    def _focus_score(self):
        self.score_edit.setFocus()
        self.score_edit.selectAll()

    def _on_start_serial_enter(self):
        """连续模式：起始学号按 Enter 跳转到结束学号"""
        self.end_serial.setFocus()
        self.end_serial.selectAll()

    def _on_end_serial_enter(self):
        self.reason_combo.setFocus()
        if not self.enable_shortcut:      # 仅当快捷模式关闭时自动弹出
            self._open_reason_dropdown()

    def _open_reason_dropdown(self):
        """展开事由下拉列表"""
        if self.reason_combo:
            self.reason_combo.showPopup()

    def eventFilter(self, obj, event):
        from PyQt6.QtCore import QEvent

        # 只处理组合框或其编辑框的事件
        if obj == self.reason_combo or obj == self.reason_combo.lineEdit():
            if event.type() == QEvent.Type.KeyPress:
                if self.enable_shortcut:
                    key = event.key()
                    # 主键盘数字 1~9 的键码范围：49~57
                    if 49 <= key <= 57:
                        idx = key - 49 + 1   # 49 -> 1, 50 -> 2, ..., 57 -> 9
                    # 小键盘数字 1~9 的键码范围：97~105
                    elif 97 <= key <= 105:
                        idx = key - 97 + 1   # 97 -> 1, 98 -> 2, ..., 105 -> 9
                    else:
                        idx = None

                    if idx is not None:
                        count = self.reason_combo.count()
                        if idx < count:
                            self.reason_combo.setCurrentIndex(idx)
                            self.on_reason_changed(self.reason_combo.currentText())
                            QTimer.singleShot(0, self._focus_score)
                        return True  # 阻止数字输入
                return False  # 其他按键放行
        return super().eventFilter(obj, event)

    def update_shortcut_config(self):
        """由外部调用，重新读取快捷开关配置"""
        self.enable_shortcut = load_config().get("enable_reason_number_shortcut", True)

    

# ======================== 排行页面（新增搜索高亮、Enter 跳转下一个） ========================
class RankingInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("rankingInterface")
        self.search_results = []      # 匹配的行号列表
        self.search_index = -1        # 当前高亮索引
        self.last_keyword = ""        # 上次搜索的关键词
        self.build_ui()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        top = QHBoxLayout()
        title = TitleLabel("📊 积分排行榜", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold;color: {_c['fg']};")
        top.addWidget(title)
        top.addStretch()

        # ---- 搜索框 ----
        top.addWidget(QLabel("搜索："))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("学号或姓名")
        self.search_edit.setFixedWidth(200)
        self.search_edit.returnPressed.connect(self.on_search_enter)   # 回车触发智能搜索/跳转
        top.addWidget(self.search_edit)

        self.search_btn = QPushButton("🔍 搜索")
        self.search_btn.clicked.connect(self.search_student)
        top.addWidget(self.search_btn)

        self.next_btn = QPushButton("下一个")
        self.next_btn.clicked.connect(self.next_match)
        self.next_btn.setEnabled(False)
        top.addWidget(self.next_btn)

        self.clear_btn = QPushButton("清除高亮")
        self.clear_btn.clicked.connect(self.clear_highlight)
        top.addWidget(self.clear_btn)

        layout.addLayout(top)

        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self.on_double_click)
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.on_header_double_clicked)
        layout.addWidget(self.table)

        self.stats_label = QLabel("")
        self.stats_label.setStyleSheet("color: #A0A0B0; font-size: 15px;")
        layout.addWidget(self.stats_label, alignment=Qt.AlignmentFlag.AlignCenter)

        back_btn = QPushButton("🔙 返回主菜单")
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.load_data()

    def refresh(self):
        """外部调用的刷新方法，重新加载全部数据并清除搜索高亮"""
        self.clear_highlight()
        self.load_data()

    def showEvent(self, event):
        super().showEvent(event)
        # 让 FluentWindow 先把页面切过去、画完首帧，再刷新数据
        QTimer.singleShot(0, self._deferred_refresh)

    def _deferred_refresh(self):
        if not self.isVisible():
            return
        self.refresh()

    def load_data(self):
        if self.app.stu_df is None:
            show_toast(self, "请先加载名单", "error")
            return

        result, err = calc_total_score(self.app.stu_df, month=None)
        if result is None:
            show_toast(self, err, "warning")
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.stats_label.setText("暂无数据")
            return

        columns = ["排名", "学号", "姓名", "日常", "听写", "背诵", "班干加分", "考试加分", "赋分总分", "累计总积分"]
        self.table.setUpdatesEnabled(False)          # ← 新增
        try:
            self.table.setColumnCount(len(columns))
            self.table.setHorizontalHeaderLabels(columns)
            self.table.setRowCount(len(result))
            for i, row in result.iterrows():
                for j, col in enumerate(columns):
                    val = row[col]
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.table.setItem(i, j, item)
            for i, row in result.iterrows():
                self.table.item(i, 1).setData(Qt.ItemDataRole.UserRole, row["学号"])
        finally:
            self.table.setUpdatesEnabled(True)       # ← 新增

        total_scores = result["累计总积分"]
        if not total_scores.empty:
            max_score = total_scores.max()
            min_score = total_scores.min()
            avg_score = total_scores.mean()
            max_row = result[result["累计总积分"] == max_score].iloc[0]
            min_row = result[result["累计总积分"] == min_score].iloc[0]
            stats_text = f"最高：{max_row['姓名']}（{max_row['学号']}）{max_score}分 | 最低：{min_row['姓名']}（{min_row['学号']}）{min_score}分 | 平均：{avg_score:.2f}分"
            self.stats_label.setText(stats_text)
        else:
            self.stats_label.setText("暂无积分数据")

        # 清除搜索高亮（因为数据重新加载）
        self.clear_highlight()

    # ---------- 搜索与高亮 ----------
    def clear_highlight(self):
        """清除所有高亮，重置搜索状态"""
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush())   # 恢复默认背景
        self.table.clearSelection()
        self.search_results = []
        self.search_index = -1
        self.next_btn.setEnabled(False)
        self.search_edit.clear()
        self.last_keyword = ""  # 重置关键词
        self.stats_label.setText(self.stats_label.text().split(" | ")[0] if " | " in self.stats_label.text() else "搜索已清除")

    def on_search_enter(self):
        """回车事件：智能搜索或跳转下一个"""
        keyword = self.search_edit.text().strip()
        if not keyword:
            self.clear_highlight()
            return

        # 若关键词未变且已有搜索结果，则跳转下一个
        if keyword == self.last_keyword and self.search_results:
            self.next_match()
        else:
            # 否则执行新搜索
            self.search_student()

    def search_student(self):
        """执行搜索，高亮第一个匹配行"""
        keyword = self.search_edit.text().strip()
        if not keyword:
            self.clear_highlight()
            return

        # 清除之前的高亮
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush())
        self.table.clearSelection()

        matches = []
        for row in range(self.table.rowCount()):
            id_item = self.table.item(row, 1)      # 学号
            name_item = self.table.item(row, 2)    # 姓名
            if id_item and name_item:
                sid = id_item.text()
                name = name_item.text()
                if keyword.lower() in sid.lower() or keyword.lower() in name.lower():
                    matches.append(row)

        self.search_results = matches
        self.search_index = -1
        self.last_keyword = keyword   # 记录本次搜索词

        if matches:
            self.next_btn.setEnabled(True)
            self.highlight_next()
            self.stats_label.setText(f"找到 {len(matches)} 条匹配")
        else:
            self.next_btn.setEnabled(False)
            self.stats_label.setText("未找到匹配项")

    def highlight_next(self):
        """高亮下一个匹配行（循环）"""
        if not self.search_results:
            return
        # 清除之前的高亮
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush())
        self.table.clearSelection()

        # 计算下一个索引
        self.search_index = (self.search_index + 1) % len(self.search_results)
        row = self.search_results[self.search_index]

        # 高亮整行
        for col in range(self.table.columnCount()):
            item = self.table.item(row, col)
            if item:
                item.setBackground(QBrush(QColor(255, 255, 0, 80)))   # 半透明黄色

        # 选中行并滚动到可见区域
        self.table.selectRow(row)
        self.table.scrollToItem(self.table.item(row, 0), QTableWidget.ScrollHint.EnsureVisible)

    def next_match(self):
        """点击“下一个”按钮或按回车调用，跳转到下一个匹配"""
        self.highlight_next()

    # ---------- 原有功能 ----------
    def on_double_click(self, index):
        row = index.row()
        stu_id_item = self.table.item(row, 1)
        if stu_id_item:
            stu_id = stu_id_item.data(Qt.ItemDataRole.UserRole)
            if stu_id:
                stu_name = self.table.item(row, 2).text()
                RecordDetailDialog(self, stu_id, stu_name, self.app).exec()

    def on_header_double_clicked(self, column):
        """处理表头双击事件，根据列索引清除对应类别的全部记录（自定义样式确认框）"""
        category_map = {
            3: "日常",
            4: "听写",
            5: "背诵",
            6: "班干加分",
            7: "考试加分",
            8: "赋分"
        }
        if column not in category_map:
            return

        category = category_map[column]

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("⚠️ 确认清除")
        msg_box.setText(f"确定要清除所有「{category}」类别的积分记录吗？\n此操作将永久删除该类别的全部数据，不可撤销！")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)   # ✅ 跟随主题
        reply = msg_box.exec()
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
            if "类别" not in df.columns:
                show_toast(self, "数据文件中没有“类别”列，无法操作", "error")
                return

            df_filtered = df[df["类别"] != category]
            removed_count = len(df) - len(df_filtered)
            if removed_count == 0:
                show_toast(self, f"没有找到「{category}」类别的记录，无需清除", "info")
                return

            df_filtered.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
            show_toast(self, f"已清除 {removed_count} 条「{category}」类别的记录", "success")
            self.refresh()
        except Exception as e:
            show_toast(self, f"清除失败：{e}", "error")

# ======================== 记录明细对话框（表格样式全局） ========================
class RecordDetailDialog(QDialog):
    def __init__(self, parent, stu_id, stu_name, app):
        super().__init__(parent)
        self.app = app
        self.stu_id = stu_id
        self.full_df = None
        self.setWindowTitle(f"📋 {stu_name} ({stu_id}) 的积分明细")
        self.resize(850, 550)
        self.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        title = QLabel(f"📋 {stu_name} ({stu_id}) 的积分明细")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setAutoFillBackground(True)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["序号", "日期", "变动类型", "类别", "分数", "赋分总分", "备注"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        self.modify_btn = QPushButton("✏️ 修改选中")
        self.modify_btn.clicked.connect(self.modify_selected)
        self.delete_btn = QPushButton("🗑️ 删除选中")
        self.delete_btn.clicked.connect(self.delete_selected)
        self.close_btn = QPushButton("关闭")
        self.close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(self.modify_btn)
        btn_layout.addWidget(self.delete_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.close_btn)
        layout.addLayout(btn_layout)

        self.load_data()

    def load_data(self):
        df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        df["学号"] = df["学号"].astype(str)
        self.full_df = df
        records = df[df["学号"] == self.stu_id]
        self.records = records
        self.table.setRowCount(len(records))
        for i, (idx, row) in enumerate(records.iterrows()):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(str(row["日期"])))
            self.table.setItem(i, 2, QTableWidgetItem(str(row["变动类型"])))
            self.table.setItem(i, 3, QTableWidgetItem(str(row["类别"])))
            self.table.setItem(i, 4, QTableWidgetItem(str(row["分数"])))
            self.table.setItem(i, 5, QTableWidgetItem(str(row["赋分总分"]) if pd.notna(row["赋分总分"]) else ""))
            self.table.setItem(i, 6, QTableWidgetItem(str(row["备注"]) if pd.notna(row["备注"]) else ""))
            self.table.item(i, 0).setData(Qt.ItemDataRole.UserRole, idx)

    def get_selected_row(self):
        current = self.table.currentRow()
        if current < 0:
            show_toast(self, "请先选择一条记录", "warning")
            return None
        return current

    def modify_selected(self):
        row = self.get_selected_row()
        if row is None:
            return
        idx = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        record = self.full_df.loc[idx].to_dict()
        dialog = ModifyRecordDialog(self, record, idx, self.app)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_data()
            show_toast(self, " 记录已修改", "success")
            self.app.refresh_ranking()

    def delete_selected(self):
        row = self.get_selected_row()
        if row is None:
            return
        idx = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        reply = QMessageBox.question(self, "确认删除", "确定要删除这条记录吗？此操作不可撤销！",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.full_df = self.full_df.drop(idx)
            self.full_df.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
            self.load_data()
            show_toast(self, " 记录已删除", "success")
            self.app.refresh_ranking()

class ModifyRecordDialog(QDialog):
    def __init__(self, parent, record, idx, app):
        super().__init__(parent)
        self.app = app
        self.record = record
        self.idx = idx
        self.setWindowTitle("✏️ 修改记录")
        self.resize(450, 450)
        self.setStyleSheet(DIALOG_STYLE)

        layout = QFormLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(14)
        layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.date_edit = QLineEdit(record["日期"])
        layout.addRow("日期 (YYYY-MM-DD):", self.date_edit)

        self.op_plus = QRadioButton("加分")
        self.op_minus = QRadioButton("扣分")
        if record["变动类型"] == "+":
            self.op_plus.setChecked(True)
        else:
            self.op_minus.setChecked(True)
        op_widget = QWidget()
        op_layout = QHBoxLayout(op_widget)
        op_layout.setContentsMargins(0, 0, 0, 0)
        op_layout.setSpacing(15)
        op_layout.addWidget(self.op_plus)
        op_layout.addWidget(self.op_minus)
        layout.addRow("变动类型:", op_widget)

        self.category_combo = ComboBox()
        self.category_combo.addItems(["日常", "听写", "背诵", "班干加分", "考试加分"])
        self.category_combo.setCurrentText(record["类别"])
        layout.addRow("类别:", self.category_combo)

        self.score_edit = QLineEdit(str(abs(record["分数"])))
        layout.addRow("分数（正数）:", self.score_edit)

        self.assign_edit = QLineEdit(str(record["赋分总分"]) if pd.notna(record["赋分总分"]) else "")
        layout.addRow("赋分总分:", self.assign_edit)

        self.note_edit = QLineEdit(str(record["备注"]) if pd.notna(record["备注"]) else "")
        layout.addRow("备注:", self.note_edit)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.save_changes)
        btn_box.rejected.connect(self.reject)
        for btn in btn_box.buttons():
            btn.setStyleSheet("""
                QPushButton {
                    border-radius: 10px;
                    padding: 6px 16px;
                    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #64B5F6,stop:1 #42A5F5);
                    color: white;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #42A5F5,stop:1 #1E88E5);
                }
            """)
        layout.addRow(btn_box)

    def save_changes(self):
        date_str = self.date_edit.text().strip()
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except:
            show_toast(self, " 日期格式应为 YYYY-MM-DD", "error")
            return
        try:
            score_abs = float(self.score_edit.text().strip())
        except:
            show_toast(self, " 分数须为数字", "error")
            return
        op = "+" if self.op_plus.isChecked() else "-"
        score = score_abs if op == "+" else -score_abs
        category = self.category_combo.currentText()
        assign_str = self.assign_edit.text().strip()
        assign_score = None
        if assign_str:
            try:
                assign_score = float(assign_str)
            except:
                show_toast(self, " 赋分总分须为数字", "error")
                return
        note = self.note_edit.text().strip()

        new_record = {
            "学号": self.record["学号"],
            "姓名": self.record["姓名"],
            "日期": date_str,
            "变动类型": op,
            "类别": category,
            "分数": score,
            "赋分总分": assign_score if assign_score is not None else None,
            "备注": note
        }

        df = pd.read_csv(DATA_CSV, encoding="utf-8-sig")
        if self.idx not in df.index:
            show_toast(self, " 记录已被删除", "error")
            return
        # 删除旧行并重置索引
        df = df.drop(self.idx).reset_index(drop=True)
        # 追加新记录
        new_df = pd.DataFrame([new_record])
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_csv(DATA_CSV, index=False, encoding="utf-8-sig")
        self.accept()

# ======================== 导出页面（按钮样式全局） ========================
class ExportInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("exportInterface")
        self.build_ui()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        title = TitleLabel("📤 导出 / 导入", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        temp_layout = QHBoxLayout()
        temp_layout.addWidget(QLabel("模板文件："))
        self.template_label = QLabel(self.app.template_path if self.app.template_path else "未选择")
        self.template_label.setStyleSheet("color: #A0A0A0;")
        temp_layout.addWidget(self.template_label)
        choose_btn = QPushButton("选择模板")
        choose_btn.clicked.connect(self.choose_template)
        clear_btn = QPushButton("清空模板")
        clear_btn.clicked.connect(self.clear_template)
        temp_layout.addWidget(choose_btn)
        temp_layout.addWidget(clear_btn)
        temp_layout.addStretch()
        layout.addLayout(temp_layout)

        

        # ---- 导出积分汇总 ----
        score_group = QGroupBox("导出积分汇总")
        score_group_layout = QVBoxLayout(score_group)
        score_desc = QLabel("选择月份可导出对应月份的积分汇总表，留空则导出全部数据")
        score_desc.setStyleSheet("color: #A0A0A0;")
        score_group_layout.addWidget(score_desc)

        month_layout = QHBoxLayout()
        month_layout.addWidget(QLabel("月份（YYYY-MM，留空为全部）："))
        self.month_edit = QLineEdit()
        self.month_edit.setPlaceholderText("例如 2026-08")
        month_layout.addWidget(self.month_edit)
        score_group_layout.addLayout(month_layout)

        export_btn_row = QHBoxLayout()
        export_btn_row.addStretch()
        export_btn = QPushButton("📥 开始导出")
        export_btn.clicked.connect(self.do_export)
        export_btn_row.addWidget(export_btn)
        score_group_layout.addLayout(export_btn_row)
        layout.addWidget(score_group)

        # ---- 新增：导出成绩总览 ----
        overview_group = QGroupBox("导出成绩总览")
        overview_layout = QVBoxLayout(overview_group)
        overview_desc = QLabel("导出本次与上次考试的对比总览表（内容参考成绩界面「总览」）")
        overview_desc.setStyleSheet("color: #A0A0A0;")
        overview_layout.addWidget(overview_desc)

        overview_btn_row = QHBoxLayout()
        overview_btn_row.addWidget(QLabel("导出范围："))
        overview_info = QLabel("本次 vs 上次（全部学生）")
        overview_info.setStyleSheet("color: {_c['fg']};")
        overview_btn_row.addWidget(overview_info)
        overview_btn_row.addStretch()
        score_export_btn = QPushButton("📊 导出成绩总览")
        score_export_btn.clicked.connect(self.do_export_score_overview)
        overview_btn_row.addWidget(score_export_btn)
        overview_layout.addLayout(overview_btn_row)
        layout.addWidget(overview_group)
        # ---- 新增结束 ----

        # ---- 导入赋分总分 ----
        import_group = QGroupBox("导入赋分总分")
        import_layout = QVBoxLayout(import_group)
        desc = QLabel("选择包含学号和赋分总分的Excel文件，可指定折扣比例")
        desc.setStyleSheet("color: #A0A0A0;")
        import_layout.addWidget(desc)
        imp_row = QHBoxLayout()
        imp_row.addWidget(QLabel("折扣比例："))
        self.import_discount = QLineEdit("100")
        self.import_discount.setFixedWidth(80)
        imp_row.addWidget(self.import_discount)
        imp_row.addWidget(QLabel("%"))
        imp_btn = QPushButton("📂 选择Excel并导入")
        imp_btn.clicked.connect(self.import_assign)
        imp_row.addWidget(imp_btn)
        import_layout.addLayout(imp_row)
        layout.addWidget(import_group)

        back_btn = QPushButton("🔙 返回主菜单")
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()

    def choose_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择模板 Excel 文件", "", "Excel 文件 (*.xlsx);;所有文件 (*)")
        if file_path:
            try:
                wb = load_workbook(file_path)
                ws = wb.active
                header = [ws.cell(row=1, column=c).value for c in range(1, min(ws.max_column+1, 20))]
                if '学号' not in header or '姓名' not in header:
                    QMessageBox.warning(self, "警告", "模板可能缺少'学号'或'姓名'列，请确认")
                save_template_config(file_path)
                self.app.template_path = file_path
                self.template_label.setText(file_path)
                show_toast(self, " 模板已选择", "success")
            except Exception as e:
                show_toast(self, f" 无法打开模板文件：{str(e)}", "error")

    def clear_template(self):
        save_template_config("")
        self.app.template_path = ""
        self.template_label.setText("未选择")
        show_toast(self, "🗑️ 模板已清空", "info")

    def do_export(self):
        month = self.month_edit.text().strip()
        if month:
            try:
                datetime.strptime(month, "%Y-%m")
            except:
                show_toast(self, " 月份格式应为 YYYY-MM", "error")
                return
        if self.app.stu_df is None:
            show_toast(self, " 请先加载学生名单", "error")
            return

        # 执行导出
        if self.app.template_path and os.path.exists(self.app.template_path):
            ok, msg = export_excel_with_template(self.app.stu_df, self.app.template_path, month if month else None)
        else:
            ok, msg = export_excel_fallback(self.app.stu_df, month if month else None)

        if ok:
            # 从 msg 中提取文件路径
            filepath = None
            if "：" in msg:
                # 取冒号后面的部分
                filepath = msg.split("：", 1)[-1].strip()
                # 如果路径包含换行或额外文字，只取第一个有效路径
                if not os.path.exists(filepath):
                    # 尝试按空格拆分取第一个可能存在的路径
                    parts = filepath.split()
                    for p in parts:
                        if os.path.exists(p):
                            filepath = p
                            break

        if ok:
            # 增加导出计数
            cfg = load_config()
            cfg["export_count"] = cfg.get("export_count", 0) + 1
            save_config(cfg)
            check_and_unlock_achievements()
            # 原有提取路径和显示对话框代码...
            if filepath and os.path.exists(filepath):
                # 弹出自定义成功对话框
                self.show_export_success_dialog(filepath)
            else:
                # 如果无法提取路径，降级使用 Toast
                show_toast(self, msg, "success")
        else:
            show_toast(self, msg, "error")

    def do_export_score_overview(self):
        """导出成绩总览对比表（与成绩界面总览视图一致）"""
        ok, msg = export_score_overview()
        if ok:
            # 增加导出计数
            cfg = load_config()
            cfg["export_count"] = cfg.get("export_count", 0) + 1
            save_config(cfg)
            check_and_unlock_achievements()

            # 从 msg 中提取文件路径
            filepath = None
            if "：" in msg:
                filepath = msg.split("：", 1)[-1].strip()
                if not os.path.exists(filepath):
                    parts = filepath.split()
                    for p in parts:
                        if os.path.exists(p):
                            filepath = p
                            break

            if filepath and os.path.exists(filepath):
                self.show_export_success_dialog(filepath)
            else:
                show_toast(self, msg, "success")
        else:
            show_toast(self, msg, "error")

    def show_export_success_dialog(self, filepath):
        """导出成功后的自定义对话框，包含三个按钮"""
        dialog = QDialog(self)
        dialog.setWindowTitle("导出成功")
        dialog.setFixedSize(500, 180)
        dialog.setStyleSheet(DIALOG_STYLE)  # 全局暗色样式

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)

        # 提示标签
        label = QLabel("✅ 导出成功！")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-size: 18px; font-weight: bold; color: #64B5F6;")
        layout.addWidget(label)

        # 第一行：打开文件所在位置（独占一行，宽度拉伸填满）
        row1 = QHBoxLayout()
        open_dir_btn = QPushButton("📂 打开文件所在位置")
        # 设置水平策略为 Expanding，使其宽度填满整行
        open_dir_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        open_dir_btn.clicked.connect(lambda: self._open_file_location(filepath))
        row1.addWidget(open_dir_btn)
        layout.addLayout(row1)

        # 第二行：打开 + 确认（两个按钮等宽，中间有间距）
        row2 = QHBoxLayout()
        row2.setSpacing(15)  # 与下面两个按钮之间的间距

        open_btn = QPushButton("📄 打开文件")
        open_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        open_btn.clicked.connect(lambda: self._open_file(filepath))

        confirm_btn = QPushButton("确认")
        confirm_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        confirm_btn.clicked.connect(dialog.accept)

        row2.addWidget(open_btn)
        row2.addWidget(confirm_btn)
        layout.addLayout(row2)

        dialog.exec()

    def _open_file_location(self, filepath):
        """打开文件所在的文件夹，并高亮选中该文件（跨平台，快速且自动聚焦）"""
        if not os.path.exists(filepath):
            show_toast(self, "文件不存在或已被移动", "warning")
            return

        try:
            if sys.platform.startswith("win"):
                filepath = os.path.normpath(filepath)
                if not self._win_open_and_select(filepath):
                    # 极端情况下的退化方案
                    subprocess.Popen(f'explorer /select,"{filepath}"')
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", filepath])
            else:
                # Linux 没有统一的高亮接口，退化为打开所在目录
                dir_path = os.path.dirname(filepath)
                QDesktopServices.openUrl(QUrl.fromLocalFile(dir_path))
        except Exception as e:
            dir_path = os.path.dirname(filepath)
            if os.path.exists(dir_path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(dir_path))
            else:
                show_toast(self, f"无法打开文件夹：{e}", "error")

    @staticmethod
    def _win_open_and_select(filepath):
        """
        Windows: 调用 Shell API SHOpenFolderAndSelectItems
        - 复用已运行的资源管理器 → 快
        - 系统会主动将窗口带到前台 → 自动聚焦
        返回 True 表示成功，False 表示需要退化处理。
        """
        try:
            import ctypes
            from ctypes import wintypes

            shell32 = ctypes.windll.shell32
            ole32 = ctypes.windll.ole32

            # 声明参数类型，避免 64 位下指针截断
            shell32.SHParseDisplayName.argtypes = [
                wintypes.LPCWSTR,                    # pszName
                ctypes.c_void_p,                     # pbc
                ctypes.POINTER(ctypes.c_void_p),     # ppidl
                wintypes.DWORD,                      # sfGAOIn
                ctypes.POINTER(wintypes.DWORD),      # psfGAOOut
            ]
            shell32.SHParseDisplayName.restype = ctypes.c_long

            shell32.SHOpenFolderAndSelectItems.argtypes = [
                ctypes.c_void_p,   # pidlFolder
                ctypes.c_uint,     # cidl
                ctypes.c_void_p,   # apidl
                wintypes.DWORD,    # dwFlags
            ]
            shell32.SHOpenFolderAndSelectItems.restype = ctypes.c_long

            # Qt 主线程通常已经初始化过 COM，这里再 CoInitialize 是引用计数安全的
            ole32.CoInitialize(None)
            try:
                pidl = ctypes.c_void_p()
                hr = shell32.SHParseDisplayName(
                    filepath, None, ctypes.byref(pidl), 0, None
                )
                if hr != 0 or not pidl:
                    return False
                try:
                    shell32.SHOpenFolderAndSelectItems(pidl, 0, None, 0)
                finally:
                    ole32.CoTaskMemFree(pidl)
                return True
            finally:
                ole32.CoUninitialize()
        except Exception:
            return False

    def _open_file(self, filepath):
        """用系统默认程序打开文件"""
        if os.path.exists(filepath):
            QDesktopServices.openUrl(QUrl.fromLocalFile(filepath))

    def import_assign(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择包含赋分总分的Excel文件", "", "Excel 文件 (*.xlsx);;所有文件 (*)")
        if not file_path:
            return
        try:
            discount = float(self.import_discount.text().strip())
        except:
            show_toast(self, " 请输入有效折扣比例", "error")
            return
        if self.app.stu_df is None:
            show_toast(self, " 请先加载学生名单", "error")
            return
        try:
            df_import = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            show_toast(self, f" 读取文件失败：{e}", "error")
            return
        if "学号" not in df_import.columns:
            show_toast(self, " Excel中没有'学号'列", "error")
            return
        score_col = None
        for col in df_import.columns:
            if str(col).strip() == "赋分总分":
                score_col = col
                break
        if score_col is None:
            show_toast(self, " 没有找到'赋分总分'列", "error")
            return

        preview_data = []
        valid_ids = set(self.app.stu_df["学号"].values)
        for _, row in df_import.iterrows():
            stu_id = str(row["学号"])
            name = self.app.id_to_name.get(stu_id, "（不在名单）")
            raw = pd.to_numeric(row[score_col], errors='coerce')
            if pd.isna(raw):
                raw = 0
            discounted = round(raw * (discount/100.0), 2)
            preview_data.append((stu_id, name, raw, discounted, stu_id in valid_ids))

        # 在 import_assign 方法中，找到 dialog 创建部分，替换为：
        dialog = QDialog(self)
        dialog.setStyleSheet(DIALOG_STYLE)
        dialog.setWindowTitle("导入赋分预览")
        dialog.resize(700, 450)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        table = QTableWidget()
        install_scrolling_header(table)
        table.setAutoFillBackground(True)  # 强制填充背景
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["学号", "姓名", "原始分数", "折扣后分数", "状态"])
        table.setRowCount(len(preview_data))
        for i, (sid, name, raw, disc, ok) in enumerate(preview_data):
            table.setItem(i, 0, QTableWidgetItem(sid))
            table.setItem(i, 1, QTableWidgetItem(name))
            table.setItem(i, 2, QTableWidgetItem(str(raw)))
            table.setItem(i, 3, QTableWidgetItem(str(disc)))
            table.setItem(i, 4, QTableWidgetItem("✅ 匹配" if ok else "❌ 不匹配"))
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(table)

        info = QLabel(f"共 {len(preview_data)} 条记录，匹配 {sum(1 for _,_,_,_,ok in preview_data if ok)} 人")
        info.setStyleSheet("color: #A0A0A0;")
        layout.addWidget(info)

        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(lambda: self.confirm_import(file_path, discount, dialog))
        btn_box.rejected.connect(dialog.reject)
        # 按钮样式
        for btn in btn_box.buttons():
            btn.setStyleSheet("""
                QPushButton {
                    border-radius: 10px;
                    padding: 6px 16px;
                    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #64B5F6,stop:1 #42A5F5);
                    color: white;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #42A5F5,stop:1 #1E88E5);
                }
            """)
        layout.addWidget(btn_box)
        dialog.exec()

    def confirm_import(self, file_path, discount, dialog):
        dialog.accept()
        success, fail = import_assign_scores(self.app.stu_df, file_path, discount)
        if success > 0:
            show_toast(self, f" 成功导入 {success} 条赋分总分记录，失败 {len(fail)} 个", "success")
            self.app.refresh_ranking()
        else:
            show_toast(self, f" 导入失败：{', '.join(fail[:5])}", "error")


# ======================== 存储空间分析 ========================
# ---------------- 分组保留条数（模块级，供主/子对话框共用） ----------------
_KEEP_GROUP_PATTERNS = {
    "exports": [
        ("班级积分月度汇总_", "班级积分月度汇总"),
        ("成绩总览对比_", "成绩总览对比"),
    ],
    "backups": [
        ("备份_积分记录_", "备份-积分记录"),
        ("名单备份_", "名单备份"),
        ("积分记录_", "积分记录"),
    ],
}

_KEEP_GROUP_FIELD = {
    "exports": "export_keep_counts",
    "backups": "backup_keep_counts",
}


def _match_keep_type(key, filename):
    for prefix, display in _KEEP_GROUP_PATTERNS.get(key, []):
        if filename.startswith(prefix):
            return display
    return "其它"


def _all_keep_type_names(key):
    names = [display for _, display in _KEEP_GROUP_PATTERNS.get(key, [])]
    names.append("其它")
    return names


def apply_keep_counts(key, dir_path):
    """按类型分组执行保留策略，返回删除的文件数。"""
    if key not in _KEEP_GROUP_FIELD:
        return 0
    if not dir_path or not os.path.exists(dir_path):
        return 0

    cfg = load_config()
    counts = cfg.get(_KEEP_GROUP_FIELD[key], {})

    files = []
    try:
        for f in os.listdir(dir_path):
            p = os.path.join(dir_path, f)
            if os.path.isfile(p):
                try:
                    files.append({
                        "name": f,
                        "path": p,
                        "mtime": os.path.getmtime(p),
                    })
                except OSError:
                    pass
    except OSError:
        return 0

    groups = {}
    for it in files:
        t = _match_keep_type(key, it["name"])
        groups.setdefault(t, []).append(it)

    deleted = 0
    for t, items in groups.items():
        keep = int(counts.get(t, 5))
        items.sort(key=lambda x: -x["mtime"])
        for it in items[keep:]:
            try:
                os.remove(it["path"])
                deleted += 1
            except OSError:
                pass
    return deleted
def _format_size_split(size_bytes):
    """把字节数拆成 (数值字符串, 单位)"""
    try:
        size_bytes = float(size_bytes)
    except (TypeError, ValueError):
        return "0", "B"
    if size_bytes < 1024:
        return f"{int(size_bytes)}", "B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f}", "KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f}", "MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f}", "GB"


def _format_size(size_bytes):
    """把字节数格式化为易读字符串"""
    num, unit = _format_size_split(size_bytes)
    return f"{num} {unit}"


# ---------------- 环形图（带进度动画） ----------------
class DonutChart(QWidget):
    """
    环形图：展示磁盘空间的三段构成
      · 蓝色 = user 已用空间
      · 紫色 = 磁盘已用空间（不含 user 部分）
      · 白色 = 磁盘可用空间
    带进度动画：打开时从 0° 缓动到目标角度（受 enable_progress_animation 控制）
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.segments = []       # [{"name":..., "size": int, "color": "#RRGGBB"}, ...]
        self.total_size = 0
        self.center_num = "0"
        self.center_unit = "B"
        self.center_sub = ""
        self.setMinimumHeight(250)

        self._progress = 0.0     # 0.0 ~ 1.0
        self._anim = None

    # ---------- Qt property，供 QPropertyAnimation 使用 ----------
    def _get_progress(self):
        return self._progress

    def _set_progress(self, v):
        self._progress = float(v)
        self.update()

    progress = pyqtProperty(float, fget=_get_progress, fset=_set_progress)

    # ---------- 对外接口 ----------
    def set_data(self, segments, center_num, center_unit, center_sub, animate=True):
        """设置数据并启动动画"""
        self.segments = segments
        self.total_size = sum(s["size"] for s in segments)
        self.center_num = center_num
        self.center_unit = center_unit
        self.center_sub = center_sub
        self._start_animation(animate)

    def _start_animation(self, animate=True):
        # 停掉旧动画
        if self._anim is not None:
            try:
                self._anim.stop()
            except Exception:
                pass
            self._anim = None

        # 读取全局开关
        try:
            enabled = load_config().get("enable_progress_animation", True)
        except Exception:
            enabled = True

        if not animate or not enabled:
            self._progress = 1.0
            self.update()
            return

        anim = QPropertyAnimation(self, b"progress")
        anim.setDuration(1400)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()
        self._anim = anim

    # ---------- 绘制 ----------
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        _c = get_theme_colors()

        w, h = self.width(), self.height()
        size = min(w, h) - 16
        if size <= 10:
            return
        cx, cy = w / 2.0, h / 2.0
        rect = QRectF(cx - size / 2.0, cy - size / 2.0, size, size)

        # 底层灰环（占位）
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(QColor(_c["border"])))
        painter.drawEllipse(rect)

        # 三段扇形
        if self.total_size > 0 and self._progress > 0:
            start_angle = 90 * 16
            p = self._progress
            for seg in self.segments:
                if seg["size"] <= 0:
                    continue
                full_span = -int(seg["size"] / self.total_size * 360 * 16)
                span = int(full_span * p)
                if span != 0:
                    painter.setPen(Qt.PenStyle.NoPen)
                    painter.setBrush(QBrush(QColor(seg["color"])))
                    painter.drawPie(rect, start_angle, span)
                start_angle += full_span

        # 中心挖空
        inner_size = size * 0.60
        inner_rect = QRectF(cx - inner_size / 2.0, cy - inner_size / 2.0, inner_size, inner_size)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(QColor(_c["bg_alt"]))) # 改为 bg_alt
        painter.drawEllipse(inner_rect)

        # ---- 中心文字 ----
        num_font = QFont("Microsoft YaHei", 24, QFont.Weight.Bold)
        unit_font = QFont("Microsoft YaHei", 11)

        painter.setFont(num_font)
        fm_num = painter.fontMetrics()
        num_w = fm_num.horizontalAdvance(self.center_num)

        painter.setFont(unit_font)
        fm_unit = painter.fontMetrics()
        unit_w = fm_unit.horizontalAdvance(" " + self.center_unit)

        total_w = num_w + unit_w
        start_x = cx - total_w / 2.0
        target_center_y = cy - inner_size * 0.06
        baseline_y = target_center_y + (fm_num.ascent() - fm_num.descent()) / 2.0

        painter.setFont(num_font)
        painter.setPen(QPen(QColor(_c["fg_bright"])))
        painter.drawText(QPointF(start_x, baseline_y), self.center_num)

        painter.setFont(unit_font)
        painter.setPen(QPen(QColor(_c["fg_dim"])))
        painter.drawText(QPointF(start_x + num_w + 3, baseline_y - 2), self.center_unit)

        painter.setFont(QFont("Microsoft YaHei", 10))
        painter.setPen(QPen(QColor(_c["fg_dim"])))
        desc_rect = QRectF(cx - inner_size / 2.0, cy + inner_size * 0.06, inner_size, inner_size * 0.22)
        painter.drawText(desc_rect, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop, self.center_sub)


class LegendWidget(QWidget):
    """图例：彩点 + 名称，水平居中"""
    def __init__(self, items=None, parent=None):
        super().__init__(parent)
        self.items = items or []   # [{"name":..., "color":..., "size":...}, ...]
        self._layout = QHBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(22)
        self._layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._populate()

    def _populate(self):
        _c = get_theme_colors()
        while self._layout.count():
            item = self._layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()

        for c in self.items:
            wrap = QWidget()
            h = QHBoxLayout(wrap)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(7)

            dot = QFrame()
            dot.setFixedSize(11, 11)
            # 白色圆点加描边，与背景区分
            if c["color"].upper() in ("#F5F5F5", "#FFFFFF"):
                dot.setStyleSheet(
                    f"background: {_c['fg_bright']}; border-radius: 5px; "
                    f"border: 1px solid {_c['border_light']};"
                )
            else:
                dot.setStyleSheet(
                    f"background: {c['color']}; border-radius: 5px; border: none;"
                )
            h.addWidget(dot)

            lbl = QLabel(c["name"])
            lbl.setStyleSheet(
                f"color: {_c['fg']}; font-size: 13px; background: transparent;"
            )
            h.addWidget(lbl)

            self._layout.addWidget(wrap)

    def update_items(self, items):
        self.items = items
        self._populate()


# ---------------- 大类卡片 ----------------
class CategoryCard(QFrame):
    """大类卡片：标题 + 右侧按钮 + 大数字 + 描述；点击整卡进入子页"""
    def __init__(self, category, on_open=None, parent=None):
        super().__init__(parent)
        _c = get_theme_colors()
        self.category = category
        self.on_open = on_open
        self.setObjectName("categoryCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet(f"""
            QFrame#categoryCard {{
                background: {_c['bg_alt']};
                border-radius: 14px;
                border: 1px solid {_c['border']};
            }}
            QFrame#categoryCard:hover {{
                background: {_c['bg_card']};
                border: 1px solid #64B5F6;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        # 第一行：标题 + 按钮
        top_row = QHBoxLayout()
        title = QLabel(category.get("name", ""))
        title.setStyleSheet(
            f"color: {_c['fg_bright']}; font-size: 16px; font-weight: 600;"
            f" background: transparent;"
        )
        top_row.addWidget(title)
        top_row.addStretch()

        button_text = category.get("button_text", "查看")
        self.btn = QPushButton(button_text)
        self.btn.setFixedSize(80, 34)
        self.btn.setStyleSheet("""
            QPushButton {
                background: rgba(100, 181, 246, 0.15);
                color: #64B5F6;
                border: 1px solid rgba(100, 181, 246, 0.35);
                border-radius: 17px;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: rgba(100, 181, 246, 0.28);
            }
            QPushButton:pressed {
                background: rgba(100, 181, 246, 0.4);
            }
        """)
        self.btn.clicked.connect(self._emit_open)
        top_row.addWidget(self.btn)
        layout.addLayout(top_row)

        # 第二行：大数字 + 单位
        num_str, unit_str = _format_size_split(category.get("size", 0))
        num_row = QHBoxLayout()
        num_row.setContentsMargins(0, 0, 0, 0)
        num_row.setSpacing(6)

        num_label = QLabel(num_str)
        num_label.setStyleSheet(
            f"color: {_c['fg_bright']}; font-size: 30px; font-weight: 700;"
            f" background: transparent;"
        )
        num_row.addWidget(num_label, 0, Qt.AlignmentFlag.AlignBottom)

        unit_label = QLabel(unit_str)
        unit_label.setStyleSheet(
            f"color: {_c['fg_dim']}; font-size: 14px; font-weight: 500;"
            f" background: transparent; padding-bottom: 6px;"
        )
        num_row.addWidget(unit_label, 0, Qt.AlignmentFlag.AlignBottom)

        num_row.addStretch()
        layout.addLayout(num_row)

        # 第三行：描述
        desc = QLabel(category.get("desc", ""))
        desc.setStyleSheet(
            f"color: {_c['fg_dim']}; font-size: 12px; background: transparent;"
        )
        desc.setWordWrap(True)
        layout.addWidget(desc)

    def _emit_open(self):
        if self.on_open:
            self.on_open(self.category)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._emit_open()
            return
        try:
            super().mousePressEvent(event)
        except RuntimeError:
            pass


# ---------------- 子页（大类下的文件列表） ----------------
class FileRow(QFrame):
    """文件行：支持双击触发回调"""
    def __init__(self, item, parent_dialog):
        super().__init__()
        _c = get_theme_colors()
        self.item = item
        self.parent_dialog = parent_dialog
        self.setStyleSheet(f"""
            QFrame {{
                background: {_c['bg_card']};
                border-radius: 10px;
                border: 1px solid {_c['border']};
            }}
            QFrame:hover {{
                background: rgba(100, 181, 246, 0.10);
                border: 1px solid rgba(100, 181, 246, 0.25);
            }}
        """)
        self.setFixedHeight(48)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        h = QHBoxLayout(self)
        h.setContentsMargins(14, 8, 14, 8)

        name = QLabel(item.get("name", ""))
        name.setStyleSheet(
            f"color: {_c['fg']}; font-size: 14px; background: transparent;"
        )
        h.addWidget(name, stretch=1)

        size_label = QLabel(_format_size(item.get("size", 0)))
        size_label.setStyleSheet(
            f"color: {_c['fg_dim']}; font-size: 13px; background: transparent;"
        )
        size_label.setFixedWidth(100)
        size_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        h.addWidget(size_label)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if getattr(self.parent_dialog, "manage_count", False):
                self.parent_dialog.open_file_menu(self.item)
        super().mouseDoubleClickEvent(event)


class GroupCountDialog(QDialog):
    """分组保留条数配置对话框：每个类型一行，各配一个 SpinBox"""
    def __init__(self, parent, title, group_names, current_counts):
        super().__init__(parent)
        self.setWindowTitle(f"⚙️ {title} · 分组保留条数")
        # 自适应高度
        h = 130 + 46 * max(1, len(group_names))
        self.setFixedSize(440, min(600, h))
        self.setStyleSheet(DIALOG_STYLE)

        self.result_counts = dict(current_counts)
        self.spins = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(14)

        info = QLabel("为每种类型分别设置保留条数：\n"
                      "同类文件超出条数时，最旧的文件会被自动删除。")
        info.setStyleSheet("color: #A0A0B0; font-size: 13px;")
        info.setWordWrap(True)
        layout.addWidget(info)

        # 每种类型一行
        for name in group_names:
            row = QHBoxLayout()
            row.setSpacing(10)

            lbl = QLabel(name)
            lbl.setStyleSheet("color: #E0E0E0; font-size: 14px;")
            lbl.setFixedWidth(190)
            row.addWidget(lbl)

            spin = QSpinBox()
            spin.setRange(1, 999)
            spin.setValue(int(current_counts.get(name, 5)))
            spin.setFixedWidth(90)
            row.addWidget(spin)

            row.addWidget(QLabel("条"))
            row.addStretch()
            layout.addLayout(row)
            self.spins[name] = spin

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("确定")
        ok_btn.setFixedWidth(110)
        ok_btn.clicked.connect(self._on_ok)
        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedWidth(110)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _on_ok(self):
        for name, spin in self.spins.items():
            self.result_counts[name] = spin.value()
        self.accept()


class StorageSubDialog(QDialog):
    """某个大类下的子文件列表：双击操作 + 按类型分组保留条数"""
    def __init__(self, category, parent=None):
        super().__init__(parent)
        self.category = category
        self.parent_dialog = parent
        self.key = category.get("key", "")
        # 只有"导出文件"和"备份文件"支持条数管理
        self.manage_count = self.key in ("exports", "backups")
        # 对应的 config 字段名（dict）
        self.counts_field = {
            "exports": "export_keep_counts",
            "backups": "backup_keep_counts",
        }.get(self.key, "")

        self.setWindowTitle(f"{category.get('icon', '📁')}  {category.get('name', '')}")
        self.resize(620, 660)
        self.setStyleSheet(DIALOG_STYLE)

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        # 打开时应用一次保留规则
        if self.manage_count and self.counts_field:
            self._apply_keep_count_grouped()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 18)
        layout.setSpacing(12)

        # ---- 顶部：标题 + 分组保留条数按钮 ----
        top_row = QHBoxLayout()
        title = QLabel(f"{category.get('icon', '📁')}  {category.get('name', '')}")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        top_row.addWidget(title)
        top_row.addStretch()

        if self.manage_count:
            count_btn = QPushButton("⚙️ 分组保留条数")
            count_btn.setFixedHeight(36)
            count_btn.clicked.connect(self._open_group_count_dialog)
            top_row.addWidget(count_btn)

        layout.addLayout(top_row)

        # 汇总
        self.summary_label = QLabel("")
        self.summary_label.setStyleSheet("color: #A0A0B0; font-size: 13px;")
        layout.addWidget(self.summary_label)

        # 描述
        if category.get("desc"):
            desc = QLabel(category["desc"])
            desc.setStyleSheet("color: #808090; font-size: 12px;")
            desc.setWordWrap(True)
            layout.addWidget(desc)

        # 提示
        if self.manage_count:
            tip = QLabel("（双击文件可进行操作；各类型超出保留条数后，最旧的文件会自动删除）")
        else:
            tip = QLabel("（只读列表，仅用于查看占用大小）")
        tip.setStyleSheet("color: #808090; font-size: 12px;")
        layout.addWidget(tip)

        # 滚动列表
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical {
                width: 6px; background: transparent; margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #4A4A5A; border-radius: 3px; min-height: 30px;
            }
            QScrollBar::handle:vertical:hover { background: #6A6A7A; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
        """)
        self.container = QWidget()
        self.container.setStyleSheet("background: transparent;")
        self.list_layout = QVBoxLayout(self.container)
        self.list_layout.setContentsMargins(0, 0, 0, 0)
        self.list_layout.setSpacing(6)
        self.scroll.setWidget(self.container)
        layout.addWidget(self.scroll, stretch=1)

        # 底部按钮
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        open_btn = QPushButton("📂 打开所在目录")
        open_btn.clicked.connect(self._open_dir)
        btn_row.addWidget(open_btn)

        close_btn = QPushButton("关闭")
        close_btn.setFixedWidth(110)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._refresh_list()

    def _all_type_names(self):
        return _all_keep_type_names(self.key)

    def _match_file_type(self, filename):
        return _match_keep_type(self.key, filename)

    def _apply_keep_count_grouped(self):
        return apply_keep_counts(self.key, self.category.get("dir_path", ""))

    def _open_group_count_dialog(self):
        """打开分组保留条数配置对话框"""
        cfg = load_config()
        counts = dict(cfg.get(self.counts_field, {}))
        # 补齐所有类型
        for name in self._all_type_names():
            counts.setdefault(name, 5)

        dlg = GroupCountDialog(self, self.category.get("name", ""),
                               self._all_type_names(), counts)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        # 保存
        cfg = load_config()
        cfg[self.counts_field] = dlg.result_counts
        save_config(cfg)

        # 立即应用清理
        deleted = self._apply_keep_count_grouped()

        # 刷新父对话框（圆环 / 卡片 / 底部）
        if self.parent_dialog is not None:
            self.parent_dialog._refresh()
        self._refresh_list()

        if deleted > 0:
            show_toast(self, f"分组保留条数已保存，删除 {deleted} 个较旧文件", "success")
        else:
            show_toast(self, "分组保留条数已保存", "success")

    # ==================== 列表填充 ====================
    def _refresh_list(self):
        # 清空旧行
        while self.list_layout.count():
            item = self.list_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

        # 从父对话框重新拿最新数据
        items = []
        if self.parent_dialog is not None:
            for c in self.parent_dialog.categories:
                if c.get("key") == self.key:
                    self.category = c
                    items = list(c.get("items", []))
                    break
        if not items:
            items = list(self.category.get("items", []))

        # 按 mtime 从新到旧
        def _mtime(x):
            p = x.get("path", "")
            try:
                return os.path.getmtime(p) if os.path.exists(p) else 0
            except OSError:
                return 0
        items.sort(key=lambda x: -_mtime(x))

        if not items:
            empty = QLabel("暂无文件")
            empty.setStyleSheet("color: #808090; font-size: 13px;")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.list_layout.addWidget(empty)
        else:
            for it in items:
                self.list_layout.addWidget(FileRow(it, self))
        self.list_layout.addStretch()

        # 更新汇总（含按类型统计）
        total = sum(x.get("size", 0) for x in items)
        num_str, unit_str = _format_size_split(total)
        summary = f"共 {len(items)} 个文件 · 总占用 {num_str} {unit_str}"

        if self.manage_count and items:
            # 按类型计数
            cfg = load_config()
            counts = cfg.get(self.counts_field, {})
            per_type = {}
            for it in items:
                t = self._match_file_type(it["name"])
                per_type[t] = per_type.get(t, 0) + 1
            parts = []
            for name in self._all_type_names():
                if name in per_type:
                    limit = counts.get(name, 5)
                    parts.append(f"{name} {per_type[name]}/{limit}")
            if parts:
                summary += "    |    " + " · ".join(parts)

        self.summary_label.setText(summary)

    # ==================== 双击操作 ====================
    def open_file_menu(self, item):
        path = item.get("path", "")
        if not path or not os.path.exists(path):
            show_toast(self, "文件不存在或已被删除", "warning")
            self._refresh_list()
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("操作文件")
        dlg.setFixedSize(480, 230)
        dlg.setStyleSheet(DIALOG_STYLE)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(25, 22, 25, 22)
        lay.setSpacing(14)

        name = os.path.basename(path)
        info = QLabel(f"文件：{name}\n大小：{_format_size(item.get('size', 0))}")
        info.setStyleSheet("color: #E0E0E0; font-size: 14px;")
        info.setWordWrap(True)
        lay.addWidget(info)

        # 第一行：打开文件所在位置
        row1 = QHBoxLayout()
        open_dir_btn = QPushButton("📂 打开文件所在位置")
        open_dir_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        open_dir_btn.clicked.connect(lambda: (dlg.accept(), self._open_file_location(path)))
        row1.addWidget(open_dir_btn)
        lay.addLayout(row1)

        # 第二行：打开 / 删除 / 取消
        row2 = QHBoxLayout()
        row2.setSpacing(10)

        open_btn = QPushButton("📄 打开")
        open_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        open_btn.clicked.connect(lambda: (dlg.accept(), self._open_file(path)))

        del_btn = QPushButton("🗑️ 删除")
        del_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        del_btn.setStyleSheet("""
            QPushButton {
                border-radius: 10px;
                padding: 8px 20px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #EF5350, stop:1 #D32F2F);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(239, 83, 80, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #D32F2F, stop:1 #B71C1C);
            }
            QPushButton:pressed { background: #B71C1C; }
        """)
        del_btn.clicked.connect(lambda: (dlg.accept(), self._delete_file(path)))

        cancel_btn = QPushButton("取消")
        cancel_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        cancel_btn.clicked.connect(dlg.reject)

        row2.addWidget(open_btn)
        row2.addWidget(del_btn)
        row2.addWidget(cancel_btn)
        lay.addLayout(row2)

        dlg.exec()

    # ==================== 单文件操作 ====================
    def _open_file_location(self, path):
        try:
            norm = os.path.normpath(path)
            if sys.platform.startswith("win"):
                if not ExportInterface._win_open_and_select(norm):
                    subprocess.Popen(f'explorer /select,"{norm}"')
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", norm])
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(norm)))
        except Exception as e:
            show_toast(self, f"无法打开位置：{e}", "error")

    def _open_file(self, path):
        if os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _delete_file(self, path):
        msg = QMessageBox(self)
        msg.setWindowTitle("确认删除")
        msg.setText(f"确定要删除文件\n“{os.path.basename(path)}” 吗？\n此操作不可恢复！")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.No)
        msg.setStyleSheet(DIALOG_STYLE)
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return
        try:
            os.remove(path)
            show_toast(self, "文件已删除", "success")
        except Exception as e:
            show_toast(self, f"删除失败：{e}", "error")
            return
        if self.parent_dialog is not None:
            self.parent_dialog._refresh()
        self._refresh_list()

    def _open_dir(self):
        dir_path = self.category.get("dir_path", "")
        if dir_path and os.path.exists(dir_path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(dir_path))


# ---------------- 主对话框 ----------------
class StorageAnalyzerDialog(QDialog):
    """存储空间主对话框：大圆环 + 图例 + 大类卡片列表（全部随滚动条滑动）"""
    def __init__(self, parent=None):
        super().__init__(parent)
        _c = get_theme_colors()
        self.setWindowTitle("💾 存储空间")
        self.resize(720, 800)
        self.setStyleSheet(DIALOG_STYLE)

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        apply_keep_counts("exports", os.path.abspath(EXPORT_DIR))
        apply_keep_counts("backups", os.path.abspath(BACKUP_DIR))

        self.categories = self._collect_storage_data()
        self.total_size = sum(c["size"] for c in self.categories)

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(24, 20, 24, 18)
        outer_layout.setSpacing(14)

        title_row = QHBoxLayout()
        title = QLabel("💾 存储空间")
        title.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {_c['fg']};")
        title_row.addWidget(title)
        title_row.addStretch()

        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.setFixedSize(110, 36)
        refresh_btn.clicked.connect(self._refresh)
        title_row.addWidget(refresh_btn)

        open_dir_btn = QPushButton("📂 打开目录")
        open_dir_btn.setFixedSize(140, 36)
        open_dir_btn.clicked.connect(self._open_user_dir)
        title_row.addWidget(open_dir_btn)

        outer_layout.addLayout(title_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("""
            QScrollArea { border: none; background: transparent; }
            QScrollBar:vertical {
                width: 6px; background: transparent; margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #4A4A5A; border-radius: 3px; min-height: 30px;
            }
            QScrollBar::handle:vertical:hover { background: #6A6A7A; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
        """)
        self.scroll = scroll

        self.content_widget = QWidget()
        self.content_widget.setStyleSheet("background: transparent;")
        content_layout = QVBoxLayout(self.content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(14)

        chart_card = QFrame()
        chart_card.setStyleSheet(f"""
            QFrame {{
                background: rgba(100, 181, 246, 0.05);
                border-radius: 16px;
                border: 1px solid {_c['border']};
            }}
        """)
        chart_layout = QVBoxLayout(chart_card)
        chart_layout.setContentsMargins(20, 16, 20, 16)
        chart_layout.setSpacing(10)

        self.donut_chart = DonutChart()
        self.donut_chart.setFixedHeight(240)
        chart_layout.addWidget(self.donut_chart)

        self.legend = LegendWidget()
        chart_layout.addWidget(self.legend)

        self._update_donut_and_legend(animate=True)
        content_layout.addWidget(chart_card)

        self.card_container = QWidget()
        self.card_container.setStyleSheet("background: transparent;")
        self.card_layout = QVBoxLayout(self.card_container)
        self.card_layout.setContentsMargins(0, 0, 0, 0)
        self.card_layout.setSpacing(12)
        self._populate_cards()
        content_layout.addWidget(self.card_container)
        content_layout.addStretch()

        scroll.setWidget(self.content_widget)
        outer_layout.addWidget(scroll, stretch=1)

        self.footer_label = QLabel("")
        self.footer_label.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 12px;")
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.footer_label.setWordWrap(True)
        outer_layout.addWidget(self.footer_label)
        self._update_footer()

        close_btn = QPushButton("关闭")
        close_btn.setFixedSize(140, 38)
        close_btn.clicked.connect(self.accept)
        outer_layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

    # ---------- 以下方法逻辑保持不变 ----------
    def _collect_category_segments(self):
        segments = []
        for c in self.categories:
            if c.get("size", 0) <= 0:
                continue
            segments.append({
                "name": c["name"],
                "size": c["size"],
                "color": c["color"],
            })
        num_str, unit_str = _format_size_split(self.total_size)
        if self.total_size > 0:
            sub = "user 总占用"
        else:
            sub = "暂无数据"
        return segments, num_str, unit_str, sub

    def _update_donut_and_legend(self, animate=True):
        segments, num_str, unit_str, sub = self._collect_category_segments()
        self.donut_chart.set_data(segments, num_str, unit_str, sub, animate=animate)
        self.legend.update_items(segments)

    def _populate_cards(self):
        while self.card_layout.count():
            item = self.card_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

        shown = [c for c in self.categories if c.get("size", 0) > 0]
        if not shown:
            _c = get_theme_colors()
            empty = QLabel("暂无数据")
            empty.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 13px;")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.card_layout.addWidget(empty)
            return

        for c in shown:
            card = CategoryCard(c, on_open=self._open_category)
            self.card_layout.addWidget(card)

    def _open_category(self, category):
        dlg = StorageSubDialog(category, self)
        dlg.exec()

    def _refresh(self):
        scrollbar = self.scroll.verticalScrollBar()
        saved_scroll = scrollbar.value() if scrollbar else 0

        apply_keep_counts("exports", os.path.abspath(EXPORT_DIR))
        apply_keep_counts("backups", os.path.abspath(BACKUP_DIR))

        self.categories = self._collect_storage_data()
        self.total_size = sum(c["size"] for c in self.categories)

        self._update_donut_and_legend(animate=True)
        self._populate_cards()
        self._update_footer()

        if scrollbar:
            scrollbar.setValue(saved_scroll)

    def _open_user_dir(self):
        path = os.path.abspath(USER_DIR)
        if os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _update_footer(self):
        try:
            import psutil
            disk = psutil.disk_usage(os.path.abspath(USER_DIR))
            text = (f"user 目录共占用 {_format_size(self.total_size)}    |    "
                    f"磁盘已用 {_format_size(disk.used)} / {_format_size(disk.total)}"
                    f"（剩余 {_format_size(disk.free)}）")
        except Exception:
            text = f"user 目录共占用 {_format_size(self.total_size)}"
        self.footer_label.setText(text)

    @staticmethod
    def _list_files(path):
        items = []
        if not os.path.exists(path):
            return items
        try:
            for f in os.listdir(path):
                p = os.path.join(path, f)
                if os.path.isfile(p):
                    try:
                        items.append({
                            "name": f,
                            "path": p,
                            "size": os.path.getsize(p)
                        })
                    except OSError:
                        pass
        except OSError:
            pass
        items.sort(key=lambda x: -x["size"])
        return items

    def _collect_storage_data(self):
        categories = []
        csv_items = [it for it in self._list_files(DATA_DIR) if it["name"].lower().endswith(".csv")]
        categories.append({
            "key": "csv", "name": "积分数据", "icon": "📊",
            "size": sum(x["size"] for x in csv_items),
            "color": "#64B5F6", "desc": "包含所有班级积分记录、成绩数据的 CSV 文件。",
            "items": csv_items, "dir_path": os.path.abspath(DATA_DIR),
        })
        excel_items = [it for it in self._list_files(DATA_DIR) if it["name"].lower().endswith((".xlsx", ".xls"))]
        categories.append({
            "key": "excel", "name": "名单文件", "icon": "📁",
            "size": sum(x["size"] for x in excel_items),
            "color": "#81C784", "desc": "包含班级名单、导出模板等 Excel 文件。",
            "items": excel_items, "dir_path": os.path.abspath(DATA_DIR),
        })
        exports_items = self._list_files(EXPORT_DIR)
        categories.append({
            "key": "exports", "name": "导出文件", "icon": "📤",
            "size": sum(x["size"] for x in exports_items),
            "color": "#FFB74D", "desc": "所有导出的 Excel 汇总表、成绩总览等文件。",
            "items": exports_items, "dir_path": os.path.abspath(EXPORT_DIR),
            "button_text": "管理",
        })
        backups_items = self._list_files(BACKUP_DIR)
        categories.append({
            "key": "backups", "name": "备份文件", "icon": "💾",
            "size": sum(x["size"] for x in backups_items),
            "color": "#F06292", "desc": "系统自动生成的名单、积分记录等备份文件。",
            "items": backups_items, "dir_path": os.path.abspath(BACKUP_DIR),
            "button_text": "管理",
        })
        json_items = [it for it in self._list_files(DATA_DIR) if it["name"].lower().endswith(".json")]
        categories.append({
            "key": "config", "name": "便签与配置", "icon": "📝",
            "size": sum(x["size"] for x in json_items),
            "color": "#BA68C8", "desc": "包含便签、系统配置、自定义名单等 JSON 文件。",
            "items": json_items, "dir_path": os.path.abspath(DATA_DIR),
        })
        return categories

class _NeuUpdateWindow(QWidget):
    """1280x720 新拟物全屏更新窗口"""
    progress_sig = pyqtSignal(int)
    status_sig   = pyqtSignal(str)
    title_sig    = pyqtSignal(str)
    ready_sig    = pyqtSignal()   # 下载完成 → 主线程调度 2 秒后重启
    failed_sig   = pyqtSignal()   # 下载失败 → 主线程调度 3 秒后关闭

    def __init__(self, version, changelog, download_url):
        super().__init__(None)
        self.setFixedSize(1280, 720)
        self.setWindowTitle("PickOS 更新程序")
        self.setObjectName("neuUpdateWindow")
        self.setStyleSheet("#neuUpdateWindow{background-color:#e0e5ec;}")

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        self.version = version
        self.changelog = changelog
        self.download_url = download_url
        self.installer_path = None

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(100, 60, 100, 60)
        self.main_layout.setSpacing(40)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.build_ui()

    def build_ui(self):
        # 标题
        self.title_label = QLabel(f"发现新版本：{self.version}")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title_label.setStyleSheet(
            "color:#5a6b8c;font-size:36pt;font-weight:700;background:transparent;"
        )
        self.main_layout.addWidget(self.title_label)

        # 内容堆叠
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("background:transparent;")

        # --- 日志视图 ---
        self.log_view = QWidget()
        log_layout = QVBoxLayout(self.log_view)
        log_layout.setContentsMargins(0, 0, 0, 0)

        self.log_browser = QTextBrowser()
        self.log_browser.setMarkdown(
            f"### 更新日志\n\n{self.changelog}\n\n---\n"
            f"**提示**：更新将保留您所有的数据 (`user` 文件夹)，请放心下载。"
        )
        self.log_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #e0e5ec;
                border-top: 4px solid #ffffff;
                border-left: 4px solid #ffffff;
                border-bottom: 4px solid #a3b1c6;
                border-right: 4px solid #a3b1c6;
                border-radius: 16px;
                color: #5a6b8c;
                font-family: "Microsoft YaHei";
                font-size: 15px;
                padding: 20px;
                selection-background-color: #b7c4d8;
            }
            QScrollBar:vertical { width: 10px; background: transparent; margin: 0; }
            QScrollBar::handle:vertical { background: #a3b1c6; border-radius: 5px; min-height: 30px; }
            QScrollBar::handle:vertical:hover { background: #8a9bb5; }
            QScrollBar::add-line, QScrollBar::sub-line { height: 0; }
        """)
        self.log_browser.setFixedSize(960, 300)
        log_layout.addWidget(self.log_browser, alignment=Qt.AlignmentFlag.AlignCenter)
        self.content_stack.addWidget(self.log_view)

        # --- 进度视图 ---
        self.progress_view = QWidget()
        prog_layout = QVBoxLayout(self.progress_view)
        prog_layout.setContentsMargins(0, 0, 0, 0)
        prog_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        prog_layout.setSpacing(30)

        self.percent_label = QLabel("0%")
        self.percent_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.percent_label.setStyleSheet(
            "color:#5a6b8c;font-size:96pt;font-weight:900;background:transparent;"
        )

        self.status_label = QLabel("正在准备下载...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(
            "color:#8a9bb5;font-size:22pt;font-weight:500;background:transparent;"
        )

        prog_layout.addWidget(self.percent_label)
        prog_layout.addWidget(self.status_label)
        self.content_stack.addWidget(self.progress_view)

        self.main_layout.addWidget(self.content_stack)

        # 按钮区
        self.btn_container = QWidget()
        self.btn_layout = QHBoxLayout(self.btn_container)
        self.btn_layout.setSpacing(60)
        self.btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.update_btn = _NeuButton("立即更新")
        self.update_btn.setFixedSize(240, 120)
        self.update_btn._radius = 24.0
        self.update_btn._blur = 10
        self.update_btn._distance = 4
        self.update_btn._cache = None
        self.update_btn._font.setPointSize(20)
        self.update_btn.clicked.connect(self.start_update)

        self.cancel_btn = _NeuButton("暂不更新")
        self.cancel_btn.setFixedSize(240, 120)
        self.cancel_btn._radius = 24.0
        self.cancel_btn._blur = 10
        self.cancel_btn._distance = 4
        self.cancel_btn._cache = None
        self.cancel_btn._font.setPointSize(20)
        self.cancel_btn.clicked.connect(self.close)

        self.btn_layout.addWidget(self.update_btn)
        self.btn_layout.addWidget(self.cancel_btn)

        self.main_layout.addWidget(self.btn_container)

        # ★ 全部信号在主线程接收
        self.progress_sig.connect(self._set_percent)
        self.status_sig.connect(self.status_label.setText)
        self.title_sig.connect(self.title_label.setText)
        self.ready_sig.connect(self._on_download_ready)
        self.failed_sig.connect(self._on_download_failed)

    # ---------- 主线程槽 ----------
    def _set_percent(self, p):
        self.percent_label.setText(f"{p}%")

    def _on_download_ready(self):
        """下载完成（主线程）：0.8 秒后启动安装程序"""
        QTimer.singleShot(800, self._apply_update)

    def _on_download_failed(self):
        """下载失败（主线程）：3 秒后关闭"""
        QTimer.singleShot(3000, self.close)

    # ---------- 子线程 ----------
    def start_update(self):
        self.btn_container.setVisible(False)
        self.content_stack.setCurrentWidget(self.progress_view)
        self.title_label.setText("正在更新 PickOS")
        threading.Thread(target=self._download_thread, daemon=True).start()

    def _download_thread(self):
        self.status_sig.emit("正在下载安装包...")

        def cb(p):
            self.progress_sig.emit(p)

        self.installer_path = UpdateManager.download_installer(
            self.download_url, cb
        )

        if self.installer_path:
            self.progress_sig.emit(100)
            self.status_sig.emit("下载完成，即将启动安装程序...")
            self.title_sig.emit("更新准备就绪")
            self.ready_sig.emit()
        else:
            self.status_sig.emit("下载失败，请检查网络连接或稍后重试。")
            self.failed_sig.emit()

    def _apply_update(self):
        if self.installer_path:
            UpdateManager.run_installer_and_exit(self.installer_path)

# ======================== 设置页面 ========================

class AnimationSettingsDialog(QDialog):
    """动画设置：开机动画 + 进度条动画（合并成一个大按钮入口）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎬 动画设置")
        self.setFixedSize(500, 380)
        self.setStyleSheet(DIALOG_STYLE)

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 22, 25, 22)
        layout.setSpacing(14)

        # ---- 标题 ----
        title = QLabel("🎬 动画设置")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        desc = QLabel("控制系统启动动画与进度条动画的开关。")
        desc.setStyleSheet("color: #A0A0B0; font-size: 13px;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 分割线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background: #2A2A3A; max-height: 1px; border: none;")
        layout.addWidget(line)

        cfg = load_config()

        # ---- 开关项 1：开机动画 ----
        row1 = QHBoxLayout()
        info1 = QVBoxLayout()
        info1.setSpacing(2)
        label1 = QLabel("开机动画")
        label1.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub1 = QLabel("程序启动时是否播放 PickOS 开机动画（重启后生效）")
        sub1.setStyleSheet("color: #808090; font-size: 12px;")
        sub1.setWordWrap(True)
        info1.addWidget(label1)
        info1.addWidget(sub1)
        row1.addLayout(info1, stretch=1)

        self.boot_switch = SwitchButton(self)
        self.boot_switch.setChecked(cfg.get("show_boot_animation", True))
        row1.addWidget(self.boot_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row1)

        # ---- 开关项 2：进度条动画 ----
        row2 = QHBoxLayout()
        info2 = QVBoxLayout()
        info2.setSpacing(2)
        label2 = QLabel("进度条动画")
        label2.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub2 = QLabel("切换到页面时，进度条与圆环从 0 缓动到当前值")
        sub2.setStyleSheet("color: #808090; font-size: 12px;")
        sub2.setWordWrap(True)
        info2.addWidget(label2)
        info2.addWidget(sub2)
        row2.addLayout(info2, stretch=1)

        self.progress_anim_switch = SwitchButton(self)
        self.progress_anim_switch.setChecked(cfg.get("enable_progress_animation", True))
        row2.addWidget(self.progress_anim_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row2)

        # ---- 开关项 3：导航栏切换动画 ----
        row3 = QHBoxLayout()
        info3 = QVBoxLayout()
        info3.setSpacing(2)
        label3 = QLabel("导航栏切换动画")
        label3.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub3 = QLabel("点击左侧导航按钮时，图标弹动；切换到设置时齿轮额外旋转")
        sub3.setStyleSheet("color: #808090; font-size: 12px;")
        sub3.setWordWrap(True)
        info3.addWidget(label3)
        info3.addWidget(sub3)
        row3.addLayout(info3, stretch=1)

        self.nav_anim_switch = SwitchButton(self)
        self.nav_anim_switch.setChecked(cfg.get("enable_nav_animation", True))
        row3.addWidget(self.nav_anim_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row3)

        layout.addStretch()

        # ---- 底部按钮 ----
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedWidth(110)
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("确定")
        ok_btn.setFixedWidth(110)
        ok_btn.clicked.connect(self._on_ok)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _on_ok(self):
        cfg = load_config()
        cfg["show_boot_animation"] = self.boot_switch.isChecked()
        cfg["enable_progress_animation"] = self.progress_anim_switch.isChecked()
        cfg["enable_nav_animation"] = self.nav_anim_switch.isChecked()
        save_config(cfg)

        if self.parent() is not None:
            show_toast(self.parent(), "动画设置已保存", "success")
        self.accept()

# ======================== 粒子效果设置对话框 ========================
class ParticleSettingsDialog(QDialog):
    """粒子效果设置：窗口关闭粒子 + 鼠标点击粒子"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("✨ 粒子效果设置")
        self.setFixedSize(500, 400)
        self.setStyleSheet(DIALOG_STYLE)

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 22, 25, 22)
        layout.setSpacing(14)

        # ---- 标题 ----
        title = QLabel("✨ 粒子效果")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        desc = QLabel("控制系统内各处粒子动画的开关，关闭后对应粒子效果将不再播放。")
        desc.setStyleSheet("color: #A0A0B0; font-size: 13px;")
        desc.setWordWrap(True)
        layout.addWidget(desc)

        # 分割线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background: #2A2A3A; max-height: 1px; border: none;")
        layout.addWidget(line)

        # ---- 开关项 1：窗口关闭粒子 ----
        cfg = load_config()

        row1 = QHBoxLayout()
        info1 = QVBoxLayout()
        info1.setSpacing(2)
        label1 = QLabel("窗口关闭粒子效果")
        label1.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub1 = QLabel("弹窗、主窗口关闭、最小化、悬浮窗双击时，播放粒子散开效果")
        sub1.setStyleSheet("color: #808090; font-size: 12px;")
        sub1.setWordWrap(True)
        info1.addWidget(label1)
        info1.addWidget(sub1)
        row1.addLayout(info1, stretch=1)

        self.close_particle_switch = SwitchButton(self)
        self.close_particle_switch.setChecked(cfg.get("enable_particle_effect", True))
        row1.addWidget(self.close_particle_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row1)

        # ---- 开关项 2：鼠标点击粒子 ----
        row2 = QHBoxLayout()
        info2 = QVBoxLayout()
        info2.setSpacing(2)
        label2 = QLabel("鼠标点击粒子效果")
        label2.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub2 = QLabel("鼠标左键点击任意位置时，在点击处溅射小粒子")
        sub2.setStyleSheet("color: #808090; font-size: 12px;")
        sub2.setWordWrap(True)
        info2.addWidget(label2)
        info2.addWidget(sub2)
        row2.addLayout(info2, stretch=1)

        self.click_particle_switch = SwitchButton(self)
        self.click_particle_switch.setChecked(cfg.get("enable_click_particle_effect", True))
        row2.addWidget(self.click_particle_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row2)

        # ---- 开关项 3：拖动拖尾粒子 ----
        row3 = QHBoxLayout()
        info3 = QVBoxLayout()
        info3.setSpacing(2)
        label3 = QLabel("鼠标拖动粒子效果")
        label3.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub3 = QLabel("按住鼠标左键拖动时，在光标路径上拖出粒子尾迹")
        sub3.setStyleSheet("color: #808090; font-size: 12px;")
        sub3.setWordWrap(True)
        info3.addWidget(label3)
        info3.addWidget(sub3)
        row3.addLayout(info3, stretch=1)

        self.drag_particle_switch = SwitchButton(self)
        self.drag_particle_switch.setChecked(cfg.get("enable_drag_particle_effect", True))
        row3.addWidget(self.drag_particle_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row3)

        # ---- 开关项 4：鼠标滚动粒子 ----
        row4 = QHBoxLayout()
        info4 = QVBoxLayout()
        info4.setSpacing(2)
        label4 = QLabel("鼠标滚轮粒子效果")
        label4.setStyleSheet("color: {_c['fg']}; font-size: 15px; font-weight: 600;")
        sub4 = QLabel("鼠标滚轮滚动时，在光标位置溅射向上飘的粒子")
        sub4.setStyleSheet("color: #808090; font-size: 12px;")
        sub4.setWordWrap(True)
        info4.addWidget(label4)
        info4.addWidget(sub4)
        row4.addLayout(info4, stretch=1)

        self.scroll_particle_switch = SwitchButton(self)
        self.scroll_particle_switch.setChecked(cfg.get("enable_scroll_particle_effect", True))
        row4.addWidget(self.scroll_particle_switch, alignment=Qt.AlignmentFlag.AlignVCenter)
        layout.addLayout(row4)

        layout.addStretch()

        # ---- 底部按钮 ----
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedWidth(110)
        cancel_btn.clicked.connect(self.reject)
        ok_btn = QPushButton("确定")
        ok_btn.setFixedWidth(110)
        ok_btn.clicked.connect(self._on_ok)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _on_ok(self):
        cfg = load_config()
        cfg["enable_particle_effect"] = self.close_particle_switch.isChecked()
        cfg["enable_click_particle_effect"] = self.click_particle_switch.isChecked()
        cfg["enable_drag_particle_effect"] = self.drag_particle_switch.isChecked()
        cfg["enable_scroll_particle_effect"] = self.scroll_particle_switch.isChecked()  # ← 新增
        save_config(cfg)

        if self.parent() is not None:
            show_toast(self.parent(), "粒子效果设置已保存", "success")
        self.accept()

# ======================== 新拟物风格控件（Neumorphism） ========================
# 用于「密码设置」对话框，整体视觉参照新拟物设计：浅色背景 + 双向柔和阴影
# 主程序是深色主题，这个对话框独立使用新拟物风格，突出层次感

class _NeuColors:
    BASE      = QColor("#e0e5ec")
    LIGHT     = QColor("#ffffff")
    DARK      = QColor("#a3b1c6")
    TEXT      = QColor("#5a6b8c")
    TEXT_SOFT = QColor("#8a9bb5")
    FOCUS     = QColor(120, 150, 200, 90)
    ACCENT    = QColor(100, 181, 246)
    DANGER    = QColor("#d9534f")


def _hash_password(pwd: str, salt: str = "PickOS_2026_v1") -> str:
    """密码哈希（SHA-256 + 固定盐）。无法逆向，忘记后不可找回。"""
    return hashlib.sha256((salt + pwd).encode("utf-8")).hexdigest()

# ======================== 智能解锁：移动磁盘识别 ========================

# ======================== 智能解锁：开发者 U 盘 ========================

DEVELOPER_USB_SERIALS = {"206DE36F"}
DEVELOPER_USB_MARKER = "PickOS_DEVELOPER.key"


def _drive_has_developer_marker(drive_path):
    """检查磁盘根目录是否有开发者标识文件"""
    try:
        marker = os.path.join(drive_path, DEVELOPER_USB_MARKER)
        return os.path.isfile(marker)
    except Exception:
        return False


def is_developer_usb(drive):
    """判断某个可移动磁盘是否为开发者 U 盘"""
    serial = str(drive.get("serial", "")).strip().upper()
    if serial and serial in DEVELOPER_USB_SERIALS:
        return True
    return _drive_has_developer_marker(drive.get("path", ""))


def find_developer_usb():
    """查找当前插入的开发者 U 盘（不受智能解锁开关影响）"""
    try:
        for d in get_removable_drives():
            if is_developer_usb(d):
                return d
    except Exception:
        pass
    return None


def find_bound_smart_unlock_device():
    """查找当前插入的、已绑定的普通智能解锁磁盘（要求智能解锁已开启）"""
    cfg = load_config()
    if not cfg.get("smart_unlock_enabled", False):
        return None
    devices = cfg.get("smart_unlock_device", [])
    if not isinstance(devices, list):
        return None
    expected = {
        str(d.get("serial", "")).strip().upper()
        for d in devices
        if isinstance(d, dict) and d.get("serial")
    }
    if not expected:
        return None
    try:
        for d in get_removable_drives():
            if str(d.get("serial", "")).strip().upper() in expected:
                return d
    except Exception:
        pass
    return None


def show_smart_unlock_toast(device, is_developer=False):
    """统一显示智能解锁成功提示"""
    label = device.get("label", "已绑定的移动磁盘")
    if is_developer:
        text = (
            "智能解锁成功\n"
            f"检测到开发者设备：{label}\n"
            "最高管理员权限"
        )
    else:
        text = f"智能解锁成功\n检测到设备：{label}"
    try:
        toast = _NeuAutoToast(text, duration=1800)
        toast.exec()
    except Exception as e:
        print(f"⚠️ 智能解锁提示显示失败：{e}")


def get_removable_drives():
    """跨平台获取当前挂载的可移动磁盘列表。
    返回 [{"path": "E:\\", "label": "U盘", "serial": "1234ABCD", "fs": "FAT32"}, ...]
    """
    drives = []
    if sys.platform.startswith("win"):
        try:
            import ctypes
            import string
            kernel32 = ctypes.windll.kernel32
            bitmask = kernel32.GetLogicalDrives()
            for i, letter in enumerate(string.ascii_uppercase):
                if not (bitmask & (1 << i)):
                    continue
                root = f"{letter}:\\"
                try:
                    # DRIVE_REMOVABLE = 2
                    if kernel32.GetDriveTypeW(root) != 2:
                        continue
                    volume_name = ctypes.create_unicode_buffer(261)
                    fs_name = ctypes.create_unicode_buffer(261)
                    serial = ctypes.c_ulong(0)
                    max_comp = ctypes.c_ulong(0)
                    flags = ctypes.c_ulong(0)
                    ok = kernel32.GetVolumeInformationW(
                        root, volume_name, len(volume_name),
                        ctypes.byref(serial), ctypes.byref(max_comp),
                        ctypes.byref(flags), fs_name, len(fs_name),
                    )
                    if not ok:
                        continue
                    drives.append({
                        "path": root,
                        "label": volume_name.value or "(无标签)",
                        "serial": f"{serial.value:08X}",
                        "fs": fs_name.value or "",
                    })
                except Exception:
                    continue
        except Exception:
            pass
    else:
        # macOS / Linux
        try:
            import psutil
            candidates = []
            if sys.platform == "darwin":
                if os.path.isdir("/Volumes"):
                    for name in os.listdir("/Volumes"):
                        mp = os.path.join("/Volumes", name)
                        if os.path.isdir(mp):
                            candidates.append((mp, name))
            else:  # Linux
                for p in psutil.disk_partitions(all=False):
                    opts = (p.opts or "").lower()
                    if ("removable" in opts
                            or p.device.startswith("/dev/sd")
                            or p.mountpoint.startswith("/media/")
                            or p.mountpoint.startswith("/run/media/")):
                        candidates.append((p.mountpoint,
                                           os.path.basename(p.mountpoint) or p.device))

            for mp, label in candidates:
                try:
                    usage = psutil.disk_usage(mp)
                    # 用总容量（16 位十六进制）作为稳定标识
                    serial = f"{usage.total:016X}"
                    drives.append({
                        "path": mp,
                        "label": label or "(无标签)",
                        "serial": serial,
                        "fs": "",
                    })
                except Exception:
                    continue
        except Exception:
            pass
    return drives


def is_smart_unlock_available():
    """兼容旧接口：开发者 U 盘或已绑定磁盘任一可用即返回 True"""
    return (find_developer_usb() is not None) or (find_bound_smart_unlock_device() is not None)

_blur_scene = None


def _get_blur_scene():
    global _blur_scene
    if _blur_scene is None:
        _blur_scene = QGraphicsScene()
    return _blur_scene


def _blur_image(img, radius):
    scene = _get_blur_scene()
    item = QGraphicsPixmapItem(QPixmap.fromImage(img))
    eff = QGraphicsBlurEffect()
    eff.setBlurRadius(radius)
    eff.setBlurHints(QGraphicsBlurEffect.BlurHint.QualityHint)
    item.setGraphicsEffect(eff)
    scene.addItem(item)

    out = QImage(img.size(), QImage.Format.Format_ARGB32_Premultiplied)
    out.fill(Qt.GlobalColor.transparent)
    p = QPainter(out)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    scene.render(p, QRectF(out.rect()), QRectF(img.rect()),
                 Qt.AspectRatioMode.IgnoreAspectRatio)
    p.end()
    scene.removeItem(item)
    return out


def _render_outer_shadow(size, radius, color, blur):
    pad = blur * 2 + 4
    w, h = size.width() + pad * 2, size.height() + pad * 2
    img = QImage(w, h, QImage.Format.Format_ARGB32_Premultiplied)
    img.fill(Qt.GlobalColor.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(color)
    p.drawRoundedRect(QRectF(pad, pad, size.width(), size.height()), radius, radius)
    p.end()
    return _blur_image(img, blur), pad


def _render_inner_shadow(size, radius, color, blur):
    pad = blur * 2 + 4
    w, h = size.width() + pad * 2, size.height() + pad * 2
    img = QImage(w, h, QImage.Format.Format_ARGB32_Premultiplied)
    img.fill(Qt.GlobalColor.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    outer = QPainterPath()
    outer.addRect(QRectF(-w, -h, w * 3, h * 3))
    inner = QPainterPath()
    inner.addRoundedRect(QRectF(pad, pad, size.width(), size.height()),
                         radius, radius)
    p.fillPath(outer.subtracted(inner), color)
    p.end()
    return _blur_image(img, blur), pad


class _NeuWidget(QWidget):
    """新拟物控件基类：凸起 / 凹陷两种状态，子类只需要实现 paint_content"""
    BODY_MIN_W = 170
    BODY_MIN_H = 60

    def __init__(self, parent=None):
        super().__init__(parent)
        self._base     = QColor(_NeuColors.BASE)
        self._light    = QColor(_NeuColors.LIGHT)
        self._dark     = QColor(_NeuColors.DARK)
        self._radius   = 16.0
        self._distance = 5
        self._blur     = 12
        self._t        = 0.0
        self._cache    = None

    @property
    def margin(self):
        return self._blur * 2 + 4 + self._distance

    def body_rect(self):
        m = self.margin
        return QRectF(m, m, self.width() - 2 * m, self.height() - 2 * m)

    def sizeHint(self):
        return QSize(2 * self.margin + self.BODY_MIN_W,
                     2 * self.margin + self.BODY_MIN_H)

    def _build_cache(self):
        br = self.body_rect()
        size = QSize(max(1, round(br.width())), max(1, round(br.height())))
        key = (size.width(), size.height(), self._radius, self._blur,
               self._distance, self._light.rgba(), self._dark.rgba())
        if self._cache and self._cache[0] == key:
            return
        light, pad = _render_outer_shadow(size, self._radius, self._light, self._blur)
        dark,  _   = _render_outer_shadow(size, self._radius, self._dark,  self._blur)
        din,   _   = _render_inner_shadow(size, self._radius, self._dark,  self._blur)
        lin,   _   = _render_inner_shadow(size, self._radius, self._light, self._blur)
        self._cache = (key, light, dark, din, lin, pad)

    def paintEvent(self, event):
        self._build_cache()
        if not self._cache:
            return
        _, light, dark, din, lin, pad = self._cache
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        p.fillRect(self.rect(), self._base)

        br = self.body_rect()
        m  = self.margin
        d  = self._distance
        t  = self._t

        path = QPainterPath()
        path.addRoundedRect(br, self._radius, self._radius)

        if t < 0.999:
            p.setOpacity(1.0 - t)
            p.drawImage(QPointF(m - pad - d, m - pad - d), light)
            p.drawImage(QPointF(m - pad + d, m - pad + d), dark)
            p.setOpacity(1.0)

        p.fillPath(path, self._base)

        if t > 0.001:
            p.save()
            p.setClipPath(path)
            p.setOpacity(t)
            p.drawImage(QPointF(m - pad + d, m - pad + d), din)
            p.drawImage(QPointF(m - pad - d, m - pad - d), lin)
            p.restore()

        self.paint_content(p, br)
        p.end()

    def paint_content(self, p, rect):
        pass


class _NeuButton(_NeuWidget):
    clicked = pyqtSignal()
    BODY_MIN_W = 140
    BODY_MIN_H = 50

    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self._text = text
        self._down = False
        self._font = QFont(self.font())
        self._font.setPointSize(11)
        self._font.setWeight(QFont.Weight.DemiBold)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumSize(160, 90)
        self._anim = QPropertyAnimation(self, b"pressT", self)
        self._anim.setDuration(130)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        # ★ 新增：兜底轮询，防止「按下后拖动、release 丢失」导致卡在凹陷状态
        self._poll_timer = QTimer(self)
        self._poll_timer.setInterval(25)
        self._poll_timer.timeout.connect(self._poll_mouse_released)

    # ---------- ★ 新增：兜底检测 ----------
    def _poll_mouse_released(self):
        if self._down and not (QApplication.mouseButtons() & Qt.MouseButton.LeftButton):
            self._down = False
            self._poll_timer.stop()
            self._animate_to(0.0)
            self._on_force_release()

    def _on_force_release(self):
        """子类可重写：release 丢失被兜底复位时执行的清理逻辑。"""
        pass

    @pyqtProperty(float)
    def pressT(self):
        return self._t

    @pressT.setter
    def pressT(self, value):
        self._t = float(value)
        self.update()

    def _animate_to(self, target):
        self._anim.stop()
        self._anim.setStartValue(self._t)
        self._anim.setEndValue(target)
        self._anim.start()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._down = True
            self._animate_to(1.0)
            self._poll_timer.start()            # ★ 只在按下期间开，静态零开销

    def mouseReleaseEvent(self, e):
        was_down = self._down
        self._down = False
        if e.button() == Qt.MouseButton.LeftButton:
            self._poll_timer.stop()             # ★ 正常路径关掉轮询
            self._animate_to(0.0)
            if was_down and self.rect().contains(e.position().toPoint()):
                self.clicked.emit()

    def paint_content(self, p, rect):
        p.setFont(self._font)
        p.setPen(QPen(_NeuColors.TEXT))
        p.drawText(rect, Qt.AlignmentFlag.AlignCenter, self._text)

class _NeuRoundButton(_NeuButton):
    """圆形新拟物按钮"""
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setMinimumSize(80, 80)
        self._font.setPointSize(18)

    def resizeEvent(self, e):
        br = self.body_rect()
        self._radius = max(1.0, min(br.width(), br.height()) / 2.0)
        self._cache = None
        super().resizeEvent(e)

    def paint_content(self, p, rect):
        p.setFont(self._font)
        p.setPen(QPen(_NeuColors.TEXT))
        p.drawText(rect, Qt.AlignmentFlag.AlignCenter, self._text)


class _NeuLineEdit(_NeuWidget):
    BODY_MIN_W = 280
    BODY_MIN_H = 56

    def __init__(self, placeholder="", password=False, parent=None):
        super().__init__(parent)
        self._t = 1.0
        self._radius = 14.0
        self._focused = False
        self.edit = QLineEdit(self)
        self.edit.setPlaceholderText(placeholder)
        self.edit.setFrame(False)
        self.edit.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        if password:
            self.edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.edit.setStyleSheet(
            "QLineEdit{background:transparent;border:none;color:#5a6b8c;"
            "font-size:11pt;selection-background-color:#b7c4d8;"
            "selection-color:#ffffff;}"
        )
        pal = self.edit.palette()
        pal.setColor(QPalette.ColorRole.PlaceholderText, QColor("#9aabc2"))
        self.edit.setPalette(pal)
        self.edit.installEventFilter(self)

    def eventFilter(self, obj, event):
        if obj is self.edit:
            if event.type() == QEvent.Type.FocusIn:
                self._focused = True
                self.update()
            elif event.type() == QEvent.Type.FocusOut:
                self._focused = False
                self.update()
        return super().eventFilter(obj, event)

    def resizeEvent(self, e):
        br = self.body_rect()
        px, py = 22, 8
        self.edit.setGeometry(int(br.x() + px), int(br.y() + py),
                              int(br.width() - 2 * px),
                              int(br.height() - 2 * py))
        super().resizeEvent(e)

    def paint_content(self, p, rect):
        if self._focused:
            p.setPen(QPen(_NeuColors.FOCUS, 1.5))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(rect.adjusted(1, 1, -1, -1),
                              self._radius, self._radius)


class _NeuSwitch(QWidget):
    """新拟物风格开关：凹陷轨道 + 凸起圆钮，可点击切换"""
    toggled = pyqtSignal(bool)

    def __init__(self, checked=False, parent=None):
        super().__init__(parent)
        self._checked = bool(checked)
        self.setFixedSize(84, 44)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def isChecked(self):
        return self._checked

    def setChecked(self, checked):
        self._checked = bool(checked)
        self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._checked = not self._checked
            self.toggled.emit(self._checked)
            self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.fillRect(self.rect(), _NeuColors.BASE)

        br = QRectF(10, 9, self.width() - 20, self.height() - 18)
        radius = br.height() / 2

        # 凹陷轨道
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(_NeuColors.DARK)
        p.drawRoundedRect(br, radius, radius)
        inner = br.adjusted(2, 2, -2, -2)
        p.setBrush(_NeuColors.DARK.darker(112))
        p.drawRoundedRect(inner, inner.height() / 2, inner.height() / 2)

        # 圆钮
        knob_size = br.height() - 8
        if self._checked:
            knob_x = br.right() - knob_size - 4
            knob_color = _NeuColors.ACCENT
        else:
            knob_x = br.left() + 4
            knob_color = _NeuColors.LIGHT

        # 圆钮阴影
        p.setBrush(_NeuColors.DARK)
        p.drawEllipse(QRectF(knob_x - 1, br.top() + 5,
                             knob_size + 2, knob_size + 2))
        # 圆钮本体
        p.setBrush(knob_color)
        p.drawEllipse(QRectF(knob_x, br.top() + 4, knob_size, knob_size))
        p.end()

class _NeuHintCard(_NeuWidget):
    """新拟物风格提示卡片（凹陷纯展示，不可点击）"""

    def __init__(self, text, parent=None):
        super().__init__(parent)
        # 凹陷状态：只画内阴影，不画外阴影
        self._t = 1.0
        self._radius = 12.0
        self._blur = 5
        self._distance = 2
        self._cache = None
        self._text = text
        self._font = QFont(self.font())
        self._font.setPointSize(9)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        # 按文字行数估算最小高度
        m = self.margin              # 5*2+4+2 = 16
        lines = text.count('\n') + 1
        self.setMinimumHeight(2 * m + lines * 22 + 12)

    def paint_content(self, p, rect):
        p.setFont(self._font)
        p.setPen(QPen(_NeuColors.TEXT_SOFT))
        text_rect = rect.adjusted(16, 6, -16, -6)
        p.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignLeft
            | Qt.AlignmentFlag.AlignVCenter
            | Qt.TextFlag.TextWordWrap,
            self._text,
        )

class _NeuTipCard(_NeuWidget):
    """新拟物风格凸起提示卡片（向上突出、位于标题与输入框之间）"""

    def __init__(self, text, parent=None, font_size=9):
        super().__init__(parent)
        # 凸起状态：只画外阴影，视觉上"浮出"背景
        self._t = 0.0
        self._radius = 16.0
        self._blur = 5
        self._distance = 3
        self._cache = None
        self._text = text
        self._font = QFont(self.font())
        self._font.setPointSize(font_size)     # ← 可配置字号
        self.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )

    def paint_content(self, p, rect):
        p.setFont(self._font)
        p.setPen(QPen(_NeuColors.TEXT))
        text_rect = rect.adjusted(18, 6, -18, -6)
        p.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap,
            self._text,
        )

class _NeuConfirmDialog(QDialog):
    """新拟物风格的确认/取消对话框，用于替换 QMessageBox.question"""

    def __init__(self, parent, title, text, ok_text="确定", cancel_text="取消"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedSize(620, 400)
        self.setObjectName("neuConfirmDialog")
        self.setStyleSheet("#neuConfirmDialog{background-color:#e0e5ec;}")
        # 关闭默认问号图标
        self.setWindowFlags(
            self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 25)
        layout.setSpacing(16)

        # ---- 标题 ----
        lbl_title = QLabel(title)
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_title.setStyleSheet(
            "color:#5a6b8c;font-size:18pt;font-weight:700;background:transparent;"
        )
        layout.addWidget(lbl_title)

        # ---- 正文卡片（凸起，与其它新拟物弹窗风格一致） ----
        card = _NeuTipCard(text, font_size=12)
        card.setFixedSize(520, 180)
        layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addStretch()

        # ---- 按钮行 ----
        btn_row = QHBoxLayout()
        btn_row.setSpacing(30)
        btn_row.addStretch()

        ok_btn = _NeuButton(ok_text)
        ok_btn.setFixedSize(160, 100)
        ok_btn._radius = 22.0
        ok_btn._blur = 10
        ok_btn._distance = 4
        ok_btn._cache = None
        ok_btn._font.setPointSize(16)
        ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(ok_btn)

        cancel_btn = _NeuButton(cancel_text)
        cancel_btn.setFixedSize(160, 100)
        cancel_btn._radius = 22.0
        cancel_btn._blur = 10
        cancel_btn._distance = 4
        cancel_btn._cache = None
        cancel_btn._font.setPointSize(16)
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

class _NeuAutoToast(QDialog):
    """新拟物风格自动关闭提示框，用于"智能解锁成功"等即时反馈。"""

    def __init__(self, text, duration=1800):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(540, 220)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(0)

        card = _NeuTipCard(text, font_size=16)
        card.setFixedSize(500, 180)
        layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        # 到点自动关闭
        QTimer.singleShot(duration, self.accept)

    def showEvent(self, event):
        super().showEvent(event)
        # 淡入效果
        try:
            eff = QGraphicsOpacityEffect(self)
            self.setGraphicsEffect(eff)
            eff.setOpacity(0.0)
            anim = QPropertyAnimation(eff, b"opacity", self)
            anim.setDuration(250)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            anim.start()
            self._fade_in_anim = anim
        except Exception:
            pass

class _ChangePasswordDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("更改密码")
        self.setFixedSize(650, 700)
        self.setObjectName("changePwdDialog")
        self.setStyleSheet("#changePwdDialog{background-color:#e0e5ec;}")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(18)

        # ===== 标题 =====
        lbl_title = QLabel("设置新密码")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_title.setStyleSheet(
            "color:#5a6b8c;font-size:18pt;font-weight:700;background:transparent;"
        )
        layout.addWidget(lbl_title)

        # ===== 标题与输入框之间：向上突起的圆角矩形说明卡片 =====
        tip = _NeuTipCard(
            "使用 4 位以上密码，可用数字或字母组合\n密码以哈希值加密储存，忘记无法找回！",
            font_size=11,
        )
        tip.setFixedSize(495, 110)
        layout.addWidget(tip, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addSpacing(10)

        # ===== 输入框：加大字号 =====
        _PWD_STYLE = (
            "QLineEdit{background:transparent;border:none;color:#5a6b8c;"
            "font-size:20pt;font-weight:600;"
            "selection-background-color:#b7c4d8;selection-color:#ffffff;}"
        )

        self.new_pwd = _NeuLineEdit("新密码（至少 4 位）", password=True)
        self.new_pwd.setFixedSize(520, 130)
        self.new_pwd._blur = 10
        self.new_pwd._distance = 4
        self.new_pwd._cache = None
        self.new_pwd.edit.setStyleSheet(_PWD_STYLE)
        layout.addWidget(self.new_pwd, alignment=Qt.AlignmentFlag.AlignCenter)

        self.confirm_pwd = _NeuLineEdit("确认新密码", password=True)
        self.confirm_pwd.setFixedSize(520, 130)
        self.confirm_pwd._blur = 10
        self.confirm_pwd._distance = 4
        self.confirm_pwd._cache = None
        self.confirm_pwd.edit.setStyleSheet(_PWD_STYLE)
        layout.addWidget(self.confirm_pwd, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addStretch()

        # ===== 三个按钮 =====
        row = QHBoxLayout()
        row.setSpacing(20)
        row.addStretch()

        BTN_W = 170
        BTN_H = 110

        self.hold_btn = _NeuButton("查看")
        self.hold_btn.setFixedSize(BTN_W, BTN_H)
        self.hold_btn._radius = 20.0
        self.hold_btn._blur = 10
        self.hold_btn._distance = 4
        self.hold_btn._cache = None
        self.hold_btn._font.setPointSize(16)
        self.hold_btn.mousePressEvent = self._hold_press
        self.hold_btn.mouseReleaseEvent = self._hold_release
        row.addWidget(self.hold_btn)

        ok_btn = _NeuButton("保存")
        ok_btn.setFixedSize(BTN_W, BTN_H)
        ok_btn._radius = 20.0
        ok_btn._blur = 10
        ok_btn._distance = 4
        ok_btn._cache = None
        ok_btn._font.setPointSize(16)
        ok_btn.clicked.connect(self._save)
        row.addWidget(ok_btn)

        cancel_btn = _NeuButton("取消")
        cancel_btn.setFixedSize(BTN_W, BTN_H)
        cancel_btn._radius = 20.0
        cancel_btn._blur = 10
        cancel_btn._distance = 4
        cancel_btn._cache = None
        cancel_btn._font.setPointSize(16)
        cancel_btn.clicked.connect(self.reject)
        row.addWidget(cancel_btn)

        row.addStretch()
        layout.addLayout(row)

        # ===== ★ 新增：Enter 键焦点切换 + 保存 =====
        # 新密码框按 Enter → 跳到确认密码框
        self.new_pwd.edit.returnPressed.connect(self._on_new_pwd_enter)
        # 确认密码框按 Enter → 直接保存
        self.confirm_pwd.edit.returnPressed.connect(self._save)

        # 打开时自动聚焦到新密码框
        QTimer.singleShot(0, self.new_pwd.edit.setFocus)

    # ---------- ★ 新增：新密码框 Enter ----------
    def _on_new_pwd_enter(self):
        """新密码框按 Enter → 焦点跳到确认密码框并全选"""
        self.confirm_pwd.edit.setFocus()
        self.confirm_pwd.edit.selectAll()

    # ---------- 按住查看 ----------
    def _hold_press(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.new_pwd.edit.setEchoMode(QLineEdit.EchoMode.Normal)
            self.confirm_pwd.edit.setEchoMode(QLineEdit.EchoMode.Normal)
        # ★ 调用父类：启动 poll_timer、做动画
        _NeuButton.mousePressEvent(self.hold_btn, e)

    def _hold_release(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.new_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)
            self.confirm_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)
        # ★ 调用父类：关闭 poll_timer、做回弹
        _NeuButton.mouseReleaseEvent(self.hold_btn, e)

    # ★ 新增：release 丢失被兜底复位时，把密文恢复回来（否则会一直明文显示）
    def _on_force_release(self):
        self.new_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)

    def _save(self):
        # 保险：按住查看时误点保存 → 先恢复密文再校验
        self.new_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_pwd.edit.setEchoMode(QLineEdit.EchoMode.Password)

        pwd1 = self.new_pwd.edit.text()
        pwd2 = self.confirm_pwd.edit.text()

        # ★ 失败时把焦点回到出错的框并全选，方便直接重输
        if not pwd1:
            show_toast(self, "密码不能为空", "warning")
            self.new_pwd.edit.setFocus()
            return
        if len(pwd1) < 4:
            show_toast(self, "密码不能少于4位", "warning")
            self.new_pwd.edit.setFocus()
            self.new_pwd.edit.selectAll()
            return
        if pwd1 != pwd2:
            show_toast(self, "两次输入的密码不一致", "warning")
            self.confirm_pwd.edit.setFocus()
            self.confirm_pwd.edit.selectAll()
            return

        cfg = load_config()
        cfg["password_hash"] = _hash_password(pwd1)
        save_config(cfg)

        # ★ 先弹成功提示，用户点「确定」后 → 关闭所有对话框并返回设置界面
        self._show_success_tip()
        self._close_all_and_back_to_settings()

    def _show_success_tip(self):
        """保存成功后的新拟物风格提示框"""
        dlg = QDialog(self)
        dlg.setWindowTitle("保存成功")
        dlg.setFixedSize(460, 260)
        dlg.setObjectName("succDlg")
        dlg.setStyleSheet("#succDlg{background-color:#e0e5ec;}")

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(30, 30, 30, 25)
        lay.setSpacing(20)

        lbl = QLabel("✅ 密码已保存")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(
            "color:#4CAF50;font-size:20pt;font-weight:700;background:transparent;"
        )
        lay.addWidget(lbl)

        lay.addStretch()

        # ★ 加大：220×100，radius 24 → 圆角矩形
        ok_btn = _NeuButton("确定")
        ok_btn.setFixedSize(220, 100)
        ok_btn._radius = 24.0
        ok_btn._blur = 10
        ok_btn._distance = 4
        ok_btn._cache = None
        ok_btn._font.setPointSize(18)
        ok_btn.clicked.connect(dlg.accept)
        lay.addWidget(ok_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        dlg.exec()

    def _close_all_and_back_to_settings(self):
        """关闭当前对话框及其所有父对话框，返回设置界面"""
        # 先收集父级 QDialog（PasswordSettingsDialog / 密码验证框等）
        parents = []
        p = self.parent()
        while p is not None:
            if isinstance(p, QDialog):
                parents.append(p)
            p = p.parent() if hasattr(p, "parent") else None

        # 关闭自身
        try:
            self.accept()
        except Exception:
            self.close()

        # 依次关闭父对话框（PasswordSettingsDialog → 设置界面）
        for d in parents:
            try:
                d.accept()
            except Exception:
                try:
                    d.reject()
                except Exception:
                    try:
                        d.close()
                    except Exception:
                        pass

class _SmartUnlockDialog(QDialog):
    """智能解锁：支持最多绑定 3 个移动磁盘（新拟物风格，650×700）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("智能解锁")
        self.setFixedSize(650, 700)
        self.setObjectName("smartUnlockDialog")
        self.setStyleSheet("#smartUnlockDialog{background-color:#e0e5ec;}")

        cfg = load_config()
        self._enabled = bool(cfg.get("smart_unlock_enabled", False))

        devices = cfg.get("smart_unlock_device", [])
        if isinstance(devices, dict):
            devices = [devices] if devices and devices.get("serial") else []
        if not isinstance(devices, list):
            devices = []
        self._devices = [
            d for d in devices
            if isinstance(d, dict) and d.get("serial")
        ][:3]

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(14)

        # ===== 标题 =====
        lbl_title = QLabel("智能解锁")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_title.setStyleSheet(
            "color:#5a6b8c;font-size:18pt;font-weight:700;background:transparent;"
        )
        layout.addWidget(lbl_title)

        # ===== 说明卡片 =====
        tip = _NeuTipCard(
            "启用后，插入已绑定的移动磁盘即可自动解锁，\n"
            "无需手动输入密码。\n"
            "最多可绑定 3 个磁盘，适合日常使用。",
            font_size=11,
        )
        tip.setFixedSize(495, 120)
        layout.addWidget(tip, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addSpacing(6)

        # ===== 开关行 =====
        switch_row = QHBoxLayout()
        switch_row.setContentsMargins(55, 0, 55, 0)
        lbl_switch = QLabel("启用智能解锁")
        lbl_switch.setStyleSheet(
            "color:#5a6b8c;font-size:15pt;font-weight:600;background:transparent;"
        )
        self.switch = _NeuSwitch(self._enabled)
        self.switch.toggled.connect(self._on_switch)
        switch_row.addWidget(lbl_switch, stretch=1)
        switch_row.addWidget(self.switch)
        layout.addLayout(switch_row)

        # ===== 状态卡片（凹陷，纯展示） =====
        self.status_card = _NeuHintCard("")
        self.status_card.setFixedSize(520, 180)
        layout.addWidget(self.status_card, alignment=Qt.AlignmentFlag.AlignCenter)

        layout.addSpacing(4)

        # ===== 操作按钮行（扫描并绑定 / 解除绑定） =====
        op_row = QHBoxLayout()
        op_row.setSpacing(20)
        op_row.addStretch()

        self.bind_btn = _NeuButton("扫描并绑定")
        self.bind_btn.setFixedSize(170, 110)
        self.bind_btn._radius = 20.0
        self.bind_btn._blur = 10
        self.bind_btn._distance = 4
        self.bind_btn._cache = None
        self.bind_btn._font.setPointSize(15)
        self.bind_btn.clicked.connect(self._on_bind)
        op_row.addWidget(self.bind_btn)

        self.unbind_btn = _NeuButton("解除绑定")
        self.unbind_btn.setFixedSize(170, 110)
        self.unbind_btn._radius = 20.0
        self.unbind_btn._blur = 10
        self.unbind_btn._distance = 4
        self.unbind_btn._cache = None
        self.unbind_btn._font.setPointSize(15)
        self.unbind_btn.clicked.connect(self._on_unbind)
        op_row.addWidget(self.unbind_btn)

        op_row.addStretch()
        layout.addLayout(op_row)

        layout.addStretch(1)

        # ===== 关闭按钮 =====
        close_row = QHBoxLayout()
        close_row.addStretch()
        close_btn = _NeuButton("关闭")
        close_btn.setFixedSize(170, 110)
        close_btn._radius = 20.0
        close_btn._blur = 10
        close_btn._distance = 4
        close_btn._cache = None
        close_btn._font.setPointSize(16)
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        close_row.addStretch()
        layout.addLayout(close_row)

        # ===== 居中 =====
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        self._refresh_status()

    # ---------- 状态显示 ----------
    def _refresh_status(self):
        has_device = bool(self._devices)
        self.unbind_btn.setEnabled(has_device)
        self.bind_btn.setEnabled(self._enabled and len(self._devices) < 3)

        if not self._enabled:
            text = (
                "智能解锁当前未启用\n"
                "开启后请插入移动磁盘并点击「扫描并绑定」"
            )
        elif not self._devices:
            text = (
                "尚未绑定移动磁盘\n"
                "请插入 U 盘 / 移动硬盘，然后点击「扫描并绑定」\n"
                "（最多可绑定 3 个）"
            )
        else:
            lines = [f"已绑定 {len(self._devices)}/3 个设备："]
            for i, d in enumerate(self._devices, 1):
                label = d.get("label", "未知设备")
                serial = d.get("serial", "")
                lines.append(f"{i}. {label}  SN: {serial}")
            lines.append("（插入任一绑定磁盘后自动解锁）")
            text = "\n".join(lines)

        self.status_card._text = text
        self.status_card.update()

    # ---------- 开关 ----------
    def _on_switch(self, checked):
        self._enabled = bool(checked)
        cfg = load_config()
        cfg["smart_unlock_enabled"] = self._enabled
        save_config(cfg)
        self._refresh_status()
        show_toast(self, "智能解锁已开启" if checked else "智能解锁已关闭",
                   "success")

    # ---------- 扫描并绑定 ----------
    def _on_bind(self):
        if not self._enabled:
            show_toast(self, "请先开启智能解锁开关", "warning")
            return
        if len(self._devices) >= 3:
            show_toast(self, "最多只能绑定 3 个移动磁盘", "warning")
            return

        drives = get_removable_drives()
        if not drives:
            QMessageBox.information(
                self, "未检测到移动磁盘",
                "未检测到可移动磁盘。\n\n请插入 U 盘 / 移动硬盘后重试。"
            )
            return

        selected = None
        if len(drives) == 1:
            selected = drives[0]
        else:
            items = [
                f"{d['path']}  {d['label']}  [SN {d['serial']}]"
                for d in drives
            ]
            choice, ok = QInputDialog.getItem(
                self, "选择磁盘", "请选择要绑定的移动磁盘：",
                items, 0, False
            )
            if not ok or not choice:
                return
            try:
                selected = drives[items.index(choice)]
            except ValueError:
                return

        if not selected or not selected.get("serial"):
            QMessageBox.warning(self, "无法绑定",
                                "无法读取该磁盘的序列号，请尝试其他设备。")
            return

        selected_serial = str(selected.get("serial", "")).strip().upper()
        for d in self._devices:
            if str(d.get("serial", "")).strip().upper() == selected_serial:
                show_toast(self, "该磁盘已经绑定过了", "info")
                return

        confirm = _NeuConfirmDialog(
            self,
            "确认绑定",
            "确定将以下磁盘绑定为智能解锁设备吗？\n\n"
            f"盘符：{selected['path']}\n"
            f"标签：{selected['label']}\n"
            f"序列号：{selected['serial']}",
            ok_text="确定绑定",
            cancel_text="取消",
        )
        if confirm.exec() != QDialog.DialogCode.Accepted:
            return

        self._devices.append(selected)
        self._devices = self._devices[:3]
        cfg = load_config()
        cfg["smart_unlock_device"] = self._devices
        save_config(cfg)
        self._refresh_status()
        show_toast(self, f"已绑定磁盘：{selected['label']}", "success")

    # ---------- 解除绑定 ----------
    def _on_unbind(self):
        if not self._devices:
            return

        if len(self._devices) == 1:
            target = self._devices[0]
        else:
            items = [
                f"{d.get('label', '未知')}  [SN {d.get('serial', '')}]"
                for d in self._devices
            ]
            choice, ok = QInputDialog.getItem(
                self, "选择要解绑的磁盘", "请选择要解除绑定的磁盘：",
                items, 0, False
            )
            if not ok or not choice:
                return
            try:
                target = self._devices[items.index(choice)]
            except ValueError:
                return

        confirm = _NeuConfirmDialog(
            self,
            "解除绑定",
            "确定要解除以下磁盘的绑定吗？\n\n"
            f"标签：{target.get('label', '未知')}\n"
            f"序列号：{target.get('serial', '')}\n\n"
            "解除后将无法通过该磁盘自动解锁。",
            ok_text="解除绑定",
            cancel_text="取消",
        )
        if confirm.exec() != QDialog.DialogCode.Accepted:
            return

        self._devices = [
            d for d in self._devices
            if str(d.get("serial", "")).strip().upper()
            != str(target.get("serial", "")).strip().upper()
        ]
        cfg = load_config()
        cfg["smart_unlock_device"] = self._devices
        save_config(cfg)
        self._refresh_status()
        show_toast(self, "已解除磁盘绑定", "info")

class PasswordPromptDialog(QDialog):
    """启动时的密码验证对话框（新拟物风格，带大号数字键盘）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔒 请输入密码")
        self.setFixedSize(1280, 720)
        self.setObjectName("pwdPromptDialog")
        self.setStyleSheet("#pwdPromptDialog{background-color:#e0e5ec;}")
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.WindowCloseButtonHint
        )

        cfg = load_config()
        self.anti_peek = cfg.get("anti_peek_enabled", False)

        # 主布局：左右分栏（边距减小，给内部腾出更多空间）
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(30)

        # ================= 左半部分：输入与操作 =================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(25)

        title = QLabel("密码验证")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color:#5a6b8c;font-size:36pt;font-weight:700;background:transparent;")
        left_layout.addWidget(title)

        desc = QLabel("PickOS 已启用密码保护\n请输入密码以继续")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setStyleSheet("color:#8a9bb5;font-size:16pt;background:transparent;")
        left_layout.addWidget(desc)

        left_layout.addStretch(1)

        # 大号输入框（减小阴影边距，让输入框可视区域更大）
        self.pwd_edit = _NeuLineEdit("请输入密码", password=True)
        self.pwd_edit.setFixedSize(460, 140)
        self.pwd_edit._blur = 10
        self.pwd_edit._distance = 4
        self.pwd_edit._cache = None
        self.pwd_edit.edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pwd_edit.edit.setStyleSheet(
            "QLineEdit{background:transparent;border:none;color:#5a6b8c;"
            "font-size:32pt;font-weight:bold;selection-background-color:#b7c4d8;"
            "selection-color:#ffffff;}"
        )
        self.pwd_edit.edit.returnPressed.connect(self._verify)
        # 防窥关闭时，输入发生变化自动验证
        self.pwd_edit.edit.textChanged.connect(self._on_text_changed)
        left_layout.addWidget(self.pwd_edit, alignment=Qt.AlignmentFlag.AlignCenter)

        # 错误提示
        self.error_label = QLabel("")
        self.error_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.error_label.setStyleSheet("color:#d9534f;font-size:16pt;font-weight:bold;background:transparent;")
        left_layout.addWidget(self.error_label)

        # 按钮行（按住查看 + 解锁）
        btn_row = QHBoxLayout()
        btn_row.setSpacing(30)
        btn_row.addStretch()

        # 按住查看（圆角矩形，减小阴影边距）
        self.hold_btn = _NeuButton("查看")
        self.hold_btn.setFixedSize(240, 120)
        self.hold_btn._radius = 24.0     
        self.hold_btn._blur = 10
        self.hold_btn._distance = 4
        self.hold_btn._cache = None
        self.hold_btn._font.setPointSize(18)
        self.hold_btn.mousePressEvent = self._hold_press
        self.hold_btn.mouseReleaseEvent = self._hold_release
        btn_row.addWidget(self.hold_btn)

        # 解锁按钮（圆角矩形，减小阴影边距）
        unlock_btn = _NeuButton("解锁")
        unlock_btn.setFixedSize(240, 120)
        unlock_btn._radius = 24.0
        unlock_btn._blur = 10
        unlock_btn._distance = 4
        unlock_btn._cache = None
        unlock_btn._font.setPointSize(18)
        unlock_btn.clicked.connect(self._verify)
        btn_row.addWidget(unlock_btn)

        btn_row.addStretch()
        left_layout.addLayout(btn_row)

        left_layout.addStretch(1)
        main_layout.addWidget(left_widget, stretch=1)

        # ================= 右半部分：大号数字键盘 =================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        grid = QGridLayout()
        grid.setSpacing(15) 
        grid.setAlignment(Qt.AlignmentFlag.AlignCenter)

        btn_size = 150
        keys = [
            ('1', 0, 0), ('2', 0, 1), ('3', 0, 2),
            ('4', 1, 0), ('5', 1, 1), ('6', 1, 2),
            ('7', 2, 0), ('8', 2, 1), ('9', 2, 2),
            ('0', 3, 0)
        ]

        for text, r, c in keys:
            btn = _NeuRoundButton(text)
            btn.setFixedSize(btn_size, btn_size)
            btn._blur = 8          
            btn._distance = 4      
            btn._cache = None      
            btn._font.setPointSize(28)
            btn.clicked.connect(lambda t=text: self._append_digit(t))
            grid.addWidget(btn, r, c)

        back_btn = _NeuButton("Back")
        back_btn.setFixedSize(btn_size * 2 + 15, btn_size)
        back_btn._radius = 40.0          
        back_btn._blur = 8
        back_btn._distance = 4
        back_btn._cache = None
        back_btn._font.setPointSize(22)
        back_btn.clicked.connect(self._backspace)
        grid.addWidget(back_btn, 3, 1, 1, 2)

        right_layout.addLayout(grid)
        main_layout.addWidget(right_widget, stretch=1)

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())
        QTimer.singleShot(100, lambda: self.pwd_edit.edit.setFocus())
        # ===== 智能解锁：实时检测 =====
        self._smart_timer = None
        self._smart_timer = QTimer(self)
        self._smart_timer.setInterval(1000)   # 1秒检测一次
        self._smart_timer.timeout.connect(self._check_smart_unlock_live)
        self._smart_timer.start()

    def _on_text_changed(self):
        # 防窥模式开启时，不自动验证
        if self.anti_peek:
            return
            
        cfg = load_config()
        pwd_hash = cfg.get("password_hash", "")
        input_pwd = self.pwd_edit.edit.text()
        if input_pwd and _hash_password(input_pwd) == pwd_hash:
            self.accept()

    def _append_digit(self, digit):
        self.pwd_edit.edit.insert(digit)
        self.pwd_edit.edit.setFocus()

    def _backspace(self):
        self.pwd_edit.edit.backspace()
        self.pwd_edit.edit.setFocus()

    def _hold_press(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.pwd_edit.edit.setEchoMode(QLineEdit.EchoMode.Normal)
        _NeuButton.mousePressEvent(self.hold_btn, e)

    def _hold_release(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.pwd_edit.edit.setEchoMode(QLineEdit.EchoMode.Password)
        _NeuButton.mouseReleaseEvent(self.hold_btn, e)

    def _on_force_release(self):
        self.pwd_edit.edit.setEchoMode(QLineEdit.EchoMode.Password)

    def _verify(self):
        cfg = load_config()
        pwd_hash = cfg.get("password_hash", "")
        input_pwd = self.pwd_edit.edit.text()
        
        if self.anti_peek:
            # 防窥模式：检查是否包含连续的正确密码
            found = False
            # 因为密码至少4位，所以我们从长度4开始截取
            for i in range(len(input_pwd)):
                for j in range(i + 4, len(input_pwd) + 1):
                    if _hash_password(input_pwd[i:j]) == pwd_hash:
                        found = True
                        break
                if found: 
                    break
            
            if found:
                self.accept()
            else:
                self.error_label.setText("密码错误（或未包含正确密码），请重试")
                self.pwd_edit.edit.clear()
                self.pwd_edit.edit.setFocus()
        else:
            if _hash_password(input_pwd) == pwd_hash:
                self.accept()
            else:
                self.error_label.setText("密码错误，请重试")
                self.pwd_edit.edit.clear()
                self.pwd_edit.edit.setFocus()

    def _check_smart_unlock_live(self):
        """密码框弹出期间实时检测：
        1. 开发者 U 盘（无需开关、无需绑定）→ 显示“最高管理员权限”
        2. 已绑定的普通智能解锁磁盘（需开关开启）
        """
        try:
            dev = find_developer_usb()
            is_developer = dev is not None

            if dev is None:
                dev = find_bound_smart_unlock_device()
                is_developer = False

            if dev is None:
                return

            # 立即停止定时器，避免重复触发
            if self._smart_timer is not None:
                self._smart_timer.stop()
                self._smart_timer = None

            show_smart_unlock_toast(dev, is_developer)

            # 自动通过验证
            self.accept()
        except Exception as e:
            print(f"⚠️ 智能解锁实时检测失败：{e}")

    def accept(self):
        if self._smart_timer is not None:
            self._smart_timer.stop()
            self._smart_timer = None
        super().accept()

    def reject(self):
        if self._smart_timer is not None:
            self._smart_timer.stop()
            self._smart_timer = None
        super().reject()

    def closeEvent(self, event):
        # 用户点关闭按钮 / Alt+F4 时也停掉
        if self._smart_timer is not None:
            self._smart_timer.stop()
            self._smart_timer = None
        super().closeEvent(event)

# ======================== 密码设置对话框（新拟物风格） ========================

class PasswordSettingsDialog(QDialog):
    """密码设置：开启/关闭会联动密码与防窥模式（新拟物风格）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔒 密码设置")
        self.setFixedSize(650, 700)                       # ← 与更改密码界面一致
        self.setObjectName("passwordDialog")
        self.setStyleSheet("#passwordDialog{background-color:#e0e5ec;}")

        cfg = load_config()
        self._pwd_enabled = bool(cfg.get("password_enabled", False))
        self._anti_peek_enabled = bool(cfg.get("anti_peek_enabled", False))
        self._pwd_hash = cfg.get("password_hash", "")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 28)
        layout.setSpacing(10)                             # 比原来紧凑，给说明腾空间

        # 标题
        title = QLabel("密码设置")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(
            "color:#5a6b8c;font-size:24pt;font-weight:700;background:transparent;"
        )
        layout.addWidget(title)

        # ===== 第一行：启用密码保护 =====
        row1 = QHBoxLayout()
        lbl1 = QLabel("启用密码保护")
        lbl1.setStyleSheet(
            "color:#5a6b8c;font-size:15pt;font-weight:600;background:transparent;"
        )
        self.pwd_switch = _NeuSwitch(self._pwd_enabled)
        self.pwd_switch.toggled.connect(self._on_pwd_switch)
        row1.addWidget(lbl1, stretch=1)
        row1.addWidget(self.pwd_switch)
        layout.addLayout(row1)

        hint1 = _NeuHintCard(
            "开启后，每次启动程序都需要输入密码才能进入。\n"
            "关闭会清除已保存的密码，下次开启需重新设置。"
        )
        layout.addWidget(hint1)

        # ===== 第二行：启用防窥模式 =====
        row2 = QHBoxLayout()
        lbl2 = QLabel("启用防窥模式")
        lbl2.setStyleSheet(
            "color:#5a6b8c;font-size:15pt;font-weight:600;background:transparent;"
        )
        self.anti_peek_switch = _NeuSwitch(self._anti_peek_enabled)
        self.anti_peek_switch.toggled.connect(self._on_anti_peek_switch)
        row2.addWidget(lbl2, stretch=1)
        row2.addWidget(self.anti_peek_switch)
        layout.addLayout(row2)

        hint2 = _NeuHintCard(
            "开启后，密码输入中可混入无关字符，\n"
            "只要包含连续正确密码即可解锁（需手动提交）。\n"
            "关闭密码保护时会自动关闭此模式。"
        )
        layout.addWidget(hint2)

        layout.addStretch(1)

        # ===== 第三行：三个等宽按钮，关闭贴最右 =====
        row3 = QHBoxLayout()
        row3.setSpacing(20)
        row3.addStretch()

        BTN_W = 180          # ← 170 → 180
        BTN_H = 110

        self.change_pwd_btn = _NeuButton("更改密码")
        self.change_pwd_btn.setFixedSize(BTN_W, BTN_H)
        self.change_pwd_btn._radius = 20.0
        self.change_pwd_btn._blur = 10
        self.change_pwd_btn._distance = 4
        self.change_pwd_btn._cache = None
        self.change_pwd_btn._font.setPointSize(16)
        self.change_pwd_btn.clicked.connect(self._on_change_pwd)
        row3.addWidget(self.change_pwd_btn)

        self.smart_unlock_btn = _NeuButton("智能解锁")
        self.smart_unlock_btn.setFixedSize(BTN_W, BTN_H)
        self.smart_unlock_btn._radius = 20.0
        self.smart_unlock_btn._blur = 10
        self.smart_unlock_btn._distance = 4
        self.smart_unlock_btn._cache = None
        self.smart_unlock_btn._font.setPointSize(16)
        self.smart_unlock_btn.clicked.connect(self._on_smart_unlock)
        row3.addWidget(self.smart_unlock_btn)

        # 中间 stretch 把关闭推到最右
        row3.addStretch()

        close_btn = _NeuButton("关闭")
        close_btn.setFixedSize(BTN_W, BTN_H)
        close_btn._radius = 20.0
        close_btn._blur = 10
        close_btn._distance = 4
        close_btn._cache = None
        close_btn._font.setPointSize(16)
        close_btn.clicked.connect(self.accept)
        row3.addWidget(close_btn)

        layout.addLayout(row3)
        layout.addStretch(1)

        

    # ---------- 开关联动（保留上一轮的行为） ----------
    def _on_pwd_switch(self, checked):
        if checked:
            # 开启：没密码 → 直接跳设置密码界面
            if not self._pwd_hash:
                dlg = _ChangePasswordDialog(self)
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    self.pwd_switch.setChecked(False)
                    return
                cfg = load_config()
                self._pwd_hash = cfg.get("password_hash", "")
                if not self._pwd_hash:
                    self.pwd_switch.setChecked(False)
                    return
            self._pwd_enabled = True
            cfg = load_config()
            cfg["password_enabled"] = True
            save_config(cfg)
            show_toast(self, "密码保护已开启", "success")
        else:
            # 关闭：清空密码 + 自动关防窥
            self._pwd_enabled = False
            self._pwd_hash = ""
            cfg = load_config()
            cfg["password_enabled"] = False
            cfg["password_hash"] = ""
            cfg["anti_peek_enabled"] = False
            save_config(cfg)
            if hasattr(self, "anti_peek_switch"):
                self.anti_peek_switch.setChecked(False)
            self._anti_peek_enabled = False
            show_toast(self, "密码保护已关闭，保存的密码已清除", "success")

    def _on_anti_peek_switch(self, checked):
        if checked and not self._pwd_hash:
            show_toast(self, "请先设置密码", "warning")
            self.anti_peek_switch.setChecked(False)
            return
        self._anti_peek_enabled = checked
        cfg = load_config()
        cfg["anti_peek_enabled"] = checked
        save_config(cfg)
        show_toast(self, "防窥模式已" + ("开启" if checked else "关闭"), "success")

    def _on_change_pwd(self):
        dlg = _ChangePasswordDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            cfg = load_config()
            self._pwd_hash = cfg.get("password_hash", "")
            if not self._pwd_enabled:
                self._pwd_enabled = True
                self.pwd_switch.setChecked(True)
                cfg["password_enabled"] = True
                save_config(cfg)

    def _on_smart_unlock(self):
        # 智能解锁依赖密码保护，若尚未设置密码则提示
        if not self._pwd_hash:
            show_toast(self, "请先启用密码保护并设置密码", "warning")
            return
        dlg = _SmartUnlockDialog(self)
        dlg.exec()

class UpdateDialog(QDialog):
    def __init__(self, parent, version, changelog, url):
        super().__init__(parent)
        self.setWindowTitle("发现新版本")
        self.resize(500, 450)
        self.setStyleSheet(DIALOG_STYLE)
        self.url = url
        self.bat_path = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        title = QLabel(f"✨ 发现新版本：{version}")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #64B5F6;")
        layout.addWidget(title)

        layout.addWidget(QLabel("更新日志："))
        log_browser = QTextBrowser()
        log_browser.setMarkdown(changelog)
        log_browser.setStyleSheet("background: #1A1A24; border: 1px solid #3A3A4A; border-radius: 8px; padding: 10px;")
        layout.addWidget(log_browser)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #A0A0B0;")
        layout.addWidget(self.status_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.update_btn = QPushButton("🚀 立即更新")
        self.update_btn.clicked.connect(self.start_download)
        self.cancel_btn = QPushButton("暂不更新")
        self.cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.update_btn)
        btn_layout.addWidget(self.cancel_btn)
        layout.addLayout(btn_layout)

    def start_download(self):
        self.update_btn.setEnabled(False)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setText("正在下载更新包...")

        def download_thread():
            def update_progress(p):
                self.progress_bar.setValue(p)

            self.bat_path = UpdateManager.download_and_apply(self.url, update_progress)
            
            if self.bat_path:
                self.status_label.setText("下载完成，准备安装...")
                QTimer.singleShot(500, self.apply_update)
            else:
                QMessageBox.critical(self, "错误", "下载或解压失败，请检查网络。")
                self.reject()

        threading.Thread(target=download_thread, daemon=True).start()

    def apply_update(self):
        if self.bat_path:
            UpdateManager.restart_and_apply(self.bat_path)

class SettingsInterface(QWidget):
    # 跨线程回传：has_update, latest_ver, changelog, url, error_msg
    _update_result_signal = pyqtSignal(bool, str, str, str, str)

    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("settingsInterface")
        self._update_result_signal.connect(self._handle_update_result)
        self.build_ui()

    def build_ui(self):
        # ========== 【新增】外层布局：只放滚动区域 ==========
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        # ========== 【新增】滚动区域 ==========
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background: transparent;
            }
            QScrollBar:vertical {
                width: 8px;
                background: transparent;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #4A4A5A;
                border-radius: 4px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background: #6A6A7A;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
        """)

        # ========== 【新增】内容容器（原来的 self 布局全部放在它里面） ==========
        content_widget = QWidget()
        content_widget.setObjectName("settingsContentWidget")   # ← 新增
        content_widget.setStyleSheet(
            "QWidget#settingsContentWidget { background: transparent; }"
        )                                                        # ← 修改

        # 主布局：垂直（原来的变量名不变，只是 parent 换成 content_widget）
        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # 标题
        title = TitleLabel("⚙️ 系统设置", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        main_layout.addWidget(title)

        # ----- 等级信息（重新设计） -----
        level_frame = QFrame(self)
        level_frame.setStyleSheet("""
            QFrame {
                background: rgba(100, 181, 246, 0.1);
                border-radius: 12px;
                border: 1px solid #3A3A4A;
                padding: 12px;
            }
        """)
        level_layout = QHBoxLayout(level_frame)
        level_layout.setSpacing(15)

        # 左侧等级数字
        self.lv_label = QLabel("Lv.6")
        self.lv_label.setStyleSheet("font-size: 36px; font-weight: bold; color: #64B5F6;")
        self.lv_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lv_label.setFixedWidth(200)   # 固定宽度保证对齐
        level_layout.addWidget(self.lv_label)

        # 中间垂直区域
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        middle_layout.setSpacing(4)

        self.title_label = QLabel("历练·逐光实践")
        self.title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        middle_layout.addWidget(self.title_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0 / 0 (0%)")
        self.progress_bar.setMinimumWidth(200)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(255,255,255,0.1);
                border-radius: 6px;
                height: 20px;
                text-align: center;
                color: #FFFFFF;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                border-radius: 6px;
            }
        """)
        middle_layout.addWidget(self.progress_bar)

        level_layout.addWidget(middle_widget, stretch=1)   # 占用剩余空间

        # ---- 右侧按钮区域（两个按钮水平排列） ----
        # 等级详情按钮
        level_detail_btn = QPushButton("📊 等级详情")
        level_detail_btn.clicked.connect(self.open_level_detail)
        level_layout.addWidget(level_detail_btn)

        # ★ 新增：成就按钮（位于等级详情按钮右侧）
        achievement_btn = QPushButton("🏆 成就")
        achievement_btn.clicked.connect(self.open_achievement)
        level_layout.addWidget(achievement_btn)

        main_layout.addWidget(level_frame)

        # ----- 两列网格布局 -----
        grid = QGridLayout()
        grid.setSpacing(18)
        grid.setContentsMargins(0, 0, 0, 0)

        # ---- 左列：外观 + 事由 ----
        left_col = QVBoxLayout()
        left_col.setSpacing(15)

        # 1. 外观组（无折叠）
        appearance_group = QGroupBox("🎨 外观")
        appearance_layout = QVBoxLayout(appearance_group)

        # 动画设置入口（合并成大按钮，具体开关在弹窗里）
        anim_row = QHBoxLayout()
        anim_row.addWidget(QLabel("动画设置："))
        anim_btn = QPushButton("🎬 动画设置")
        anim_btn.clicked.connect(self.open_animation_settings)
        anim_row.addWidget(anim_btn)
        anim_row.addWidget(QLabel("（开机动画、进度条动画）"))
        anim_row.addStretch()
        appearance_layout.addLayout(anim_row)

        # 粒子效果入口（保持不变）
        particle_row = QHBoxLayout()
        particle_row.addWidget(QLabel("粒子效果："))
        particle_btn = QPushButton("✨ 粒子效果设置")
        particle_btn.clicked.connect(self.open_particle_settings)
        particle_row.addWidget(particle_btn)
        particle_row.addWidget(QLabel("（窗口、鼠标的粒子效果）"))
        particle_row.addStretch()
        appearance_layout.addLayout(particle_row)

        # ---- 密码保护入口（新拟物风格弹窗） ----
        pwd_row = QHBoxLayout()
        pwd_row.addWidget(QLabel("密码保护："))
        pwd_btn = QPushButton("🔒 密码设置")
        pwd_btn.clicked.connect(self.open_password_settings)
        pwd_row.addWidget(pwd_btn)
        pwd_row.addWidget(QLabel("（启动密码验证）"))
        pwd_row.addStretch()
        appearance_layout.addLayout(pwd_row)

        # ===== 主题色 =====
        theme_row = QHBoxLayout()
        theme_row.addWidget(QLabel("主题色："))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["深色", "浅色（默认）", "自定义"])
        _cfg_t = load_config()
        _mode_map = {"dark": "深色", "light": "浅色（默认）", "custom": "自定义"}
        _cur_mode = _cfg_t.get("theme_mode", "light")
        self.theme_combo.setCurrentText(_mode_map.get(_cur_mode, "浅色（默认）"))
        self.theme_combo.currentTextChanged.connect(self.on_theme_changed)
        theme_row.addWidget(self.theme_combo)
        theme_row.addStretch()
        appearance_layout.addLayout(theme_row)

        # 自定义颜色（仅自定义时显示）
        self.theme_custom_widget = QWidget()
        custom_row = QHBoxLayout(self.theme_custom_widget)
        custom_row.setContentsMargins(0, 0, 0, 0)
        custom_row.addWidget(QLabel("窗口颜色："))
        self.theme_bg_btn = QPushButton()
        self.theme_bg_btn.setFixedSize(90, 30)
        self.theme_bg_btn.clicked.connect(self.pick_theme_bg)
        custom_row.addWidget(self.theme_bg_btn)
        custom_row.addWidget(QLabel("文字颜色："))
        self.theme_fg_btn = QPushButton()
        self.theme_fg_btn.setFixedSize(90, 30)
        self.theme_fg_btn.clicked.connect(self.pick_theme_fg)
        custom_row.addWidget(self.theme_fg_btn)
        custom_row.addStretch()
        appearance_layout.addWidget(self.theme_custom_widget)

        self._refresh_theme_color_preview()
        self.theme_custom_widget.setVisible(_cur_mode == "custom")

        left_col.addWidget(appearance_group)

        # 2. 事由组
        reason_group = QGroupBox("📝 事由")
        reason_layout = QVBoxLayout(reason_group)
        # 事由管理
        reason_row1 = QHBoxLayout()
        reason_row1.addWidget(QLabel("事由管理："))
        reason_btn = QPushButton("📝 管理事由")
        reason_btn.clicked.connect(self.open_reason_manager)
        reason_row1.addWidget(reason_btn)
        reason_row1.addStretch()
        reason_layout.addLayout(reason_row1)
        # 数字快捷选择
        shortcut_row = QHBoxLayout()
        shortcut_row.addWidget(QLabel("事由数字快捷选择："))
        self.shortcut_switch = SwitchButton(self)
        self.shortcut_switch.setChecked(load_config().get("enable_reason_number_shortcut", True))
        self.shortcut_switch.checkedChanged.connect(self.toggle_shortcut)
        shortcut_row.addWidget(self.shortcut_switch)
        shortcut_row.addWidget(QLabel("（开启后，焦点到事由框时按数字键选择）"))
        shortcut_row.addStretch()
        reason_layout.addLayout(shortcut_row)
        left_col.addWidget(reason_group)
                # 3. 最近记录组（新增）
        record_group = QGroupBox("📋 最近记录")
        record_layout = QVBoxLayout(record_group)

        limit_row = QHBoxLayout()
        limit_row.addWidget(QLabel("显示最近条数："))
        self.record_limit_spin = QSpinBox()
        self.record_limit_spin.setRange(10, 1000)
        self.record_limit_spin.setSingleStep(10)
        self.record_limit_spin.setValue(load_config().get("recent_records_limit", 100))
        self.record_limit_spin.valueChanged.connect(self.on_record_limit_changed)
        limit_row.addWidget(self.record_limit_spin)
        limit_row.addWidget(QLabel("条"))
        limit_row.addStretch()
        record_layout.addLayout(limit_row)

        tip_label = QLabel("（在录入界面的「📋 录入记录」窗口中生效）")
        tip_label.setStyleSheet("color: #A0A0A0; font-size: 12px;")
        record_layout.addWidget(tip_label)

        left_col.addWidget(record_group)

        # 将左列加入网格（第0行，第0列）
        grid.addLayout(left_col, 0, 0)

        # ---- 右列：总览 + 数据管理 ----
        right_col = QVBoxLayout()
        right_col.setSpacing(15)

        # 3. 总览组
        overview_group = QGroupBox("📊 总览")
        overview_layout = QVBoxLayout(overview_group)
        # 显示人数
        n_row = QHBoxLayout()
        n_row.addWidget(QLabel("进步/退步显示人数："))
        self.overview_spin = QSpinBox()
        self.overview_spin.setRange(1, 30)
        self.overview_spin.setValue(load_config().get("overview_top_n", 5))
        self.overview_spin.valueChanged.connect(self.on_overview_n_changed)
        n_row.addWidget(self.overview_spin)
        n_row.addStretch()
        overview_layout.addLayout(n_row)
        # 排名依据
        rank_row = QHBoxLayout()
        rank_row.addWidget(QLabel("进退步依据："))
        self.rank_combo = QComboBox()
        self.rank_combo.addItems(["赋分总分", "赋分校名", "赋分班名"])
        self.rank_combo.setCurrentText(load_config().get("overview_rank_by", "赋分总分"))
        self.rank_combo.currentTextChanged.connect(self.on_rank_by_changed)
        rank_row.addWidget(self.rank_combo)
        rank_row.addStretch()
        overview_layout.addLayout(rank_row)
        right_col.addWidget(overview_group)

        # 4. 数据管理组
        data_group = QGroupBox("🗄️ 数据管理")
        data_layout = QVBoxLayout(data_group)
        # 切换名单
        change_row = QHBoxLayout()
        change_row.addWidget(QLabel("名单管理："))
        change_btn = QPushButton("📂 切换名单")
        change_btn.clicked.connect(self.app.show_change_file)
        change_row.addWidget(change_btn)
        change_row.addStretch()
        data_layout.addLayout(change_row)
        # 重置积分
        reset_row = QHBoxLayout()
        reset_row.addWidget(QLabel("积分管理："))
        reset_btn = QPushButton("🔄 重置所有积分")
        reset_btn.clicked.connect(self.app.reset_scores)
        reset_row.addWidget(reset_btn)
        reset_row.addStretch()
        data_layout.addLayout(reset_row)
        # 格式化数据
        format_row = QHBoxLayout()
        format_row.addWidget(QLabel("数据管理："))
        format_btn = QPushButton("🗑️ 格式化所有数据")
        format_btn.clicked.connect(self.format_data)
        format_row.addWidget(format_btn)
        format_row.addStretch()
        data_layout.addLayout(format_row)
        # 存储空间
        storage_row = QHBoxLayout()
        storage_row.addWidget(QLabel("存储管理："))
        storage_btn = QPushButton("💾 存储空间")
        storage_btn.clicked.connect(self.open_storage_analyzer)
        storage_row.addWidget(storage_btn)
        storage_row.addStretch()
        data_layout.addLayout(storage_row)
        right_col.addWidget(data_group)
        # ---- 5. 帮助组（新增） ----
        help_group = QGroupBox("📖 帮助")
        help_layout = QVBoxLayout(help_group)

        # ---- 检查更新 ----
        update_row = QHBoxLayout()
        update_row.addWidget(QLabel("版本更新："))
        check_update_btn = QPushButton("🔄 检查更新")
        check_update_btn.clicked.connect(self.check_for_updates)
        update_row.addWidget(check_update_btn)
        update_row.addStretch()
        help_layout.addLayout(update_row)

        # ---- 使用指南 ----
        guide_row = QHBoxLayout()
        guide_row.addWidget(QLabel("使用指南："))
        guide_btn = QPushButton("📖 查看使用指南")
        guide_btn.clicked.connect(self.open_user_guide)
        guide_row.addWidget(guide_btn)
        guide_row.addStretch()
        help_layout.addLayout(guide_row)

        # ---- 新手引导（新增）----
        onboard_row = QHBoxLayout()
        onboard_row.addWidget(QLabel("新手引导："))
        onboard_btn = QPushButton("🎯 重新查看新手引导")
        onboard_btn.clicked.connect(self.open_onboarding)
        onboard_row.addWidget(onboard_btn)
        onboard_row.addStretch()
        help_layout.addLayout(onboard_row)

        tip_label2 = QLabel("（使用指南包含所有功能模块的详细说明、快捷键、常见问题等）")
        tip_label2.setStyleSheet("color: #A0A0A0; font-size: 12px;")
        help_layout.addWidget(tip_label2)

        right_col.addWidget(help_group)

        # 将右列加入网格（第0行，第1列）
        grid.addLayout(right_col, 0, 1)

        # 将网格加入主布局
        main_layout.addLayout(grid)

        # 返回按钮（居中对齐）
        back_btn = QPushButton("🔙 返回主菜单")
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        main_layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        main_layout.addStretch()

        # ========== 【新增】把内容装进滚动区域 ==========
        scroll.setWidget(content_widget)
        outer_layout.addWidget(scroll)

    # ---- 新增：成就按钮回调 ----
    def open_achievement(self):
        dialog = AchievementDialog(self)
        dialog.exec()

    def check_for_updates(self):
        show_toast(self, "正在检查更新，请稍候...", "info")

        def check_thread():
            try:
                has_update, latest_ver, changelog, url = UpdateManager.check_update()
                self._update_result_signal.emit(
                    bool(has_update), latest_ver or "", changelog or "", url or "", ""
                )
            except Exception as e:
                self._update_result_signal.emit(False, "", "", "", str(e))

        threading.Thread(target=check_thread, daemon=True).start()

    def _handle_update_result(self, has_update, latest_ver, changelog, url, error_msg):
        """在主线程里收到信号后处理 UI"""
        if error_msg:
            show_toast(self, error_msg, "error", duration=5000)
            return
        if has_update:
            self._update_window = _NeuUpdateWindow(latest_ver, changelog, url)
            self._update_window.show()
        else:
            if latest_ver:
                show_toast(self, f"当前已是最新版本 ({CURRENT_VERSION})", "success")
            else:
                show_toast(self, "获取版本信息失败，请查看控制台日志", "error")

    def open_user_guide(self):
        """打开使用指南对话框"""
        dialog = UserGuideDialog(self)
        dialog.exec()

    def open_onboarding(self):
        """重新打开新手引导（引导完成后不隐藏主窗口）"""
        main_win = self.app  # SettingsInterface.app 就是 MainWindow

        # 如果已经打开着引导窗口，先关掉旧的，避免重复叠加
        old = getattr(self, "_onboarding", None)
        if old is not None:
            try:
                old.close()
            except Exception:
                pass
            self._onboarding = None

        def _on_finish():
            # 引导关闭（无论完成还是 Alt+F4）都走这里
            self._onboarding = None
            # 引导过程中可能改动了名单，刷新一下
            try:
                if hasattr(main_win, "inputInterface"):
                    main_win.inputInterface.update_student_list()
            except Exception:
                pass
            try:
                if hasattr(main_win, "refresh_ranking"):
                    main_win.refresh_ranking()
            except Exception:
                pass
            # 把焦点还给主窗口
            try:
                main_win.raise_()
                main_win.activateWindow()
            except Exception:
                pass

        try:
            self._onboarding = OnboardingWindow(main_win, on_finish=_on_finish)
            self._onboarding.show()
            self._onboarding.raise_()
            self._onboarding.activateWindow()
        except Exception as e:
            import traceback
            print("❌ 打开新手引导失败：")
            traceback.print_exc()
            QMessageBox.warning(self, "错误", f"打开新手引导失败：{e}")

    def open_storage_analyzer(self):
        """打开存储空间分析对话框"""
        dialog = StorageAnalyzerDialog(self)
        dialog.exec()

    def on_theme_changed(self, text):
        mode_map = {"深色": "dark", "浅色（默认）": "light", "自定义": "custom"}
        mode = mode_map.get(text, "light")
        cfg = load_config()
        cfg["theme_mode"] = mode
        save_config(cfg)
        self.theme_custom_widget.setVisible(mode == "custom")
        if mode == "custom":
            self._refresh_theme_color_preview()
        # ★ 无论是深色、浅色还是自定义，切换后都询问是否重启
        self._ask_restart_for_theme()
            

    def _refresh_theme_color_preview(self):
        cfg = load_config()
        bg = cfg.get("theme_bg", "#16161E")
        fg = cfg.get("theme_fg", "#E0E0E0")
        self.theme_bg_btn.setText(bg)
        self.theme_bg_btn.setStyleSheet(
            f"background: {bg}; color: {fg}; "
            f"border: 1px solid #888; border-radius: 6px; font-size: 12px;"
        )
        self.theme_fg_btn.setText(fg)
        self.theme_fg_btn.setStyleSheet(
            f"background: {fg}; color: {bg}; "
            f"border: 1px solid #888; border-radius: 6px; font-size: 12px;"
        )

    def pick_theme_bg(self):
        from PyQt6.QtWidgets import QColorDialog
        cfg = load_config()
        current = QColor(cfg.get("theme_bg", "#16161E"))
        color = QColorDialog.getColor(current, self, "选择窗口颜色")
        if color.isValid():
            cfg["theme_bg"] = color.name()
            save_config(cfg)
            self._refresh_theme_color_preview()
            self._ask_restart_for_theme()

    def pick_theme_fg(self):
        from PyQt6.QtWidgets import QColorDialog
        cfg = load_config()
        current = QColor(cfg.get("theme_fg", "#E0E0E0"))
        color = QColorDialog.getColor(current, self, "选择文字颜色")
        if color.isValid():
            cfg["theme_fg"] = color.name()
            save_config(cfg)
            self._refresh_theme_color_preview()
            self._ask_restart_for_theme()


    

    def open_reason_manager(self):
        dialog = ReasonManagerDialog(self, self.app)
        dialog.exec()

    def toggle_shortcut(self):
        checked = self.shortcut_switch.isChecked()
        cfg = load_config()
        cfg["enable_reason_number_shortcut"] = checked
        save_config(cfg)
        show_toast(self, "设置已保存", "info")
        # 通知录入页面更新配置
        self.app.refresh_input_shortcut()

    def on_record_limit_changed(self, value):
        cfg = load_config()
        cfg["recent_records_limit"] = value
        save_config(cfg)
        show_toast(self, f"最近记录显示条数已设为 {value}", "info")

    def on_overview_n_changed(self, value):
        cfg = load_config()
        cfg["overview_top_n"] = value
        save_config(cfg)
        show_toast(self, f"总览显示人数已设为 {value}", "info")
        if self.app.scoreInterface and getattr(self.app.scoreInterface, 'current_exam', None) == "总览":
            self.app.scoreInterface.display_overview()

    def on_rank_by_changed(self, value):
        cfg = load_config()
        cfg["overview_rank_by"] = value
        save_config(cfg)
        show_toast(self, f"排行榜依据已切换至：{value}", "info")
        # 若当前在总览视图，立即刷新
        if self.app.scoreInterface and getattr(self.app.scoreInterface, 'current_exam', None) == "总览":
            self.app.scoreInterface.display_overview()

    def format_data(self):
        """格式化所有数据（三次确认后删除 user 文件夹并重启）"""
        # ---- 第一次确认 ----
        msg_box1 = QMessageBox(self)
        msg_box1.setWindowTitle("确认格式化")
        msg_box1.setText("⚠️ 是否确定要格式化所有数据？\n此操作将 **永久删除** 所有用户数据（包括名单、积分、配置、便签等），并重启程序。")
        msg_box1.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box1.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box1.setStyleSheet(DIALOG_STYLE)   # ✅ 跟随主题
        if msg_box1.exec() != QMessageBox.StandardButton.Yes:
            return

        # ---- 第二次确认 ----
        msg_box2 = QMessageBox(self)
        msg_box2.setWindowTitle("二次确认")
        msg_box2.setText("格式化将删除全部数据，确认继续？")
        msg_box2.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box2.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box2.setStyleSheet(DIALOG_STYLE)   # ✅ 直接复用，不需要再引用 msg_box1
        if msg_box2.exec() != QMessageBox.StandardButton.Yes:
            return

        # ---- 第三次：输入验证 ----
        text, ok = QInputDialog.getText(
            self,
            "输入验证",
            "请输入“确认格式化”以继续：",
            QLineEdit.EchoMode.Normal
        )
        if not ok or text.strip() != "确认格式化":
            show_toast(self, "输入错误，操作取消", "warning")
            return

        # ---- 执行删除 ----
        try:
            if os.path.exists(USER_DIR):
                shutil.rmtree(USER_DIR)
                show_toast(self, "✅ 数据已格式化，程序即将重启", "success")
            else:
                show_toast(self, "user 文件夹不存在，无需格式化", "info")
                return
        except Exception as e:
            show_toast(self, f"格式化失败：{e}", "error")
            return

        # ========== 关键修复：立即重建目录和默认配置文件 ==========
        os.makedirs(USER_DIR, exist_ok=True)
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(EXPORT_DIR, exist_ok=True)
        os.makedirs(BACKUP_DIR, exist_ok=True)

        # 生成默认配置（与 load_config() 的默认值保持一致）
        default_cfg = {
            "student_file": DEFAULT_STUDENT_EXCEL,
            "template_file": "",
            "last_export_month": "",
            "default_discount": 100,
            "default_category": "日常",
            "show_boot_animation": True,
            "install_time": time.time(),
            "backup_dir": "backups",
            "max_backups": 10,
            "last_backup_date": "",
            "custom_reasons": [
                {"name": "迟到", "type": "-", "score": 2},
                {"name": "早退", "type": "-", "score": 2},
                {"name": "卫生", "type": "-", "score": 1},
                {"name": "纪律", "type": "-", "score": 1},
                {"name": "作业", "type": "-", "score": 2},
                {"name": "课堂表现", "type": "+", "score": 2},
            ],
            "last_input_category": "日常",
            "last_input_reason": "",
            "last_input_discount": "100",
            "fix_params_serial": False,
            "enable_reason_number_shortcut": True,
            "countdown_events": [],
            "selected_countdown_index": -1,
            "overview_top_n": 5,
            "overview_rank_by": "赋分班名",
            "total_records_count": 0,
            "total_experience": 0,
            "today_experience": 0,
            "last_experience_date": "",
            "sign_in_streak": 0,
            "last_sign_in_date": "",
            "online_seconds_today": 0,
            "last_online_update": 0.0,
            "last_device_day": "",
            "achievements": {},
            "device_total_days_recorded": 0,
            "import_count": 0,
            "export_count": 0,
            "backup_count": 0,
            "recent_records_limit": 100,   # 最近录入记录显示条数
            "enable_progress_animation": True,
            "enable_particle_effect": True,      # ← 新增：弹窗关闭粒子效果
            "enable_click_particle_effect": True,   # ← 新增：鼠标点击粒子
            "enable_drag_particle_effect": True,   # ← 新增：拖动拖尾粒子
            "enable_scroll_particle_effect": True, # ← 新增：鼠标滚动粒子
            "enable_nav_animation": True,          # ← 新增：导航栏切换动画
            # ---- 主题 ----
            "theme_mode": "light",          # dark / light / custom
            "theme_bg": "#16161E",         # 自定义模式下的窗口颜色
            "theme_fg": "#E0E0E0",         # 自定义模式下的文字颜色
            # 按类型分别保留条数（每种类型各自独立计数）
            "export_keep_counts": {
                "班级积分月度汇总": 5,
                "成绩总览对比": 5,
                "其它": 5,
            },
            "backup_keep_counts": {
                "备份-积分记录": 5,
                "名单备份": 5,
                "积分记录": 5,
                "其它": 5,
            },
        # ---- 密码保护 ----
        "password_enabled": False,     # ← 新增
        "password_hash": "",           # ← 新增
        "anti_peek_enabled": False,    # ← 新增：防窥模式
        # ---- 智能解锁（U 盘） ----
        "smart_unlock_enabled": False,
        "smart_unlock_device": [],        
        "onboarding_completed": False,   # ← 新增：是否已完成新手引导
        }
        with open(CONFIG_JSON, "w", encoding="utf-8") as f:
            json.dump(default_cfg, f, ensure_ascii=False, indent=2)

        # ---- 显示重启提示（使用统一的 QMessageBox 样式） ----
        # ---- 最后的"操作完成"提示 ----
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("操作完成")
        msg_box.setText("✅ 数据已删除，程序将自动重启。")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Ok)
        msg_box.setStyleSheet(DIALOG_STYLE)   # ✅ 跟随主题（同时自动修掉原代码的 #42A5F6 笔误）
        msg_box.exec()

        # 延迟重启
        QTimer.singleShot(1000, self._restart_app)

    def _restart_app(self):
        """重启当前应用程序（修复打包后重启双开问题）"""
        import subprocess
        import sys
        if getattr(sys, "frozen", False):
            # 打包后的 exe 环境：sys.argv[0] 是 exe 本身，不能再次传入，否则会启动两个进程
            subprocess.Popen([sys.executable] + sys.argv[1:])
        else:
            # 源码运行环境：sys.argv[0] 是脚本名，必须传入
            subprocess.Popen([sys.executable] + sys.argv)
        QApplication.quit()

    def _ask_restart_for_theme(self):
        """主题修改后弹出确认框：左边重启、右边取消"""
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("主题已保存")
        dialog.setFixedSize(460, 180)
        dialog.setStyleSheet(f"""
            QDialog {{
                background: {_c['bg']};
                border-radius: 18px;
            }}
            QLabel {{
                color: {_c['fg']};
                font-size: 15px;
            }}
            QPushButton {{
                border-radius: 10px;
                padding: 8px 24px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(100, 181, 246, 0.3);
                min-width: 100px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #42A5F5, stop:1 #1E88E5);
                border-color: #64B5F6;
            }}
            QPushButton:pressed {{
                background: #1E88E5;
            }}
            QPushButton#cancelBtn {{
                background: #555555;
            }}
            QPushButton#cancelBtn:hover {{
                background: #6A6A7A;
            }}
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(30, 25, 30, 25)
        layout.setSpacing(18)

        label = QLabel("主题颜色已保存，是否立即重启程序以应用新主题？")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setWordWrap(True)
        layout.addWidget(label)

        # 按钮行：左「重启」、右「取消」
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        btn_layout.addStretch()

        restart_btn = QPushButton("🔄 重启")
        restart_btn.setFixedWidth(120)
        restart_btn.clicked.connect(dialog.accept)

        cancel_btn = QPushButton("取消")
        cancel_btn.setObjectName("cancelBtn")
        cancel_btn.setFixedWidth(120)
        cancel_btn.clicked.connect(dialog.reject)

        btn_layout.addWidget(restart_btn)
        btn_layout.addWidget(cancel_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        dialog.move(screen.center() - dialog.rect().center())

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._restart_app()

    

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh_level_display()

    def refresh_level_display(self):
        if not hasattr(self, 'lv_label') or not hasattr(self, 'title_label') or not hasattr(self, 'progress_bar'):
            return
        cfg = load_config()
        total_exp = cfg.get("total_experience", 0)
        total_records = cfg.get("total_records_count", 0)
        level = calculate_level()
        title = LEVEL_TITLES.get(level, "未知")
        level_range = LEVEL_RANGES.get(level, "")
        combined = f"{level_range}·{title}" if level_range else title

        # 更新左侧等级数字
        self.lv_label.setText(f"Lv.{level}")

        # 更新称号
        self.title_label.setText(combined)

        # 更新进度条（带缓动动画）
        exp_current, exp_needed, progress = get_level_progress(total_exp)
        if level >= MAX_LEVEL:
            animate_progress_bar(self.progress_bar, 100)
            self.progress_bar.setFormat("已满级")
        else:
            animate_progress_bar(self.progress_bar, int(progress))
            self.progress_bar.setFormat(f"{exp_current:.0f} / {exp_current + exp_needed:.0f}  ({int(progress)}%)")

    def open_level_detail(self):
        """打开等级详情对话框"""
        dialog = LevelDetailDialog(self)
        dialog.exec()

    def open_animation_settings(self):
        """打开动画设置对话框"""
        dialog = AnimationSettingsDialog(self)
        dialog.exec()

    def open_particle_settings(self):
        """打开粒子效果设置对话框"""
        dialog = ParticleSettingsDialog(self)
        dialog.exec()

    def open_password_settings(self):
        """打开密码设置对话框（新拟物风格）"""
        cfg = load_config()
        if cfg.get("password_hash", ""):
            # 如果已有密码，必须先通过身份验证
            dlg = PasswordPromptDialog(self)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
        dialog = PasswordSettingsDialog(self)
        dialog.exec()

    

# ======================== 使用指南对话框 ========================
USER_GUIDE_TEXT = """===============================
PickOS 班级管理系统使用指南
===============================
版本：v6.4.0.110
适用平台：Windows / macOS / Linux（正在努力适配HarmonyOS，敬请期待...）
联系创作者：2892065430@qq.com

【目录】
1. 系统概览与设置入口
2. 启动与首页
3. 录入积分
4. 积分排行榜
5. 成绩总览
6. 名单管理
7. 导出与导入
8. 设置详解
9. 等级与经验系统
10. 成就系统
11. 键盘快捷键
12. 数据存储与备份
13. 悬浮窗与托盘
14. 常见问题
15. 彩蛋

─────────────────────────────────
一、系统概览与设置入口
─────────────────────────────────
PickOS 是一个班级积分与成绩管理系统，主要模块如下：

  首页   : 时钟、签到、便签、日历、倒计时、设备监控、积分统计
  录入   : 单条/连续录入积分、事由管理、最近记录
  排行   : 积分排行榜、搜索高亮、明细查看
  成绩   : 导入成绩、考试对比、考试加分
  名单   : 自定义名单、班干加分
  导出   : 导出 Excel、导入赋分
  设置   : 所有系统配置、等级/成就、数据管理

设置入口：左侧导航栏底部 ⚙️ 设置。

─────────────────────────────────
二、启动与首页
─────────────────────────────────
2.1 启动过程
  · 首次运行自动创建 user/ 目录结构。
  · 默认播放约 3 秒开机动画（显示版本号、版权信息）。
  · 相关设置：设置 → 🎨 外观 → 开机动画

2.2 顶部区域
  · 实时时钟：显示星期、月份、日期、时分秒，秒跳动时有字符波浪动画。
  · 问候语：根据当前时段自动变化。
  · 快捷按钮：录入、排行、成绩、名单 —— 一键跳转对应页面。
  · 📅 签到按钮：
      - 每天可签到一次。
      - 连续签到天数越多，经验越高：第1天+2，第2天+6，第3天+10，
        第4天+10，第5天+15，第6天+15，第7天及以后固定+20。
      - 若今日经验已达上限（120），签到仍成功，但不再获得经验。
      - 断签后连续天数重置为1。

2.3 便签卡片
  · 左侧显示便签标题和内容编辑区，右侧为便签列表。
  · 添加：点击右上角方框，输入标题即可。
  · 重命名/删除：双击列表中的便签，弹出操作菜单。
  · 自动保存：编辑内容后自动保存到 user/data/notes.json。
  · 默认包含一条欢迎便签。

2.4 日历与倒计时卡片
  · 月份切换：
      - 单击月份左右小方框：上一月/下一月。
      - 单击年月文字：分别输入年份和月份。
      - Ctrl + 鼠标滚轮：在日历区域滚动切换月份。
      - 右上角小方框：回到今天并高亮。
  · 单击任意日期：日期选择，高亮显示。
  · 倒计时：
      - 双击倒计时区域打开倒计时管理器。
      - 可添加多个事件（名称 + 目标日期），支持编辑、删除、设为当前。
      - 默认事件为"国庆假期"。

2.5 设备监控卡片
  · 实时显示 CPU 使用率、内存占用、磁盘可用空间、电池状态、网络速率。

2.6 积分统计卡片
  · 今日已录入：当天写入的积分记录条数。
  · 累计录入总条数：从配置中读取，若配置为0且CSV有数据会自动迁移。
  · 积分冠军：累计总积分最高的学生。

2.7 系统状态卡片
  · 本次运行时间（自程序启动）。
  · 设备总运行时间（自安装时间起）。
  · 版本号、版权信息。

─────────────────────────────────
三、录入积分
─────────────────────────────────
3.1 界面元素
  连续模式     : 勾选后进入批量录入模式
  学号         : 可输入纯学号、学号+姓名、纯姓名（重名取第一个）
  姓名标签     : 自动显示匹配的学生姓名，重名会提示
  日期         : 默认今天，格式 YYYY-MM-DD
  类型         : 加分 / 扣分
  类别         : 日常 / 听写 / 背诵 / 班干加分 / 考试加分
  分数         : 支持小数
  赋分总分     : 可选，用于录入考试成绩
  折扣比例     : 默认100%，实际分数 = 分数 × 折扣%
  事由         : 下拉选择或手动输入，选择后自动填充类型/分数/类别
  保留学号     : 提交后不清空学号
  保留事由     : 提交后不清空事由
  确认提交     : 提交当前录入
  清空表单     : 重置表单

3.2 单条录入流程
  1. 输入学号（支持学号、姓名），系统自动匹配信息。
  2. 选择类型、类别，输入分数。
  3. （可选）输入赋分总分、折扣比例。
  4. 选择或输入事由。
  5. 点击【确认提交】或按 Ctrl+Enter。
  6. 成功后提示，若未勾选保留项则清空相应字段。

3.3 连续模式详解
  · 勾选【连续模式】后，界面显示起始学号、结束学号输入框。
  · 输入起止学号，点击提交后开始连续录入。
  · 状态显示：显示当前待录学号。
  · 按钮：
      - 暂停/继续：暂停自动跳转，手动继续。
      - 跳过：跳过当前学号，不写入记录，并记录到跳过名单。
      - 停止：结束连续录入，弹出确认。
  · 固定参数：勾选后，学号自动递增，其他参数保持不变。

3.4 事由管理
  · 添加/编辑/删除事由：设置 → 📝 事由 → 📝 管理事由。
  · 每个事由包含：名称、类型、默认分数、匹配类别。
  · 数字快捷选择：设置 → 📝 事由 → 事由数字快捷选择。

3.5 最近记录
  · 点击右上角【📋 录入记录】打开最近记录窗口。
  · 显示条数：设置 → 📋 最近记录 → 显示最近条数，默认100。
  · 操作：点击第一列复选框勾选可批量删除；双击任意行可修改/删除。

─────────────────────────────────
四、积分排行榜
─────────────────────────────────
4.1 显示内容
  · 列：排名、学号、姓名、日常、听写、背诵、班干加分、考试加分、
        赋分总分、累计总积分。
  · 按累计总积分降序，同分按学号升序。
  · 底部统计：最高分、最低分、平均分。

4.2 搜索与高亮
  · 输入学号或姓名，点击 🔍 搜索 或按 Enter。
  · 匹配行以蓝色高亮，并自动选中滚动到可见区域。
  · 再次按 Enter 或点【下一个】循环跳转匹配项。

4.3 双击操作
  · 双击某一行：打开该学生的积分明细对话框。
  · 双击表头某类别：批量清除该类别全部记录（不可撤销）。

─────────────────────────────────
五、成绩总览
─────────────────────────────────
5.1 导入成绩
  · 点击【📂 导入成绩】，选择 Excel。
  · 必需列：学号、姓名。
  · 推荐列：语文、数学、英语、物理、化学赋分、生物赋分、总分、
            原分校名、赋分校名、赋分班名等。
  · 导入时，若已有"本次"数据，会自动转为"上次"。

5.2 考试成绩
  · 按总分由高到低排列
  · 底部显示最高分、最低分、平均分、单科王
  · 若有相同分数则并列显示

5.3 考试切换
  · 本次考试：当前导入的数据。
  · 上次考试：上一次导入的数据。
  · 总览：本次与上次对比。

5.4 总览对比视图
  · 学号排序，显示两次成绩。
  · 自动计算差值：较上次赋分总分、较上次校名、较上次班名。
  · 底部显示【进步前 N】和【退步前 N】。
  · 相关设置：设置 → 📊 总览

5.5 考试加分
  · 点击【🎯 考试加分】打开规则管理。
  · 规则类型：赋分总分、赋分校名、赋分班名、进步排名、退步排名、单科王。
  · 点击【📤 提交加分】后，预览加分名单，确认后写入积分系统。

─────────────────────────────────
六、名单管理
─────────────────────────────────
6.1 自定义名单
  · 用于管理班干部、值日组等。
  · 新建 / 重命名 / 删除名单。

6.2 条目管理
  · 条目字段：类别、成员（支持搜索、多选）、加分。
  · 添加条目、双击编辑、删除选中。

6.3 班干加分
  · 点击【🎯 班干加分】，系统汇总当前名单所有条目的成员和加分。
  · 弹出确认对话框，确认后批量写入积分系统。

─────────────────────────────────
七、导出与导入
─────────────────────────────────
7.1 导出 Excel
  · 模板文件（可选）：需包含列标题 班级、学号、姓名。
  · 月份（可选）：格式 YYYY-MM，留空导出全部。
  · 点击【📥 开始导出】。
  · 导出文件名会自动带上时间戳（如"班级积分月度汇总_2026-05_20260501_181230.xlsx"）。
  · 成功后弹窗：打开文件夹 / 打开文件 / 确认。

7.2 导出成绩总览
  · 导出"本次 vs 上次"的对比总览表，含学号、姓名、上次/本次各项
    分数与名次、以及差值。
  · 文件名自动带时间戳（如"成绩总览对比_20260501_181230.xlsx"）。

7.3 导入赋分总分
  · 准备 Excel，至少含 学号 和 赋分总分 两列。
  · 设置折扣比例（默认100%）。
  · 点击【📂 选择Excel并导入】，预览匹配结果后确认。

─────────────────────────────────
八、设置详解
─────────────────────────────────
8.1 等级信息面板
  · 显示当前等级、称号、经验进度条。
  · 📊 等级详情：查看累计录入、今日经验、总经验、经验明细。
  · 🏆 成就：打开成就总览面板。

8.2 🎨 外观
  · 开机动画：控制启动时是否播放 3 秒动画。
  · 进度条动画：切换到页面时进度条/圆环是否从 0 缓动到当前值。
  · 粒子效果：点击「✨ 粒子效果设置」打开弹窗，共 4 个独立开关：
      - 窗口关闭粒子效果：弹窗、主窗口关闭/最小化、悬浮窗双击时，播放粒子散开效果。
      - 鼠标点击粒子效果：鼠标左键点击任意位置时，在点击处溅射小粒子。
      - 鼠标拖动粒子效果：按住鼠标左键拖动时，在光标路径上拖出粒子尾迹。
      - 鼠标滚动粒子效果：鼠标滚轮滚动时，在光标位置溅射向上飘的粒子。

8.3 📝 事由
  · 📝 管理事由：增删改事由。
  · 事由数字快捷选择：开启后录入界面按 1-9 快速选择事由。

8.4 📋 最近记录
  · 显示最近条数：10-1000，默认100。

8.5 📊 总览
  · 进步/退步显示人数：1-30，默认5。
  · 进退步依据：赋分总分 / 赋分校名 / 赋分班名。

8.6 🗄️ 数据管理
  · 📂 切换名单：更换班级名单 Excel。
  · 🔄 重置所有积分：清空积分记录，自动备份。
  · 💾 存储空间：打开"存储空间"对话框，查看各分类占用、管理文件。
  · 🗑️ 格式化所有数据：三次确认后删除 user/ 目录并重启，谨慎使用。

8.7 📖 帮助
  · 📖 查看使用指南：打开本文件（支持关键字搜索、字号调节、复制全文）。

─────────────────────────────────
九、存储空间
─────────────────────────────────
入口：设置 → 🗄️ 数据管理 → 💾 存储空间

9.1 主界面组成
  · 顶部标题行：💾 存储空间 + 🔄 刷新 + 📂 打开目录
  · 大圆环：按分类着色，展示 user/ 目录各部分的占用比例。
      中心显示 user 目录的总占用大小。
      打开对话框时圆环从 0° 顺时针缓动展开（时长约 1.4 秒）。
      是否动画由"设置 → 🎨 外观 → 进度条动画"控制。
  · 图例：列出各分类的名称与颜色（哪个分类为 0 就不显示）。
  · 分类卡片：每张卡片显示一个大类的名称、大数字、单位、描述，
      以及一个右侧按钮（普通分类为"查看"，导出/备份为"管理"）。
  · 圆环、图例、分类卡片都在同一个滚动区域内，整体一起上下滑动。
  · 底部信息：user 目录总占用 + 所在磁盘已用/总量/剩余。

9.2 分类说明
  积分数据     : user/data/ 下所有 CSV 文件（积分记录、成绩数据）
  名单文件     : user/data/ 下所有 Excel 文件（班级名单、导出模板）
  导出文件     : user/exports/ 下所有文件（含时间戳的汇总表、总览等）
  备份文件     : user/backups/ 下所有文件（名单、积分、Excel 备份）
  便签与配置   : user/data/ 下所有 JSON 文件（便签、配置、自定义名单）

9.3 打开分类（子页）
  · 点击分类卡片或它的按钮，弹出该分类的文件列表。
  · 子页顶部显示："共 N 个文件 · 总占用 X"。
  · 文件行：左侧文件名，右侧大小，按修改时间从新到旧排序。
  · 底部按钮：📂 打开所在目录 / 关闭。

9.4 双击文件操作（仅"导出文件"和"备份文件"）
  · 在"导出文件"或"备份文件"子页中，双击任意一行弹出操作菜单。
  · 菜单共四个操作：
      - 📂 打开文件所在位置（独占一行，宽度填满）
        Windows 会打开资源管理器并高亮该文件；macOS 会在 Finder 中定位；
        其他平台打开所在目录。
      - 📄 打开（用系统默认程序打开该文件）
      - 🗑️ 删除（二次确认后删除该文件）
      - 取消
  · 其它分类（积分数据、名单文件、便签与配置）的子页是只读列表，
    双击没有任何反应，也不会显示"分组保留条数"按钮。

9.5 分组保留条数（仅"导出文件"和"备份文件"）
  · 两个子页右上角有一个"⚙️ 分组保留条数"按钮，点击打开配置对话框。
  · 配置对话框按类型分别设置保留条数（每个类型一个 SpinBox，范围 1-999，默认 5）：
      导出文件：班级积分月度汇总 / 成绩总览对比 / 其它
      备份文件：备份-积分记录 / 名单备份 / 积分记录 / 其它
  · 点击"确定"后：
      1. 立即把设置写入 config.json。
      2. 立刻按新规则清理目录：每种类型各自独立按修改时间从新到旧
         保留对应条数，超出的最旧文件会被删除。
      3. 主对话框的圆环、卡片、底部信息同步刷新。
      4. 子页里的文件列表也重排，顶部汇总会显示每种类型的当前数量
         （如"班级积分月度汇总 5/5 · 成绩总览对比 3/5"）。
  · 打开导出/备份子页时也会自动按配置的规则清理一次，
    保证进入时就已经是合规状态。
  · 打开存储空间主对话框时，以及点击"🔄 刷新"时，也会自动清理一次。

9.6 刷新
  · 点击"🔄 刷新"：
      1. 先按保留规则清理导出/备份目录（超出的旧文件立即删除）。
      2. 重新扫描所有分类。
      3. 圆环与图例重放动画。
      4. 分类卡片重建。
      5. 底部信息更新。
      6. 保持当前滚动位置不变。

9.7 打开目录
  · 点击"📂 打开目录"：在系统文件管理器中打开 user/ 目录。

─────────────────────────────────
十、等级与经验系统
─────────────────────────────────
每日总上限 120。

  来源      : 规则                        : 每日上限
  签到      : 连续1-6天递增，第7天起+20   : —
  在线计时  : 每2分钟+1                   : 20
  录入记录  : 每条+2                       : 60
  设备运行  : 每天+50                      : 不计入每日上限

等级表（1-20级）：
  Lv.1  0     微光初现      Lv.11 3000  星途进阶
  Lv.2  50    渐露锋芒      Lv.12 4000  灿星专务
  Lv.3  150   拾光前行      Lv.13 6000  熠星荣誉
  Lv.4  300   星芒见习      Lv.14 8000  繁星资深
  Lv.5  500   逐光实践      Lv.15 10000 慧星筹策
  Lv.6  800   耀光实干      Lv.16 13000 星河开拓
  Lv.7  1200  聚光执行      Lv.17 16000 星宇共建
  Lv.8  1600  凝光统筹      Lv.18 20000 瀚星构筑
  Lv.9  2000  流光协领      Lv.19 25000 极星领航
  Lv.10 2500  明光骨干      Lv.20 30000 万象星穹

查看方式：
  · 主窗口标题栏左侧实时显示 Lv.X。
  · 设置 → 等级信息面板 查看进度条和详情。
  · 升级时弹出全局 Toast 通知。

─────────────────────────────────
十一、成就系统
─────────────────────────────────
11.1 成就分类
  🏅 升级成就   : Lv.2 ~ Lv.20 每级一个
  📆 使用历程   : 初次启幕、十日相伴、月度同行、百秩征途
  📂 导入成绩   : 初导名册、三番汇册、五度整编、十载汇录
  📝 录入记录   : 落笔初记、百册留痕、…、鸿篇巨录
  📅 签到       : 七日不渝、月余坚守、半岁恒心
  📤 导出/备份  : 数表初成、存鉴留档

11.2 查看与解锁
  · 查看：设置 → 🏆 成就，显示所有成就进度。
  · 解锁时弹出通知，经验不占用每日上限。

─────────────────────────────────
十二、键盘快捷键
─────────────────────────────────
  Ctrl + Enter             提交录入
  Esc                      从录入/排行等页面返回首页
  Enter（学号框）           跳转事由框
  Enter（事由框）           选中事由并跳到分数框
  Enter（分数框）           提交
  1-9（事由框，需开启快捷）  快速选择对应事由
  Ctrl + 鼠标滚轮（日历）   切换月份
  双击（文件行）            在导出/备份子页弹出操作菜单
  鼠标滚轮滚动              在光标处溅射粒子（需开启"鼠标滚动粒子效果"）

─────────────────────────────────
十三、数据存储与备份
─────────────────────────────────
13.1 目录结构
  user/
   ├── data/
   │    ├── config.json             系统配置
   │    ├── notes.json              便签
   │    ├── custom_lists.json       自定义名单
   │    ├── 班级积分记录.csv         积分原始记录
   │    ├── 全班名单.xlsx            班级名单
   │    ├── current_scores.csv      本次成绩
   │    └── previous_scores.csv     上次成绩
   ├── exports/                     导出文件（每次带时间戳）
   └── backups/
        ├── 备份_积分记录_YYYYMMDD_HHMMSS.csv   （重置积分前备份）
        ├── 备份_积分记录_YYYYMMDD_HHMMSS.xlsx  （重置积分前备份）
        ├── 名单备份_YYYYMMDD_HHMMSS.csv        （切换名单时备份）
        └── 积分记录_YYYYMMDD.csv               （每日自动备份）

13.2 自动备份
  · 切换名单时自动备份旧名单。
  · 重置积分前自动备份积分记录和 Excel。
  · 每天自动备份积分记录一次。

13.3 保留策略（可按类型分别设置）
  · 导出文件按"班级积分月度汇总 / 成绩总览对比 / 其它"三类各自保留 N 条。
  · 备份文件按"备份-积分记录 / 名单备份 / 积分记录 / 其它"四类各自保留 N 条。
  · 默认每类 5 条，超过时最旧的自动删除。
  · 修改方式：存储空间 → 打开分类 → 右上角"⚙️ 分组保留条数"。
  · 触发时机：打开存储空间主对话框时、点击"🔄 刷新"时、
              打开导出/备份子页时、修改分组保留条数后。

─────────────────────────────────
十四、悬浮窗与托盘
─────────────────────────────────
  · 最小化主窗口：自动隐藏到系统托盘，并显示一个可拖动的悬浮图标。
  · 悬浮图标：
      - 双击恢复主窗口。
      - 右键菜单：显示主窗口、查看积分排行榜、退出。
  · 托盘图标：
      - 双击显示主窗口。
      - 右键菜单：显示主窗口、退出。
  · 关闭主窗口：完全退出程序。

─────────────────────────────────
十五、常见问题
─────────────────────────────────
Q1：图标加载失败？
    Base64 图标字符串未填写或损坏，不影响功能。

Q2：启动卡在开机动画？
    设置 → 🎨 外观 → 开机动画 关闭。

Q3：Excel 导入失败？
    确认包含 学号、姓名 列（成绩还需科目列），学号建议为纯文本。

Q4：如何彻底清空所有数据？
    设置 → 🗄️ 数据管理 → 🗑️ 格式化所有数据。

Q5：等级经验不涨？
    检查今日总经验是否达 120，或对应子项是否达上限。

Q6：最小化后程序不见了？
    程序变为右下角悬浮图标，双击恢复。

Q7：如何查看某学生所有记录？
    排行页面双击该学生行，或录入页面"录入记录"中双击。

Q8：连续模式如何跳过缺考学生？
    连续模式下点击【跳过】。

Q9：事由数字快捷键无效？
    检查 设置 → 📝 事由 → 事由数字快捷选择 是否开启。

Q10：进度条/圆环不动画？
    检查 设置 → 🎨 外观 → 进度条动画 是否开启。

Q11：鼠标没有粒子？
    检查 设置 → 🎨 外观 → ✨ 粒子效果设置 → 粒子效果 是否开启。
    另外，粒子触发有频率间隔保护，快速触发时不会每次都出粒子。

Q12：导出的文件会覆盖旧的吗？
    不会。每次导出的文件名都带有时间戳（如"班级积分月度汇总_20260501_181230.xlsx"）。
    但超过"分组保留条数"设置的最旧文件会被自动删除。

Q13：如何调整导出/备份文件保留条数？
    打开存储空间 → 进入"导出文件"或"备份文件" → 点右上角"⚙️ 分组保留条数"，
    为每个类型分别设置保留条数。

Q14：为什么超出的文件要打开"存储空间"才会自动删除？
    不会。"存储空间"主对话框打开时、点击"🔄 刷新"时、打开导出/备份子页时，
    都会自动触发清理。你也可以通过点击"🔄 刷新"主动触发。

Q15：双击文件没有反应？
    只有"导出文件"和"备份文件"子页支持双击操作。其它分类（积分数据、
    名单文件、便签与配置）是只读列表。

Q16：圆环的每一段代表什么？
    每个分类一段，颜色与卡片/图例中的颜色一致。圆环中心显示 user/
    目录的总占用大小。

─────────────────────────────────
十六、彩蛋
─────────────────────────────────
在【关于页面】快速点击版本号 5 次，触发满屏爱心动画 💖

===============================
感谢您的使用！
PickOS
有你更出色
Copyright © 2026 wuweijia-star  ·  GPLv3
===============================
"""


class UserGuideDialog(QDialog):
    """使用指南对话框（支持多结果搜索、字号调整）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📖 使用指南")
        self.resize(1000, 760)
        self.setStyleSheet(DIALOG_STYLE)

        # ---- 搜索状态 ----
        self.search_matches = []      # 所有匹配位置的字符索引
        self.search_index = -1        # 当前高亮的匹配索引
        self.last_keyword = ""        # 上次搜索的关键词
        self.current_font_size = 13   # 当前字号

        # 居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # ---- 标题 ----
        title = QLabel("📖 PickOS 班级管理系统使用指南")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        # ---- 工具行 ----
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        toolbar.addWidget(QLabel("搜索："))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入关键字（回车定位，再按回车跳下一个）")
        self.search_edit.setFixedWidth(280)
        self.search_edit.returnPressed.connect(self._do_search)
        toolbar.addWidget(self.search_edit)

        find_btn = QPushButton("🔍 查找")
        find_btn.setMinimumWidth(110)      # ← 改：自适应宽度
        find_btn.setMinimumHeight(34)
        find_btn.clicked.connect(self._do_search)
        toolbar.addWidget(find_btn)

        self.next_btn = QPushButton("下一个")
        self.next_btn.setMinimumWidth(100)  # ← 改
        self.next_btn.setMinimumHeight(34)
        self.next_btn.setEnabled(False)
        self.next_btn.clicked.connect(self._next_match)
        toolbar.addWidget(self.next_btn)

        clear_btn = QPushButton("清除高亮")
        clear_btn.setMinimumWidth(120)      # ← 改：四个中文字要宽一点
        clear_btn.setMinimumHeight(34)
        clear_btn.clicked.connect(self._clear_highlight)
        toolbar.addWidget(clear_btn)

        toolbar.addStretch()

        # ---- 状态标签 ----
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #A0A0B0; font-size: 13px;")
        toolbar.addWidget(self.status_label)

        toolbar.addStretch()

        # ---- 字号 ----
        toolbar.addWidget(QLabel("字号："))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(10, 24)
        self.font_size_spin.setValue(13)
        self.font_size_spin.setFixedWidth(70)
        self.font_size_spin.valueChanged.connect(self._apply_font_size)
        toolbar.addWidget(self.font_size_spin)

        layout.addLayout(toolbar)

        # ---- 指南内容浏览器 ----
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)
        self.browser.setPlainText(USER_GUIDE_TEXT)
        layout.addWidget(self.browser)
        # 统一通过 _apply_font_size 初始化样式（含 font-size，保证字号可控）
        self._apply_font_size(self.current_font_size)

        # ---- 底部按钮 ----
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        copy_btn = QPushButton("📋 复制全文")
        copy_btn.setMinimumWidth(130)      # ← 改
        copy_btn.setMinimumHeight(34)
        copy_btn.clicked.connect(self._copy_all)
        btn_layout.addWidget(copy_btn)

        close_btn = QPushButton("关闭")
        close_btn.setMinimumWidth(90)      # ← 改
        close_btn.setMinimumHeight(34)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    # ==================== 搜索逻辑 ====================
    def _do_search(self):
        """执行搜索：关键词变了就重新搜索全部匹配；否则跳到下一个"""
        keyword = self.search_edit.text().strip()
        if not keyword:
            self._clear_highlight()
            return

        # 同一个关键词 → 跳到下一个
        if keyword == self.last_keyword and self.search_matches:
            self._next_match()
            return

        # 新关键词 → 清空旧结果并重新搜索
        self._clear_highlight(keep_keyword=True)
        self.last_keyword = keyword

        # 找到所有匹配（字符索引）
        text = self.browser.toPlainText()
        matches = []
        start = 0
        while True:
            idx = text.find(keyword, start)
            if idx == -1:
                break
            matches.append(idx)
            start = idx + len(keyword)

        self.search_matches = matches
        self.search_index = -1

        if matches:
            self.status_label.setText(f"找到 {len(matches)} 条匹配")
            self.next_btn.setEnabled(True)
            self._next_match()
        else:
            self.status_label.setText(f"未找到：{keyword}")
            self.next_btn.setEnabled(False)
            show_toast(self, f"未找到：{keyword}", "info")

    def _next_match(self):
        """高亮并跳到下一个匹配"""
        if not self.search_matches:
            return
        self.search_index = (self.search_index + 1) % len(self.search_matches)
        self._highlight_all_matches()

        # 滚动到当前匹配
        idx = self.search_matches[self.search_index]
        cursor = QTextCursor(self.browser.document())
        cursor.setPosition(idx)
        cursor.setPosition(idx + len(self.last_keyword),
                            QTextCursor.MoveMode.KeepAnchor)
        self.browser.setTextCursor(cursor)
        self.browser.ensureCursorVisible()

        # 更新状态
        self.status_label.setText(
            f"第 {self.search_index + 1} / {len(self.search_matches)} 条匹配"
        )

    def _highlight_all_matches(self):
        """高亮全部匹配：当前项金黄，其余蓝色"""
        keyword = self.last_keyword
        if not keyword:
            self.browser.setExtraSelections([])
            return

        selections = []
        for i, idx in enumerate(self.search_matches):
            cursor = QTextCursor(self.browser.document())
            cursor.setPosition(idx)
            cursor.setPosition(idx + len(keyword),
                                QTextCursor.MoveMode.KeepAnchor)
            sel = QTextEdit.ExtraSelection()
            sel.cursor = cursor
            if i == self.search_index:
                # 当前匹配项 —— 金黄
                sel.format.setBackground(QColor("#FFD700"))
                sel.format.setForeground(QColor("#000000"))
            else:
                # 其它匹配项 —— 蓝色
                sel.format.setBackground(QColor("#64B5F6"))
                sel.format.setForeground(QColor("#FFFFFF"))
            selections.append(sel)

        self.browser.setExtraSelections(selections)

    def _clear_highlight(self, keep_keyword=False):
        """清除所有高亮"""
        self.browser.setExtraSelections([])
        self.search_matches = []
        self.search_index = -1
        self.next_btn.setEnabled(False)
        self.status_label.setText("")
        if not keep_keyword:
            self.last_keyword = ""

    # ==================== 字号调整 ====================
    def _apply_font_size(self, size):
        """调整字号：通过 QSS 强制覆盖，100% 生效"""
        self.current_font_size = size
        _c = get_theme_colors()

        # 保存滚动位置
        scrollbar = self.browser.verticalScrollBar()
        saved_scroll = scrollbar.value() if scrollbar else 0

        # 通过 QSS 修改 font-size —— 使用主题色
        self.browser.setStyleSheet(f"""
            QTextBrowser {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border']};
                border-radius: 10px;
                color: {_c['fg']};
                padding: 12px;
                selection-background-color: #64B5F6;
                selection-color: white;
                font-family: "Consolas", "Microsoft YaHei", monospace;
                font-size: {size}pt;
            }}
        """)

        # 恢复滚动位置
        if scrollbar:
            scrollbar.setValue(saved_scroll)

        # 重新应用搜索高亮
        if self.search_matches:
            self._highlight_all_matches()

    # ==================== 复制 ====================
    def _copy_all(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(USER_GUIDE_TEXT)
        show_toast(self, "使用指南已复制到剪贴板", "success")


class LevelDetailDialog(QDialog):
    """等级详情对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🎖️ 等级详情")
        self.resize(420, 320)
        self.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(16)

        # 加载数据
        cfg = load_config()
        total_exp = cfg.get("total_experience", 0)
        total_records = cfg.get("total_records_count", 0)
        today_exp = cfg.get("today_experience", 0)
        level = calculate_level_by_exp(total_exp)
        level_range = LEVEL_RANGES.get(level, "")
        title = LEVEL_TITLES.get(level, "未知")
        combined = f"{level_range}·{title}" if level_range else title
        level_title = QLabel(f"Lv.{level}  「{combined}」")

        # 等级与称号（大号）
        level_title = QLabel(f"Lv.{level}  「{title}」")
        level_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        level_title.setStyleSheet("font-size: 28px; font-weight: bold; color: #64B5F6;")
        layout.addWidget(level_title)

        # 经验进度条
        exp_current, exp_needed, progress = get_level_progress(total_exp)
        if level >= MAX_LEVEL:
            progress_text = "已满级"
        else:
            progress_text = f"{exp_current:.0f} / {exp_current + exp_needed:.0f}"
        progress_bar = QProgressBar()
        progress_bar.setRange(0, 100)
        progress_bar.setValue(0)   # 起点先设为 0，动画从 0 缓动
        progress_bar.setFormat(f"{progress_text}  ({int(progress)}%)")
        progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background: rgba(255,255,255,0.1);
                border-radius: 6px;
                height: 20px;
                text-align: center;
                color: #FFFFFF;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                            stop:0 #64B5F6, stop:1 #42A5F5);
                border-radius: 6px;
            }
        """)
        layout.addWidget(progress_bar)
        # 延迟触发动画：等对话框显示后再启动，否则动画可能在不可见时跑完
        QTimer.singleShot(100, lambda pb=progress_bar, v=int(progress):
                          animate_progress_bar(pb, v))

        # 升级所需经验（若未满级）
        if level < MAX_LEVEL:
            need_label = QLabel(f"升级还需 {exp_needed:.0f} 经验")
            need_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            need_label.setStyleSheet("color: #A0A0B0; font-size: 14px;")
            layout.addWidget(need_label)

        # 分割线
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("background: #3A3A4A;")
        layout.addWidget(line)

        # 统计信息
        stats_layout = QFormLayout()
        stats_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        stats_layout.setSpacing(10)

        stats_layout.addRow("累计录入：", QLabel(str(total_records)))
        stats_layout.addRow("今日经验：", QLabel(f"{today_exp} / {DAILY_LIMIT_TOTAL}"))
        stats_layout.addRow("总经验：", QLabel(f"{total_exp:.0f}"))

        layout.addLayout(stats_layout)

        # 查看明细按钮
        detail_btn = QPushButton("📋 查看经验明细")
        detail_btn.clicked.connect(self.show_experience_log)
        layout.addWidget(detail_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # 确保窗口居中
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

    def show_experience_log(self):
        """弹出经验明细对话框"""
        dialog = ExperienceLogDialog(self)
        dialog.exec()

class AchievementDialog(QDialog):
    """成就总览面板"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🏆 成就系统")
        self.resize(700, 500)
        self.setStyleSheet(DIALOG_STYLE)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title = QLabel("🏆 成就系统")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        scroll_widget = QWidget()
        scroll_widget.setStyleSheet("background: transparent;")   # 新增：背景透明
        scroll_widget.setAutoFillBackground(False)               # 新增：禁用自动填充背景
        self.grid_layout = QGridLayout(scroll_widget)
        self.grid_layout.setSpacing(8)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.load_achievements()

    def load_achievements(self):
        cfg = load_config()
        unlocked = cfg.get("achievements", {})
        context = {
            "level": calculate_level(),
            "days_since_install": int((time.time() - cfg.get("install_time", time.time())) // 86400),
            "total_records": cfg.get("total_records_count", 0),
            "sign_in_streak": cfg.get("sign_in_streak", 0),
            "import_count": cfg.get("import_count", 0),
            "export_count": cfg.get("export_count", 0),
            "backup_count": cfg.get("backup_count", 0),
        }

        # 分组
        groups = {}
        for ach in ACHIEVEMENT_LIST:
            prefix = ach["aid"].split("_")[0] if "_" in ach["aid"] else "other"
            groups.setdefault(prefix, []).append(ach)

        group_titles = {
            "lv": "🏅 升级成就",
            "walk": "📆 使用历程",
            "import": "📂 导入成绩",
            "record": "📝 录入记录",
            "sign": "📅 签到",
            "export": "📤 导出备份",
            "backup": "💾 备份",
            "first": "🌟 初次启幕"
        }

        row = 0
        for group_name, group_achs in groups.items():
            title_label = QLabel(group_titles.get(group_name, group_name))
            title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #64B5F6; margin-top: 10px;")
            self.grid_layout.addWidget(title_label, row, 0, 1, 2)
            row += 1
            for ach in group_achs:
                aid = ach["aid"]
                unlocked_flag = unlocked.get(aid, False)
                progress = self._get_progress(aid, context)
                if unlocked_flag:
                    status_text = "✅ 已解锁"
                    color = "#4CAF50"
                elif progress is not None:
                    status_text = f"进度: {progress}"
                    color = "#FF9800"
                else:
                    status_text = "🔒 未解锁"
                    color = "#888888"

                card = QFrame()
                card.setStyleSheet("""
                    QFrame {
                        background: rgba(100, 181, 246, 0.1);
                        border-radius: 12px;
                        border: 1px solid #3A3A4A;
                        padding: 8px;
                    }
                """)
                card_layout = QHBoxLayout(card)
                card_layout.setContentsMargins(10, 5, 10, 5)
                name_label = QLabel(ach['name'])
                name_label.setStyleSheet("font-weight: bold; color: {_c['fg']};")
                desc_label = QLabel(f"{ach['desc']} (+{ach['exp']}经验)")
                desc_label.setStyleSheet("color: #A0A0B0; font-size: 12px;")
                status_label = QLabel(status_text)
                status_label.setStyleSheet(f"color: {color}; font-weight: bold;")
                card_layout.addWidget(name_label)
                card_layout.addWidget(desc_label)
                card_layout.addStretch()
                card_layout.addWidget(status_label)
                self.grid_layout.addWidget(card, row, 0, 1, 2)
                row += 1

    def _get_progress(self, aid, context):
        if aid.startswith("lv_"):
            lv = int(aid.split("_")[1])
            return f"{context['level']}/{lv}" if context['level'] < lv else None
        elif aid.startswith("walk_"):
            days = int(aid.split("_")[1].replace("day", ""))
            return f"{context['days_since_install']}/{days}" if context['days_since_install'] < days else None
        elif aid.startswith("import_"):
            count = int(aid.split("_")[1])
            return f"{context['import_count']}/{count}" if context['import_count'] < count else None
        elif aid.startswith("record_"):
            count = 1 if aid == "record_first" else int(aid.split("_")[1])
            return f"{context['total_records']}/{count}" if context['total_records'] < count else None
        elif aid.startswith("sign_"):
            count = int(aid.split("_")[1])
            return f"{context['sign_in_streak']}/{count}" if context['sign_in_streak'] < count else None
        elif aid == "export_first":
            return f"{context['export_count']}/1" if context['export_count'] < 1 else None
        elif aid == "backup_first":
            return f"{context['backup_count']}/1" if context['backup_count'] < 1 else None
        return None

class ExperienceLogDialog(QDialog):
    """今日经验汇总卡片视图"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📋 今日经验明细")
        self.resize(500, 500)
        _c = get_theme_colors()
        self.setStyleSheet(f"""
            QDialog {{
                background: {_c['bg']};
                border-radius: 18px;
            }}
            QFrame {{
                background: {_c['bg_alt']};
                border-radius: 12px;
                border: 1px solid {_c['border_light']};
            }}
            QLabel {{
                color: {_c['fg']};
            }}
            QProgressBar {{
                border: none;
                background: rgba(150, 150, 150, 0.2);
                border-radius: 9px;
                text-align: center;
                color: {_c['fg_bright']};
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
                border-radius: 9px;
            }}
            QPushButton {{
                border-radius: 10px;
                padding: 8px 20px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
                color: white;
                font-weight: 600;
                border: 1px solid rgba(100, 181, 246, 0.3);
                font-size: 15px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
                border-color: #64B5F6;
            }}
            QPushButton:pressed {{
                background: #1E88E5;
            }}
            QScrollArea {{
                border: none;
                background: transparent;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title = QLabel("📋 今日经验明细")
        title.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        scroll_widget = QWidget()
        scroll_widget.setStyleSheet("background: transparent;")
        self.card_layout = QVBoxLayout(scroll_widget)
        self.card_layout.setSpacing(10)
        self.card_layout.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)

        self.total_widget = QWidget()
        total_layout = QHBoxLayout(self.total_widget)
        self.total_label = QLabel("今日总经验：")
        self.total_label.setStyleSheet(f"font-weight: bold; font-size: 16px; color: {_c['fg']};")
        total_layout.addWidget(self.total_label)
        self.total_progress = QProgressBar()
        self.total_progress.setRange(0, DAILY_LIMIT_TOTAL)
        self.total_progress.setFixedHeight(18)
        total_layout.addWidget(self.total_progress)
        self.total_value = QLabel("0 / 120")
        self.total_value.setStyleSheet("font-weight: bold; color: #64B5F6;")
        total_layout.addWidget(self.total_value)
        layout.addWidget(self.total_widget)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        clear_btn = QPushButton("🗑️ 清空日志")
        clear_btn.clicked.connect(self.clear_log)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.load_data()

    def load_data(self):
        while self.card_layout.count():
            item = self.card_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        cfg = load_config()
        today = get_today_date()
        today_exp = cfg.get("today_experience", 0)
        sources = []

        last_sign = cfg.get("last_sign_in_date", "")
        streak = cfg.get("sign_in_streak", 0)
        if last_sign == today:
            sign_exp = SIGN_IN_DAILY_BONUS if streak >= 7 else SIGN_IN_REWARDS.get(streak, 2)
            sign_progress = f"连续签到 {streak} 天"
            sign_earned = sign_exp
        else:
            sign_progress = "今日未签到"
            sign_earned = 0
        sources.append({"icon": "📅", "name": "签到", "progress": sign_progress, "earned": sign_earned, "max": None, "detail": f"+{sign_exp}" if last_sign == today else "未签到"})

        online_exp = cfg.get("today_online_exp", 0)
        online_sec = cfg.get("online_seconds_today", 0)
        online_min = online_sec // 60
        sources.append({"icon": "⏱️", "name": "在线计时", "progress": f"已在线 {online_min} 分钟", "earned": online_exp, "max": DAILY_LIMIT_ONLINE, "detail": f"{online_exp}/{DAILY_LIMIT_ONLINE}"})

        record_exp = cfg.get("today_record_exp", 0)
        record_count = record_exp // RECORD_EXP_PER_ENTRY
        sources.append({"icon": "📝", "name": "录入记录", "progress": f"已录入 {record_count} 条记录", "earned": record_exp, "max": DAILY_LIMIT_RECORD, "detail": f"{record_exp}/{DAILY_LIMIT_RECORD}"})

        last_device = cfg.get("last_device_day", "")
        device_exp = DEVICE_RUNTIME_DAILY_EXP if last_device == today else 0
        device_progress = "今日已领取" if last_device == today else "今日未领取"
        sources.append({"icon": "📱", "name": "设备运行", "progress": device_progress, "earned": device_exp, "max": None, "detail": f"+{DEVICE_RUNTIME_DAILY_EXP}" if last_device == today else "未领取"})

        for src in sources:
            card = self.create_card(src)
            self.card_layout.addWidget(card)

        QTimer.singleShot(100, lambda pb=self.total_progress, v=int(today_exp): animate_progress_bar(pb, v))
        self.total_value.setText(f"{today_exp} / {DAILY_LIMIT_TOTAL}")
        self.total_label.setText(f"今日总经验：{today_exp}")
        self.card_layout.addStretch()

    def create_card(self, src):
        _c = get_theme_colors()
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background: {_c['bg_alt']};
                border-radius: 12px;
                border: 1px solid {_c['border_light']};
                padding: 10px;
            }}
            QLabel {{
                color: {_c['fg']};
            }}
        """)
        layout = QHBoxLayout(card)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(12)

        icon_label = QLabel(src["icon"])
        icon_label.setStyleSheet("font-size: 28px; background: transparent;")
        layout.addWidget(icon_label)

        middle = QVBoxLayout()
        name_label = QLabel(src["name"])
        name_label.setStyleSheet(f"font-weight: bold; font-size: 15px; color: {_c['fg_bright']}; background: transparent;")
        progress_label = QLabel(src["progress"])
        progress_label.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 13px; background: transparent;")
        middle.addWidget(name_label)
        middle.addWidget(progress_label)
        layout.addLayout(middle, stretch=1)

        earned = src["earned"]
        max_val = src.get("max")
        detail_text = f"{earned}/{max_val}" if max_val is not None else src.get("detail", f"+{earned}")
        color = "#4CAF50" if earned > 0 else _c["fg_dim"]
        exp_label = QLabel(detail_text)
        exp_label.setStyleSheet(f"font-weight: bold; font-size: 15px; color: {color}; background: transparent;")
        layout.addWidget(exp_label)
        return card

    def clear_log(self):
        reply = QMessageBox.question(self, "确认清空", "确定要清空所有经验记录吗？此操作不可撤销！",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            cfg = load_config()
            cfg["experience_log"] = []
            save_config(cfg)
            show_toast(self, "经验日志已清空", "info")
    

# ======================== 事由管理对话框 ========================
class ReasonManagerDialog(QDialog):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.setWindowTitle("📝 管理事由")
        self.resize(750, 550)
        self.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        title = QLabel("📝 管理事由")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: {_c['fg']};")
        layout.addWidget(title)

        info = QLabel("事由用于积分录入时自动填充类型、分数和类别。")
        info.setStyleSheet("color: #A0A0A0; font-size: 14px;")
        layout.addWidget(info)

        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setAutoFillBackground(True)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["事由名称", "类型", "默认分数", "匹配类别"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)   # 隐藏行号
        layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(12)
        add_btn = QPushButton("➕ 添加")
        add_btn.clicked.connect(self.add_reason)
        edit_btn = QPushButton("✏️ 编辑")
        edit_btn.clicked.connect(self.edit_reason)
        delete_btn = QPushButton("🗑️ 删除")
        delete_btn.clicked.connect(self.delete_reason)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.load_data()

    def load_data(self):
        cfg = load_config()
        self.reasons = cfg.get("custom_reasons", [])
        # 迁移：为旧数据补上 category 字段
        for r in self.reasons:
            if "category" not in r:
                r["category"] = ""
        self.table.setRowCount(len(self.reasons))
        for i, r in enumerate(self.reasons):
            self.table.setItem(i, 0, QTableWidgetItem(r["name"]))
            self.table.setItem(i, 1, QTableWidgetItem(r["type"]))
            self.table.setItem(i, 2, QTableWidgetItem(str(r["score"])))
            self.table.setItem(i, 3, QTableWidgetItem(r.get("category", "")))

    def add_reason(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("添加事由")
        dialog.resize(350, 300)
        layout = QFormLayout(dialog)
        name_edit = QLineEdit()
        layout.addRow("事由名称：", name_edit)
        type_plus = QRadioButton("加分")
        type_minus = QRadioButton("扣分")
        type_plus.setChecked(True)
        type_widget = QWidget()
        type_layout = QHBoxLayout(type_widget)
        type_layout.addWidget(type_plus)
        type_layout.addWidget(type_minus)
        layout.addRow("类型：", type_widget)
        score_edit = QLineEdit()
        layout.addRow("默认分数：", score_edit)
        # 类别下拉框
        category_combo = QComboBox()
        category_combo.addItems(["（不匹配）", "日常", "听写", "背诵", "班干加分", "考试加分"])
        category_combo.setCurrentIndex(0)
        layout.addRow("匹配类别：", category_combo)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addRow(btn_box)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_edit.text().strip()
            if not name:
                show_toast(self, "名称不能为空", "error")
                return
            try:
                score = float(score_edit.text().strip())
            except:
                show_toast(self, "分数须为数字", "error")
                return
            if any(r["name"] == name for r in self.reasons):
                show_toast(self, "已存在同名事由", "warning")
                return
            type_ = "+" if type_plus.isChecked() else "-"
            category = category_combo.currentText()
            if category == "（不匹配）":
                category = ""
            self.reasons.append({"name": name, "type": type_, "score": score, "category": category})
            self.save_and_refresh()

    def edit_reason(self):
        row = self.table.currentRow()
        if row < 0:
            show_toast(self, "请先选择一条事由", "warning")
            return
        old = self.reasons[row]
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑事由")
        dialog.resize(350, 300)
        layout = QFormLayout(dialog)
        name_edit = QLineEdit(old["name"])
        layout.addRow("事由名称：", name_edit)
        type_plus = QRadioButton("加分")
        type_minus = QRadioButton("扣分")
        if old["type"] == "+":
            type_plus.setChecked(True)
        else:
            type_minus.setChecked(True)
        type_widget = QWidget()
        type_layout = QHBoxLayout(type_widget)
        type_layout.addWidget(type_plus)
        type_layout.addWidget(type_minus)
        layout.addRow("类型：", type_widget)
        score_edit = QLineEdit(str(old["score"]))
        layout.addRow("默认分数：", score_edit)
        category_combo = QComboBox()
        category_combo.addItems(["（不匹配）", "日常", "听写", "背诵", "班干加分", "考试加分"])
        current_cat = old.get("category", "")
        index = 0
        if current_cat:
            try:
                index = ["日常", "听写", "背诵", "班干加分", "考试加分"].index(current_cat) + 1
            except ValueError:
                index = 0
        category_combo.setCurrentIndex(index)
        layout.addRow("匹配类别：", category_combo)
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addRow(btn_box)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            name = name_edit.text().strip()
            if not name:
                show_toast(self, "名称不能为空", "error")
                return
            try:
                score = float(score_edit.text().strip())
            except:
                show_toast(self, "分数须为数字", "error")
                return
            if any(i != row and r["name"] == name for i, r in enumerate(self.reasons)):
                show_toast(self, "已存在同名事由", "warning")
                return
            type_ = "+" if type_plus.isChecked() else "-"
            category = category_combo.currentText()
            if category == "（不匹配）":
                category = ""
            self.reasons[row] = {"name": name, "type": type_, "score": score, "category": category}
            self.save_and_refresh()

    def delete_reason(self):
        row = self.table.currentRow()
        if row < 0:
            show_toast(self, "请先选择一条事由", "warning")
            return
        if QMessageBox.question(self, "确认删除", f"确定要删除事由“{self.reasons[row]['name']}”吗？",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            del self.reasons[row]
            self.save_and_refresh()

    def save_and_refresh(self):
        cfg = load_config()
        # 确保所有事由都有 category 字段
        for r in self.reasons:
            if "category" not in r:
                r["category"] = ""
        cfg["custom_reasons"] = self.reasons
        save_config(cfg)
        self.load_data()
        self.app.refresh_input_reasons()

class StudentSelectionDialog(QDialog):
    """从班级名单中选择学生（支持搜索、多选），按学号排序，已选成员置顶并高亮"""
    def __init__(self, parent, stu_df, id_to_name, selected_ids=None):
        super().__init__(parent)
        _c = get_theme_colors()
        self.stu_df = stu_df
        self.id_to_name = id_to_name
        self.selected_ids = set(selected_ids or [])
        self.result_ids = []
        self.setWindowTitle("选择学生")
        self.resize(500, 450)
        self.setStyleSheet(DIALOG_STYLE + f"""
            QListWidget {{
                background: {_c['bg_alt']};
                border: 1px solid {_c['border_light']};
                border-radius: 8px;
                outline: none;
            }}
            QListWidget::item {{
                padding: 6px;
                border-bottom: 1px solid {_c['border']};
                color: {_c['fg']};
            }}
            QListWidget::item:selected {{
                background: #64B5F6;
                color: #FFFFFF;
                font-weight: bold;
                border-left: 3px solid #42A5F5;
            }}
            QListWidget::item:hover {{
                background: {_c['bg_card']};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # 搜索框
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("搜索："))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("输入学号或姓名")
        self.search_edit.textChanged.connect(self.filter_students)
        search_layout.addWidget(self.search_edit)
        layout.addLayout(search_layout)

        # 学生列表（多选）
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        layout.addWidget(self.list_widget)

        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(self.accept_selection)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        # 加载所有学生并按学号排序
        self.all_students = []
        for _, row in stu_df.iterrows():
            sid = str(row["学号"])
            name = row["姓名"]
            self.all_students.append((sid, name))
        self.all_students.sort(key=lambda x: int(x[0]))

        self.filter_students()

    def filter_students(self):
        """根据搜索关键字过滤，并将已选成员置顶"""
        keyword = self.search_edit.text().strip().lower()
        matched = []
        for sid, name in self.all_students:
            if keyword and keyword not in sid.lower() and keyword not in name.lower():
                continue
            matched.append((sid, name))
        # 分离已选和未选
        selected = [item for item in matched if item[0] in self.selected_ids]
        unselected = [item for item in matched if item[0] not in self.selected_ids]
        sorted_items = selected + unselected
        self.list_widget.clear()
        for sid, name in sorted_items:
            item = QListWidgetItem(f"{sid}  {name}")
            item.setData(Qt.ItemDataRole.UserRole, sid)
            if sid in self.selected_ids:
                item.setSelected(True)
            self.list_widget.addItem(item)

    def accept_selection(self):
        selected_items = self.list_widget.selectedItems()
        self.result_ids = [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]
        self.accept()

class CategoryEditDialog(QDialog):
    """编辑类别、成员、加分的对话框"""
    def __init__(self, parent, category_data, stu_df, id_to_name):
        """
        category_data: dict with keys "category", "members" (list of student ids), "bonus"
        """
        super().__init__(parent)
        self.parent = parent
        self.stu_df = stu_df
        self.id_to_name = id_to_name
        self.category_data = category_data.copy()
        self.result = None   # 返回修改后的数据

        self.setWindowTitle("编辑条目")
        self.resize(450, 280)
        self.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        # 类别名称
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.category_edit = QLineEdit(self.category_data.get("category", ""))
        form_layout.addRow("类别名称：", self.category_edit)

        # 成员（显示已选学生）
        self.member_display = QLabel()
        self.member_display.setWordWrap(True)
        self.member_display.setStyleSheet("color: #A0A0A0; border: 1px solid #3A3A4A; border-radius: 6px; padding: 6px;")
        self.update_member_display()
        member_btn = QPushButton("选择成员")
        member_btn.clicked.connect(self.choose_members)
        member_layout = QHBoxLayout()
        member_layout.addWidget(self.member_display, stretch=1)
        member_layout.addWidget(member_btn)
        form_layout.addRow("成员：", member_layout)

        # 加分
        self.bonus_edit = QSpinBox()
        self.bonus_edit.setRange(-999, 999)
        self.bonus_edit.setValue(self.category_data.get("bonus", 0))
        form_layout.addRow("加分：", self.bonus_edit)

        layout.addLayout(form_layout)

        # 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.accepted.connect(self.accept_changes)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def choose_members(self):
        """弹出学生选择对话框，多选"""
        if self.stu_df is None or self.stu_df.empty:
            show_toast(self, "请先加载班级名单", "error")
            return
        current_ids = set(self.category_data.get("members", []))
        dialog = StudentSelectionDialog(self, self.stu_df, self.id_to_name, selected_ids=current_ids)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.category_data["members"] = dialog.result_ids
            self.update_member_display()

    def update_member_display(self):
        """更新成员显示标签，按学号排序"""
        ids = self.category_data.get("members", [])
        if not ids:
            self.member_display.setText("（无）")
            return
        # 按学号数字排序
        sorted_ids = sorted(ids, key=lambda x: int(x))
        names = []
        for sid in sorted_ids:
            name = self.id_to_name.get(sid, sid)
            names.append(f"{name}（{sid}）")
        self.member_display.setText("、".join(names))

    def accept_changes(self):
        category = self.category_edit.text().strip()
        if not category:
            show_toast(self, "类别名称不能为空", "error")
            return
        bonus = self.bonus_edit.value()
        self.result = {
            "category": category,
            "members": self.category_data.get("members", []),
            "bonus": bonus
        }
        self.accept()

class ListManagerInterface(QWidget):
    """自定义名单管理页面（类似成绩总览布局，支持类别条目，含序号列）"""
    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.setObjectName("listManagerInterface")
        self.custom_lists = load_custom_lists()
        self.current_index = -1
        self.build_ui()
        self.refresh_combo()

    def build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        # ---- 顶部行：标题 + 下拉框 + 名单操作按钮 ----
        top = QHBoxLayout()
        title = TitleLabel("📋 名单管理", self)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: {_c['fg']};")
        top.addWidget(title)
        top.addStretch()

        top.addWidget(QLabel("名单："))
        self.combo = QComboBox()
        self.combo.currentIndexChanged.connect(self.on_combo_changed)
        top.addWidget(self.combo)

        add_list_btn = QPushButton("➕ 新建")
        add_list_btn.clicked.connect(self.add_list)
        top.addWidget(add_list_btn)

        del_list_btn = QPushButton("🗑️ 删除")
        del_list_btn.clicked.connect(self.delete_list)
        top.addWidget(del_list_btn)

        rename_list_btn = QPushButton("✏️ 重命名")
        rename_list_btn.clicked.connect(self.rename_list)
        top.addWidget(rename_list_btn)

        bonus_btn = QPushButton("🎯 班干加分")
        bonus_btn.clicked.connect(self.add_bonus_for_list)
        top.addWidget(bonus_btn)

        layout.addLayout(top)

        # ---- 条目表格（增加序号列） ----
        self.table = QTableWidget()
        install_scrolling_header(self.table)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["序号", "类别", "成员", "加分"])

        # ★ 关键：一创建就开启自动换行，不要等到 setItem 之后
        self.table.setWordWrap(True)
        # ★ 关键：不允许省略号，保证内容完整换行显示
        self.table.setTextElideMode(Qt.TextElideMode.ElideNone)

        # ★ 关键：分列设置宽度策略，让"成员"列拿到剩余空间，序号/加分列给固定宽度
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)        # 序号
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)  # 类别
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)      # 成员
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)        # 加分
        self.table.setColumnWidth(0, 60)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(3, 80)

        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self.on_table_double_click)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # ---- 底部：统计 + 条目操作按钮 ----
        bottom = QHBoxLayout()
        self.item_count_label = QLabel("共 0 项")
        self.item_count_label.setStyleSheet("color: #A0A0B0; font-size: 15px;")
        bottom.addWidget(self.item_count_label)

        bottom.addStretch()

        add_item_btn = QPushButton("➕ 添加条目")
        add_item_btn.clicked.connect(self.add_item)
        bottom.addWidget(add_item_btn)

        del_item_btn = QPushButton("🗑️ 删除选中")
        del_item_btn.clicked.connect(self.delete_item)
        bottom.addWidget(del_item_btn)

        layout.addLayout(bottom)

        # ---- 返回按钮 ----
        back_btn = QPushButton("🔙 返回主菜单")
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        self.clear_table()

    # ---------- 名单下拉框 ----------
    def refresh_combo(self):
        current_name = self.combo.currentText() if self.combo.count() > 0 else ""
        self.combo.clear()
        for item in self.custom_lists:
            self.combo.addItem(item["name"])
        if current_name:
            idx = self.combo.findText(current_name)
            if idx >= 0:
                self.combo.setCurrentIndex(idx)
            else:
                self.combo.setCurrentIndex(0 if self.combo.count() > 0 else -1)
        else:
            self.combo.setCurrentIndex(0 if self.combo.count() > 0 else -1)

        if self.combo.currentIndex() >= 0:
            self.current_index = self.combo.currentIndex()
            self.show_items(self.current_index)
        else:
            self.current_index = -1
            self.clear_table()

    def on_combo_changed(self, index):
        if index < 0 or index >= len(self.custom_lists):
            self.current_index = -1
            self.clear_table()
            return
        self.current_index = index
        self.show_items(index)

    # ---------- 表格显示 ----------
    def show_items(self, index):
        items = self.custom_lists[index]["items"]
        self.table.setRowCount(len(items))
        id_to_name = self.app.id_to_name
        for i, item in enumerate(items):
            # 序号
            no_item = QTableWidgetItem(str(i + 1))
            no_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 0, no_item)

            # 类别
            cat_item = QTableWidgetItem(item["category"])
            cat_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 1, cat_item)

            # 成员（按学号排序显示）
            members = item.get("members", [])
            if members:
                sorted_members = sorted(members, key=lambda x: int(x))
                names = [f"{id_to_name.get(sid, sid)}（{sid}）" for sid in sorted_members]
                member_str = "、".join(names)
            else:
                member_str = ""
            member_item = QTableWidgetItem(member_str)
            # ★ 关键：成员列左对齐，配合换行显示更自然
            member_item.setTextAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            self.table.setItem(i, 2, member_item)

            # 加分
            bonus_item = QTableWidgetItem(str(item.get("bonus", 0)))
            bonus_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 3, bonus_item)

        self.table.setUpdatesEnabled(False)              # ← 新增
        try:
            self.table.setRowCount(len(items))
            for i, item in enumerate(items):
                no_item = QTableWidgetItem(str(i + 1))
                no_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(i, 0, no_item)

                cat_item = QTableWidgetItem(item["category"])
                cat_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(i, 1, cat_item)

                members = item.get("members", [])
                if members:
                    sorted_members = sorted(members, key=lambda x: int(x))
                    names = [f"{id_to_name.get(sid, sid)}（{sid}）" for sid in sorted_members]
                    member_str = "、".join(names)
                else:
                    member_str = ""
                member_item = QTableWidgetItem(member_str)
                member_item.setTextAlignment(
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                )
                self.table.setItem(i, 2, member_item)

                bonus_item = QTableWidgetItem(str(item.get("bonus", 0)))
                bonus_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.table.setItem(i, 3, bonus_item)
        finally:
            self.table.setUpdatesEnabled(True)           # ← 新增

        self.item_count_label.setText(f"共 {len(items)} 项")

        QTimer.singleShot(0, self._adjust_row_heights)
        QTimer.singleShot(50, self._adjust_row_heights)

    def _adjust_row_heights(self):
        """根据当前列宽和内容重新计算行高"""
        if self.table.rowCount() == 0:
            return
        self.table.resizeRowsToContents()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        # 窗口尺寸变化时，延迟重算行高，防止成员内容被裁切
        if not hasattr(self, "_row_height_timer") or self._row_height_timer is None:
            self._row_height_timer = QTimer(self)
            self._row_height_timer.setSingleShot(True)
            self._row_height_timer.timeout.connect(self._adjust_row_heights)
        self._row_height_timer.start(50)   # 50ms 去抖

    def clear_table(self):
        self.table.setRowCount(0)
        self.item_count_label.setText("共 0 项")

    # ---------- 名单操作 ----------
    def add_list(self):
        name, ok = QInputDialog.getText(self, "新建名单", "请输入名单名称：")
        if ok and name.strip():
            name = name.strip()
            if any(item["name"] == name for item in self.custom_lists):
                show_toast(self, "已存在同名名单", "warning")
                return
            self.custom_lists.append({"name": name, "items": []})
            save_custom_lists(self.custom_lists)
            self.refresh_combo()
            idx = self.combo.findText(name)
            if idx >= 0:
                self.combo.setCurrentIndex(idx)
            show_toast(self, f"名单“{name}”已创建", "success")

    def delete_list(self):
        if self.current_index < 0:
            show_toast(self, "请先选择一个名单", "warning")
            return
        name = self.custom_lists[self.current_index]["name"]
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("确认删除")
        msg_box.setText(f"确定要删除名单“{name}”吗？")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)  # 应用暗色样式
        reply = msg_box.exec()
        if reply == QMessageBox.StandardButton.Yes:
            del self.custom_lists[self.current_index]
            save_custom_lists(self.custom_lists)
            self.current_index = -1
            self.refresh_combo()
            show_toast(self, f"名单“{name}”已删除", "info")

    def rename_list(self):
        if self.current_index < 0:
            show_toast(self, "请先选择一个名单", "warning")
            return
        old_name = self.custom_lists[self.current_index]["name"]
        new_name, ok = QInputDialog.getText(self, "重命名名单", "请输入新名称：", text=old_name)
        if ok and new_name.strip():
            new_name = new_name.strip()
            if any(item["name"] == new_name for item in self.custom_lists if item["name"] != old_name):
                show_toast(self, "已存在同名名单", "warning")
                return
            self.custom_lists[self.current_index]["name"] = new_name
            save_custom_lists(self.custom_lists)
            self.refresh_combo()
            idx = self.combo.findText(new_name)
            if idx >= 0:
                self.combo.setCurrentIndex(idx)
            show_toast(self, f"已重命名为“{new_name}”", "success")

    def add_bonus_for_list(self):
        """为当前名单的所有成员批量添加班干加分"""
        if self.current_index < 0:
            show_toast(self, "请先选择一个名单", "warning")
            return
        if self.app.stu_df is None or self.app.stu_df.empty:
            show_toast(self, "请先加载班级名单", "error")
            return

        items = self.custom_lists[self.current_index]["items"]
        if not items:
            show_toast(self, "当前名单没有条目", "warning")
            return

        bonus_records = []  # 每个元素 (学号, 姓名, 加分值, 类别)
        for item in items:
            category = item.get("category", "")
            members = item.get("members", [])
            bonus = item.get("bonus", 0)
            if not members or bonus == 0:
                continue
            for sid in members:
                if sid not in self.app.id_to_name:
                    continue
                name = self.app.id_to_name[sid]
                bonus_records.append((sid, name, bonus, category))

        if not bonus_records:
            show_toast(self, "没有有效的加分项（成员为空或加分为0）", "warning")
            return

        # 弹出确认对话框
        self.show_bonus_confirm_dialog(bonus_records)

    def show_bonus_confirm_dialog(self, bonus_records):
        """显示确认加分明细的对话框"""
        _c = get_theme_colors()
        dialog = QDialog(self)
        dialog.setWindowTitle("确认班干加分")
        dialog.resize(600, 400)
        dialog.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        info = QLabel(f"将为以下 {len(bonus_records)} 名学生添加班干加分，请确认：")
        info.setStyleSheet(f"color: {_c['fg']}; font-size: 15px;")
        layout.addWidget(info)

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["学号", "姓名", "加分", "类别"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setRowCount(len(bonus_records))
        for i, (sid, name, bonus, category) in enumerate(bonus_records):
            table.setItem(i, 0, QTableWidgetItem(sid))
            table.setItem(i, 1, QTableWidgetItem(name))
            table.setItem(i, 2, QTableWidgetItem(str(bonus)))
            table.setItem(i, 3, QTableWidgetItem(category))
            table.verticalHeader().setVisible(False)
        layout.addWidget(table)

        total_people = len(bonus_records)
        total_score = sum(b for _, _, b, _ in bonus_records)
        stats = QLabel(f"共 {total_people} 人，总加分 {total_score} 分")
        stats.setStyleSheet(f"color: {_c['fg_dim']}; font-size: 14px;")
        layout.addWidget(stats)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(dialog.reject)
        ok_btn = QPushButton("✅ 确认提交")
        ok_btn.clicked.connect(dialog.accept)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.commit_bonus(bonus_records)

    def commit_bonus(self, bonus_records):
        """将加分记录写入积分系统"""
        date_str = datetime.now().strftime("%Y-%m-%d")
        success_count = 0
        for sid, name, bonus, category in bonus_records:
            note = f"班干加分（{category}）"
            ok, msg = save_score_record(
                stu_id=sid,
                name=name,
                date_str=date_str,
                op="+",
                category="班干加分",
                score=bonus,
                note=note,
                assign_score=None
            )
            if ok:
                success_count += 1
            else:
                print(f"加分失败：{sid} {name}，{msg}")  # 可选日志

        show_toast(self, f"成功为 {success_count} 名学生添加班干加分", "success")
        # 刷新排行榜（如果主窗口有刷新方法）
        if hasattr(self.app, 'refresh_ranking'):
            self.app.refresh_ranking()

    # ---------- 条目操作 ----------
    def add_item(self):
        if self.current_index < 0:
            show_toast(self, "请先选择一个名单", "warning")
            return
        new_item = {"category": "新类别", "members": [], "bonus": 0}
        dialog = CategoryEditDialog(self, new_item, self.app.stu_df, self.app.id_to_name)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            self.custom_lists[self.current_index]["items"].append(dialog.result)
            save_custom_lists(self.custom_lists)
            self.show_items(self.current_index)
            show_toast(self, "已添加新条目", "success")

    def delete_item(self):
        if self.current_index < 0:
            show_toast(self, "请先选择一个名单", "warning")
            return
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            show_toast(self, "请先选择要删除的条目", "warning")
            return
        items = self.custom_lists[self.current_index]["items"]
        to_remove = sorted([idx.row() for idx in selected_rows], reverse=True)
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("确认删除")
        msg_box.setText(f"确定要删除选中的 {len(to_remove)} 个条目吗？")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)
        msg_box.setStyleSheet(DIALOG_STYLE)
        reply = msg_box.exec()
        if reply == QMessageBox.StandardButton.Yes:
            for idx in to_remove:
                if idx < len(items):
                    del items[idx]
            save_custom_lists(self.custom_lists)
            self.show_items(self.current_index)
            show_toast(self, f"已删除 {len(to_remove)} 个条目", "success")

    def on_table_double_click(self, index):
        """双击表格行，弹出编辑对话框"""
        if self.current_index < 0:
            return
        row = index.row()
        items = self.custom_lists[self.current_index]["items"]
        if row < 0 or row >= len(items):
            return
        item = items[row].copy()
        dialog = CategoryEditDialog(self, item, self.app.stu_df, self.app.id_to_name)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            items[row] = dialog.result
            save_custom_lists(self.custom_lists)
            self.show_items(self.current_index)
            show_toast(self, "条目已更新", "success")

# ======================== 关于页面 ========================

class AboutInterface(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        _c = get_theme_colors()                 # ★ 新增
        self.app = parent
        self.setObjectName("aboutInterface")
        self.click_count = 0
        self.last_click_time = time.time()
        self._restore_timer = None

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(40, 40, 40, 40)
        main_layout.setSpacing(15)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title_label = QLabel("📖 关于本系统", self)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # ★ 修复：原字符串少了 f 前缀
        title_label.setStyleSheet(
            f"font-family: 'Segoe UI'; font-size: 32px; "
            f"font-weight: bold; color: {_c['fg']};"
        )
        main_layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignHCenter)

        info_frame = QWidget(self)
        info_layout = QVBoxLayout(info_frame)
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(4)

        # ★ 修改：color 缺省时使用主题前景色，避免浅色主题下白字看不见
        def add_line(text, font_size=24, color=None, bold=False, clickable=False):
            if color is None:
                color = _c["fg"]
            label = QLabel(text, info_frame)
            label.setAlignment(Qt.AlignmentFlag.AlignLeft)
            weight = "bold" if bold else "normal"
            label.setStyleSheet(
                f"font-family: 'Microsoft YaHei'; font-size: {font_size}px; "
                f"font-weight: {weight}; color: {color};"
            )
            if clickable:
                label.setCursor(Qt.CursorShape.PointingHandCursor)
            info_layout.addWidget(label)
            return label

        def add_empty_line(height=14):
            empty = QLabel("", info_frame)
            empty.setFixedHeight(height)
            info_layout.addWidget(empty)

        add_line("【系统名称】：PickOS 班级管理系统")
        add_empty_line(14)

        self.version_label = add_line(
            f"【版本】：{SYSTEM_VERSION}",
            color=_c["fg_bright"],
            bold=True,
            clickable=True
        )
        self.version_label.mousePressEvent = self.on_version_click
        self.default_text_color = _c["fg_bright"]      # ★ 原为 "#FFFFFF"

        add_empty_line(14)
        add_line("【系统状态】：运行中")
        add_empty_line(14)
        add_line("Copyright © 2026 wuweijia-star  ·  GPLv3")
        add_empty_line(14)
        add_line("                            中国制造")
        add_line("                      MADE IN CHINA")

        main_layout.addWidget(info_frame, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.time_label = QLabel("", self)
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_label.setStyleSheet(
            "font-family: 'Microsoft YaHei'; font-size: 22px; color: #4A90E2;"
        )
        main_layout.addWidget(self.time_label, alignment=Qt.AlignmentFlag.AlignHCenter)

        back_btn = PushButton("🔙 返回", self)
        back_btn.clicked.connect(lambda: self.app.switchTo(self.app.homeInterface))
        main_layout.addWidget(back_btn, alignment=Qt.AlignmentFlag.AlignHCenter)

        self.update_time()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.destroyed.connect(self.on_destroy)

    def on_version_click(self, event):
        now = time.time()
        if now - self.last_click_time > 1.0:
            self.click_count = 0
        self.click_count += 1
        self.last_click_time = now

        # 高亮蓝色（保持原有的彩蛋交互）
        self.version_label.setStyleSheet(
            "font-family: 'Microsoft YaHei'; font-size: 24px; "
            "color: #00BFFF; font-weight: bold;"
        )
        if self._restore_timer is not None:
            self._restore_timer.stop()
        self._restore_timer = QTimer(self)
        self._restore_timer.setSingleShot(True)
        self._restore_timer.timeout.connect(self.restore_version_color)
        self._restore_timer.start(200)

        if self.click_count >= 5:
            self.click_count = 0
            self.app.show_heart_animation()

    def restore_version_color(self):
        # ★ 修复：使用主题前景色，不再固定为白色
        self.version_label.setStyleSheet(
            f"font-family: 'Microsoft YaHei'; font-size: 24px; "
            f"color: {self.default_text_color}; font-weight: bold;"
        )
        self._restore_timer = None

    # update_time / on_destroy 原样保留

    def update_time(self):
        cfg = load_config()
        device_install_ts = cfg.get("install_time", time.time())

        elapsed = time.time() - self.app.start_time
        h1 = int(elapsed // 3600)
        m1 = int((elapsed % 3600) // 60)
        s1 = int(elapsed % 60)

        total_system = time.time() - INSTALL_TIMESTAMP
        d_sys = int(total_system // 86400)
        h_sys = int((total_system % 86400) // 3600)
        m_sys = int((total_system % 3600) // 60)
        s_sys = int(total_system % 60)

        total_device = time.time() - device_install_ts
        d_dev = int(total_device // 86400)
        h_dev = int((total_device % 86400) // 3600)
        m_dev = int((total_device % 3600) // 60)
        s_dev = int(total_device % 60)

        text = (
            f"本次运行时间：{h1}小时{m1}分{s1}秒\n"
            f"本设备运行总时间：{d_dev}天{h_dev}小时{m_dev}分{s_dev}秒\n"
            f"PickOS创立时间：{d_sys}天{h_sys}小时{m_sys}分{s_sys}秒"
        )
        self.time_label.setText(text)

    def on_destroy(self):
        if self._restore_timer is not None:
            self._restore_timer.stop()
            self._restore_timer = None
        if self.timer is not None:
            self.timer.stop()

# ======================== 新手引导（新拟物 + Hello 动画） ========================
# 说明：本模块仅首次启动时显示，引导完成会写 config["onboarding_completed"]=True

class _HelloWidget(QWidget):
    """苹果风格 Hello 文字动画：逐字从下方滑入 + 淡入，单色（符合新拟物）"""
    COLOR = "#5a6b8c"     # 与 _NeuColors.TEXT 一致，柔和蓝灰
    CHARS = ["H", "e", "l", "l", "o"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background:transparent;")
        self._progresses = [0.0] * 5
        self._anims = []
        self._font = QFont("Segoe UI")
        self._font.setPixelSize(200)
        self._font.setWeight(QFont.Weight.Bold)
        self._finished_callback = None

    def start_animation(self, on_finished=None):
        self._progresses = [0.0] * 5
        self._finished_callback = on_finished
        self._anims.clear()
        self.update()

        for i in range(5):
            anim = QVariantAnimation(self)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setDuration(650)
            anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            anim.valueChanged.connect(lambda v, idx=i: self._set_progress(idx, v))
            QTimer.singleShot(i * 150, anim.start)
            self._anims.append(anim)

        total_ms = 4 * 150 + 650 + 300
        QTimer.singleShot(total_ms, self._on_anim_done)

    def _on_anim_done(self):
        if self._finished_callback:
            self._finished_callback()

    def _set_progress(self, idx, v):
        self._progresses[idx] = float(v)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setFont(self._font)

        fm = painter.fontMetrics()
        char_widths = [fm.horizontalAdvance(c) for c in self.CHARS]
        total_width = sum(char_widths)

        x = (self.width() - total_width) / 2
        baseline_y = (self.height() + fm.ascent() - fm.descent()) / 2

        base_color = QColor(self.COLOR)
        for i, char in enumerate(self.CHARS):
            progress = self._progresses[i]
            if progress <= 0:
                x += char_widths[i]
                continue
            c = QColor(base_color)
            c.setAlpha(int(255 * min(1.0, progress)))
            painter.setPen(c)
            offset_y = (1 - progress) * 45
            painter.drawText(QPointF(x, baseline_y + offset_y), char)
            x += char_widths[i]


def _make_round_neu_button(text, size=180, font_size=28):
    """快速创建一个 180×180 的圆形新拟物按钮"""
    btn = _NeuRoundButton(text)
    # ★ 先设置参数，再 setFixedSize：resizeEvent 会用新参数算 radius
    btn._blur = 8
    btn._distance = 4
    btn._font.setPointSize(font_size)
    btn.setFixedSize(size, size)   # → radius = (180 - 2*(16+4+4))/2 = 66 ✔ 真圆
    btn._cache = None
    return btn


def _fade_in_widget(widget, duration=600):
    """让 widget 淡入（不需要时请保留引用，避免动画被 GC）"""
    eff = QGraphicsOpacityEffect(widget)
    widget.setGraphicsEffect(eff)
    eff.setOpacity(0.0)
    anim = QPropertyAnimation(eff, b"opacity")
    anim.setStartValue(0.0)
    anim.setEndValue(1.0)
    anim.setDuration(duration)
    anim.setEasingCurve(QEasingCurve.Type.OutCubic)
    anim.start()
    widget._fade_anim = anim
    return anim


class _HelloStage(QWidget):
    def __init__(self, on_next):
        super().__init__()
        self.setStyleSheet("background:transparent;")
        self.on_next = on_next
        self._btn_anim = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 20, 0, 20)
        layout.setSpacing(20)

        # 上方弹性 → Hello → 弹性，保证 Hello 永远垂直居中
        layout.addStretch(1)
        self.hello_widget = _HelloWidget()
        self.hello_widget.setFixedSize(900, 350)
        layout.addWidget(self.hello_widget, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch(1)

        # ★ 关键：按钮放在固定高度容器里。
        #   按钮 setVisible(False) ↔ True 时容器高度不变 → Hello 位置不动。
        btn_container = QWidget()
        btn_container.setFixedHeight(200)
        btn_container.setStyleSheet("background:transparent;")
        btn_inner = QHBoxLayout(btn_container)
        btn_inner.setContentsMargins(0, 0, 0, 0)
        btn_inner.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.next_btn = _make_round_neu_button("→", 180, 42)
        self.next_btn.clicked.connect(self.on_next)
        self.next_btn.setVisible(False)
        btn_inner.addWidget(self.next_btn)

        layout.addWidget(btn_container)

    def start(self):
        self.next_btn.setVisible(False)
        if self.next_btn.graphicsEffect():
            self.next_btn.setGraphicsEffect(None)
        self.hello_widget.start_animation(on_finished=self._show_button)

    def _show_button(self):
        self.next_btn.setVisible(True)
        self._btn_anim = _fade_in_widget(self.next_btn, 600)


class _ShortcutStage(QWidget):
    def __init__(self, on_yes, on_no):
        super().__init__()
        self.setStyleSheet("background:transparent;")
        self.on_yes = on_yes
        self.on_no = on_no

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 40, 0, 40)
        layout.setSpacing(20)

        # 上方：凸起的新拟物圆角卡片，文字居中
        self.tip_card = _NeuTipCard("是否创建桌面快捷方式？\n""（仅支持Windows）", font_size=18)
        self.tip_card.setFixedSize(620, 160)

        layout.addStretch(1)
        layout.addWidget(self.tip_card, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch(1)

        # 下方：两个按钮水平居中
        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        btn_row.setSpacing(60)

        yes_btn = _make_round_neu_button("创建", 180, 28)
        yes_btn.clicked.connect(self.on_yes)
        btn_row.addWidget(yes_btn)

        no_btn = _make_round_neu_button("稍后", 180, 28)
        no_btn.clicked.connect(self.on_no)
        btn_row.addWidget(no_btn)

        layout.addLayout(btn_row)

    def start(self):
        pass


class _ClassStage(QWidget):
    """配置名单阶段：凸起卡片包文字 + 📂 主按钮 + 稍后配置链接"""

    def __init__(self, on_pick, on_skip):
        super().__init__()
        self.setStyleSheet("background:transparent;")
        self.on_pick = on_pick
        self.on_skip = on_skip

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 30, 0, 40)
        layout.setSpacing(20)

        # 上方：凸起卡片，两行文字居中
        self.tip_card = _NeuTipCard(
            "请配置班级名单文件\n"
            "Excel 格式，需包含「学号」和「姓名」两列\n"
            "（可在设置-数据里进行配置与更改）",
            font_size=14,
        )
        self.tip_card.setFixedSize(780, 200)

        layout.addStretch(1)
        layout.addWidget(self.tip_card, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch(1)

        # 下方：「配置」+「稍后」两个同尺寸圆形按钮，左右排布
        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        btn_row.setSpacing(60)

        self.pick_btn = _make_round_neu_button("配置", 180, 28)
        self.pick_btn.clicked.connect(self.on_pick)
        btn_row.addWidget(self.pick_btn)

        self.skip_btn = _make_round_neu_button("稍后", 180, 28)
        self.skip_btn.clicked.connect(self.on_skip)
        btn_row.addWidget(self.skip_btn)

        layout.addLayout(btn_row)

    def start(self):
        # 整体淡入由 OnboardingWindow._switch_to 完成
        pass


class _PasswordStage(QWidget):
    def __init__(self, on_yes, on_no):
        super().__init__()
        self.setStyleSheet("background:transparent;")
        self.on_yes = on_yes
        self.on_no = on_no

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 40, 0, 40)
        layout.setSpacing(20)

        # 上方：凸起卡片
        self.tip_card = _NeuTipCard(
            "是否设置启动密码？\n（保护您的数据安全）",
            font_size=16,
        )
        self.tip_card.setFixedSize(660, 200)

        layout.addStretch(1)
        layout.addWidget(self.tip_card, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch(1)

        # 下方：两个按钮
        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignmentFlag.AlignCenter)
        btn_row.setSpacing(60)

        yes_btn = _make_round_neu_button("是", 180, 28)
        yes_btn.clicked.connect(self.on_yes)
        btn_row.addWidget(yes_btn)

        no_btn = _make_round_neu_button("否", 180, 28)
        no_btn.clicked.connect(self.on_no)
        btn_row.addWidget(no_btn)

        layout.addLayout(btn_row)

    def start(self):
        pass


class OnboardingWindow(QWidget):
    """新手引导窗口（1280×720，新拟物风格，苹果 Hello 动画）"""

    def __init__(self, main_window, on_finish=None):
        super().__init__(None)
        self.main_window = main_window
        self._on_finish_callback = on_finish
        self._closing = False
        self._active_anims = []
        self._current_index = -1

        self.setWindowTitle("欢迎使用 PickOS")
        self.setFixedSize(1280, 720)
        self.setWindowFlags(
            Qt.WindowType.Window
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.FramelessWindowHint
        )
        self.setObjectName("onboardingWindow")
        self.setStyleSheet("#onboardingWindow{background-color:#e0e5ec;}")

        screen = QApplication.primaryScreen().availableGeometry()
        self.move(screen.center() - self.rect().center())

        self.container = QWidget(self)
        self.container.setGeometry(0, 0, 1280, 720)
        self.container.setStyleSheet("background-color:#e0e5ec;")

        self.hello_stage = _HelloStage(self._on_hello_next)
        self.shortcut_stage = _ShortcutStage(self._on_shortcut_yes, self._on_shortcut_no)
        self.class_stage = _ClassStage(self._on_class_pick, self._on_class_skip)
        self.password_stage = _PasswordStage(self._on_password_yes, self._on_password_no)

        self.stages = [
            self.hello_stage,
            self.shortcut_stage,
            self.class_stage,
            self.password_stage,
        ]

        for w in self.stages:
            w.setParent(self.container)
            w.setGeometry(0, 0, 1280, 720)
            w.setVisible(False)

        QTimer.singleShot(100, lambda: self._switch_to(0, animate=False))

    # ---------- 阶段切换 ----------
    def _switch_to(self, index, animate=True):
        if index == self._current_index:
            return

        old_index = self._current_index
        self._current_index = index
        new_widget = self.stages[index]
        new_widget.setVisible(True)
        new_widget.raise_()

        if old_index >= 0:
            old_widget = self.stages[old_index]
            if animate:
                old_eff = QGraphicsOpacityEffect(old_widget)
                old_widget.setGraphicsEffect(old_eff)
                old_eff.setOpacity(1.0)
                out_anim = QPropertyAnimation(old_eff, b"opacity")
                out_anim.setStartValue(1.0)
                out_anim.setEndValue(0.0)
                out_anim.setDuration(300)
                out_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
                out_anim.finished.connect(
                    lambda w=old_widget: (w.setVisible(False), w.setGraphicsEffect(None))
                )
                out_anim.start()
                self._active_anims.append(out_anim)
            else:
                old_widget.setVisible(False)
                old_widget.setGraphicsEffect(None)

        if animate:
            new_eff = QGraphicsOpacityEffect(new_widget)
            new_widget.setGraphicsEffect(new_eff)
            new_eff.setOpacity(0.0)
            in_anim = QPropertyAnimation(new_eff, b"opacity")
            in_anim.setStartValue(0.0)
            in_anim.setEndValue(1.0)
            in_anim.setDuration(400)
            in_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            in_anim.finished.connect(
                lambda w=new_widget: w.setGraphicsEffect(None)
            )
            in_anim.finished.connect(lambda: self._start_stage(index))
            in_anim.start()
            self._active_anims.append(in_anim)
        else:
            new_widget.setGraphicsEffect(None)
            self._start_stage(index)

    def _start_stage(self, index):
        stage = self.stages[index]
        if hasattr(stage, "start"):
            stage.start()

    # ---------- 各阶段回调 ----------
    def _on_hello_next(self):
        self._switch_to(1)

    def _on_shortcut_yes(self):
        self._create_desktop_shortcut()
        self._switch_to(2)

    def _on_shortcut_no(self):
        self._switch_to(2)

    def _on_class_pick(self):
        """选择班级名单：成功才进入下一阶段；失败 / 取消都留在当前阶段。"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择班级名单 Excel 文件", "",
            "Excel 文件 (*.xlsx *.xls);;所有文件 (*)"
        )
        if not file_path:
            # 用户主动取消 → 留在当前阶段（不静默放行）
            return

        try:
            if not self.main_window._save_as_class_list(file_path):
                # 失败：弹 QMessageBox 明确告知，不放行
                QMessageBox.warning(
                    self, "名单文件不符合要求",
                    "请确认 Excel 中包含「学号」和「姓名」两列。\n\n"
                    "请重新选择正确文件，或点击下方「稍后配置」跳过。"
                )
                return
            if not self.main_window.load_student_data(CLASS_LIST_FILE):
                QMessageBox.warning(
                    self, "名单加载失败",
                    "名单加载失败，请重新选择。\n\n"
                    "或点击下方「稍后配置」跳过。"
                )
                return
        except Exception as e:
            QMessageBox.warning(
                self, "处理名单失败",
                f"{e}\n\n请重新选择，或点击下方「稍后配置」跳过。"
            )
            return

        # 成功 → 进入密码阶段
        self._switch_to(3)

    def _on_class_skip(self):
        """用户明确点击「稍后配置」→ 直接进入密码阶段。"""
        self._switch_to(3)

    def _on_password_yes(self):
        self._open_password_settings()
        self._finish()

    def _on_password_no(self):
        self._finish()

    # ---------- 密码设置（调用前面的更改密码界面） ----------
    def _open_password_settings(self):
        try:
            dlg = _ChangePasswordDialog(self)
            dlg.exec()
            cfg = load_config()
            if cfg.get("password_hash"):
                cfg["password_enabled"] = True
                save_config(cfg)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"设置密码失败：{e}")

    # ---------- 桌面快捷方式（仅 Windows） ----------
    def _create_desktop_shortcut(self):
        import sys
        import os
        import traceback

        if not sys.platform.startswith("win"):
            QMessageBox.information(self, "提示", "当前平台暂不支持自动创建桌面快捷方式，\n请手动创建。")
            return

        try:
            # 1. 获取桌面路径
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            if not os.path.exists(desktop):
                desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
            if not os.path.exists(desktop):
                QMessageBox.warning(self, "提示", "未找到桌面目录")
                return

            target_path = os.path.join(desktop, "PickOS.lnk")

            # ★ 新增：检查是否已经存在（安装包可能已经创建过了）
            if os.path.exists(target_path):
                show_toast(self, "桌面上已存在 PickOS 快捷方式，无需重复创建", "info")
                return  # 直接结束，不往下覆盖

            # 2. 动态区分打包与源码环境
            if getattr(sys, "frozen", False):
                exe_path = sys.executable
                work_dir = os.path.dirname(exe_path)
                arguments = ""
            else:
                exe_path = sys.executable
                work_dir = os.path.dirname(os.path.abspath(__file__))
                script_path = os.path.abspath(__file__)
                arguments = f'"{script_path}"'

            # 3. 使用 win32com 原生创建快捷方式
            try:
                import win32com.client
            except ImportError:
                QMessageBox.warning(self, "提示", "缺少依赖库 pywin32，请先执行 pip install pywin32\n或手动创建快捷方式。")
                return

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(target_path)
            shortcut.TargetPath = exe_path
            shortcut.Arguments = arguments
            shortcut.WorkingDirectory = work_dir
            shortcut.IconLocation = f"{exe_path},0"
            shortcut.Description = "PickOS 班级管理系统"
            shortcut.Save()

            self._show_shortcut_success_tip()

        except Exception as e:
            traceback.print_exc()
            QMessageBox.warning(
                self,
                "提示",
                f"创建快捷方式失败：{e}\n请手动创建。"
            )

    def _show_shortcut_success_tip(self):
        """创建快捷方式成功后的新拟物风格提示框（与修改密码成功弹窗一致）"""
        dlg = QDialog(self)
        dlg.setWindowTitle("创建成功")
        dlg.setFixedSize(460, 260)
        dlg.setObjectName("shortcutSuccDlg")
        dlg.setStyleSheet("#shortcutSuccDlg{background-color:#e0e5ec;}")

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(30, 30, 30, 25)
        lay.setSpacing(20)

        lbl = QLabel("✅ 桌面快捷方式已创建")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(
            "color:#4CAF50;font-size:20pt;font-weight:700;background:transparent;"
        )
        lay.addWidget(lbl)

        lay.addStretch()

        # 与修改密码成功弹窗相同的按钮尺寸和参数
        ok_btn = _NeuButton("确定")
        ok_btn.setFixedSize(220, 100)
        ok_btn._radius = 24.0
        ok_btn._blur = 10
        ok_btn._distance = 4
        ok_btn._cache = None
        ok_btn._font.setPointSize(18)
        ok_btn.clicked.connect(dlg.accept)
        lay.addWidget(ok_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        dlg.exec()

    # ---------- 结束引导 ----------
    def _finish(self):
        if self._closing:
            return
        self._closing = True

        cfg = load_config()
        cfg["onboarding_completed"] = True
        save_config(cfg)

        eff = QGraphicsOpacityEffect(self.container)
        self.container.setGraphicsEffect(eff)
        eff.setOpacity(1.0)
        anim = QPropertyAnimation(eff, b"opacity")
        anim.setStartValue(1.0)
        anim.setEndValue(0.0)
        anim.setDuration(500)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.finished.connect(self._do_close)
        anim.start()
        self._close_anim = anim

    def _do_close(self):
        self.close()
        if self._on_finish_callback:
            self._on_finish_callback()

    def closeEvent(self, event):
        # 强制关闭（Alt+F4）也标记完成，避免下次重复引导
        if not self._closing:
            self._closing = True
            cfg = load_config()
            cfg["onboarding_completed"] = True
            save_config(cfg)
            if self._on_finish_callback:
                self._on_finish_callback()
        super().closeEvent(event)

# ======================== 程序入口 ========================
if __name__ == "__main__":
    import sys
    from PyQt6.QtWidgets import QApplication

    print("🚀 [1] PickOS 进入 main")
    print(f"📋 Python: {sys.version}")
    print(f"📋 CWD: {os.getcwd()}")

    # 检测打包环境变量，若存在则直接退出，避免弹出 GUI
    if os.environ.get("PACKAGING_MODE") == "1":
        print("检测到 PyInstaller 打包环境，跳过 GUI 启动。")
        sys.exit(0)
    # 确保上述 import 已在文件开头，此处只展示核心逻辑
    # 创建必要的用户文件夹
    for folder in [USER_DIR, DATA_DIR, EXPORT_DIR, BACKUP_DIR]:
        os.makedirs(folder, exist_ok=True)
    app = QApplication(sys.argv)
    print("🚀 [2] QApplication 已创建")

    # ===== 启动密码验证（若已开启密码保护） =====
    def _check_password_gate():
        cfg = load_config()
        if not cfg.get("password_enabled", False):
            print("🔓 密码保护未启用")
            return True
        if not cfg.get("password_hash", ""):
            print("🔓 密码未设置，跳过验证")
            return True

        # ---- 智能解锁：先检测开发者 U 盘，再检测已绑定的普通磁盘 ----
        try:
            # 1) 开发者 U 盘：不受智能解锁开关影响
            dev = find_developer_usb()
            if dev is not None:
                print("开发者智能解锁成功：最高管理员权限")
                show_smart_unlock_toast(dev, is_developer=True)
                return True

            # 2) 普通智能解锁：要求开关已开启且磁盘已绑定
            bound = find_bound_smart_unlock_device()
            if bound is not None:
                print("智能解锁成功：检测到已绑定的移动磁盘")
                show_smart_unlock_toast(bound, is_developer=False)
                return True
        except Exception as e:
            print(f"⚠️ 智能解锁检测失败：{e}")

        try:
            print("🔒 正在弹出密码验证框...")
            dlg = PasswordPromptDialog()
            ok = dlg.exec() == QDialog.DialogCode.Accepted
            print(f"🔒 密码验证结果：{ok}")
            return ok
        except Exception as e:
            import traceback
            print("❌ 密码验证框异常：")
            traceback.print_exc()
            return True


    # ===== 应用主题到全局对话框样式 DIALOG_STYLE =====
    globals()["DIALOG_STYLE"] = _replace_theme_colors(DIALOG_STYLE)


    # ---------- 新增：全局样式（使 QInputDialog 背景为黑色） ----------
    app.setStyleSheet(_replace_theme_colors("""
        QInputDialog {
            background: #16161E;
            border-radius: 18px;
        }
        QInputDialog QLabel {
            color: #E0E0E0;
            font-size: 15px;
        }
        QInputDialog QLineEdit {
            background: #1A1A24;
            color: #F0F0F0;
            border: 1px solid #3A3A4A;
            border-radius: 8px;
            padding: 8px 12px;
            selection-background-color: #64B5F6;
        }
        QInputDialog QLineEdit:focus {
            border: 2px solid #64B5F6;
            background: #1E1E2A;
        }
        QInputDialog QPushButton {
            border-radius: 10px;
            padding: 8px 20px;
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #64B5F6, stop:1 #42A5F5);
            color: white;
            font-weight: 600;
            border: 1px solid rgba(100, 181, 246, 0.3);
            font-size: 15px;
            min-width: 80px;
        }
        QInputDialog QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
            border-color: #64B5F6;
        }
        QInputDialog QPushButton:pressed {
            background: #1E88E5;
        }
        QInputDialog QPushButton:default {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #42A5F5, stop:1 #1E88E5);
        }
    """))
    # ---------- 新增结束 ---------

    from PyQt6.QtGui import QPalette, QColor

    def apply_theme_palette(app):
        c = get_theme_colors()
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window,          QColor(c["bg"]))
        palette.setColor(QPalette.ColorRole.WindowText,      QColor(c["fg"]))
        palette.setColor(QPalette.ColorRole.Base,            QColor(c["bg_alt"]))
        palette.setColor(QPalette.ColorRole.Text,            QColor(c["fg"]))
        palette.setColor(QPalette.ColorRole.Button,          QColor(c["bg_card"]))
        palette.setColor(QPalette.ColorRole.ButtonText,      QColor(c["fg"]))
        palette.setColor(QPalette.ColorRole.Highlight,       QColor(100, 181, 246))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.ToolTipBase,     QColor(c["bg_card"]))
        palette.setColor(QPalette.ColorRole.ToolTipText,     QColor(c["fg"]))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor(c["fg_dim"]))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text,       QColor(c["fg_dim"]))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor(c["fg_dim"]))
        app.setPalette(palette)

    # 调用函数
    apply_theme_palette(app)

    # ---------- 安装弹窗粒子事件过滤器 ----------
    particle_filter = ParticleEventFilter(app)
    app.installEventFilter(particle_filter)
    
    setThemeColor("#64B5F6")
    setTheme(get_qfluent_theme())

    cfg = load_config()
    show_boot = cfg.get("show_boot_animation", True)

    # 全局变量用于保存窗口引用，避免被垃圾回收
    main_window = None
    onboarding_window = None   # ← 新增

    def fade_in_main_window(win):
        """让主窗口从透明度 0 平滑淡入到 1，同时触发粒子效果"""
        win.setWindowOpacity(0.0)
        win.show()

        def _center_and_resize():
            try:
                # ★ 用窗口当前所在屏幕（多显示器友好），而不是 primaryScreen
                scr = win.screen() or QApplication.primaryScreen()
                if scr is None:
                    return
                avail = scr.availableGeometry()

                # ★ 目标尺寸不超过屏幕可用尺寸，否则 WM 会强制压缩 / 移出屏幕
                target_w = min(1280, avail.width())
                target_h = min(720, avail.height())
                win.resize(target_w, target_h)

                # ★ 用窗口"实际尺寸"居中（不是常量 1280×720）
                x = avail.x() + (avail.width() - win.width()) // 2
                y = avail.y() + (avail.height() - win.height()) // 2
                win.move(x, y)
            except Exception as e:
                print(f"⚠️ 主窗口重新定位失败: {e}")

        # show 之后延迟一步：让 WM 先完成窗口创建/接管
        QTimer.singleShot(30, _center_and_resize)
        # 保险起见，稍晚一点再确认一次（部分系统/主题动画会二次调整窗口）
        QTimer.singleShot(120, _center_and_resize)

        # 主窗口出现时触发粒子
        QTimer.singleShot(60, lambda: _spawn_close_particles(win))

        anim = QPropertyAnimation(win, b"windowOpacity", win)
        anim.setDuration(700)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        anim.start()
        win._fade_in_anim = anim

    def preload_main_window_silently():
        """在开机动画版权页期间静默预加载主窗口（不显示）"""
        global main_window
        if main_window is not None:
            return
        try:
            main_window = MainWindow()
            main_window.hide()      # 确保不会提前显示
        except Exception as e:
            import traceback
            print(f"⚠️ 主窗口预加载失败：{e}")
            traceback.print_exc()
            main_window = None

    
    def on_boot_finished():
        """开机动画结束回调：密码验证 → 新手引导（首次）→ 显示主窗口"""
        global main_window, onboarding_window

        # ---- 主窗口构造 ----
        if main_window is None:
            try:
                main_window = MainWindow()
                main_window.hide()
            except Exception as e:
                import traceback
                print("=" * 60)
                print("❌ 主窗口构造失败，异常堆栈：")
                traceback.print_exc()
                print("=" * 60)
                QApplication.quit()
                return

        # ---- 密码验证 ----
        try:
            if not _check_password_gate():
                QApplication.quit()
                return
        except Exception as e:
            import traceback
            print("=" * 60)
            print("❌ 密码验证失败，异常堆栈：")
            traceback.print_exc()
            print("=" * 60)
            print("⚠️ 已跳过密码验证，继续启动")

        # ---- 新手引导 / 直接显示主窗口 ----
        cfg = load_config()
        print(f"📋 启动状态：onboarding_completed = {cfg.get('onboarding_completed', False)}")

        def _safe_post_boot():
            try:
                main_window.start_post_boot_tasks()
            except Exception as e:
                import traceback
                print(f"⚠️ 启动后置任务失败：{e}")
                traceback.print_exc()

        def _launch_main_window():
            """主窗口淡入 + 主窗口真正出现后再跑启动后置任务"""
            fade_in_main_window(main_window)

            # 延迟 450ms 再跑后置任务：
            #   · 淡入动画 700ms，450ms 时窗口已基本可见（opacity ≈ 0.7~0.9）
            #   · 避免"窗口还是半透明时，通知已经完全不透明地飘出来"的错位感
            #   · 后置任务本身有 _post_boot_started 去重，多次调用安全
            QTimer.singleShot(450, _safe_post_boot)

        if not cfg.get("onboarding_completed", False):
            def _on_onboarding_finish():
                global onboarding_window
                onboarding_window = None
                try:
                    if hasattr(main_window, "inputInterface"):
                        main_window.inputInterface.update_student_list()
                except Exception:
                    pass
                try:
                    if hasattr(main_window, "refresh_ranking"):
                        main_window.refresh_ranking()
                except Exception:
                    pass
                # ★ 新手引导完成 → 显示主窗口 → 再跑后置任务
                _launch_main_window()

            try:
                print("🎬 正在创建新手引导窗口...")
                onboarding_window = OnboardingWindow(
                    main_window, on_finish=_on_onboarding_finish
                )
                onboarding_window.show()
                onboarding_window.raise_()
                onboarding_window.activateWindow()
                print(f"✅ 新手引导窗口已显示，isVisible={onboarding_window.isVisible()}")
            except Exception as e:
                import traceback
                print("=" * 60)
                print("❌ 新手引导窗口失败，异常堆栈：")
                traceback.print_exc()
                print("=" * 60)
                print("⚠️ 直接进入主界面")
                onboarding_window = None
                # 兜底：标记引导已完成，避免下次再卡在这
                cfg["onboarding_completed"] = True
                save_config(cfg)
                # ★ 引导失败兜底 → 显示主窗口 → 再跑后置任务
                _launch_main_window()
        else:
            print("🚀 常规启动，直接显示主窗口")
            # ★ 常规启动 → 显示主窗口 → 再跑后置任务
            _launch_main_window()

    # ======================== 启动流程 ========================
    if show_boot:
        print("🎬 [3] 即将创建 BootSplash")
        splash = BootSplash(on_boot_finished, enable_animation=True)
        print(f"🎬 [4] BootSplash 已创建，isVisible={splash.isVisible()}, geometry={splash.geometry()}")
        QTimer.singleShot(50, preload_main_window_silently)
    else:
        print("🎬 [3] 无开机动画，直接进入 on_boot_finished")
        on_boot_finished()

    sys.exit(app.exec())
