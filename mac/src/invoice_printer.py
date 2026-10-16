# -*- coding: utf-8 -*-
"""
发票打印助手
============
把报销 Excel 里嵌入的发票 / 采购凭证图片自动提取出来, 按订单分组,
勾选后自动排版并直接发送到打印机, 节省纸张。

排版模式: 单面 / 双面  ×  整页(每页1张) / 二分之一(每页2张) / 四分之一(每页4张)
双面打印为长边翻页(书本式), 需要打印机支持自动双面。

用法:  python invoice_printer.py
"""

import io
import ctypes
import json
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
import posixpath
import threading
import traceback
from dataclasses import dataclass
from xml.etree import ElementTree as ET

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw, ImageFont, ImageTk

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pypdfium2
    HAS_PDFIUM = True
except ImportError:
    HAS_PDFIUM = False

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()                       # HEIC/HEIF(iPhone 照片)
    try:                                         # 旧版 Pillow 无原生 AVIF 时补上
        from pillow_heif import register_avif_opener
        register_avif_opener()
    except Exception:
        pass
    HAS_HEIF = True
except Exception:
    HAS_HEIF = False

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

DND_LOG_FILE = os.path.join(tempfile.gettempdir(), '发票打印助手_拖放日志.txt')

# 平台标识: Windows 全功能; macOS 走 CUPS 打印, OCR/剪贴板/COM 部分能力受限
IS_MAC = (sys.platform == 'darwin')
IS_WIN = (os.name == 'nt')


def _dnd_log(msg):
    """拖放诊断日志: 记录注册/收包/解析全过程, 拖放无反应时供远程排查"""
    try:
        with open(DND_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(time.strftime('[%m-%d %H:%M:%S] ') + msg + '\n')
    except Exception:
        pass


def _clipboard_files():
    """读剪贴板中的文件列表(CF_HDROP), 供 Ctrl+V 导入"""
    out = []
    if IS_MAC:
        return out        # mac 版暂不支持文件粘贴(拖入可用)
    try:
        u32, s32 = ctypes.windll.user32, ctypes.windll.shell32
        CF_HDROP = 15
        # 64 位句柄必须显式声明原型: 默认 c_int 会把大句柄符号扩展成非法指针
        u32.IsClipboardFormatAvailable.argtypes = [ctypes.c_uint]
        u32.OpenClipboard.argtypes = [ctypes.c_void_p]
        u32.GetClipboardData.argtypes = [ctypes.c_uint]
        u32.GetClipboardData.restype = ctypes.c_void_p
        s32.DragQueryFileW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                       ctypes.c_wchar_p, ctypes.c_uint]
        if not u32.IsClipboardFormatAvailable(CF_HDROP):
            return out
        opened = False
        for _ in range(4):            # 剪贴板可能被其他程序短暂占用, 重试
            if u32.OpenClipboard(None):
                opened = True
                break
            time.sleep(0.03)
        if not opened:
            return out
        try:
            h = u32.GetClipboardData(CF_HDROP)
            if not h:
                return out
            n = s32.DragQueryFileW(h, 0xFFFFFFFF, None, 0)
            buf = ctypes.create_unicode_buffer(260)
            for i in range(n):
                if s32.DragQueryFileW(h, i, buf, 260):
                    out.append(buf.value)
        finally:
            u32.CloseClipboard()
    except Exception as ex:
        _dnd_log('剪贴板读取失败: %r' % ex)
    return out

APP_TITLE = '发票打印助手'
APP_VERSION = 'v3.8'
PRINT_DPI = 300    # 送打印机时页面渲染精度
PDF_DPI = 200      # 导出 PDF 的精度(体积/清晰度折中)
PREVIEW_DPI = 60   # 预览窗口精度

PAPERS_MM = {'A4': (210.0, 297.0), 'Letter': (215.9, 279.4)}


def app_dir():
    """程序所在目录(兼容 PyInstaller 打包后的 exe)"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def app_data_dir():
    """用户数据目录: Win=%APPDATA%\\发票打印助手, mac=~/Library/Application Support
    设置/缓存不放 exe 旁边——exe 在桌面时会往桌面掉 json 文件,
    容易被误当成病毒生成的可疑文件"""
    if IS_MAC:
        base = os.path.expanduser('~/Library/Application Support')
    else:
        base = os.environ.get('APPDATA') or app_dir()
    d = os.path.join(base, '发票打印助手')
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = app_dir()
    return d


def _migrate_old_data():
    """老版本把 settings/ocr_cache 放在 exe 旁边, 静默搬到新数据目录(一次性)"""
    src_dir = app_dir()
    if src_dir == app_data_dir():
        return
    for name in ('settings.json', 'ocr_cache.json'):
        try:
            src = os.path.join(src_dir, name)
            dst = os.path.join(app_data_dir(), name)
            if os.path.isfile(src) and not os.path.isfile(dst):
                shutil.copyfile(src, dst)
        except Exception:
            pass


_migrate_old_data()

SETTINGS_FILE = os.path.join(app_data_dir(), 'settings.json')

TEMPLATE_FILE = '备用金报销模版.xlsx'


def embedded_template_path():
    """内置备用金报销模版路径: 打包后在解包目录, 源码运行在程序目录"""
    candidates = []
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        candidates.append(os.path.join(sys._MEIPASS, TEMPLATE_FILE))
    candidates.append(os.path.join(app_dir(), TEMPLATE_FILE))
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None

# ---------------------------------------------------------------- OPC(xlsx) 解析

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_DRW = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_ETC = 'http://www.wps.cn/officeDocument/2017/etCustomData'


def q(ns, tag):
    return '{%s}%s' % (ns, tag)


def norm_path(base_dir, target):
    """把 OPC rels 的相对 Target 归一化为包内路径"""
    t = (target or '').replace('\\', '/')
    if t.startswith('/'):
        return t.lstrip('/')
    return posixpath.normpath(posixpath.join(base_dir, t))


def rels_path_for(path):
    d, b = posixpath.split(path)
    return posixpath.join(d, '_rels', b + '.rels')


def extract_cellimages(zf, names):
    """WPS“单元格内嵌图片”清单 (xl/cellimages.xml): {图片ID: 包内媒体路径}"""
    if 'xl/cellimages.xml' not in names:
        return {}
    root = ET.fromstring(zf.read('xl/cellimages.xml'))
    rels = {}
    if 'xl/_rels/cellimages.xml.rels' in names:
        rels = {r.get('Id'): r.get('Target')
                for r in ET.fromstring(zf.read('xl/_rels/cellimages.xml.rels'))}
    out = {}
    for ci in root.iter(q(NS_ETC, 'cellImage')):
        name = rid = None
        for cnv in ci.iter(q(NS_DRW, 'cNvPr')):
            name = cnv.get('name')
            break
        for blip in ci.iter(q(NS_A, 'blip')):
            rid = blip.get(q(NS_R, 'embed'))
            break
        tgt = rels.get(rid)
        if name and tgt:
            out[name] = norm_path('xl', tgt)
    return out


DISPIMG_RE = re.compile(r'DISPIMG\(\s*"([^"]+)"')


def col_letters_to_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def extract_dispimg_cells(sheet_root):
    """扫描工作表里 =DISPIMG("ID_xxx") 公式单元格, 返回 [(行, 列, 图片ID)]"""
    out = []
    for c in sheet_root.iter(q(NS_MAIN, 'c')):
        f = c.find(q(NS_MAIN, 'f'))
        if f is None or not f.text or 'DISPIMG' not in f.text:
            continue
        m = DISPIMG_RE.search(f.text)
        ref = re.match(r'([A-Z]+)(\d+)', c.get('r') or '')
        if m and ref:
            out.append((int(ref.group(2)),
                        col_letters_to_num(ref.group(1)), m.group(1)))
    return out


def extract_excel_images(path):
    """读取 xlsx, 返回嵌入图片列表 [{sheet, row, col, filename, data}, ...]
    row/col 为图片所在单元格(1 基)。支持两种存放方式:
    1. 标准浮动图片 (xl/drawings 锚点)
    2. WPS 单元格内嵌图片 (=DISPIMG 公式 + xl/cellimages.xml)"""
    out = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        wbx = ET.fromstring(zf.read('xl/workbook.xml'))
        wbrels = {r.get('Id'): r.get('Target')
                  for r in ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))}
        cellmap = extract_cellimages(zf, names)
        for sh in wbx.find(q(NS_MAIN, 'sheets')):
            sheet_name = sh.get('name')
            target = wbrels.get(sh.get(q(NS_R, 'id')))
            if not target:
                continue
            sheet_path = norm_path('xl', target)
            if sheet_path not in names:
                continue
            sroot = ET.fromstring(zf.read(sheet_path))
            order = 0

            def add(row, col, media):
                nonlocal order
                if media and media in names:
                    out.append({'sheet': sheet_name, 'row': row, 'col': col,
                                'filename': posixpath.basename(media),
                                'data': zf.read(media), 'order': order})
                    order += 1

            # 方式1: 标准浮动图片
            dnode = sroot.find(q(NS_MAIN, 'drawing'))
            if dnode is not None:
                sr_path = rels_path_for(sheet_path)
                if sr_path in names:
                    srels = {r.get('Id'): r.get('Target')
                             for r in ET.fromstring(zf.read(sr_path))}
                    drawing_path = norm_path(
                        posixpath.dirname(sheet_path),
                        srels.get(dnode.get(q(NS_R, 'id')), ''))
                    if drawing_path in names:
                        drels_path = rels_path_for(drawing_path)
                        drels = {}
                        if drels_path in names:
                            drels = {r.get('Id'): r.get('Target')
                                     for r in ET.fromstring(zf.read(drels_path))}
                        droot = ET.fromstring(zf.read(drawing_path))
                        for anchor in droot:
                            row = col = 0
                            frm = anchor.find(q(NS_DRW, 'from'))
                            if frm is not None:
                                col = int(frm.findtext(q(NS_DRW, 'col'), '0') or 0)
                                row = int(frm.findtext(q(NS_DRW, 'row'), '0') or 0)
                            seen = set()
                            for blip in anchor.iter(q(NS_A, 'blip')):
                                rid = blip.get(q(NS_R, 'embed'))
                                if not rid or rid in seen:
                                    continue
                                seen.add(rid)
                                add(row + 1, col + 1,
                                    norm_path(posixpath.dirname(drawing_path),
                                              drels.get(rid, '')))
            # 方式2: WPS 单元格内嵌图片
            if cellmap:
                for row, col, img_id in extract_dispimg_cells(sroot):
                    add(row, col, cellmap.get(img_id))
    return out


# ---------------------------------------------------------------- 订单分组

ORDER_HINT_GROUPS = (
    ('订单号', '订单编号', '订单', '单号', '合同号', '工单号', '申请单', '报销单'),
    ('编号', '序号'),        # 弱线索: 没有订单号列时用序号列区分行
)
AMOUNT_RE = re.compile(r'^[\d.,%\s¥￥元:：/-]+$')
ID_RE1 = re.compile(r'[A-Za-z]\d{3,}')          # 如 SO20260101-001
ID_RE2 = re.compile(r'(?<!\d)\d{6,}(?!\d)')      # 6 位以上连续数字


def build_merged_map(ws):
    """合并单元格内坐标 -> 左上角坐标"""
    m = {}
    for rng in ws.merged_cells.ranges:
        tl = (rng.min_row, rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                m[(r, c)] = tl
    return m


def _cell_text(ws, row, col, mmap):
    rr, cc = mmap.get((row, col), (row, col))
    v = ws.cell(row=rr, column=cc).value
    return '' if v is None else str(v).strip()


def detect_order_col(ws):
    """在前 20 行 × 40 列里找表头含“订单/单号/序号”等关键词的列
    返回 (列号, 表头行号), 找不到返回 (None, 0)"""
    lim_r = min(ws.max_row or 1, 20)
    lim_c = min(ws.max_column or 1, 40)
    for hints in ORDER_HINT_GROUPS:
        for r in range(1, lim_r + 1):
            for c in range(1, lim_c + 1):
                v = ws.cell(row=r, column=c).value
                if v is None:
                    continue
                s = str(v).strip()
                for h in hints:
                    if h in s:
                        return (c, r)
    return (None, 0)


def find_label(ws, mmap, ocol, orow, row):
    """确定某行图片的订单标签"""
    if ocol:
        # 合并单元格的订单号通常写在组首行, 从图片行向上找最近的非空值
        for r in range(row, orow, -1):
            t = _cell_text(ws, r, ocol, mmap)
            if t:
                return t
    lim_c = min(ws.max_column or 1, 40)
    fallback = ''
    for c in range(1, lim_c + 1):
        t = _cell_text(ws, row, c, mmap)
        if not t or len(t) > 60:
            continue
        if 'DISPIMG' in t or 'CELLIMAGE' in t:   # WPS 内嵌图公式串
            continue
        if AMOUNT_RE.match(t):          # 金额/日期之类不像单号
            continue
        if ID_RE1.search(t) or ID_RE2.search(t):
            return t
        if not fallback:
            fallback = t
    return fallback or None


DETAIL_HINTS = ('名称', '用途', '金额', '发票详情')


def scan_headers(ws, orow):
    """扫描表头区: 返回 ({关键词: 列号}, {列号: 表头文字})"""
    lim_r = min(ws.max_row or 1, orow if orow else 8)
    lim_c = min(ws.max_column or 1, 40)
    dcols, hmap = {}, {}
    for r in range(1, lim_r + 1):
        for c in range(1, lim_c + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            s = str(v).strip()
            if s:
                hmap[c] = s.replace('\n', ' ')[:12]   # 靠近数据的行优先
            for h in DETAIL_HINTS:
                if h in s and h not in dcols:
                    dcols[h] = c
    return dcols, hmap


def _fmt_amount(t):
    if re.fullmatch(r'\d+(\.\d+)?', t.replace(',', '')):
        return '¥' + t.replace(',', '')
    return t


def build_detail(ws, mmap, dcols, row, is_invoice=False):
    """拼该行的关键信息: 名称 / 用途 / 金额 / 发票详情。
    发票详情列内粘有发票图片时, 无论单元格文字是什么都记为“有”"""
    parts = []
    for h in DETAIL_HINTS:
        if h == '发票详情':
            if is_invoice:
                parts.append('发票详情:有')
            else:
                c = dcols.get(h)
                if c:
                    t = _cell_text(ws, row, c, mmap)
                    if t:
                        parts.append('发票详情:%s' % t.replace('\n', ' ')[:12])
            continue
        c = dcols.get(h)
        if not c:
            continue
        t = _cell_text(ws, row, c, mmap)
        if not t:
            continue
        if h == '金额':
            t = _fmt_amount(t)
        parts.append('%s:%s' % (h, t.replace('\n', ' ')[:28]))
    return ' / '.join(parts)


def _detail_amount(detail):
    m = re.search(r'金额:([^ /｜]+)', detail)
    return m.group(1) if m else ''


def group_display(chunk):
    """分组行(序号行): 序号 【名称】【用途】【金额】【发票详情】【购买方】"""
    it = chunk[0]
    parts = [it.label]
    m = re.search(r'名称:([^/｜]+)', it.detail)
    if m:
        parts.append('【%s】' % m.group(1).strip())
    m = re.search(r'用途:([^/｜]+)', it.detail)
    if m:
        parts.append('【%s】' % m.group(1).strip())
    m = re.search(r'金额:([^/｜]+)', it.detail)
    if m:
        parts.append('【%s】' % m.group(1).strip())
    m = re.search(r'发票详情:([^/｜]+)', it.detail)
    if m:
        parts.append('【发票详情:%s】' % m.group(1).strip())
    buyer = next((x.buyer for x in chunk if x.buyer), '')
    if buyer:
        parts.append('【购买方:%s】' % buyer)
    return it.label + (' ' + ''.join(parts[1:]) if len(parts) > 1 else '')


# ---------------------------------------------------------------- 发票购买方识别

PS_OCR_PS1 = r'''
param([string]$Path)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Foundation, ContentType=WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($WinRtTask, $ResultType) {
    $netTask = $asTask.MakeGenericMethod($ResultType).Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($engine) {
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $lines = @($result.Lines | ForEach-Object {
        $words = @($_.Words | ForEach-Object {
            $r = $_.BoundingRect
            [pscustomobject]@{ x = [int]$r.X; y = [int]$r.Y
                               w = [int]$r.Width; h = [int]$r.Height }
        })
        if ($words.Count -gt 0) {
            $minX = ($words.x | Measure-Object -Minimum).Minimum
            $minY = ($words.y | Measure-Object -Minimum).Minimum
            $maxR = ($words | ForEach-Object { $_.x + $_.w } |
                     Measure-Object -Maximum).Maximum
            $maxB = ($words | ForEach-Object { $_.y + $_.h } |
                     Measure-Object -Maximum).Maximum
            [pscustomobject]@{ t = $_.Text
                               x = [int]$minX; y = [int]$minY
                               w = [int]($maxR - $minX); h = [int]($maxB - $minY) }
        } else {
            [pscustomobject]@{ t = $_.Text; x = 0; y = 0; w = 0; h = 0 }
        }
    })
    ConvertTo-Json -InputObject $lines -Compress
}
'''

_ocr_ps1_path = None


def ocr_image_lines(data, ext='.png', timeout=30):
    """用 Windows 自带 OCR 识别图片, 返回带坐标的行列表
    [{'t': 文本, 'x','y','w','h': 像素框}], 失败返回 []"""
    global _ocr_ps1_path
    if os.name != 'nt':
        return []
    try:
        import subprocess
        if _ocr_ps1_path is None:
            _ocr_ps1_path = os.path.join(tempfile.gettempdir(),
                                         'invoice_printer_ocr.ps1')
            with open(_ocr_ps1_path, 'w', encoding='ascii', errors='ignore') as f:
                f.write(PS_OCR_PS1)
        fd, img = tempfile.mkstemp(suffix=ext)
        try:
            with os.fdopen(fd, 'wb') as f:
                f.write(data)
            r = subprocess.run(
                ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                 '-File', _ocr_ps1_path, '-Path', img],
                capture_output=True, timeout=timeout,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            out = r.stdout.decode('utf-8', 'ignore').strip()
            parsed = json.loads(out)
            if isinstance(parsed, dict):
                parsed = [parsed]
            return [p for p in parsed if isinstance(p, dict) and p.get('t')]
        finally:
            try:
                os.remove(img)
            except OSError:
                pass
    except Exception:
        return []


BUYER_STOP = (r'纳税人识别|统一社会信用|登记号|地址|电话|开户|账号|密码区|'
              r'发票号码|发票代码|开票日期|价税合计|合计|税额|校验码|机器编号|'
              r'项目名称|货物|规格|销售方|销方名称|购方名称|收款人|复核|开票人|'
              r'备注|名称[:：]|$')
# 发票左侧“购买方”是竖排, OCR 常把它和“名称”混成 购名称/购买名称 等
BUYER_RES = (
    re.compile(r'购买方名称[:：]?(.{2,40}?)(' + BUYER_STOP + ')'),
    re.compile(r'购名称[:：]?(.{2,40}?)(' + BUYER_STOP + ')'),
    re.compile(r'购买名称[:：]?(.{2,40}?)(' + BUYER_STOP + ')'),
    re.compile(r'购买方[:：]?(.{2,40}?)(' + BUYER_STOP + ')'),
)
_BUYER_NOISE = '信销购码了之的及与等方'
_ORG_SUFFIX = re.compile(
    r'(公司|企业|集团|中心|商行|经营部|门市部|厂|店|部|所|馆|院|'
    r'合作社|工作室|事务所|研究院)$')


def _clean_buyer(v):
    v = re.sub(r'^名称[:：]', '', v).strip('：:，,、;；。()（）「」 ')
    v = v.lstrip('·•◦')                   # OCR 噪声: 开头的间隔点
    # 发票标题等 OCR 噪声常粘连在名称后面: 从尾往回找第一个组织后缀
    for cut in range(0, min(20, len(v) - 3)):
        cand = v[:len(v) - cut]
        if _ORG_SUFFIX.search(cand):
            return cand
    for _ in range(4):
        if not v or _ORG_SUFFIX.search(v):
            break
        if v[-1] in _BUYER_NOISE:   # OCR 粘连的噪声尾字
            v = v[:-1]
        else:
            break
    return v


def extract_buyer(ocr_text):
    """从发票 OCR 全文里提取购买方主体名称(兜底用)"""
    if not ocr_text:
        return ''
    compact = re.sub(r'\s+', '', ocr_text)
    for rx in BUYER_RES:
        m = rx.search(compact)
        if m:
            v = _clean_buyer(m.group(1))
            if 2 <= len(v) <= 40 and re.search(r'[一-鿿A-Za-z]', v) \
                    and v not in ('信息', '名称', '购买方', '销售方'):
                return v
    return ''


def extract_buyer_from_lines(lines):
    """按位置提取购买方: 发票版式固定, 购买方在左上角区域"""
    if not lines:
        return ''
    W = max((p['x'] + p['w'] for p in lines), default=1) or 1
    H = max((p['y'] + p['h'] for p in lines), default=1) or 1
    cands = []
    for p in lines:
        tc = p['t'].replace(' ', '')       # OCR 常在汉字间插空格
        xc = (p['x'] + p['w'] / 2) / W
        yc = (p['y'] + p['h'] / 2) / H
        if yc <= 0.40 and xc <= 0.62:      # 购买方信息框(左上区域)
            cands.append((yc, tc))
    cands.sort(key=lambda t: t[0])
    for _yc, tc in cands:                   # 只认“名称:xxx”行——
        if '名称' in tc:                    # 支付截图的收款方行没有“名称”标签
            m = re.search(r'名称[:：]?(.+)', tc)
            if m:
                v = _clean_buyer(m.group(1))
                if 4 <= len(v) <= 40 and re.search(r'[一-鿿]', v):
                    return v
    return ''


def ocr_extract_buyer(data, ext='.png'):
    """图片 -> 购买方名称: 先按位置, 再用全文正则兜底"""
    lines = ocr_image_lines(data, ext)
    if not lines:
        return ''
    return (extract_buyer_from_lines(lines)
            or extract_buyer(' '.join(p['t'] for p in lines)))


OCR_CACHE_FILE = os.path.join(app_data_dir(), 'ocr_cache.json')

OCR_FAIL = '开票主体读取失败'   # 有发票但识别不出开票主体时的占位提示


def load_ocr_cache():
    try:
        with open(OCR_CACHE_FILE, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_ocr_cache(cache):
    try:
        with open(OCR_CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False)
    except Exception:
        pass


def wrap_text(text, font, maxw, max_lines=3):
    """按像素宽度贪心折行, 超过 max_lines 行截断并加省略号"""
    lines = []
    cur = ''
    overflow = False
    for ch in text:
        if cur and font.getlength(cur + ch) > maxw:
            lines.append(cur)
            cur = ch
            if len(lines) == max_lines:
                overflow = True
                break
        else:
            cur += ch
    if not overflow and cur:
        lines.append(cur)
    if overflow:
        last = lines[-1]
        while last and font.getlength(last + '…') > maxw:
            last = last[:-1]
        lines[-1] = last + '…'
    return lines


def group_items(path, raw):
    """把提取到的图片按工作表排序并附上订单标签, 返回 ImageItem 列表"""
    by_sheet = {}
    for r in raw:
        by_sheet.setdefault(r['sheet'], []).append(r)
    wb = openpyxl.load_workbook(path, data_only=True) if HAS_OPENPYXL else None
    items = []
    for sheet, rows in by_sheet.items():
        ws = wb[sheet] if (wb and sheet in wb.sheetnames) else None
        mmap = build_merged_map(ws) if ws else {}
        ocol, orow = detect_order_col(ws) if ws else (None, 0)
        dcols, hmap = scan_headers(ws, orow) if ws else ({}, {})
        for r in sorted(rows, key=lambda x: (x['row'], x['col'], x['order'])):
            label = (find_label(ws, mmap, ocol, orow, r['row'])
                     if ws else None) or '第%d行' % r['row']
            # 图片粘在“发票详情”列内 = 这张图就是发票
            is_invoice = (dcols.get('发票详情') is not None
                          and r['col'] == dcols['发票详情'])
            detail = build_detail(ws, mmap, dcols, r['row'],
                                  is_invoice) if ws else ''
            items.append(ImageItem(sheet=sheet, row=r['row'], col=r['col'],
                                   filename=r['filename'], data=r['data'],
                                   label=label, detail=detail,
                                   col_header=hmap.get(r['col'], ''),
                                   is_invoice=is_invoice))
    return items


# ---------------------------------------------------------------- 图片对象

@dataclass
class ImageItem:
    sheet: str
    row: int
    col: int
    filename: str
    data: bytes
    label: str
    detail: str = ''        # 名称/用途/金额/发票详情 等关键信息(打印标签用)
    col_header: str = ''    # 图片所在列的表头(如“凭证”), 列表展示用
    buyer: str = ''         # 从发票图片识别出的购买方主体名称
    is_invoice: bool = False  # 图片是否位于“发票详情”列(即发票本身)
    pil: Image.Image = None
    selected: bool = False  # 导入后默认全不选, 由用户勾选

    @property
    def width(self):
        return self.pil.width

    @property
    def height(self):
        return self.pil.height


_NON_IMAGE_EXTS = {'.pdf', '.ps', '.eps'}   # 文档类, 不作为图片导入


def _image_extensions():
    """动态收集当前环境支持的全部图片扩展名(以 Pillow 注册表为准)"""
    exts = set(Image.registered_extensions().keys())
    if HAS_HEIF:
        exts |= {'.heic', '.heics', '.heif', '.heifs'}
    return sorted(e.lower() for e in exts if e.lower() not in _NON_IMAGE_EXTS)


def _load_pdf_pages(path, max_pages=10):
    """PDF -> 按页渲染(约200dpi)为图片项。
    视为发票: is_invoice=True, data 存 PNG 字节供后台 OCR 购买方。"""
    import pypdfium2 as pdfium
    items = []
    stem = os.path.splitext(os.path.basename(path))[0]
    pdf = pdfium.PdfDocument(path)
    try:
        for i in range(min(len(pdf), max_pages)):
            pil = pdf[i].render(scale=200 / 72).to_pil().convert('RGB')
            png = io.BytesIO()
            pil.save(png, format='PNG')
            items.append(ImageItem(
                sheet='PDF', row=i + 1, col=0,
                filename='%s 第%d页.png' % (stem, i + 1),
                data=png.getvalue(), label=stem,
                is_invoice=True, pil=pil))
    finally:
        pdf.close()
    return items


def load_pil_image(data, filename):
    """bytes -> RGB PIL.Image, 透明底贴白底, 解析失败给占位图"""
    try:
        img = Image.open(io.BytesIO(data))
        img.load()
    except Exception:
        img = Image.new('RGB', (800, 500), (250, 245, 235))
        d = ImageDraw.Draw(img)
        d.text((20, 20), '无法解析: %s' % filename, fill=(180, 60, 60))
        return img
    if img.mode in ('RGBA', 'LA', 'P'):
        img = img.convert('RGBA')
        bg = Image.new('RGB', img.size, (255, 255, 255))
        bg.paste(img, mask=img.getchannel('A'))
        img = bg
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    if max(img.size) > 4000:
        img.thumbnail((4000, 4000), Image.Resampling.LANCZOS)
    return img


# ---------------------------------------------------------------- 排版引擎

_font_cache = {}


def get_font(size):
    if size not in _font_cache:
        f = None
        for p in ('C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf',
                  'C:/Windows/Fonts/simsun.ttc'):
            try:
                f = ImageFont.truetype(p, size)
                break
            except Exception:
                continue
        _font_cache[size] = f or ImageFont.load_default()
    return _font_cache[size]


def mm2px(mm, dpi):
    return int(round(mm / 25.4 * dpi))


def _best_half_grid(chunk, W, H, m, g):
    """自动排列: 在 上下堆(2,1) / 左右排(1,2) 里选让本页图片
    总显示面积更大的那种 —— 谁大图片就印得更大更清晰"""
    best, best_area = (2, 1), -1.0
    for r, c in ((2, 1), (1, 2)):
        cw = (W - 2 * m - (c - 1) * g) / c
        ch = (H - 2 * m - (r - 1) * g) / r
        area = 0.0
        for it in chunk:
            s = min(cw / it.width, ch / it.height)
            area += it.width * it.height * s * s
        if area > best_area:
            best, best_area = (r, c), area
    return best


def page_grid(layout, arrangement, chunk=None, W=0, H=0, m=0, g=0):
    """layout: full/half/quarter/eighth, 返回该页 (行, 列)"""
    if layout == 'full':
        return (1, 1)
    if layout == 'half':
        if arrangement == '上下':
            return (2, 1)
        if arrangement == '左右':
            return (1, 2)
        return _best_half_grid(chunk, W, H, m, g)
    if layout == 'eighth':
        return (4, 2)
    return (2, 2)  # quarter


_LAYOUT_SLOTS = {'full': 1, 'half': 2, 'quarter': 4, 'eighth': 8}


def _dzfp_filename_buyer(filename):
    """电子发票下载文件名自带购买方: dzfp_发票号_购买方名称_日期时间.pdf
    OCR 读不出时的兜底来源(仅识别 dzfp 命名规范, 避免误判)"""
    m = re.match(r'^dzfp_\d+_(.+?)_\d{14}', filename or '')
    return m.group(1) if m else ''


def plan_sheets(items, opts):
    """把图片分配到纸张(仅在“存在不同购买方的发票”的隔离模式下调用):
    - 发票(有购买方): 同购买方尽量同纸; 不同购买方严格不同纸
    - 其余图片: 作为填充分配到有空位的纸, 优先离原位置最近的纸,
      把隔离产生的空白吃掉, 节省纸张
    返回 [{'items': [ImageItem...], 'anchor': 原始起始序号}, ...]"""
    slots_page = _LAYOUT_SLOTS[opts['layout']]
    slots_sheet = slots_page * (2 if opts.get('side', '单面') == '双面' else 1)
    orig_idx = {id(it): k for k, it in enumerate(items)}
    sheets = []
    by_buyer = {}                       # 购买方 -> 当前所属纸张

    def new_sheet(anchor):
        s = {'items': [], 'buyers': set(), 'anchor': anchor}
        sheets.append(s)
        return s

    unknown = []                        # 读不出购买方的发票: 每张独立成纸
    for it in items:                    # 1) 发票先落位(保持原顺序)
        if not it.is_invoice:
            continue
        if not it.buyer:
            # OCR 未完成/失败: 先试文件名兜底; 仍没有则独立成纸,
            # 绝不静默丢弃(曾经因此导致排版预览少一张发票)
            it.buyer = _dzfp_filename_buyer(it.filename)
            if not it.buyer:
                unknown.append(it)
                continue
        s = by_buyer.get(it.buyer)
        if s is not None and len(s['items']) >= slots_sheet:
            s = None                    # 该购买方的纸满了 → 另开一张
        if s is None:
            s = new_sheet(orig_idx[id(it)])
            by_buyer[it.buyer] = s
        s['items'].append(it)

    for it in unknown:                  # 未知购买方发票: 独占纸张(保守隔离)
        s = new_sheet(orig_idx[id(it)])
        s['items'].append(it)

    for it in items:                    # 2) 其余图片填空(优先就近)
        if it.is_invoice:
            continue
        k = orig_idx[id(it)]
        cands = [s for s in sheets if len(s['items']) < slots_sheet]
        if not cands:                   # 全满 → 开新填充纸
            cands = [new_sheet(k)]
        s = min(cands, key=lambda s: abs(s['anchor'] - k))
        s['items'].append(it)
    return sheets


def plan_pages(items, opts):
    """把图片列表按排版切成页: [(rows, cols, [ImageItem...]), ...]
    仅当选中图片中存在两种及以上不同购买方的发票时才启用隔离:
    不同购买方不进同一张纸(正反面算同一张), 双面不足整页补空白反面页;
    其他情况一律按常规排版, 不产生空白页。"""
    slots_page = _LAYOUT_SLOTS[opts['layout']]
    duplex = opts.get('side', '单面') == '双面'
    pw, ph = PAPERS_MM[opts['paper']]
    W, H = mm2px(pw, 72), mm2px(ph, 72)
    m = mm2px(opts['margin'], 72)
    g = mm2px(opts['gap'], 72)
    pages = []

    def pack(seq):
        i = 0
        while i < len(seq):
            chunk = seq[i:i + slots_page]
            r, c = page_grid(opts['layout'], opts['arrangement'],
                             chunk, W, H, m, g)
            pages.append((r, c, chunk))
            i += slots_page

    distinct = {it.buyer for it in items if it.buyer}
    if len(distinct) >= 2:                  # 有不一致的购买方 → 隔离
        for sheet in plan_sheets(items, opts):
            start = len(pages)
            pack(sheet['items'])
            if duplex and (len(pages) - start) % 2 == 1:
                pages.append((1, 1, []))    # 空白反面页, 保住隔离
    else:
        pack(items)                         # 常规排版
    return pages


def render_page(plan, opts, dpi):
    """渲染一页: 图片等比缩放居中放入各自格子, 可选分隔线和订单号标签"""
    r, c, chunk = plan
    pw, ph = PAPERS_MM[opts['paper']]
    W, H = mm2px(pw, dpi), mm2px(ph, dpi)
    m = mm2px(opts['margin'], dpi)
    g = mm2px(opts['gap'], dpi)
    img = Image.new('RGB', (W, H), 'white')
    d = ImageDraw.Draw(img)
    cw = (W - 2 * m - (c - 1) * g) / c
    ch = (H - 2 * m - (r - 1) * g) / r
    rects = []
    caps = []
    f = get_font(mm2px(3.6, dpi)) if opts['label'] else None
    lh = int(mm2px(3.6, dpi) * 1.35) if opts['label'] else 0
    for idx, it in enumerate(chunk):
        rr, cc = divmod(idx, c)
        x0 = m + cc * (cw + g)
        y0 = m + rr * (ch + g)
        rects.append((x0, y0, x0 + cw, y0 + ch))
        # 备注条带: 独占格子顶部, 图片缩小让路, 两者绝不重叠
        lines = []
        if opts['label']:
            cap = it.label
            if it.detail:
                cap += ' ' + it.detail
            if it.buyer:
                cap += ' 购买方:' + it.buyer
            lines = wrap_text(cap, f, cw - mm2px(2, dpi), 3)
        cap_h = (len(lines) * lh + mm2px(1.4, dpi)) if lines else 0
        caps.append((x0, y0, lines))
        # 图片: 备注条带以下的剩余区域, 等比缩放居中
        iw, ih = cw, ch - cap_h
        if iw <= 2 or ih <= 2:
            continue
        im = it.pil
        scale = min(iw / im.width, ih / im.height)
        nw, nh = max(1, int(im.width * scale)), max(1, int(im.height * scale))
        im2 = im.resize((nw, nh), Image.Resampling.LANCZOS)
        px = int(x0 + (iw - nw) / 2)
        py = int(y0 + cap_h + (ih - nh) / 2)
        img.paste(im2, (px, py))

    # 备注文字画在各自预留条带内(图片已避开, 无需衬底)
    for x0, y0, lines in caps:
        tx = x0 + mm2px(1, dpi)
        ty = y0 + mm2px(0.6, dpi)
        for ln in lines:
            d.text((tx, ty), ln, fill=(70, 70, 70), font=f)
            ty += lh
    # 分隔框最后画在最上层: 保证四条边完整可见(修“左边线消失”)
    if opts['sep']:
        for (x0, y0, x1, y1) in rects:
            d.rectangle([x0, y0, x1, y1], outline=(150, 150, 150),
                        width=max(1, mm2px(0.25, dpi)))
    return img


def iter_pages(items, opts, dpi):
    for plan in plan_pages(items, opts):
        yield render_page(plan, opts, dpi)


# ---------------------------------------------------------------- 打印机

def _write_printer_diag(diag):
    """打印机枚举为空时写诊断日志, 便于远程排查"""
    try:
        with open(os.path.join(app_data_dir(), '打印机诊断日志.txt'), 'a',
                  encoding='utf-8') as f:
            f.write('\n'.join(diag) + '\n\n')
    except Exception:
        pass


# ---------------- winspool/GDI 常量与结构(纯 ctypes, 取代 pywin32) ----------------
DMORIENT_PORTRAIT = 1
DMDUP_SIMPLEX, DMDUP_VERTICAL = 1, 2
DMCOLOR_MONOCHROME, DMCOLOR_COLOR = 1, 2
DM_ORIENTATION, DM_DUPLEX, DM_COLOR = 0x0001, 0x1000, 0x0800
DM_IN_BUFFER, DM_OUT_BUFFER = 8, 2
PHYSICALWIDTH, PHYSICALHEIGHT = 110, 111
PHYSICALOFFSETX, PHYSICALOFFSETY = 112, 113
DIB_RGB_COLORS, SRCCOPY, COLORONCOLOR = 0, 0x00CC0020, 3
DC_DUPLEX = 17
PRINTER_ENUM_LOCAL_CONN, PRINTER_ENUM_NETWORK = 0x6, 0x20


class _DOCINFOW(ctypes.Structure):
    _fields_ = [('cbSize', ctypes.c_uint32), ('lpszDocName', ctypes.c_wchar_p),
                ('lpszOutput', ctypes.c_wchar_p),
                ('lpszDatatype', ctypes.c_wchar_p),
                ('fwType', ctypes.c_uint32)]


class _BMIH(ctypes.Structure):
    _fields_ = [('biSize', ctypes.c_uint32), ('biWidth', ctypes.c_int32),
                ('biHeight', ctypes.c_int32), ('biPlanes', ctypes.c_uint16),
                ('biBitCount', ctypes.c_uint16),
                ('biCompression', ctypes.c_uint32),
                ('biSizeImage', ctypes.c_uint32),
                ('biXPelsPerMeter', ctypes.c_int32),
                ('biYPelsPerMeter', ctypes.c_int32),
                ('biClrUsed', ctypes.c_uint32),
                ('biClrImportant', ctypes.c_uint32)]


class _PRINTER_INFO_4W(ctypes.Structure):
    _fields_ = [('pPrinterName', ctypes.c_wchar_p),
                ('pServerName', ctypes.c_wchar_p),
                ('Attributes', ctypes.c_uint32)]


class _PRINTER_INFO_2W(ctypes.Structure):
    _fields_ = [('pServerName', ctypes.c_wchar_p),
                ('pPrinterName', ctypes.c_wchar_p),
                ('pShareName', ctypes.c_wchar_p),
                ('pPortName', ctypes.c_wchar_p),
                ('pDriverName', ctypes.c_wchar_p),
                ('pComment', ctypes.c_wchar_p),
                ('pLocation', ctypes.c_wchar_p),
                ('pDevMode', ctypes.c_void_p),
                ('pSepFile', ctypes.c_wchar_p),
                ('pPrintProcessor', ctypes.c_wchar_p),
                ('pDatatype', ctypes.c_wchar_p),
                ('pParameters', ctypes.c_wchar_p),
                ('pSecurityDescriptor', ctypes.c_void_p),
                ('Attributes', ctypes.c_uint32),
                ('Priority', ctypes.c_uint32),
                ('DefaultPriority', ctypes.c_uint32),
                ('StartTime', ctypes.c_uint32),
                ('UntilTime', ctypes.c_uint32),
                ('Status', ctypes.c_uint32),
                ('cJobs', ctypes.c_uint32),
                ('AveragePPM', ctypes.c_uint32)]


def _winspool_enum_printers(diag, add):
    """ctypes 枚举本机/连接/网络打印机, 供诊断与兜底"""
    w = ctypes.WinDLL('winspool.drv')
    for flags, label in ((PRINTER_ENUM_LOCAL_CONN, 'LOCAL|CONNECTIONS'),
                         (PRINTER_ENUM_NETWORK, 'NETWORK')):
        try:
            need, got = ctypes.c_uint32(0), ctypes.c_uint32(0)
            w.EnumPrintersW(flags, None, 4, None, 0,
                            ctypes.byref(need), ctypes.byref(got))
            if not need.value:
                diag.append('EnumPrinters(%s): 0 台' % label)
                continue
            buf = ctypes.create_string_buffer(need.value)
            got2 = ctypes.c_uint32(0)
            if not w.EnumPrintersW(flags, None, 4, buf, need.value,
                                   ctypes.byref(need), ctypes.byref(got2)):
                diag.append('EnumPrinters(%s) 二次调用失败' % label)
                continue
            arr = ctypes.cast(buf, ctypes.POINTER(
                _PRINTER_INFO_4W * got2.value)).contents
            names = [x.pPrinterName or '' for x in arr]
            diag.append('EnumPrinters(%s): %d 台 %s' % (label, len(names),
                                                        names))
            for n in names:
                add(n)
        except Exception as e:
            diag.append('EnumPrinters(%s) 异常: %r' % (label, e))


def list_printers():
    """枚举打印机: 打印 API 优先, 注册表兜底; 全空时写诊断日志"""
    if IS_MAC:
        out = []
        try:
            r = subprocess.run(['lpstat', '-p'], capture_output=True,
                               text=True, timeout=10)
            for ln in (r.stdout or '').splitlines():
                if ln.startswith('printer '):
                    name = ln[8:].split()[0].rstrip(':')
                    if name and name not in out:
                        out.append(name)
        except Exception:
            pass
        return out
    diag = ['--- 打印机诊断 %s ---' % time.strftime('%Y-%m-%d %H:%M:%S')]
    out, seen = [], set()

    def add(name):
        name = (name or '').strip()
        if not name:
            return
        diag.append('  发现: %s' % name)
        if name not in seen:
            seen.add(name)
            out.append(name)

    # 注意: NETWORK 不能与 LOCAL 合并枚举(部分系统会返回空), 分开补齐
    _winspool_enum_printers(diag, add)
    try:
        dp = default_printer()
        diag.append('默认打印机: %r' % dp)
        add(dp)
    except Exception as e:
        diag.append('默认打印机读取异常: %r' % e)

    # 注册表兜底(不依赖打印 API, 安全软件/权限异常时仍能列出)
    try:
        import winreg
        for root, path, mode in (
                (winreg.HKEY_CURRENT_USER,
                 r'Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts',
                 'values'),
                (winreg.HKEY_CURRENT_USER,
                 r'Software\Microsoft\Windows NT\CurrentVersion\Devices',
                 'values'),
                (winreg.HKEY_LOCAL_MACHINE,
                 r'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
                 r'\Print\Printers',
                 'keys')):
            try:
                k = winreg.OpenKey(root, path)
            except OSError as e:
                diag.append('注册表 %s 打不开: %r' % (path, e))
                continue
            i, cnt = 0, 0
            while True:
                try:
                    if mode == 'keys':
                        name = winreg.EnumKey(k, i)
                    else:
                        name, _v, _t = winreg.EnumValue(k, i)
                except OSError:
                    break
                i += 1
                cnt += 1
                add(name)
            winreg.CloseKey(k)
            diag.append('注册表 %s: %d 项' % (path, cnt))
    except Exception as e:
        diag.append('注册表枚举异常: %r' % e)

    if not out:
        diag.append('!! 所有路径均未发现打印机')
        _write_printer_diag(diag)
    else:
        diag.append('共 %d 台' % len(out))
    return out


def default_printer():
    """读系统默认打印机 (Win=ctypes, mac=CUPS)"""
    if IS_MAC:
        try:
            r = subprocess.run(['lpstat', '-d'], capture_output=True,
                               text=True, timeout=10)
            for ln in (r.stdout or '').splitlines():
                if 'default destination' in ln:
                    return ln.split(':', 1)[1].strip()
        except Exception:
            pass
        return ''
    try:
        w = ctypes.WinDLL('winspool.drv')
        n = ctypes.c_uint32(0)
        w.GetDefaultPrinterW(None, ctypes.byref(n))
        buf = ctypes.create_unicode_buffer(n.value or 1)
        if w.GetDefaultPrinterW(buf, ctypes.byref(n)):
            return buf.value
    except Exception:
        pass
    return ''


def printer_supports_duplex(printer):
    """查询打印机是否支持双面; 查询失败时不阻止打印(返回 True)"""
    if IS_MAC:
        return True      # CUPS 对不支持的选项自动忽略, 不必预检
    try:
        w = ctypes.WinDLL('winspool.drv')
        h = ctypes.c_void_p()
        if not w.OpenPrinterW(printer, ctypes.byref(h), None):
            return True
        try:
            need = ctypes.c_uint32(0)
            w.GetPrinterW(h, 2, None, 0, ctypes.byref(need))
            if not need.value:
                return True
            buf = ctypes.create_string_buffer(need.value)
            got = ctypes.c_uint32(0)
            if not w.GetPrinterW(h, 2, buf, need.value, ctypes.byref(got)):
                return True
            info = ctypes.cast(buf, ctypes.POINTER(_PRINTER_INFO_2W)).contents
            cap = w.DeviceCapabilitiesW(printer, info.pPortName or '',
                                        DC_DUPLEX, None, None)
            return bool(cap)
        finally:
            w.ClosePrinter(h)
    except Exception:
        return True  # 查询失败时不阻止打印


def convert_xls_to_xlsx(xls_path):
    """调用本机 Excel / WPS 把 .xls 另存为 .xlsx, 返回新文件路径(临时目录)"""
    if IS_MAC:
        raise RuntimeError(
            'mac 版暂不支持 .xls 自动转换。\n'
            '请先用 Excel/WPS 把文件另存为 .xlsx 再导入。')
    try:
        import pythoncom
        import win32com.client
    except Exception:
        raise RuntimeError(
            '本机 Python 组件(pywin32)加载失败, 无法自动转换 .xls。\n'
            '请先用 Excel/WPS 把文件另存为 .xlsx 再导入。')
    pythoncom.CoInitialize()
    app = None
    for progid in ('Excel.Application', 'KET.Application', 'ET.Application'):
        try:
            app = win32com.client.Dispatch(progid)
            break
        except Exception:
            continue
    if app is None:
        raise RuntimeError(
            '本机没有可用的 Excel 或 WPS, 无法自动转换 .xls。\n'
            '请先用 Excel/WPS 把文件另存为 .xlsx 再导入。')
    out = os.path.join(
        tempfile.gettempdir(),
        '发票打印助手_%s.xlsx' % time.strftime('%H%M%S'))
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(xls_path), 0, True)  # 只读打开
        try:
            wb.SaveAs(os.path.abspath(out), 51)  # 51 = xlsx
        finally:
            wb.Close(False)
    finally:
        app.Quit()
    return out


def extract_excel_attachments(path):
    """收集 xl/embeddings 里的 PDF 附件:
    直接的 .pdf, 或藏在 OLE 对象(.bin)里的 PDF(按 %PDF- 标记截取)"""
    out = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not name.startswith('xl/embeddings/'):
                continue
            base = posixpath.basename(name)
            low = base.lower()
            try:
                data = zf.read(name)
            except Exception:
                continue
            if low.endswith('.pdf'):
                out.append((base, data))
            elif low.endswith('.bin'):
                i = data.find(b'%PDF-')
                if i >= 0:
                    j = data.rfind(b'%%EOF')
                    pdf = data[i:j + 5] if j > i else data[i:]
                    out.append((posixpath.splitext(base)[0] + '.pdf', pdf))
    return out


def render_pdf_pages(data, target_w=2200, max_pages=5):
    """把 PDF 逐页渲染为 PIL 图片(电子发票通常只有 1 页)"""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(data)
    try:
        imgs = []
        for i in range(min(len(doc), max_pages)):
            page = doc[i]
            scale = max(0.5, target_w / max(page.get_width(), 1.0))
            imgs.append(page.render(scale=scale).to_pil().convert('RGB'))
        return imgs
    finally:
        doc.close()


class _POINTL(ctypes.Structure):
    _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]


class _DEVMODE_U(ctypes.Union):
    """16 字节联合: 8 个 short(orientation..printquality) 与 POINTL/dword 重叠"""
    _fields_ = [('dmPosition', _POINTL),
                ('dmDisplayOrientation', ctypes.c_uint32),
                ('dmDisplayFixedOutput', ctypes.c_uint32),
                ('_shorts', ctypes.c_int16 * 8)]


class _DEVMODEW(ctypes.Structure):
    """DEVMODEW 完整布局(2 字节压缩对齐, 总 220 字节,
    偏移经驱动实际缓冲反向校验)"""
    _pack_ = 2
    _fields_ = [
        ('dmDeviceName', ctypes.c_wchar * 32),
        ('dmSpecVersion', ctypes.c_uint16),
        ('dmDriverVersion', ctypes.c_uint16),
        ('dmSize', ctypes.c_uint16),
        ('dmDriverExtra', ctypes.c_uint16),
        ('dmFields', ctypes.c_uint32),          # 72
        ('u', _DEVMODE_U),                      # 76..92
        ('dmColor', ctypes.c_int16),            # 92
        ('dmDuplex', ctypes.c_int16),           # 94
        ('dmYResolution', ctypes.c_int16),      # 96
        ('dmTTOption', ctypes.c_int16),         # 98
        ('dmCollate', ctypes.c_int16),          # 100
        ('dmFormName', ctypes.c_wchar * 32),    # 102
        ('dmLogPixels', ctypes.c_uint16),       # 134
        ('dmBitsPerPel', ctypes.c_uint32),
        ('dmPelsWidth', ctypes.c_uint32),
        ('dmPelsHeight', ctypes.c_uint32),
        ('dmDisplayFlags', ctypes.c_uint32),
        ('dmDisplayFrequency', ctypes.c_uint32),
        ('dmICMMethod', ctypes.c_uint32),
        ('dmICMIntent', ctypes.c_uint32),
        ('dmMediaType', ctypes.c_uint32),
        ('dmDitherType', ctypes.c_uint32),
        ('dmReserved1', ctypes.c_uint32),
        ('dmReserved2', ctypes.c_uint32),
        ('dmPanningWidth', ctypes.c_uint32),
        ('dmPanningHeight', ctypes.c_uint32),
    ]


def _devmode_buf(printer, hprinter):
    """取驱动默认 DEVMODE 完整缓冲(公共+私有), 校验布局并返回 (buf, dm)"""
    winspool = ctypes.WinDLL('winspool.drv')
    need = winspool.DocumentPropertiesW(None, hprinter, printer,
                                        None, None, 0)
    if not need:
        raise RuntimeError('读取打印机 [%s] 设置失败(驱动可能异常或未正确安装)'
                           % printer)
    buf = ctypes.create_string_buffer(need)
    winspool.DocumentPropertiesW(None, hprinter, printer, buf, None,
                                 DM_OUT_BUFFER)
    dm = ctypes.cast(buf, ctypes.POINTER(_DEVMODEW)).contents
    if dm.dmSize > ctypes.sizeof(_DEVMODEW):
        raise RuntimeError('DEVMODE 布局与驱动不匹配(dmSize=%d)'
                           % dm.dmSize)
    return buf, dm


def mac_print(plans, opts, printer, duplex, copies, progress=None, color='彩色'):
    """macOS 打印: 页面渲染成多页 PDF, 交给 CUPS(lp) 发送。
    双面/黑白由 CUPS 选项控制(-o sides= / print-color-mode=)。
    返回 (双面是否生效, 黑白/彩色是否生效) —— CUPS 接受选项即视为生效。"""
    pages = []
    total = len(plans)
    for k, plan in enumerate(plans):
        page = render_page(plan, opts, PRINT_DPI)
        if color == '黑白':
            page = page.convert('L').convert('RGB')
        pages.append(page)
        if progress:
            progress(k + 1, total)
    pdf = os.path.join(tempfile.gettempdir(),
                       '发票打印助手_%s.pdf' % time.strftime('%H%M%S'))
    pages[0].save(pdf, save_all=True, append_images=pages[1:],
                  resolution=PRINT_DPI)
    cmd = ['lp', '-d', printer, '-n', str(copies),
           '-o', 'sides=' + ('two-sided-long-edge' if duplex
                             else 'one-sided'),
           '-o', 'print-color-mode=' + ('monochrome' if color == '黑白'
                                        else 'color'),
           pdf]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    finally:
        try:
            os.remove(pdf)
        except Exception:
            pass
    if r.returncode != 0:
        raise RuntimeError('发送打印任务失败 [%s]\n%s'
                           % (printer, (r.stderr or r.stdout or '').strip()))
    return True, True


def gdi_print(plans, opts, printer, duplex, copies, progress=None, color=None,
              output=None):
    """纯 ctypes GDI 打印(不依赖 pywin32, 任何电脑打印按钮都可用)。
    双面/黑白/彩色写入 DEVMODE 并经驱动 DocumentProperties 校验合并;
    页面图像用 StretchDIBits 直绘。返回 (双面是否生效, 黑白/彩色是否生效)。"""
    winspool = ctypes.WinDLL('winspool.drv')
    gdi32 = ctypes.windll.gdi32
    gdi32.CreateDCW.restype = ctypes.c_void_p
    # 句柄/指针参数必须显式声明原型(c_int 会截断 64 位句柄)
    gdi32.StartDocW.restype = ctypes.c_int
    gdi32.StartDocW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    for fn in ('StartPage', 'EndPage', 'EndDoc', 'DeleteDC'):
        f = getattr(gdi32, fn)
        f.restype = ctypes.c_int
        f.argtypes = [ctypes.c_void_p]
    gdi32.GetDeviceCaps.restype = ctypes.c_int
    gdi32.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
    gdi32.SetStretchBltMode.restype = ctypes.c_int
    gdi32.SetStretchBltMode.argtypes = [ctypes.c_void_p, ctypes.c_int]
    gdi32.StretchDIBits.restype = ctypes.c_int
    gdi32.StretchDIBits.argtypes = [
        ctypes.c_void_p, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.c_int, ctypes.c_char_p, ctypes.c_void_p,
        ctypes.c_uint, ctypes.c_uint]

    hprinter = ctypes.c_void_p()
    if not winspool.OpenPrinterW(printer, ctypes.byref(hprinter), None):
        raise RuntimeError('无法连接打印机 [%s]\n请确认打印机已开机、'
                           '名称正确且本机有使用权限。' % printer)
    try:
        buf, dm = _devmode_buf(printer, hprinter)
        dm.u._shorts[0] = DMORIENT_PORTRAIT   # dmOrientation 在联合体里
        dm.dmFields |= DM_ORIENTATION
        dm.dmDuplex = (DMDUP_VERTICAL if duplex else DMDUP_SIMPLEX)
        dm.dmFields |= DM_DUPLEX
        if color == '黑白':
            dm.dmColor = DMCOLOR_MONOCHROME
            dm.dmFields |= DM_COLOR
        elif color == '彩色':
            dm.dmColor = DMCOLOR_COLOR
            dm.dmFields |= DM_COLOR
        # 交给驱动校验合并(保留私有数据), 部分驱动需要两次合并
        winspool.DocumentPropertiesW(None, hprinter, printer, buf, buf,
                                     DM_IN_BUFFER | DM_OUT_BUFFER)
        winspool.DocumentPropertiesW(None, hprinter, printer, buf, buf,
                                     DM_IN_BUFFER | DM_OUT_BUFFER)
        want_duplex = DMDUP_VERTICAL if duplex else DMDUP_SIMPLEX
        want_color = (DMCOLOR_MONOCHROME if color == '黑白'
                      else DMCOLOR_COLOR)
        duplex_ok = dm.dmDuplex == want_duplex
        color_ok = (color == '黑白') or dm.dmColor == want_color
    finally:
        winspool.ClosePrinter(hprinter)

    hdc = gdi32.CreateDCW('WINSPOOL', printer, None, buf)
    if not hdc:
        raise RuntimeError('创建打印任务失败 [%s]\n'
                           '该打印机驱动可能不支持程序直接打印, '
                           '可尝试「导出 PDF」后手动打印。' % printer)
    total = len(plans)
    try:
        di = _DOCINFOW(ctypes.sizeof(_DOCINFOW),
                       '%s %s' % (APP_TITLE, time.strftime('%Y%m%d-%H%M%S')),
                       output, None, 0)
        if gdi32.StartDocW(hdc, ctypes.byref(di)) <= 0:
            raise RuntimeError('打印机 [%s] 拒绝了打印任务。' % printer)
        try:
            gdi32.SetStretchBltMode(hdc, COLORONCOLOR)
            pw = gdi32.GetDeviceCaps(hdc, PHYSICALWIDTH)
            ph = gdi32.GetDeviceCaps(hdc, PHYSICALHEIGHT)
            ox = gdi32.GetDeviceCaps(hdc, PHYSICALOFFSETX)
            oy = gdi32.GetDeviceCaps(hdc, PHYSICALOFFSETY)
            for copy in range(copies):
                for k, plan in enumerate(plans):
                    page = render_page(plan, opts, PRINT_DPI)
                    if color == '黑白':
                        # 软件灰度: 部分驱动(如爱普生)不接受驱动级黑白
                        page = page.convert('L').convert('RGB')
                    if gdi32.StartPage(hdc) <= 0:
                        raise RuntimeError('打印机 [%s] 开始新页失败。'
                                           % printer)
                    w, hgt = page.size
                    bmi = _BMIH(ctypes.sizeof(_BMIH), w, -hgt, 1, 32, 0,
                                0, 0, 0, 0, 0)     # 负高度 = 自上而下
                    bits = page.tobytes('raw', 'BGRX')
                    # GDI 坐标原点在"可打印区"左上角: 必须减去不可打印边距,
                    # 页面图像才能与纸张物理边缘对齐(否则内容整体右移/下移,
                    # 左侧留白翻倍、右侧被裁掉)
                    gdi32.StretchDIBits(hdc, -ox, -oy, pw, ph, 0, 0, w, hgt,
                                        bits, ctypes.byref(bmi),
                                        DIB_RGB_COLORS, SRCCOPY)
                    gdi32.EndPage(hdc)
                    page = None
                    if progress:
                        progress(copy * total + k + 1, total * copies)
        finally:
            gdi32.EndDoc(hdc)
    finally:
        try:
            gdi32.DeleteDC(hdc)
        except Exception:
            pass
    return duplex_ok, color_ok


# ---------------------------------------------------------------- 预览窗口

class PagePreviewWindow(tk.Toplevel):
    def __init__(self, master, items, opts, side):
        super().__init__(master)
        self.title('排版预览')
        self.geometry('760x900')
        self._pages = []          # 渲染好的 PIL 页面(渲染一次, 显示时缩放)
        self._labels = []
        self._job = None

        canvas = tk.Canvas(self, highlightthickness=0, bg='#6b6b6b')
        self._canvas = canvas
        vs = ttk.Scrollbar(self, orient='vertical', command=canvas.yview)
        frame = ttk.Frame(canvas)
        frame.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        self._win = canvas.create_window((0, 0), window=frame, anchor='nw')
        canvas.bind('<Configure>', self._on_canvas_resize)
        canvas.configure(yscrollcommand=vs.set)
        canvas.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        # 滚轮: 用 bind_all 才能捕获子控件上的事件, 只响应本窗口
        self.bind_all('<MouseWheel>', self._on_wheel)

        loading = ttk.Label(frame, text='正在渲染页面 ...')
        loading.pack(pady=24)
        self._loading = loading
        self.update()

        plans = plan_pages(items, opts)
        for k, plan in enumerate(plans):
            if side == '双面':
                cap = '第 %d 张纸 · %s' % (k // 2 + 1,
                                          '正面' if k % 2 == 0 else '反面')
            else:
                cap = '第 %d 页' % (k + 1)
            if not plan[2]:
                cap += '（空白）'        # 隔离产生的空白反面页
            tk.Label(frame, text=cap, bg='#5a5a5a', fg='white',
                     font=('Microsoft YaHei UI', 11, 'bold')).pack(
                fill='x', pady=(10, 2))
            page = render_page(plan, opts, 100)   # 高清渲染一次, 显示时再缩放
            self._pages.append(page)
            lbl = tk.Label(frame, bg='#5a5a5a')
            lbl.pack(pady=(0, 10))
            self._labels.append(lbl)
        tk.Label(frame, text='共 %d 页' % len(plans), bg='#5a5a5a',
                 fg='white').pack(fill='x', pady=8)

    def _display_pages(self):
        """页面按当前窗口宽度等比缩放并居中显示(四周留白)"""
        self._job = None
        try:
            if not self.winfo_exists():
                return
            if getattr(self, '_loading', None) is not None:
                self._loading.destroy()
                self._loading = None
            cw = max(self._canvas.winfo_width() - 40, 200)  # 两侧各留 20px
            for page, lbl in zip(self._pages, self._labels):
                scale = min(cw / page.width, 1.0)
                nw = max(1, int(page.width * scale))
                nh = max(1, int(page.height * scale))
                ph = ImageTk.PhotoImage(
                    page.resize((nw, nh), Image.Resampling.LANCZOS))
                lbl.configure(image=ph)
                lbl.image = ph
        except Exception:
            pass

    def _on_canvas_resize(self, e):
        """窗口宽度变化: 内容区跟随 + 防抖后重排页面"""
        try:
            self._canvas.itemconfigure(self._win, width=e.width)
        except Exception:
            pass
        if self._job:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
        self._job = self.after(150, self._display_pages)


# ---------------------------------------------------------------- 主界面

    def _on_wheel(self, e):
        """排版预览滚轮滚动"""
        try:
            if not self.winfo_exists() or \
                    e.widget.winfo_toplevel() is not self:
                return
            self._canvas.yview_scroll(-1 * (e.delta // 120 or 1), 'units')
        except Exception:
            pass


class App:
    def __init__(self, root):
        self.root = root
        root.title('%s %s' % (APP_TITLE, APP_VERSION))
        # 高分屏缩放系数: 按 DPI 将所有像素尺寸等比放大(字体为磅值, Tk 自动缩放)
        try:
            self.scale = max(1.0, root.winfo_fpixels('1i') / 96.0)
        except Exception:
            self.scale = 1.0
        root.geometry('%dx%d' % (int(1360 * self.scale),
                                 int(800 * self.scale)))
        style = ttk.Style()
        try:
            style.theme_use('vista')
        except tk.TclError:
            pass
        style.configure('Treeview', rowheight=int(30 * self.scale),
                        font=('Microsoft YaHei UI', 10))

        # 全局字体统一调大 1 号(系统默认 9 -> 10): 未显式指定字体的控件走这些命名字体
        try:
            import tkinter.font as tkfont
            for fname in ('TkDefaultFont', 'TkTextFont', 'TkHeadingFont',
                          'TkMenuFont', 'TkTooltipFont'):
                f = tkfont.nametofont(fname)
                sz = f.actual('size')
                f.configure(size=sz + 1 if sz > 0 else sz - 1)
        except Exception:
            pass

        self.items = []          # ImageItem 有序列表(即打印顺序)
        self.iid_map = {}        # tree iid -> items 下标
        self.runs_map = {}       # 组 iid -> (起, 止) 下标区间
        self._preview_img = None
        self._preview_item = None    # 当前预览的图片
        self._preview_job = None     # 预览重绘防抖
        self._lightbox = None        # 大图查看弹窗
        self._col_resize = None      # 表头列宽拖动状态
        self._hover_after = None      # 悬停定时器
        self._hover_popup = None      # 悬停小图弹窗
        self._hover_shown_for = None  # 当前弹窗对应的图片 id()
        self._hover_cache = {}        # id(item) -> 缩略图 PhotoImage
        self._ocr_cache = load_ocr_cache()
        self._ocr_running = False

        # ---- 状态变量
        self.xlsx_path = tk.StringVar(value='(未导入)')
        self.side_var = tk.StringVar(value='单面')
        self.layout_var = tk.StringVar(value='half')
        self.arr_var = tk.StringVar(value='自动')
        self.paper_var = tk.StringVar(value='A4')
        self.margin_var = tk.IntVar(value=5)
        self.gap_var = tk.IntVar(value=5)
        self.sep_var = tk.BooleanVar(value=True)
        self.label_var = tk.BooleanVar(value=True)
        self.copies_var = tk.IntVar(value=1)
        self.color_var = tk.StringVar(value='彩色')
        self.printer_var = tk.StringVar()

        self._load_settings()
        self._build_ui()
        self._register_drop()   # 必须在全部控件建好后注册, 否则子控件收不到拖放
        self.root.bind('<Control-v>', self.on_paste_files)   # 备用导入通道
        self.root.bind('<Control-V>', self.on_paste_files)
        self.refresh_printers()
        self.refresh_summary()
        root.protocol('WM_DELETE_WINDOW', self.on_close)

    # ---------------- 界面搭建

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=(8, 6))
        top.pack(fill='x')
        ttk.Button(top, text='📂 导入 Excel',
                   command=self.do_import).pack(side='left', padx=(0, 8))
        ttk.Button(top, text='🖼 添加图片文件',
                   command=self.add_image_files).pack(side='left', padx=(0, 8))
        ttk.Button(top, text='📥 备用金报销模版下载',
                   command=self.download_template).pack(side='left', padx=(0, 8))
        ttk.Label(top, textvariable=self.xlsx_path,
                  foreground='#555').pack(side='left', padx=8)

        # 中部: 左列表 | 竖线A(拖动控左宽) | 弹性空隙 | 竖线B(拖动控右宽) | 右预览
        # 两根竖线各自独立, 拖动互不影响另一侧的宽度
        body = ttk.Frame(self.root)
        self.body = body
        body.pack(fill='both', expand=True, padx=8, pady=(0, 4))

        # ---- 左侧: 图片列表 (宽度 = 窗口的 62%, 随窗口等比伸缩)
        # 中间不留任何可见分隔线: 两侧边框各自描述边界, 中间就是背景留白
        left = tk.Frame(body, highlightthickness=2,
                        highlightbackground='#a8a8a8')
        left.place(relx=0, rely=0, relwidth=0.62, relheight=1.0, width=-6)
        self.left_panel = left

        btns = ttk.Frame(left)
        btns.pack(fill='x', padx=3, pady=(3, 0))
        for text, cmd in (('全选', lambda: self.select_all(True)),
                          ('全不选', lambda: self.select_all(False)),
                          ('反选', self.invert_selection),
                          ('清空', self.clear_all)):
            ttk.Button(btns, text=text, width=8, command=cmd).pack(
                side='left', padx=(0, 4))
        # 汇总: 与按钮同一行, 顶着表格右侧显示
        self.summary_var = tk.StringVar(value='未导入数据')
        ttk.Label(btns, textvariable=self.summary_var, foreground='#0a6',
                  anchor='e').pack(side='right', fill='x', expand=True,
                                   padx=(0, 18))

        cols = ('info',)
        self.tree = ttk.Treeview(left, columns=cols, show='tree headings')
        self.tree.heading('#0', text='订单 / 图片 (点行勾选 · 点信息列预览)')
        self.tree.heading('info', text='信息')
        self.tree.column('#0', width=480, stretch=False)
        self.tree.column('info', width=230, stretch=False)
        vs = ttk.Scrollbar(left, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side='left', fill='both', expand=True,
                       padx=(3, 0), pady=3)
        vs.pack(side='left', fill='y', padx=(0, 3))

        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<ButtonRelease-1>', self._on_tree_release)
        self.tree.bind('<Configure>', self._on_tree_configure)
        self.tree.bind('<space>', self.on_tree_space)
        self.tree.bind('<Motion>', self.on_tree_motion)
        self.tree.bind('<Leave>', self.on_tree_leave)
        self.tree.bind('<MouseWheel>', self.on_tree_wheel)
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)


        # ---- 右侧: 预览 (宽度 = 窗口的 38%, 随窗口等比伸缩)
        right = tk.Frame(body, highlightthickness=2,
                         highlightbackground='#a8a8a8')
        right.place(relx=0.62, x=10, rely=0, relwidth=0.38, relheight=1.0,
                    width=-10)
        self.right_panel = right

        pf = ttk.Labelframe(right, text='图片预览 (点击左侧列表查看)')
        pf.pack(fill='both', expand=True, padx=3, pady=3)
        self.preview = tk.Canvas(pf, bg='white', highlightthickness=0)
        self.preview.pack(fill='both', expand=True, padx=4, pady=4)
        self.preview.bind('<Configure>', self._on_preview_resize)
        self.preview.bind('<Button-1>', self._on_preview_click)

        # ---- 底部通栏: 排版设置(宽度随左表) + 打印(宽度随右预览)
        bottom = ttk.Frame(self.root, height=128)
        bottom.pack(fill='x', padx=8, pady=(0, 4))
        bottom.pack_propagate(False)

        self.bottom_left = ttk.Frame(bottom)
        self.bottom_left.place(relx=0, rely=0, relwidth=0.62, relheight=1.0,
                               width=-6)
        sf = ttk.Labelframe(self.bottom_left, text='排版设置')
        sf.pack(fill='both', expand=True)
        row1 = ttk.Frame(sf)
        row1.pack(fill='x', padx=6, pady=3)
        ttk.Label(row1, text='打印面:').pack(side='left')
        for v in ('单面', '双面'):
            ttk.Radiobutton(row1, text=v, value=v, variable=self.side_var,
                            command=self.refresh_summary).pack(
                side='left', padx=(6, 2))
        ttk.Label(row1, text='   每页张数:').pack(side='left')
        for text, v in (('整页', 'full'), ('1/2', 'half'),
                        ('1/4', 'quarter'), ('1/8', 'eighth')):
            ttk.Radiobutton(row1, text=text, value=v, variable=self.layout_var,
                            command=self.refresh_summary).pack(
                side='left', padx=(6, 2))

        row2 = ttk.Frame(sf)
        row2.pack(fill='x', padx=6, pady=3)
        ttk.Label(row2, text='排列:').pack(side='left')
        for v in ('自动', '上下', '左右'):
            ttk.Radiobutton(row2, text=v, value=v, variable=self.arr_var,
                            command=self.refresh_summary).pack(
                side='left', padx=(6, 2))
        ttk.Label(row2, text='   纸张:').pack(side='left')
        ttk.Combobox(row2, textvariable=self.paper_var, width=7,
                     values=list(PAPERS_MM), state='readonly').pack(
            side='left', padx=(4, 12))
        ttk.Label(row2, text='页边距mm:').pack(side='left')
        ttk.Spinbox(row2, textvariable=self.margin_var, from_=0, to=30,
                    width=4, command=self.refresh_summary).pack(
            side='left', padx=(4, 12))
        ttk.Label(row2, text='间距mm:').pack(side='left')
        ttk.Spinbox(row2, textvariable=self.gap_var, from_=0, to=30,
                    width=4, command=self.refresh_summary).pack(side='left', padx=4)

        row3 = ttk.Frame(sf)
        row3.pack(fill='x', padx=6, pady=3)
        ttk.Checkbutton(row3, text='绘制分隔框线', variable=self.sep_var,
                        command=self.refresh_summary).pack(side='left')
        ttk.Checkbutton(row3, text='打印信息标签(名称/用途/金额/发票)',
                        variable=self.label_var,
                        command=self.refresh_summary).pack(side='left', padx=12)
        ttk.Label(row3, text='份数:').pack(side='left')
        ttk.Spinbox(row3, textvariable=self.copies_var, from_=1, to=99,
                    width=4).pack(side='left', padx=4)

        self.bottom_right = ttk.Frame(bottom)
        self.bottom_right.place(relx=0.62, x=10, rely=0, relwidth=0.38,
                                relheight=1.0)
        pf2 = ttk.Labelframe(self.bottom_right, text='打印')
        pf2.pack(fill='both', expand=True)
        prow = ttk.Frame(pf2)
        prow.pack(fill='x', padx=6, pady=3)
        ttk.Label(prow, text='打印机:').pack(side='left')
        self.printer_cb = ttk.Combobox(prow, textvariable=self.printer_var,
                                       state='readonly')
        self.printer_cb.pack(side='left', fill='x', expand=True, padx=4)
        ttk.Button(prow, text='刷新', width=5,
                   command=self.refresh_printers).pack(side='left')

        crow = ttk.Frame(pf2)
        crow.pack(fill='x', padx=6, pady=(0, 3))
        ttk.Label(crow, text='打印色彩:').pack(side='left')
        for v in ('彩色', '黑白'):
            ttk.Radiobutton(crow, text=v, value=v,
                            variable=self.color_var).pack(side='left', padx=(6, 2))
        ttk.Label(crow, text='(黑白由打印机驱动转换)',
                  foreground='#888').pack(side='left', padx=(8, 0))

        brow = ttk.Frame(pf2)
        brow.pack(fill='x', padx=6, pady=(3, 6))
        ttk.Button(brow, text='👁 预览排版',
                   command=self.show_page_preview).pack(side='left')
        ttk.Button(brow, text='导出 PDF',
                   command=self.export_pdf).pack(side='left', padx=8)
        self.print_btn = ttk.Button(brow, text='🖨 开始打印',
                                    command=self.start_print)
        self.print_btn.pack(side='right')
        self.status_var = tk.StringVar(value='就绪 · %s%s' % (
            APP_VERSION, ' · 可拖入 Excel/图片导入' if HAS_DND else ''))
        ttk.Label(self.root, textvariable=self.status_var, anchor='w',
                  relief='sunken', padding=(6, 2)).pack(fill='x', side='bottom')

    # ---------------- 数据操作

    def _status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def load_items(self, items, src_label):
        for it in items:
            if it.pil is None:
                it.pil = load_pil_image(it.data, it.filename)
            # 保留原始字节(it.data), 供后台 OCR 识别购买方使用
        self.items.extend(items)
        self.xlsx_path.set(src_label)
        self.rebuild_tree()
        self.refresh_summary()
        self._status('已加载 %d 张图片' % len(items))
        if self.items:
            first = [iid for iid in self.tree.get_children('')]
            if first:
                kids = self.tree.get_children(first[0])
                if kids:
                    self.tree.selection_set(kids[0])
        self.start_ocr_backfill()

    def do_import(self):
        path = filedialog.askopenfilename(
            title='选择报销 Excel',
            filetypes=[('Excel 工作簿', '*.xlsx *.xlsm *.xls'),
                       ('所有文件', '*.*')])
        if path:
            self._import_path(path)

    def _import_path(self, path):
        if not os.path.isfile(path):
            messagebox.showerror('导入失败', '文件不存在:\n%s' % path)
            return
        # 线路分工: 「导入 Excel」按钮只认 Excel/PDF, 图片请走添加按钮或拖入
        ext = os.path.splitext(path)[1].lower()
        if ext in _image_extensions():
            messagebox.showinfo(
                '提示', '这是图片文件。\n请用「🖼 添加图片文件」按钮添加, '
                        '或直接拖入窗口。')
            self._status('「导入 Excel」不支持图片, 已忽略')
            return
        if path.lower().endswith('.pdf'):
            self._import_pdf(path)
            return
        src_label = os.path.basename(path)
        try:
            self._status('正在解析 Excel ...')
            self.root.config(cursor='watch')
            self.root.update()
            if path.lower().endswith('.xls'):
                self._status('正在转换 .xls (需要本机装有 Excel 或 WPS) ...')
                self.root.update()
                path = convert_xls_to_xlsx(path)
            raw = extract_excel_images(path)
            items = group_items(path, raw) if raw else []
            n_att = self._load_attachments(path)
            if not items and not n_att:
                messagebox.showinfo(
                    '未找到内容', '该 Excel 里没有嵌入图片, 也没有 PDF 附件。\n'
                    '注意: 以“链接文件”方式插入的图片无法提取, 请直接嵌入。')
                self._status('未找到内容')
                return
            self.load_items(items, '📄 ' + src_label + ' (%d 张图片%s)'
                            % (len(items),
                               ', %d 页 PDF 附件' % n_att if n_att else ''))
        except Exception:
            messagebox.showerror('导入失败', traceback.format_exc())
            self._status('导入失败')
        finally:
            self.root.config(cursor='')

    def _load_attachments(self, path):
        """把 Excel 内嵌 PDF 附件渲染成图片追加到列表, 返回页数"""
        try:
            atts = extract_excel_attachments(path)
        except Exception:
            return 0
        if not atts:
            return 0
        if not HAS_PDFIUM:
            messagebox.showwarning(
                '发现 PDF 附件',
                'Excel 里有 %d 个 PDF 附件, 但未安装 pypdfium2, '
                '无法转成图片打印。\n可执行: pip install pypdfium2' % len(atts))
            return 0
        added = 0
        for name, data in atts:
            try:
                pages = render_pdf_pages(data)
            except Exception:
                continue
            label = 'PDF·%s' % os.path.splitext(name)[0][:30]
            for i, pil in enumerate(pages):
                self.items.append(ImageItem(
                    sheet='PDF附件', row=0, col=0,
                    filename='%s 第%d页' % (name, i + 1),
                    data=b'', pil=pil, label=label))
                added += 1
        return added

    def add_image_files(self):
        paths = filedialog.askopenfilenames(
            title='选择图片/PDF(可多选)',
            filetypes=[('图片/PDF 文件',
                        ' '.join('*' + e for e in _image_extensions()) + ' *.pdf'),
                       ('所有文件', '*.*')])
        self._append_image_files(paths)

    def _import_pdf(self, path):
        """PDF 发票: 按页渲染导入, 后台自动 OCR 购买方"""
        try:
            self._status('正在解析 PDF ...')
            self.root.config(cursor='watch')
            self.root.update()
            items = _load_pdf_pages(path)
            if not items:
                messagebox.showinfo('提示', '该 PDF 没有可显示的页面。')
                return
            self.load_items(items, '📄 ' + os.path.basename(path)
                            + ' (PDF %d 页)' % len(items))
            self._status('PDF 已导入: 发票图片将自动识别购买方')
        except Exception:
            messagebox.showerror('导入失败', traceback.format_exc())
            self._status('PDF 导入失败')
        finally:
            self.root.config(cursor='')

    def _append_image_files(self, paths):
        """兼容旧调用: 交由统一入口处理(PDF 会按页展开)"""
        # 线路分工: 「添加图片文件」不收 Excel, 给出明确指引
        excels = [p for p in paths
                  if os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm',
                                                        '.xls')]
        if excels:
            messagebox.showinfo(
                '提示', 'Excel 文件请用「📂 导入 Excel」按钮或直接拖入窗口,\n'
                        '已忽略 %d 个 Excel。' % len(excels))
            self._status('「添加图片文件」不支持 Excel, 已忽略')
        paths = [p for p in paths
                 if os.path.splitext(p)[1].lower() not in ('.xlsx', '.xlsm',
                                                           '.xls')]
        pdfs = [p for p in paths if os.path.splitext(p)[1].lower() == '.pdf']
        rest = [p for p in paths if os.path.splitext(p)[1].lower() != '.pdf']
        if pdfs:
            items = []
            for p in pdfs:
                items += _load_pdf_pages(p)
            if items:
                self.load_items(items, '📄 已导入 %d 个 PDF (%d 页)'
                                % (len(pdfs), len(items)))
        if rest:
            self._append_image_files_inner(rest)

    def _append_image_files_inner(self, paths):
        items = []
        for p in paths:
            try:
                with open(p, 'rb') as f:
                    data = f.read()
            except OSError:
                continue
            items.append(ImageItem(sheet='手动添加', row=0, col=0,
                                   filename=os.path.basename(p), data=data,
                                   label='手动添加'))
        if items:
            self.load_items(items, '已追加 %d 张本地图片' % len(items))

    def download_template(self):
        """导出内置的备用金报销模版"""
        src = embedded_template_path()
        if not src:
            messagebox.showerror(
                '模板缺失',
                '软件内未找到内置模板文件(%s)。\n请联系软件提供者重新打包。'
                % TEMPLATE_FILE)
            return
        dst = filedialog.asksaveasfilename(
            title='保存备用金报销模版',
            initialfile=TEMPLATE_FILE,
            defaultextension='.xlsx',
            filetypes=[('Excel 工作簿', '*.xlsx')])
        if not dst:
            return
        try:
            shutil.copyfile(src, dst)
        except Exception as e:
            messagebox.showerror('保存失败', str(e))
            return
        if messagebox.askyesno(
                '模板下载完成', '已保存到:\n%s\n\n是否立即打开所在文件夹?' % dst):
            try:
                os.startfile(os.path.dirname(dst))
            except Exception:
                pass

    # ---------------- 列表勾选/排序

    def rebuild_tree(self):
        self._hover_cancel()
        self._hide_hover()
        self.tree.delete(*self.tree.get_children())
        self.iid_map.clear()
        self.runs_map.clear()
        i = gi = 0
        n = len(self.items)
        while i < n:
            j = i
            label = self.items[i].label
            while j < n and self.items[j].label == label:
                j += 1
            chunk = self.items[i:j]
            gid = 'g%d' % gi
            gi += 1
            self.runs_map[gid] = (i, j)
            nsel = sum(1 for it in chunk if it.selected)
            glyph = '☑' if nsel == len(chunk) else ('▣' if nsel else '☐')
            self.tree.insert('', 'end', gid,
                             text=' %s  %s' % (glyph, group_display(chunk)),
                             values=('%d 张' % len(chunk),), open=True)
            for k in range(i, j):
                it = self.items[k]
                iid = 'i%d' % k
                self.iid_map[iid] = k
                g = '☑' if it.selected else '☐'
                fname = it.filename
                parts = []
                if it.col_header:
                    parts.append(it.col_header[:10])
                if it.buyer:
                    parts.append(it.buyer)      # 购买方显示在信息栏
                info = ' · '.join(parts)
                self.tree.insert(gid, 'end', iid,
                                 text='   %s %s' % (g, fname),
                                 values=(info,))
            i = j

    def toggle_iid(self, iid):
        if iid in self.runs_map:
            st, en = self.runs_map[iid]
            newv = not all(it.selected for it in self.items[st:en])
            for it in self.items[st:en]:
                it.selected = newv
        elif iid in self.iid_map:
            it = self.items[self.iid_map[iid]]
            it.selected = not it.selected
        else:
            return
        self.rebuild_tree()
        self.refresh_summary()

    def select_all(self, val):
        for it in self.items:
            it.selected = val
        self.rebuild_tree()
        self.refresh_summary()

    def invert_selection(self):
        for it in self.items:
            it.selected = not it.selected
        self.rebuild_tree()
        self.refresh_summary()

    def clear_all(self):
        """一键清空已加载的表格数据"""
        self.items.clear()
        self._preview_item = None
        self._preview_img = None
        self._hover_cache.clear()
        self.preview.delete('all')
        self.rebuild_tree()
        self.refresh_summary()
        self._status('已清空全部数据, 可重新导入')

    # ---------------- 事件

    def _near_col_boundary(self, x):
        """鼠标是否在表头“订单/图片 | 信息”列分隔线附近(±5px)"""
        edge = self.tree.winfo_width() - self.tree.column('info', 'width')
        return abs(x - edge) <= 5

    def _on_tree_release(self, _e):
        self._col_resize = None

    def _on_tree_configure(self, _e):
        """面板宽度变化时, 让“订单/图片”列自动填满(信息列固定)"""
        if getattr(self, '_col_resize', None):
            return
        w = self.tree.winfo_width() - self.tree.column('info', 'width') - 2
        if w > 150 and abs(self.tree.column('#0', 'width') - w) > 2:
            self.tree.column('#0', width=w)

    def on_tree_click(self, e):
        self._hover_cancel()
        self._hide_hover()
        # 表头分隔线: 进入列宽拖动模式(ttk 原生不支持, 手动实现)
        try:
            region = self.tree.identify_region(e.x, e.y)
        except Exception:
            region = ''
        if region == 'separator' or (
                region == 'heading' and self._near_col_boundary(e.x)):
            self._col_resize = '#0'
            return 'break'
        iid = self.tree.identify_row(e.y)
        col = self.tree.identify_column(e.x)
        if not iid:
            return
        if col == '#0':
            # 点标题列任意位置 = 勾选/取消 (组=整组切换, 子项=单选)
            self.toggle_iid(iid)
            if iid in self.iid_map:
                self.show_item_preview(self.items[self.iid_map[iid]])
            elif iid in self.runs_map:
                self.show_item_preview(self.items[self.runs_map[iid][0]])
            return 'break'
        # 点信息列 = 仅选中预览, 不改变勾选

    def on_tree_space(self, _e):
        self._hover_cancel()
        self._hide_hover()
        sel = self.tree.selection()
        if sel:
            self.toggle_iid(sel[0])

    def on_tree_select(self, _e=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        if iid in self.iid_map:
            self.show_item_preview(self.items[self.iid_map[iid]])
        elif iid in self.runs_map:      # 选中组时预览第一张
            st, _en = self.runs_map[iid]
            self.show_item_preview(self.items[st])

    # ---------------- 拖放导入

    def _register_drop(self):
        """整个窗口(含所有子控件)注册为拖放目标; 必须在建完全部 UI 后调用"""
        n_ok = n_fail = 0
        first_err = None
        if HAS_DND:

            def walk(w):
                yield w
                for c in w.winfo_children():
                    yield from walk(c)

            for w in list(walk(self.root)):
                try:
                    w.drop_target_register(DND_FILES)
                    # 必须用 dnd_bind! tkdnd 事件带自定义 % 替换字段
                    # (见 _substitute_dnd), 标准 bind 的解析器会直接抛
                    # "expected integer but got %#" 导致回调永远不执行
                    w.dnd_bind('<<Drop>>', self.on_drop)
                    # 悬停反馈: 拖着文件进入窗口时状态栏提示, 便于确认事件到达
                    w.dnd_bind('<<DropEnter>>', self._on_drag_enter)
                    w.dnd_bind('<<DropLeave>>', self._on_drag_leave)
                    n_ok += 1
                except Exception as ex:
                    n_fail += 1
                    if first_err is None:
                        first_err = repr(ex)
        is_admin = False
        try:
            is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            pass
        _dnd_log('启动 %s HAS_DND=%s 注册成功=%d 失败=%d %s管理员运行=%s'
                 % (APP_VERSION, HAS_DND, n_ok, n_fail,
                    ('首错:%s ' % first_err) if first_err else '', is_admin))
        if is_admin:
            # UIPI: 普通权限的资源管理器向管理员进程拖放会被 Windows 拦截,
            # 逐窗口放行 WM_DROPFILES(0x233)/WM_COPYGLOBALDATA(0x4A)
            allowed = 0
            try:
                u32 = ctypes.windll.user32

                def walk2(w):
                    yield w
                    for c in w.winfo_children():
                        yield from walk2(c)

                for w in walk2(self.root):
                    try:
                        hwnd = w.winfo_id()
                        if u32.ChangeWindowMessageFilterEx(hwnd, 0x0233, 1,
                                                           None):
                            allowed += 1
                        u32.ChangeWindowMessageFilterEx(hwnd, 0x004A, 1, None)
                    except Exception:
                        pass
            except Exception:
                pass
            _dnd_log('管理员模式: 对 %d 个窗口放行拖放消息(0x233/0x4A)' % allowed)
            self.root.after(
                1200, self._status,
                '管理员模式: 若拖入无效请改从文件夹拖放, 或用 Ctrl+V 粘贴导入')

    def _on_drag_enter(self, _e):
        self._status('松开鼠标即可导入')

    def _on_drag_leave(self, _e):
        self._status('就绪')

    def _parse_drop_paths(self, data):
        """解析 tkdnd 的文件列表: 带空格的路径是 {花括号} 组, 其余取无空白连续段.
        不能用 tk.splitlist —— 无空格路径里的反斜杠会被 Tcl 当转义序列吃掉
        (实测 D:\26 -> 0x16), 导致 os.path.isfile 全部失败而静默丢弃"""
        paths = []
        for m in re.finditer(r'\{([^{}]*)\}|(\S+)', data):
            p = m.group(1) if m.group(1) is not None else m.group(2)
            if p and p not in paths:
                paths.append(p)
        return paths

    def _collect_dir_files(self, folder, out):
        """递归收集文件夹内支持的文件(xlsx/pdf/图片), 追加到 out, 返回收集数"""
        n0 = len(out)
        for root, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if not d.startswith(('.', '$'))]
            for fn in sorted(files):
                if len(out) - n0 >= 500:   # 单次拖入上限, 防误拖整个盘
                    self._status('文件夹内容过多, 仅导入前 500 个文件')
                    return len(out) - n0
                ext = os.path.splitext(fn)[1].lower()
                if (ext in ('.xlsx', '.xlsm', '.xls', '.pdf')
                        or ext in _image_extensions()):
                    out.append(os.path.join(root, fn))
        return len(out) - n0

    def on_drop(self, e):
        """拖入文件或文件夹: Excel 走导入, PDF 按页展开, 图片走追加;
        拖入文件夹时递归展开里面全部支持文件(含子文件夹)"""
        _dnd_log('收到拖放事件: %r' % getattr(e, 'data', '<无data>'))
        try:
            paths = []
            for p in self._parse_drop_paths(e.data):
                if os.path.isfile(p):
                    paths.append(p)
                elif os.path.isdir(p):
                    n = self._collect_dir_files(p, paths)
                    _dnd_log('展开文件夹 %s: 收集 %d 个文件' % (p, n))
        except Exception as ex:
            self._status('拖入处理失败: %r' % ex)
            return
        if not paths:
            _dnd_log('拖放被丢弃: 无有效文件路径')
            self._status('拖入的内容不是有效文件')
            return
        self._route_import(paths)

    def _route_import(self, paths):
        """按类型分流导入: Excel / PDF / 图片 (拖放与 Ctrl+V 共用)"""
        excels = [p for p in paths
                  if os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm', '.xls')]
        pdfs = [p for p in paths if os.path.splitext(p)[1].lower() == '.pdf']
        images = [p for p in paths
                  if os.path.splitext(p)[1].lower() in _image_extensions()]
        _dnd_log('分流导入: excel=%d pdf=%d 图片=%d'
                 % (len(excels), len(pdfs), len(images)))
        try:
            # 三类各自独立处理: 混合拖入(如 PDF+图片)时全部导入, 不互斥
            if excels:
                for p in excels:
                    self._import_path(p)
            if pdfs:
                items = []
                for p in pdfs:
                    items += _load_pdf_pages(p)
                if items:
                    self.load_items(items, '📄 已导入 %d 个 PDF (%d 页)'
                                    % (len(pdfs), len(items)))
            if images:
                self._append_image_files(images)
            if not (excels or pdfs or images):
                messagebox.showwarning('不支持的文件',
                                       '请拖入 Excel(.xlsx/.xls)、PDF 或图片文件。')
        except Exception:
            messagebox.showerror('拖入处理失败', traceback.format_exc())

    def on_paste_files(self, _e=None):
        """Ctrl+V: 导入剪贴板中的文件 (拖放不可用时的备用通道)"""
        paths = _clipboard_files()
        if not paths:
            _dnd_log('Ctrl+V: 剪贴板无文件(未复制文件或读取失败)')
            return 'break'
        _dnd_log('Ctrl+V 导入: %r' % paths)
        self._route_import([p for p in paths if os.path.isfile(p)])
        return 'break'

    # ---------------- 发票购买方后台识别

    def start_ocr_backfill(self):
        """按统一模板规则后台识别购买方:
        发票粘在“发票详情”列内 -> 识别该发票图片;
        旧格式(发票混在凭证列) -> 兼容识别; 读不出则标记“开票主体读取失败”"""
        if IS_MAC:
            return    # mac 无 Windows OCR, 跳过购买方识别(不影响打印)
        if self._ocr_running:
            return
        runs = []
        i, n = 0, len(self.items)
        while i < n:
            j = i
            while j < n and self.items[j].label == self.items[i].label:
                j += 1
            chunk = self.items[i:j]
            has_inv = any(it.is_invoice for it in chunk)
            if (has_inv or '发票详情:有' in chunk[0].detail) and \
                    not any(x.buyer for x in chunk):
                runs.append(chunk)
            i = j
        if not runs:
            return
        self._ocr_running = True
        threading.Thread(target=self._ocr_worker, args=(runs,),
                         daemon=True).start()

    def _ocr_worker(self, runs):
        total = len(runs)
        done = 0
        empty_streak = 0
        found_any = False
        disabled = False

        def status(msg):
            def _s():
                try:
                    self._status(msg)
                except Exception:
                    pass
            try:
                self.root.after(0, _s)
            except Exception:
                pass

        status('正在识别发票购买方 (0/%d) ...' % total)
        for chunk in runs:
            if disabled:
                break
            done += 1
            inv_items = [it for it in chunk if it.is_invoice]
            # 统一模板: 发票详情列内的图片就是发票; 旧格式兼容识别凭证列
            targets = (inv_items or chunk)[:4]
            buyer = None
            for it in targets:
                h = hashlib.md5(it.data).hexdigest()
                buyer = self._ocr_cache.get(h)
                if buyer is None:
                    lines = ocr_image_lines(
                        it.data, os.path.splitext(it.filename)[1] or '.png')
                    if not lines:
                        empty_streak += 1
                        if empty_streak >= 2 and not found_any:
                            disabled = True
                            status('未启用发票识别(系统缺少中文 OCR 语言包)')
                            break
                    buyer = (extract_buyer_from_lines(lines)
                             or extract_buyer(' '.join(p['t'] for p in lines)))
                    self._ocr_cache[h] = buyer
                    save_ocr_cache(self._ocr_cache)
                if buyer:
                    found_any = True
                    break
            if disabled:
                break
            if not buyer:
                # 兜底: 电子发票下载文件名自带购买方(dzfp_号码_购买方_日期时间)
                for it in (inv_items or chunk[:1]):
                    buyer = _dzfp_filename_buyer(it.filename)
                    if buyer:
                        break
            if not buyer:
                buyer = OCR_FAIL       # 有发票但读不出开票主体
            for it in (inv_items or chunk[:1]):
                it.buyer = buyer
            if buyer != OCR_FAIL:
                status('识别到购买方: %s (%d/%d) ...' % (buyer[:16], done, total))
            else:
                status('开票主体读取失败 (%d/%d) ...' % (done, total))
        self._ocr_running = False
        if not disabled:
            status('购买方识别完成' if found_any
                   else '购买方识别完成(未在发票中找到购买方)')

        def refresh():
            try:
                self.rebuild_tree()
                self.refresh_summary()
            except Exception:
                pass
        try:
            self.root.after(0, refresh)
        except Exception:
            pass

    def show_item_preview(self, it):
        self._preview_item = it
        self._render_preview()

    def _render_preview(self):
        """预览图自适应完整显示: 竖图以上下为准(左右留白),
        宽图以左右为准(上下留白), 居中并跟随面板大小缩放"""
        self._preview_job = None
        it = getattr(self, '_preview_item', None)
        if it is None or it.pil is None:
            return
        try:
            cw = max(self.preview.winfo_width(), 50)
            ch = max(self.preview.winfo_height(), 50)
            im = it.pil
            scale = min(cw / im.width, ch / im.height)   # 完整显示
            nw = max(1, int(im.width * scale))
            nh = max(1, int(im.height * scale))
            im2 = im.resize((nw, nh), Image.Resampling.LANCZOS)
            self._preview_img = ImageTk.PhotoImage(im2)
            self.preview.delete('all')
            self.preview.create_image(cw // 2, ch // 2,
                                      image=self._preview_img,
                                      anchor='center')
        except Exception:
            pass

    def _on_preview_resize(self, _e):
        """预览面板尺寸变化 -> 防抖后重新铺满"""
        if self._preview_job:
            try:
                self.root.after_cancel(self._preview_job)
            except Exception:
                pass
        self._preview_job = self.root.after(120, self._render_preview)

    def _on_preview_click(self, _e):
        it = getattr(self, '_preview_item', None)
        if it is not None:
            self._open_lightbox(it)

    def _open_lightbox(self, it):
        """点击图片: 全屏深色背景 + 大图居中查看。
        点击图片以外的深色区域 / ✕ 按钮 / Esc 关闭; 点击图片本身不关闭。"""
        if getattr(self, '_lightbox', None) is not None:
            return
        try:
            sw, sh = self.root.winfo_screenwidth(), \
                self.root.winfo_screenheight()
            im = it.pil
            scale = min(sw * 0.92 / im.width, sh * 0.92 / im.height, 3.0)
            im2 = im.resize((max(1, int(im.width * scale)),
                             max(1, int(im.height * scale))),
                            Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im2)

            mask = tk.Toplevel(self.root)
            mask.overrideredirect(True)
            mask.geometry('%dx%d+0+0' % (sw, sh))
            mask.configure(bg='#1a1a1a')
            try:
                mask.attributes('-topmost', True)
            except Exception:
                pass

            def close(_e=None):
                if _e is not None and \
                        getattr(_e, 'widget', None) is img_lbl:
                    return        # 点击图片本身: 不关闭, 便于查看
                try:
                    mask.destroy()
                except Exception:
                    pass
                self._lightbox = None

            # 大图(点击图片本身不关闭, 便于查看)
            box = tk.Frame(mask, bg='#1a1a1a')
            img_lbl = tk.Label(box, image=ph, bg='#111')
            img_lbl.pack(padx=4, pady=4)
            cap = (it.filename + (' · 购买方:' + it.buyer if it.buyer else ''))
            tk.Label(box, text=cap[:60], bg='#1a1a1a', fg='#ddd',
                     font=('Microsoft YaHei UI', 11)).pack(
                padx=10, pady=(0, 8))
            box.place(relx=0.5, rely=0.5, anchor='center')
            box._keep = ph                       # 防止图片被回收

            # 关闭按钮
            tk.Button(mask, text='✕ 关闭 (Esc)', command=close,
                      bg='#333', fg='white', relief='flat',
                      font=('Microsoft YaHei UI', 10),
                      activebackground='#555',
                      activeforeground='white').place(
                relx=1.0, x=-24, y=14, anchor='ne')

            # 点击图片以外的深色区域关闭; 点击图片本身不关闭
            mask.bind('<Button-1>', close)
            box.bind('<Button-1>', close)
            mask.bind('<Escape>', close)
            box.bind('<Escape>', close)
            mask.focus_set()
            self._lightbox = mask
            self._status('查看大图: 点击深色区域或按 Esc 关闭')
        except Exception:
            self._lightbox = None

    # ---------------- 悬停小图预览

    def on_tree_motion(self, e):
        if getattr(self, '_col_resize', None) == '#0':
            # 拖动表头分隔线: 调整“订单/图片”列宽(信息列跟随)
            self.tree.column('#0', width=max(280, min(1000, int(e.x))))
            return
        # 悬停到列分隔线附近时提示可拖动
        try:
            if self.tree.identify_region(e.x, e.y) == 'heading' and \
                    self._near_col_boundary(e.x):
                self.tree.configure(cursor='sb_h_double_arrow')
            else:
                self.tree.configure(cursor='')
        except Exception:
            pass
        iid = self.tree.identify_row(e.y)
        if not iid or iid not in self.iid_map:
            self._hover_cancel()
            self._hide_hover()
            return
        it = self.items[self.iid_map[iid]]
        if self._hover_popup and self._hover_shown_for == id(it):
            return                      # 已在显示这一张
        # 立即弹出, 挪开即消失
        self._hover_cancel()
        self._hide_hover()
        self._show_hover(it, e.x_root, e.y_root)

    def on_tree_leave(self, _e):
        self._hover_cancel()
        self._hide_hover()

    def on_tree_wheel(self, _e):
        self._hover_cancel()
        self._hide_hover()              # 不拦截, 滚动照常

    def _hover_cancel(self):
        if self._hover_after:
            try:
                self.root.after_cancel(self._hover_after)
            except Exception:
                pass
            self._hover_after = None

    def _hide_hover(self):
        if self._hover_popup:
            try:
                self._hover_popup.destroy()
            except Exception:
                pass
            self._hover_popup = None
        self._hover_shown_for = None

    def _show_hover(self, it, x, y):
        """在光标附近弹出该行图片的小窗预览"""
        self._hover_after = None
        self._hide_hover()
        try:
            ph = self._hover_cache.get(id(it))
            if ph is None:
                im = it.pil.copy()
                im.thumbnail((300, 300), Image.Resampling.LANCZOS)
                ph = ImageTk.PhotoImage(im)
                self._hover_cache[id(it)] = ph
            pop = tk.Toplevel(self.root)
            pop.wm_overrideredirect(True)       # 无边框提示框样式
            try:
                pop.attributes('-topmost', True)
            except Exception:
                pass
            pop.wm_withdraw()
            box = tk.Frame(pop, bg='#444')
            tk.Label(box, image=ph, bg='white').pack(padx=3, pady=3)
            txt = it.filename
            if it.buyer:
                txt += ' · 购买方:' + it.buyer
            amt = _detail_amount(it.detail)
            if amt:
                txt = '%s · %s' % (amt, txt)
            tk.Label(box, text=txt[:44], bg='#444', fg='white',
                     font=('Microsoft YaHei UI', 10)).pack(padx=8, pady=(0, 5))
            box.pack()
            pop.update_idletasks()
            w, h = pop.winfo_reqwidth(), pop.winfo_reqheight()
            sw, sh = pop.winfo_screenwidth(), pop.winfo_screenheight()
            px = x + 18 if x + 18 + w < sw - 8 else x - w - 12
            py = y + 18 if y + 18 + h < sh - 8 else y - h - 12
            pop.wm_geometry('+%d+%d' % (max(4, px), max(4, py)))
            pop.wm_deiconify()
            self._hover_popup = pop
            self._hover_shown_for = id(it)
        except Exception:
            self._hide_hover()

    # ---------------- 汇总/打印/导出

    def current_opts(self):
        return dict(layout=self.layout_var.get(), arrangement=self.arr_var.get(),
                    paper=self.paper_var.get(), margin=int(self.margin_var.get()),
                    gap=int(self.gap_var.get()), sep=self.sep_var.get(),
                    label=self.label_var.get(), side=self.side_var.get())

    # ---------------- 设置记忆

    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding='utf-8') as f:
                s = json.load(f)
        except Exception:
            return

        def put(var, key, cast=None):
            if key in s:
                try:
                    var.set(cast(s[key]) if cast else s[key])
                except Exception:
                    pass
        put(self.side_var, 'side')
        put(self.layout_var, 'layout')
        put(self.arr_var, 'arrangement')
        put(self.paper_var, 'paper')
        put(self.margin_var, 'margin', int)
        put(self.gap_var, 'gap', int)
        put(self.sep_var, 'sep', bool)
        put(self.label_var, 'caption', bool)
        put(self.copies_var, 'copies', int)
        put(self.color_var, 'color')
        put(self.printer_var, 'printer')
        g = s.get('geometry')
        if g and 'x' in g:
            self.root.geometry(g)

    def on_close(self):
        try:
            s = dict(side=self.side_var.get(), layout=self.layout_var.get(),
                     arrangement=self.arr_var.get(), paper=self.paper_var.get(),
                     margin=int(self.margin_var.get()),
                     gap=int(self.gap_var.get()),
                     sep=bool(self.sep_var.get()),
                     caption=bool(self.label_var.get()),
                     copies=int(self.copies_var.get()),
                     color=self.color_var.get(),
                     printer=self.printer_var.get(),
                     geometry=self.root.winfo_geometry())
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except Exception:
            pass
        self.root.destroy()

    def selected_items(self):
        return [it for it in self.items if it.selected]

    def refresh_summary(self):
        items = self.selected_items()
        if not items:
            self.summary_var.set('未选择任何图片')
            return
        opts = self.current_opts()
        pages = len(plan_pages(items, opts))
        duplex = self.side_var.get() == '双面'
        sheets = (pages + 1) // 2 if duplex else pages
        per_sheet = _LAYOUT_SLOTS[opts['layout']] * (2 if duplex else 1)
        need = -(-len(items) // per_sheet)
        iso = ' · 已按购买方隔离' if sheets > need else ''
        if duplex:
            self.summary_var.set(
                '已选 %d/%d 张 · %d 页 → 双面 %d 张纸%s'
                % (len(items), len(self.items), pages, sheets, iso))
        else:
            self.summary_var.set(
                '已选 %d/%d 张 · 单面 %d 页 / %d 张纸%s'
                % (len(items), len(self.items), pages, pages, iso))

    def refresh_printers(self):
        ps = list_printers()
        self.printer_cb['values'] = ps
        if not ps:
            self._status('未发现打印机: 请把《打印机诊断日志.txt》'
                         '发给软件开发者排查')
            return
        if self.printer_var.get() in ps:
            return  # 保留上次使用(设置里记忆)的打印机
        d = default_printer()
        self.printer_var.set(d if d in ps else ps[0])

    def _isolated(self, items):
        """当前选择是否触发了购买方隔离"""
        return len({it.buyer for it in items if it.buyer}) >= 2

    def _wait_ocr_done(self):
        """隔离依赖购买方识别结果, 识别未完成时给导出/打印一个等待机会"""
        if not self._ocr_running:
            return
        if not messagebox.askyesno(
                '购买方识别中',
                '发票购买方识别尚未完成, 现在继续可能无法完整隔离。\n'
                '是否等待识别完成后再继续?(通常几十秒)'):
            return
        t0 = time.time()
        self.root.config(cursor='watch')
        try:
            while self._ocr_running and time.time() - t0 < 180:
                self.root.update()
                time.sleep(0.05)
        finally:
            self.root.config(cursor='')

    def show_page_preview(self):
        items = self.selected_items()
        if not items:
            messagebox.showinfo('提示', '请先勾选要打印的图片。')
            return
        self._wait_ocr_done()
        try:
            self._status('正在渲染排版预览 ...')
            self.root.config(cursor='watch')
            self.root.update()
            PagePreviewWindow(self.root, items, self.current_opts(),
                              self.side_var.get())
            self._status('排版预览已打开')
        except Exception:
            messagebox.showerror('预览失败', traceback.format_exc())
            self._status('预览失败')
        finally:
            self.root.config(cursor='')

    def export_pdf(self):
        items = self.selected_items()
        if not items:
            messagebox.showinfo('提示', '请先勾选要打印的图片。')
            return
        self._wait_ocr_done()
        path = filedialog.asksaveasfilename(
            title='导出 PDF', defaultextension='.pdf',
            initialfile='报销凭证-%s.pdf' % time.strftime('%Y%m%d-%H%M%S'),
            filetypes=[('PDF', '*.pdf')])
        if not path:
            return
        opts = self.current_opts()
        try:
            self.root.config(cursor='watch')
            self._status('正在生成 PDF ...')
            self.root.update()
            pages = list(iter_pages(items, opts, PDF_DPI))
            pages[0].save(path, save_all=True, append_images=pages[1:],
                          resolution=float(PDF_DPI))
            self._status('已导出: %s (%d 页%s)' % (
                path, len(pages),
                ' · 已按购买方隔离' if self._isolated(items) else ''))
        except Exception:
            messagebox.showerror('导出失败', traceback.format_exc())
            self._status('导出失败')
        finally:
            self.root.config(cursor='')

    def start_print(self):
        items = self.selected_items()
        if not items:
            messagebox.showinfo('提示', '请先勾选要打印的图片。')
            return
        self._wait_ocr_done()
        printer = self.printer_var.get()
        if not printer:
            messagebox.showwarning('提示', '未找到可用打印机。')
            return
        opts = self.current_opts()
        plans = plan_pages(items, opts)
        duplex = self.side_var.get() == '双面'
        copies = max(1, int(self.copies_var.get()))
        if duplex and not printer_supports_duplex(printer):
            if not messagebox.askyesno(
                    '打印机可能不支持双面',
                    '系统报告 [%s] 不支持自动双面打印。\n'
                    '继续则按打印机默认方式输出, 是否继续?' % printer):
                return
        sheets = (len(plans) + 1) // 2 if duplex else len(plans)
        per_sheet = _LAYOUT_SLOTS[opts['layout']] * (2 if duplex else 1)
        need = -(-len(items) // per_sheet)
        iso = ('\n已按购买方隔离: 不同购买方的发票不会印在同一张纸上'
               if sheets > need else '')
        msg = ('打印机: %s\n图片: %d 张 → %d 页\n'
               '方式: %s%s · %s\n纸张: %s · 份数: %d\n预计消耗: 约 %d 张纸%s\n\n确认打印?'
               % (printer, len(items), len(plans),
                  self.side_var.get(), '' if self.layout_var.get() == 'full'
                  else {'half': ' · 二分之一排版', 'quarter': ' · 四分之一排版',
                        'eighth': ' · 八分之一排版'}[self.layout_var.get()],
                  self.color_var.get(),
                  opts['paper'], copies, sheets * copies, iso))
        if not messagebox.askyesno('确认打印', msg):
            return

        self.print_btn.config(state='disabled')
        t = threading.Thread(target=self._print_worker,
                             args=(plans, opts, printer, duplex, copies,
                                   self.color_var.get()),
                             daemon=True)
        t.start()

    def _print_worker(self, plans, opts, printer, duplex, copies, color='彩色'):
        def progress(done, total):
            self.root.after(0, self._status,
                            '正在发送 %d/%d 页到 [%s] ...' % (done, total, printer))
        try:
            if IS_MAC:
                duplex_ok, color_ok = mac_print(plans, opts, printer, duplex,
                                                copies, progress, color)
            else:
                duplex_ok, color_ok = gdi_print(plans, opts, printer, duplex,
                                                copies, progress, color)
            notes = []
            if duplex and not duplex_ok:
                notes.append('打印机未接受双面设置, 实际单面输出')
            if not color_ok:
                notes.append('打印机未接受%s设置' %
                             ('黑白' if color == '黑白' else '彩色'))
            note = ('(注意: %s)' % '; '.join(notes)) if notes else ''
            self.root.after(0, self._status,
                            '打印完成: %d 页已发送至 %s %s'
                            % (len(plans) * copies, printer, note))
            if notes:
                self.root.after(0, lambda n='; '.join(notes):
                                messagebox.showwarning('打印设置未完全生效',
                                                       '注意: ' + n))
        except Exception:
            err = traceback.format_exc()
            self.root.after(0, lambda: messagebox.showerror('打印失败', err))
            self.root.after(0, self._status, '打印失败')
        finally:
            self.root.after(0, lambda: self.print_btn.config(state='normal'))


def run_selftest():
    """--selftest: 验证依赖与打印机(打包 exe 后快速自检用)"""
    lines = ['平台: %s' % ('macOS (打印走 CUPS)' if IS_MAC else 'Windows')]
    for label, modname in (('Pillow', 'PIL'), ('openpyxl', 'openpyxl'),
                           ('pypdfium2', 'pypdfium2')):
        try:
            __import__(modname)
            lines.append('%s OK' % label)
        except Exception as e:
            lines.append('%s FAIL: %s' % (label, e))
    if IS_WIN:
        try:   # 打印管线: 纯 ctypes, 不再依赖 pywin32
            ctypes.WinDLL('winspool.drv')
            ctypes.windll.gdi32
            lines.append('打印管线(winspool+GDI) OK')
        except Exception as e:
            lines.append('打印管线 FAIL: %s' % e)
        try:   # .xls 老格式转换用的 COM 组件(可选, 缺失只影响 .xls)
            __import__('win32com.client')
            lines.append('XLS转换组件 OK')
        except Exception as e:
            lines.append('XLS转换组件 不可用(%s), 仅影响 .xls 自动转换' % e)
    else:
        lp = shutil.which('lp')
        lines.append('打印管线(CUPS lp) %s' % ('OK: ' + lp if lp
                                              else '未找到 lp 命令'))
        lines.append('XLS转换组件: mac 版不支持 .xls, 请另存为 .xlsx')
    try:   # 拖放支持: 需真正加载 tkdnd 扩展
        from tkinterdnd2 import TkinterDnD, DND_FILES
        r = TkinterDnD.Tk()
        r.withdraw()
        r.drop_target_register(DND_FILES)
        r.destroy()
        lines.append('tkinterdnd2 OK')
    except Exception as e:
        lines.append('tkinterdnd2 FAIL: %s' % e)
    try:   # 内置备用金报销模版
        src = embedded_template_path()
        if src:
            lines.append('内置模板 OK (%d 字节)' % os.path.getsize(src))
        else:
            lines.append('内置模板 FAIL: 未找到 %s' % TEMPLATE_FILE)
    except Exception as e:
        lines.append('内置模板 FAIL: %s' % e)
    try:
        from pillow_heif import register_heif_opener  # noqa
        register_heif_opener()
        lines.append('pillow-heif OK (图片格式支持 %d 种)'
                     % len(_image_extensions()))
    except Exception as e:
        lines.append('pillow-heif FAIL: %s' % e)
    ps = list_printers()
    lines.append('打印机 %d 台: %s' % (len(ps), ', '.join(ps) or '(无)'))
    out = os.path.join(tempfile.gettempdir(), '发票打印助手_自检.txt')
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('\n'.join(lines))
    return 0 if not any('FAIL' in x for x in lines) else 1


def main():
    if '--selftest' in sys.argv:
        sys.exit(run_selftest())
    # 高分屏适配: 声明 DPI 感知(必须在创建任何窗口之前)
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # per-monitor
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    if HAS_DND:
        try:
            root = TkinterDnD.Tk()
        except Exception as ex:
            _dnd_log('TkinterDnD.Tk() 失败, 回退普通 Tk(拖放不可用): %r' % ex)
            root = tk.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
