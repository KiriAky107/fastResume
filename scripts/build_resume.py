#!/usr/bin/env python3
"""Build an editable resume from local YAML/JSON; no network or AI service needed."""
from __future__ import annotations

import argparse
import base64
import binascii
import copy
import importlib.util
import io
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
from urllib.parse import urlsplit

import yaml
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageOps
from pypdf import PdfReader


class ResumeError(ValueError):
    """A readable configuration or export error."""


DEFAULT_THEME = {
    "paper": "A4", "accent": "#0D7187", "ink": "#13242D",
    "muted": "#52636D", "background": "#EFF8FA", "border": "#D3E2E7",
    "font": "Calibri", "east_asia_font": "Microsoft YaHei",
    "body_size": 8.5, "line_height": 12.6,
    "margin_mm": {"top": 10, "right": 14, "bottom": 9, "left": 14},
}
TITLES = {
    "profile": "求职摘要", "education": "教育背景", "skills": "专业技能",
    "work_experiences": "工作经历", "project_experiences": "项目经历",
    "internship_experiences": "实习经历",
    "certificates": "竞赛奖项与证书",
}
BASE_FIELDS = ("name", "position", "subtitle", "eyebrow", "tagline", "phone",
               "email", "location", "status", "details", "photo")
PHOTO_LIMIT = 2 * 1024 * 1024


def _mapping(value, path):
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ResumeError(f"{path}: 必须是对象 / mapping")
    return value


def _list(value, path):
    if value is None:
        return []
    if not isinstance(value, list):
        raise ResumeError(f"{path}: 必须是列表")
    return value


def _text(value, path):
    if value is None:
        return ""
    if not isinstance(value, str):
        raise ResumeError(f"{path}: 必须是文本；日期、电话请加引号")
    # XML 1.0 cannot represent these characters.
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]", value):
        raise ResumeError(f"{path}: 包含不支持的控制字符")
    return value.strip()


def _texts(value, path):
    texts = [_text(v, f"{path}[{i}]") for i, v in enumerate(_list(value, path))]
    return [text for text in texts if text]


def _fields(value, fields, path):
    obj = _mapping(value, path)
    return {key: _text(obj.get(key, ""), f"{path}.{key}") for key in fields}


def _url(value, path):
    value = _text(value, path)
    if not value:
        return ""
    try:
        parts = urlsplit(value)
        if parts.scheme.lower() not in {"https", "http", "mailto", "tel"}:
            raise ValueError()
        if parts.scheme.lower() in {"https", "http"} and not parts.hostname:
            raise ValueError()
        if not parts.netloc and not parts.path:
            raise ValueError()
        if any(c.isspace() for c in value):
            raise ValueError()
    except ValueError:
        raise ResumeError(f"{path}: 请使用完整的 http/https/mailto/tel 链接") from None
    return value


def _number(value, path, low, high):
    if (isinstance(value, bool) or not isinstance(value, (int, float))
            or not math.isfinite(value) or not low <= value <= high):
        raise ResumeError(f"{path}: 必须是 {low}–{high} 之间的有限数值")
    return float(value)


def _photo(value, path="basics.photo"):
    value = _text(value, path)
    if not value:
        return ""
    match = re.fullmatch(r"data:image/(png|jpeg|webp);base64,([A-Za-z0-9+/=]+)", value)
    if not match or len(match[2]) > (PHOTO_LIMIT + 2) // 3 * 4:
        raise ResumeError(f"{path}: 请上传不超过 2 MB 的 PNG / JPEG / WebP 图片")
    try:
        content = base64.b64decode(match[2], validate=True)
        if len(content) > PHOTO_LIMIT:
            raise ValueError("too large")
        with Image.open(io.BytesIO(content)) as picture:
            if picture.format not in {"PNG", "JPEG", "WEBP"} or picture.width * picture.height > 20_000_000:
                raise ValueError("unsupported image")
            picture.verify()
    except (ValueError, OSError, SyntaxError, binascii.Error, Image.DecompressionBombError) as exc:
        raise ResumeError(f"{path}: 图片无效或像素过大，请换一张清晰的小图") from exc
    return value


def _photo_file(path):
    path = Path(path)
    try:
        if path.stat().st_size > PHOTO_LIMIT:
            raise ResumeError("照片不能超过 2 MB")
        content = path.read_bytes()
        with Image.open(io.BytesIO(content)) as picture:
            mime = {"JPEG": "jpeg", "PNG": "png", "WEBP": "webp"}.get(picture.format)
        if not mime:
            raise ResumeError("请使用 PNG / JPEG / WebP 照片")
        return _photo("data:image/" + mime + ";base64," + base64.b64encode(content).decode("ascii"))
    except (OSError, ValueError, SyntaxError, Image.DecompressionBombError) as exc:
        raise ResumeError(f"无法读取照片：{exc}") from exc


def normalize_resume(raw):
    """Validate supported fields and return a new, complete configuration."""
    raw = _mapping(raw, "resume")
    result = {"basics": _fields(raw.get("basics", {}), BASE_FIELDS, "basics")}
    if not result["basics"]["name"]:
        raise ResumeError("basics.name: 不能为空")
    result["basics"]["photo"] = _photo(result["basics"]["photo"])
    focus = _list(_mapping(raw.get("basics"), "basics").get("photo_focus"), "basics.photo_focus") or [0.5, 0.5]
    if len(focus) != 2:
        raise ResumeError("basics.photo_focus: 需要 [水平焦点, 垂直焦点] 两个 0–1 数值")
    result["basics"]["photo_focus"] = [_number(v, f"basics.photo_focus[{i}]", 0, 1) for i, v in enumerate(focus)]
    result["code"] = []
    for i, item in enumerate(_list(raw.get("code", []), "code")):
        path = f"code[{i}]"
        item = _mapping(item, path)
        link = _fields(item, ("label", "icon"), path)
        link["url"] = _url(item.get("url", ""), path + ".url")
        if link["label"] or link["url"]:
            result["code"].append(link)
    result["profile"] = _text(raw.get("profile", ""), "profile")
    result["education"] = [
        _fields(v, ("school", "major", "date", "detail"), f"education[{i}]")
        for i, v in enumerate(_list(raw.get("education", []), "education"))
    ]
    skill_raw = _mapping(raw.get("skills", {}), "skills")
    result["skills"] = {
        "summary": _text(skill_raw.get("summary", ""), "skills.summary"),
        "items": _texts(skill_raw.get("items", []), "skills.items"),
        "groups": [_fields(v, ("label", "value"), f"skills.groups[{i}]")
                   for i, v in enumerate(_list(skill_raw.get("groups", []), "skills.groups"))],
    }
    for key in ("internship_experiences", "work_experiences", "project_experiences"):
        result[key] = []
        for i, item in enumerate(_list(raw.get(key, []), key)):
            path = f"{key}[{i}]"
            item = _mapping(item, path)
            record = _fields(item, ("name", "date", "role", "description", "tech_stack"), path)
            record["url"] = _url(item.get("url", ""), path + ".url")
            record["highlights"] = _texts(item.get("highlights", []), path + ".highlights")
            if any(record.values()):
                result[key].append(record)
    result["certificates"] = []
    for i, item in enumerate(_list(raw.get("certificates", []), "certificates")):
        path = f"certificates[{i}]"
        result["certificates"].append(
            _text(item, path) if isinstance(item, str)
            else _fields(item, ("name", "result"), path)
        )
    result["extra_sections"] = []
    for i, item in enumerate(_list(raw.get("extra_sections", []), "extra_sections")):
        path = f"extra_sections[{i}]"
        item = _mapping(item, path)
        entry = _fields(item, ("title",), path)
        entry["items"] = _texts(item.get("items", []), path + ".items")
        result["extra_sections"].append(entry)
    result["section_titles"] = TITLES.copy()
    title_raw = _mapping(raw.get("section_titles", {}), "section_titles")
    for key in TITLES:
        if key in title_raw:
            result["section_titles"][key] = _text(title_raw[key], f"section_titles.{key}") or TITLES[key]
    theme_raw = _mapping(raw.get("theme", {}), "theme")
    theme = copy.deepcopy(DEFAULT_THEME)
    for key in ("accent", "ink", "muted", "background", "border"):
        value = _text(theme_raw.get(key, theme[key]), f"theme.{key}")
        if not re.fullmatch(r"#[0-9a-fA-F]{6}", value):
            raise ResumeError(f"theme.{key}: 需要 #RRGGBB 颜色")
        theme[key] = value.upper()
    for key in ("font", "east_asia_font", "paper"):
        theme[key] = _text(theme_raw.get(key, theme[key]), f"theme.{key}")
        if not theme[key]:
            raise ResumeError(f"theme.{key}: 不能为空")
    if theme["paper"] not in {"A4", "Letter"}:
        raise ResumeError("theme.paper: 仅支持 A4 或 Letter")
    theme["body_size"] = _number(theme_raw.get("body_size", theme["body_size"]), "theme.body_size", 7, 18)
    theme["line_height"] = _number(theme_raw.get("line_height", theme["line_height"]), "theme.line_height", 9, 36)
    if theme["line_height"] < theme["body_size"] * 1.1:
        raise ResumeError("theme.line_height: 至少为 body_size 的 1.1 倍，避免正文被裁切")
    margin_raw = _mapping(theme_raw.get("margin_mm", {}), "theme.margin_mm")
    for side in theme["margin_mm"]:
        theme["margin_mm"][side] = _number(margin_raw.get(side, theme["margin_mm"][side]), f"theme.margin_mm.{side}", 5, 45)
    result["theme"] = theme
    return result


def load_resume(path):
    path = Path(path)
    try:
        source = path.read_text(encoding="utf-8-sig")
        if path.suffix.lower() == ".json":
            raw = json.loads(source)
        elif path.suffix.lower() in {".yml", ".yaml"}:
            raw = yaml.safe_load(source)
        else:
            raise ResumeError("配置文件需要 .yaml / .yml / .json 后缀")
        if isinstance(raw, dict) and isinstance(raw.get("basics"), dict):
            photo = raw["basics"].get("photo")
            if isinstance(photo, str) and photo.strip() and not photo.startswith("data:"):
                photo_path = Path(photo).expanduser()
                if not photo_path.is_absolute():
                    photo_path = path.parent / photo_path
                raw = copy.deepcopy(raw)
                raw["basics"]["photo"] = _photo_file(photo_path)
        return normalize_resume(raw)
    except (OSError, UnicodeError, yaml.YAMLError, json.JSONDecodeError) as exc:
        raise ResumeError(f"读取配置失败：{exc}") from exc


def _xml(parent, name, **attrs):
    element = OxmlElement("w:" + name)
    for key, value in attrs.items():
        element.set(qn("w:" + key), str(value))
    parent.append(element)
    return element


class ResumeDocument:
    """Two-column hero, spaced skill cards and bordered experience cards."""

    def __init__(self, data):
        self.data = data
        self.theme = t = data["theme"]
        self.doc = Document()
        # Start with neutral styles even if a distributor customizes default.docx.
        for style in self.doc.styles:
            p_pr = style.element.find(qn("w:pPr"))
            if p_pr is not None:
                for border in list(p_pr.findall(qn("w:pBdr"))):
                    p_pr.remove(border)
        self.doc.core_properties.author = "fastResume"
        self.doc.core_properties.last_modified_by = "fastResume"
        self.doc.core_properties.title = data["basics"]["name"] + " Resume"
        self.doc.core_properties.subject = "Local configurable resume"
        self.doc.core_properties.comments = ""
        section = self.doc.sections[0]
        section.page_width, section.page_height = (Mm(210), Mm(297)) if t["paper"] == "A4" else (Inches(8.5), Inches(11))
        for side, value in t["margin_mm"].items():
            setattr(section, side + "_margin", Mm(value))
        section.header_distance = Mm(4)
        section.footer_distance = Mm(5)
        for name in ("Normal", "Title", "Heading 1", "Heading 2"):
            style = self.doc.styles[name]
            style.font.name = t["font"]
            style.font.size = Pt(t["body_size"])
            style.font.color.rgb = RGBColor.from_string(t["ink"][1:])
            style.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), t["east_asia_font"])
            style.paragraph_format.space_before = Pt(0)
            style.paragraph_format.space_after = Pt(0)
            style.paragraph_format.line_spacing = Pt(t["line_height"])
        normal = self.doc.styles["Normal"].paragraph_format
        normal.widow_control = True
        settings = self.doc.settings.element
        _xml(settings, "autoHyphenation", val="0")
        self.width_mm = section.page_width.mm - section.left_margin.mm - section.right_margin.mm

    @staticmethod
    def tint(color, white=0.5):
        rgb = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
        return "#" + "".join(f"{round(value + (255 - value) * white):02X}" for value in rgb)

    def table(self, widths, parent=None):
        if parent is None:
            table = self.doc.add_table(rows=1, cols=len(widths))
        else:
            table = parent.add_table(rows=1, cols=len(widths))
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _xml(table._tbl.tblPr, "tblInd", w=0, type="dxa")
        width = table._tbl.tblPr.find(qn("w:tblW"))
        width.set(qn("w:type"), "dxa")
        width.set(qn("w:w"), str(round(sum(widths) / 25.4 * 1440)))
        borders = _xml(table._tbl.tblPr, "tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            _xml(borders, side, val="nil")
        for index, mm in enumerate(widths):
            table.columns[index].width = Mm(mm)
            table.cell(0, index).width = Mm(mm)
        _xml(table.rows[0]._tr.get_or_add_trPr(), "cantSplit")
        return table

    def cell(self, cell, *, fill=None, padding=(0, 0, 0, 0), border=False, left=0):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        properties = cell._tc.get_or_add_tcPr()
        if fill:
            _xml(properties, "shd", fill=fill[1:], val="clear")
        margins = _xml(properties, "tcMar")
        for side, value in zip(("top", "right", "bottom", "left"), padding):
            _xml(margins, side, w=round(value * 20), type="dxa")
        edges = _xml(properties, "tcBorders")
        for side in ("top", "right", "bottom", "left"):
            _xml(edges, side, val="single" if border else "nil", sz=4,
                 color=self.theme["border"][1:])
        if left:
            edge = edges.find(qn("w:left"))
            edge.set(qn("w:val"), "single")
            edge.set(qn("w:sz"), str(round(left * 8)))
            edge.set(qn("w:color"), self.tint(self.theme["accent"], 0.45)[1:])

    def cell_p(self, cell, *, first=False, line=None, before=0, after=0, keep=False):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.line_spacing = Pt(line or self.theme["line_height"])
        p.paragraph_format.keep_with_next = keep
        return p

    def gap(self, points=3):
        p = self.paragraph(after=0)
        p.paragraph_format.line_spacing = Pt(points)
        self.run(p, "", size=1)
        p.paragraph_format.space_before = Pt(0)
        return p

    def run(self, p, text, *, bold=False, size=None, color=None):
        r = p.add_run(text)
        r.bold = bold
        r.font.name = self.theme["font"]
        r.font.size = Pt(size or self.theme["body_size"])
        r.font.color.rgb = RGBColor.from_string((color or self.theme["ink"])[1:])
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), self.theme["east_asia_font"])
        return r

    def rich(self, p, text, *, prefix=False, **kwargs):
        if prefix and "**" not in text:
            match = re.match(r"^([^：\n]{1,20}：)(.*)$", text, re.S)
            if match:
                self.run(p, match[1], bold=True, **kwargs)
                self.run(p, match[2], **kwargs)
                return
        for i, part in enumerate(re.split(r"\*\*(.*?)\*\*", text, flags=re.S)):
            if part:
                self.run(p, part, bold=bool(i % 2), **kwargs)

    def paragraph(self, text="", *, after=2.5, before=0, keep=False, **kwargs):
        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        p.paragraph_format.keep_with_next = keep
        if text:
            self.rich(p, text, **kwargs)
        return p

    def icon(self, p, kind, color=None, size=0.13, badge=False):
        """Original, locally drawn line icons; no font or remote asset dependency."""
        color = "#FFFFFF" if badge else (color or self.theme["accent"])
        im = Image.new("RGBA", (96, 96))
        d = ImageDraw.Draw(im)
        kw = {"fill": color, "width": 6}
        if kind == "mail":
            d.rounded_rectangle((12, 23, 84, 74), radius=5, outline=color, width=6)
            d.line([(14, 25), (48, 52), (82, 25)], **kw)
        elif kind == "phone":
            d.rounded_rectangle((28, 10, 68, 86), radius=8, outline=color, width=6)
            d.line((40, 73, 56, 73), **kw)
        elif kind == "location":
            d.ellipse((24, 10, 72, 58), outline=color, width=6)
            d.line([(26, 45), (48, 85), (70, 45)], **kw)
            d.ellipse((40, 26, 56, 42), outline=color, width=5)
        elif kind in {"status", "calendar"}:
            d.rounded_rectangle((15, 21, 81, 82), radius=5, outline=color, width=6)
            d.line((16, 39, 80, 39), **kw)
            d.line((32, 12, 32, 28), **kw)
            d.line((64, 12, 64, 28), **kw)
            d.line([(32, 58), (44, 69), (66, 49)], **kw)
        elif kind == "education":
            d.polygon([(8, 35), (48, 14), (88, 35), (48, 56)], outline=color, width=5)
            d.line([(24, 46), (24, 68), (48, 80), (72, 68), (72, 46)], **kw)
        elif kind in {"projects", "work"}:
            d.rounded_rectangle((13, 24, 83, 81), radius=6, outline=color, width=6)
            d.line([(33, 24), (33, 13), (63, 13), (63, 24)], **kw)
            d.line((14, 46, 82, 46), **kw)
            d.line((43, 44, 53, 44), fill=color, width=11)
        elif kind == "profile":
            d.ellipse((35, 9, 61, 35), outline=color, width=6)
            d.arc((19, 37, 77, 100), 180, 360, fill=color, width=6)
            d.line((19, 70, 77, 70), **kw)
        elif kind == "awards":
            d.polygon([(48, 9), (59, 32), (84, 35), (65, 53), (70, 80), (48, 67), (26, 80), (31, 53), (12, 35), (37, 32)], outline=color, width=5)
        elif kind in {"code", "skills", "github"}:
            d.line([(33, 22), (11, 48), (33, 74)], **kw)
            d.line([(63, 22), (85, 48), (63, 74)], **kw)
            d.line((56, 15, 40, 81), **kw)
        else:
            d.ellipse((12, 12, 84, 84), outline=color, width=6)
            d.ellipse((32, 12, 64, 84), outline=color, width=5)
            d.line((13, 48, 83, 48), **kw)
        if badge:
            tile = Image.new("RGBA", (96, 96))
            ImageDraw.Draw(tile).rounded_rectangle((0, 0, 95, 95), radius=16, fill=self.theme["accent"])
            tile.alpha_composite(im.resize((72, 72), Image.Resampling.LANCZOS), (12, 12))
            im = tile
        stream = io.BytesIO()
        im.save(stream, format="PNG")
        stream.seek(0)
        p.add_run().add_picture(stream, width=Inches(size))
        self.run(p, "  ")

    def link(self, p, label, url, **kwargs):
        if not url:
            return self.run(p, label, **kwargs)
        relation = self.doc.part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
        el = OxmlElement("w:hyperlink")
        el.set(qn("r:id"), relation)
        run = self.run(p, label, **kwargs)
        p._p.remove(run._r)
        el.append(run._r)
        p._p.append(el)

    def heading(self, key, icon):
        p = self.paragraph(after=3, before=5.5, keep=True)
        p.style = self.doc.styles["Heading 1"]
        p.paragraph_format.line_spacing = Pt(16)
        self.icon(p, icon, size=0.17, badge=True)
        self.run(p, self.data["section_titles"].get(key, key), bold=True, size=11, color=self.theme["accent"])

    def header(self):
        b = self.data["basics"]
        has_photo = bool(b["photo"])
        right_width = self.width_mm * 83 / 182
        photo_width = self.width_mm * 29 / 182 if has_photo else 0
        left_width = self.width_mm - right_width - photo_width
        hero = self.table((photo_width, left_width, right_width) if has_photo else (left_width, right_width))
        if has_photo:
            photo_cell, left, right = hero.rows[0].cells
            self.cell(photo_cell, fill=self.theme["background"], padding=(6.25, 3.75, 6, 6.25), border=True, left=2.75)
            edges = photo_cell._tc.get_or_add_tcPr().find(qn("w:tcBorders"))
            edges.find(qn("w:left")).set(qn("w:color"), self.theme["accent"][1:])
            edges.find(qn("w:right")).set(qn("w:val"), "nil")
            p = self.cell_p(photo_cell, first=True)
            p.paragraph_format.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            raw = base64.b64decode(b["photo"].split(",", 1)[1])
            with Image.open(io.BytesIO(raw)) as image:
                foreground = ImageOps.exif_transpose(image).convert("RGBA")
                image = Image.new("RGBA", foreground.size, "white")
                image.alpha_composite(foreground)
                image = image.convert("RGB")
                visible_width = min(image.width, image.height * 3 / 4)
                visible_height = min(image.height, image.width * 4 / 3)
                centering = []
                for full, visible, focus in zip(image.size, (visible_width, visible_height), b["photo_focus"]):
                    centering.append(0.5 if full == visible else max(0, min(1, (focus * full - visible / 2) / (full - visible))))
                image = ImageOps.fit(image, (360, 480), Image.Resampling.LANCZOS, centering=tuple(centering))
                stream = io.BytesIO()
                image.save(stream, format="JPEG", quality=95)
            stream.seek(0)
            p.add_run().add_picture(stream, width=Mm(24), height=Mm(32))
        else:
            left, right = hero.rows[0].cells
        self.cell(left, fill=self.theme["background"], padding=(6.25, 4, 6, 9.5 if not has_photo else 3.75), border=True, left=0 if has_photo else 2.75)
        self.cell(right, fill=self.theme["background"], padding=(6.25, 6.5, 6, 3.75), border=True)
        left._tc.get_or_add_tcPr().find(qn("w:tcBorders")).find(qn("w:right")).set(qn("w:val"), "nil")
        right._tc.get_or_add_tcPr().find(qn("w:tcBorders")).find(qn("w:left")).set(qn("w:val"), "nil")
        left_edge = left._tc.get_or_add_tcPr().find(qn("w:tcBorders")).find(qn("w:left"))
        if has_photo:
            left_edge.set(qn("w:val"), "nil")
        else:
            left_edge.set(qn("w:color"), self.theme["accent"][1:])
        first = True
        if b["eyebrow"]:
            p = self.cell_p(left, first=True, line=8, after=4)
            r = self.run(p, b["eyebrow"], bold=True, size=6.5, color=self.theme["muted"])
            _xml(r._r.get_or_add_rPr(), "spacing", val=22)
            first = False
        p = self.cell_p(left, first=first, line=27, after=4)
        p.style = self.doc.styles["Title"]
        self.run(p, b["name"], bold=True, size=24)
        if b["position"]:
            p = self.cell_p(left, line=15, after=2)
            self.run(p, b["position"], bold=True, size=10.5, color=self.theme["accent"])
        if b["subtitle"]:
            p = self.cell_p(left, line=9)
            self.run(p, b["subtitle"], bold=True, size=6.5, color=self.theme["muted"])
        contacts = []
        for key in ("phone", "email"):
            if b[key]:
                url = ("tel:" if key == "phone" else "mailto:") + b[key]
                try:
                    url = _url(url, "contact")
                except ResumeError:
                    url = ""
                contacts.append(("mail" if key == "email" else key, b[key], url, False))
        contacts.extend((c["icon"] or "globe", c["label"] or c["url"], c["url"], False) for c in self.data["code"])
        contacts.extend((key, b[key], "", True) for key in ("location", "status") if b[key])
        if contacts:
            inner_width = right_width - (6.5 + 3.75) / 72 * 25.4
            grid = self.table((inner_width * 38 / 79, inner_width * 41 / 79), parent=right)
            for index in range(0, len(contacts), 2):
                row = grid.rows[0] if index == 0 else grid.add_row()
                for offset, cell in enumerate(row.cells):
                    self.cell(cell, padding=(0, 0.5, 0, 0))
                    p = self.cell_p(cell, first=True, line=14)
                    if index + offset < len(contacts):
                        icon, text, url, bold = contacts[index + offset]
                        self.icon(p, icon, size=0.10)
                        self.link(p, text, url, size=8, bold=bold,
                                  color=self.theme["accent"] if bold else self.theme["muted"])
            # Nested tables require an ending paragraph; keep it unobtrusive.
            for p in (right.paragraphs[0], right.paragraphs[-1]):
                p.paragraph_format.line_spacing = Pt(1)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
        else:
            right.paragraphs[0].paragraph_format.line_spacing = Pt(1)
        if b["details"]:
            p = self.cell_p(right, line=10.2, before=2)
            self.run(p, b["details"], size=7, color=self.theme["muted"])
        if b["tagline"]:
            p = self.cell_p(right, line=11, before=2)
            self.run(p, b["tagline"], size=7.5, bold=True, color=self.theme["accent"])

    def education(self):
        entries = [e for e in self.data["education"] if any(e.values())]
        if not entries:
            return
        self.heading("education", "education")
        for entry in entries:
            date_width = self.width_mm * 33 / 182
            major_width = self.width_mm * 66 / 182
            has_header = any(entry[key] for key in ("date", "school", "major"))
            if has_header:
                table = self.table((date_width, self.width_mm - date_width - major_width, major_width))
                for cell in table.rows[0].cells:
                    self.cell(cell)
                for index, (key, size, bold, color) in enumerate((
                    ("date", 8, True, self.theme["muted"]),
                    ("school", 9.5, True, self.theme["ink"]),
                    ("major", 8.5, False, self.theme["muted"]),
                )):
                    p = self.cell_p(table.cell(0, index), first=True, line=13)
                    if index == 2:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    self.run(p, entry[key], size=size, bold=bold, color=color)
            if entry["detail"]:
                p = self.paragraph(entry["detail"], after=2, size=8, color=self.theme["muted"])
                p.paragraph_format.left_indent = Mm(date_width if has_header else 0)
                p.paragraph_format.line_spacing = Pt(11.5)

    def skills(self):
        skills = self.data["skills"]
        groups = [v for v in skills["groups"] if v["label"] or v["value"]]
        if not (groups or skills["summary"] or any(skills["items"])):
            return
        self.heading("skills", "skills")
        if skills["summary"]:
            self.paragraph(skills["summary"])
        for index in range(0, len(groups), 2):
            if index:
                self.gap(3)
            card_width = (self.width_mm - 3) / 2
            table = self.table((card_width, 3, card_width))
            for cell in table.rows[0].cells:
                self.cell(cell)
            for offset, cell_index in ((0, 0), (1, 2)):
                cell = table.cell(0, cell_index)
                if index + offset >= len(groups):
                    continue
                group = groups[index + offset]
                self.cell(cell, fill=self.tint(self.theme["background"], 0.5),
                          padding=(2.75, 4.5, 2.75, 4.75), border=True, left=1.875)
                if group["label"]:
                    p = self.cell_p(cell, first=True, line=11, after=2)
                    self.run(p, group["label"], bold=True, size=8.5, color=self.theme["accent"])
                if group["value"]:
                    p = self.cell_p(cell, first=not group["label"], line=11.4)
                    self.rich(p, group["value"], size=8)
        for text in skills["items"]:
            self.paragraph("• " + text)

    def projects(self, key):
        if not self.data[key]:
            return
        self.heading(key, "projects" if key == "project_experiences" else "work")
        for index, entry in enumerate(self.data[key]):
            if index:
                self.gap(3)
            card = self.table((self.width_mm,))
            cell = card.cell(0, 0)
            self.cell(cell, fill=self.tint(self.theme["background"], 0.88),
                      padding=(3.25, 5.75, 3.25, 6.25), border=True, left=2)
            first = True
            title = entry["name"] or entry["url"]
            if title or entry["date"]:
                p = self.cell_p(cell, first=True, line=14.3)
                first = False
                p.paragraph_format.tab_stops.add_tab_stop(Mm(self.width_mm) - Pt(12), WD_TAB_ALIGNMENT.RIGHT)
                if title:
                    self.link(p, title + (" ↗" if entry["url"] else ""), entry["url"], bold=True, size=9.5)
                if entry["date"]:
                    self.run(p, "\t" + entry["date"], bold=True, size=8, color=self.theme["muted"])
            if entry["role"]:
                p = self.cell_p(cell, first=first, line=10.8, before=1, after=1.8)
                first = False
                self.run(p, entry["role"], size=7.5, color=self.theme["muted"])
            if entry["description"]:
                p = self.cell_p(cell, first=first, line=11.8, after=1.8)
                first = False
                label = {"project_experiences": "项目简述", "internship_experiences": "实习内容"}.get(key, "工作内容")
                self.run(p, label + "  ", bold=True, size=8, color=self.theme["accent"])
                self.rich(p, entry["description"], size=8)
            for text in entry["highlights"]:
                p = self.cell_p(cell, first=first, line=self.theme["line_height"], after=1.4)
                first = False
                p.paragraph_format.left_indent = Pt(8)
                p.paragraph_format.first_line_indent = Pt(-8)
                self.run(p, "• ", color=self.theme["accent"])
                self.rich(p, text, prefix=True)
            if entry["tech_stack"]:
                p = self.cell_p(cell, first=first, line=10.5, before=1.2)
                self.run(p, "技术栈  ", size=7.5, bold=True, color=self.theme["accent"])
                self.run(p, entry["tech_stack"], size=7.5, color=self.theme["muted"])

    def awards(self):
        entries = [item for item in self.data["certificates"] if (item if isinstance(item, str) else any(item.values()))]
        if not entries:
            return
        self.heading("certificates", "awards")
        table = self.table((self.width_mm * 0.59, self.width_mm * 0.41))
        for index, entry in enumerate(entries):
            row = table.rows[0] if index == 0 else table.add_row()
            if isinstance(entry, str) or not (entry["name"] and entry["result"]):
                cell = row.cells[0].merge(row.cells[1])
                self.cell(cell, fill=self.tint(self.theme["background"], 0.5), padding=(2, 4, 2, 4), border=True)
                p = self.cell_p(cell, first=True, line=11.5)
                result_only = isinstance(entry, dict) and bool(entry["result"])
                if result_only:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                value = entry if isinstance(entry, str) else (entry["name"] or entry["result"])
                if result_only:
                    self.run(p, value, size=8, bold=True, color=self.theme["accent"])
                else:
                    self.rich(p, value, size=8)
            else:
                for cell_index, key in enumerate(("name", "result")):
                    cell = row.cells[cell_index]
                    self.cell(cell, fill=self.tint(self.theme["background"], 0.5), padding=(2, 4, 2, 4), border=True)
                    p = self.cell_p(cell, first=True, line=11.5)
                    if cell_index:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    self.run(p, entry[key], size=8, bold=bool(cell_index),
                             color=self.theme["accent"] if cell_index else self.theme["ink"])

    def build(self, path):
        self.header()
        if self.data["profile"]:
            self.heading("profile", "profile")
            p = self.paragraph(self.data["profile"], after=0)
            p.paragraph_format.left_indent = Pt(3)
        self.education()
        self.skills()
        self.projects("internship_experiences")
        self.projects("work_experiences")
        self.projects("project_experiences")
        self.awards()
        for section in self.data["extra_sections"]:
            if any(section["items"]):
                self.heading(section["title"] or "补充信息", "code")
                for text in section["items"]:
                    self.paragraph(text, after=2)
        self.doc.save(path)


def _convert_pdf(docx, pdf, backend="auto", office_path=None, timeout=120, workdir=None):
    path = Path(__file__).with_name("pdf_export.py")
    spec = importlib.util.spec_from_file_location("fastresume_pdf_export", path)
    if spec is None or spec.loader is None:
        raise ResumeError("找不到 scripts/pdf_export.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    try:
        module.convert_pdf(Path(docx), Path(pdf), backend, office_path, timeout, workdir)
    except (OSError, RuntimeError) as exc:
        raise ResumeError(str(exc)) from exc


def generate_resume(data, output, formats=("docx", "pdf"), pdf_backend="auto",
                    office_path=None, max_pages=None, force=False, timeout=120):
    data = normalize_resume(data)
    if not formats or any(f not in {"docx", "pdf"} for f in formats):
        raise ResumeError("formats: 至少选择 docx 或 pdf")
    formats = tuple(dict.fromkeys(formats))
    if pdf_backend not in {"auto", "word", "libreoffice"}:
        raise ResumeError("pdf_backend: 请选择 auto / word / libreoffice")
    timeout = _number(timeout, "timeout", 1, 3600)
    if max_pages is not None:
        if isinstance(max_pages, bool) or not isinstance(max_pages, int) or max_pages < 1:
            raise ResumeError("max_pages: 需要正整数")
        if "pdf" not in formats:
            raise ResumeError("max_pages: 必须同时生成 PDF")
    prefix = Path(output).expanduser().absolute()
    if prefix.suffix.lower() in {".docx", ".pdf"}:
        prefix = prefix.with_suffix("")
    targets = {f: Path(str(prefix) + "." + f) for f in formats}
    for path in targets.values():
        if path.is_dir() or path.is_symlink():
            raise ResumeError(f"输出不是普通文件：{path}")
        if path.exists() and not force:
            raise ResumeError(f"文件已存在：{path}；更换名称或使用 --force")
    prefix.parent.mkdir(parents=True, exist_ok=True)
    pages = None
    try:
        with tempfile.TemporaryDirectory(prefix=".fastresume-", dir=prefix.parent) as temporary:
            staging = Path(temporary)
            docx = staging / "resume.docx"
            pdf = staging / "resume.pdf"
            ResumeDocument(data).build(docx)
            if "pdf" in formats:
                _convert_pdf(docx, pdf, backend=pdf_backend, office_path=office_path,
                             timeout=timeout, workdir=staging)
                if not pdf.is_file() or not pdf.stat().st_size:
                    raise ResumeError("PDF 转换未生成有效文件")
                pages = len(PdfReader(pdf).pages)
                if pages == 0:
                    raise ResumeError("PDF 没有页面")
                if max_pages is not None and pages > max_pages:
                    raise ResumeError(f"PDF 共 {pages} 页，超出 --max-pages {max_pages}；请精简内容或调整主题")
            # Stage validated bytes. Back up only outputs explicitly replaced by --force.
            backups = {}
            published = []
            try:
                for fmt, target in targets.items():
                    if target.exists():
                        if not force:
                            raise ResumeError(f"输出在生成期间已出现：{target}")
                        backup = staging / ("previous." + fmt)
                        shutil.copy2(target, backup)
                        backups[target] = backup
                    staged = docx if fmt == "docx" else pdf
                    if force:
                        published.append(target)
                        os.replace(staged, target)
                    else:
                        # Exclusive create closes the preflight race without overwriting a file.
                        with target.open("xb") as stream:
                            published.append(target)
                            stream.write(staged.read_bytes())
                        continue
            except BaseException:
                recovered = []
                for target in reversed(published):
                    if target in backups:
                        try:
                            os.replace(backups[target], target)
                        except OSError as restore_error:
                            # A locked target must not cause its original bytes to be
                            # discarded when the staging directory is cleaned.
                            descriptor, recovery = tempfile.mkstemp(
                                prefix=target.stem + ".recovered-", suffix=target.suffix,
                                dir=target.parent,
                            )
                            os.close(descriptor)
                            shutil.copy2(backups[target], recovery)
                            recovered.append((recovery, restore_error))
                    else:
                        target.unlink(missing_ok=True)
                if recovered:
                    paths = "；".join(path for path, _ in recovered)
                    raise ResumeError(f"恢复原文件失败；原件备份已保留：{paths}") from recovered[0][1]
                raise
    except ResumeError:
        raise
    except Exception as exc:
        raise ResumeError(f"生成失败：{exc}") from exc
    return {**{fmt: str(path) for fmt, path in targets.items()}, "pages": pages}


def main(argv=None):
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=root / "examples/resume.template.yaml", help="本地 YAML / JSON 配置")
    parser.add_argument("--output", type=Path, default=Path("output/resume"), help="输出文件名前缀")
    parser.add_argument("--formats", nargs="+", choices=("docx", "pdf"), default=["docx"], help="默认仅 docx，无需 Office")
    parser.add_argument("--pdf-backend", choices=("auto", "word", "libreoffice"), default="auto")
    parser.add_argument("--office-path", type=Path, help="LibreOffice soffice 可执行文件路径")
    parser.add_argument("--photo", type=Path, help="可选证件照 PNG / JPEG / WebP，最多 2 MB")
    parser.add_argument("--timeout", type=float, default=120, help="单次 PDF 转换超时秒数")
    parser.add_argument("--max-pages", type=int, help="PDF 最大页数，超限不发布输出")
    parser.add_argument("--force", action="store_true", help="替换本次指定的已有输出")
    parser.add_argument("--validate-only", action="store_true", help="只验证配置，不生成文件")
    args = parser.parse_args(argv)
    try:
        data = load_resume(args.source)
        if args.photo:
            data["basics"]["photo"] = _photo_file(args.photo)
            data["basics"]["photo_focus"] = [0.5, 0.5]
        if args.validate_only:
            print("配置检查通过")
            return 0
        result = generate_resume(data, args.output, args.formats, args.pdf_backend,
                                 args.office_path, args.max_pages, args.force, args.timeout)
        for fmt in ("docx", "pdf"):
            if fmt in result:
                print(f"{fmt.upper()}: {result[fmt]}")
        if result["pages"] is not None:
            print(f"PDF: {result['pages']} 页")
        return 0
    except ResumeError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
