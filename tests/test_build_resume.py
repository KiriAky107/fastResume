"""Behavior checks for the configurable resume exporter; no Office is required."""

from __future__ import annotations

import base64
import copy
import contextlib
import importlib.util
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree

sys.dont_write_bytecode = True

import yaml
from PIL import Image
from pypdf import PdfWriter

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "easycv_build_resume_under_test", ROOT / "scripts" / "build_resume.py"
)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("Cannot load the resume exporter")
generator = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = generator
SPEC.loader.exec_module(generator)


def sample_resume() -> dict:
    """Use fictional labels and reserved example contact details only."""
    return {
        "basics": {
            "name": "匿名键盘手",
            "position": "软件开发实习生",
            "email": "member@example.com",
            "phone": "",
            "status": "档期可议",
            "tagline": "可以迷路，调用链不能迷路。",
        },
        "code": [{"label": "作品站", "url": "https://example.com/portfolio"}],
        "profile": "希望承担接口开发、联调、测试与交付工作。",
        "education": [
            {"school": "示例学院甲", "major": "计算机相关专业", "date": "20XX - 至今"},
            {"school": "示例学院乙", "major": "通用课程", "date": "20XX - 20XX"},
        ],
        "skills": {
            "groups": [
                {"label": "技能甲", "value": "Python"},
                {"label": "技能乙", "value": "TypeScript"},
                {"label": "技能丙", "value": "Git"},
            ]
        },
        "project_experiences": [
            {
                "name": f"示例项目{index}",
                "date": "20XX - 至今",
                "role": "核心开发",
                "url": f"https://example.com/project-{index}",
                "description": f"项目{index}的用途与服务对象。",
                "highlights": [f"功能实现：完成示例功能{index}。"],
                "tech_stack": "Python · SQLite",
            }
            for index in range(1, 6)
        ],
        "work_experiences": [
            {
                "name": "示例工作室",
                "role": "示例职位",
                "description": "虚构的任职条目。",
                "highlights": ["完成接口开发与联调。"],
            }
        ],
        "certificates": ["示例证书", {"name": "示例竞赛", "result": "示例奖项"}],
        "extra_sections": [{"title": "协作偏好", "items": ["清楚记录接口约定。"]}],
    }


def docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("word/document.xml"))
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    return "\n".join(node.text or "" for node in root.findall(".//w:t", namespace))


class ResumeExporterTests(unittest.TestCase):
    def assert_no_staging_left(self, root: Path, output: Path) -> None:
        allowed_directories = {
            parent for parent in (output.parent, *output.parent.parents) if parent != root
        }
        leftovers = [
            str(path.relative_to(root))
            for path in root.rglob("*")
            if path.is_file() or path not in allowed_directories
        ]
        self.assertEqual([], leftovers, "Failed exports must not leave generated files")

    def test_yaml_and_json_have_equivalent_content(self) -> None:
        source = sample_resume()
        with tempfile.TemporaryDirectory(prefix="easycv-config-check-") as temporary:
            root = Path(temporary)
            yaml_path = root / "resume.yaml"
            json_path = root / "resume.json"
            yaml_path.write_text(
                yaml.safe_dump(source, allow_unicode=True, sort_keys=False), encoding="utf-8"
            )
            json_path.write_text(json.dumps(source, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(
                generator.load_resume(yaml_path), generator.load_resume(json_path)
            )

    def test_load_resolves_photo_relative_to_source_without_changing_source_files(self) -> None:
        raw = {
            "basics": {
                "name": "匿名候选人", "photo": "../images/avatar.png",
                "photo_focus": [0.25, 0.75],
            }
        }
        with tempfile.TemporaryDirectory(prefix="easycv-relative-photo-check-") as temporary:
            root = Path(temporary)
            config_dir, image_dir = root / "config", root / "images"
            config_dir.mkdir()
            image_dir.mkdir()
            source = config_dir / "resume.yaml"
            photo = image_dir / "avatar.png"
            source.write_text(yaml.safe_dump(raw, allow_unicode=True), encoding="utf-8")
            Image.new("RGB", (6, 8), color=(80, 130, 150)).save(photo, format="PNG")
            original_source, original_photo = source.read_bytes(), photo.read_bytes()

            loaded = generator.load_resume(source)

            self.assertTrue(loaded["basics"]["photo"].startswith("data:image/png;base64,"))
            embedded = base64.b64decode(loaded["basics"]["photo"].split(",", 1)[1], validate=True)
            self.assertEqual(original_photo, embedded)
            self.assertEqual([0.25, 0.75], loaded["basics"]["photo_focus"])
            self.assertEqual(original_source, source.read_bytes())
            self.assertEqual(original_photo, photo.read_bytes())
            self.assertEqual({config_dir, image_dir, source, photo}, set(root.rglob("*")))
        self.assertEqual(
            [0.5, 0.5],
            generator.normalize_resume({"basics": {"name": "匿名候选人"}})["basics"]["photo_focus"],
        )

    def test_docx_retains_arbitrary_counts_and_optional_sections(self) -> None:
        source = sample_resume()
        untouched = copy.deepcopy(source)
        with tempfile.TemporaryDirectory(prefix="easycv-docx-check-") as temporary:
            prefix = Path(temporary) / "anonymous"
            with patch.object(generator, "_convert_pdf") as convert:
                generator.generate_resume(source, prefix, formats=("docx",))
            convert.assert_not_called()
            output = prefix.with_suffix(".docx")
            self.assertTrue(output.is_file())
            text = docx_text(output)
            for label in [
                "匿名键盘手",
                "示例学院甲",
                "示例学院乙",
                "技能甲",
                "技能乙",
                "技能丙",
                "示例工作室",
                "示例证书",
                "示例竞赛",
                "示例奖项",
                "协作偏好",
                "可以迷路，调用链不能迷路。",
                *[f"示例项目{index}" for index in range(1, 6)],
            ]:
                with self.subTest(label=label):
                    self.assertIn(label, text)
            self.assertEqual({output}, set(Path(temporary).rglob("*")))
        self.assertEqual(untouched, source, "Export must not mutate the caller's data")

    def test_empty_optional_sections_do_not_produce_empty_headings(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-minimal-check-") as temporary:
            prefix = Path(temporary) / "minimal"
            generator.generate_resume(
                {"basics": {"name": "匿名候选人"}}, prefix, formats=("docx",)
            )
            text = docx_text(prefix.with_suffix(".docx"))
            self.assertIn("匿名候选人", text)
            for empty_heading in ["求职摘要", "教育背景", "专业技能", "实习经历", "项目经历", "工作经历", "竞赛奖项与证书"]:
                with self.subTest(heading=empty_heading):
                    self.assertNotIn(empty_heading, text)

    def test_internships_and_optional_details_are_retained_in_section_order(self) -> None:
        source = sample_resume()
        source["basics"]["details"] = "角色资料：虚构设定 · 档期可议"
        source["internship_experiences"] = [
            {
                "name": f"示例实习单位{index}",
                "date": f"20X{index}.01 - 20X{index}.06",
                "role": f"示例实习岗位{index}",
                "url": f"https://example.com/internship-{index}",
                "description": f"虚构实习职责说明{index}。",
                "highlights": [f"实现记录：完成虚构接口{index}。"],
                "tech_stack": f"Python · 演示工具{index}",
            }
            for index in range(1, 5)
        ]
        with tempfile.TemporaryDirectory(prefix="easycv-internship-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            with patch.object(generator, "_convert_pdf") as convert:
                generator.generate_resume(source, prefix, formats=("docx",))
            convert.assert_not_called()
            output = prefix.with_suffix(".docx")
            text = docx_text(output)
            self.assertIn(source["basics"]["details"], text)
            self.assertEqual(1, text.count("实习经历"))
            self.assertLess(text.index("专业技能"), text.index("实习经历"))
            self.assertLess(text.index("实习经历"), text.index("项目经历"))
            with zipfile.ZipFile(output) as archive:
                relationships = archive.read("word/_rels/document.xml.rels").decode("utf-8")
            for entry in source["internship_experiences"]:
                for field in ("name", "date", "role", "description", "tech_stack"):
                    with self.subTest(name=entry["name"], field=field):
                        self.assertIn(entry[field], text)
                self.assertIn(entry["highlights"][0].split("：", 1)[1], text)
                self.assertIn(entry["url"], relationships)
            self.assertEqual({output}, set(root.rglob("*")))

    def test_explicitly_blank_records_do_not_produce_empty_headings(self) -> None:
        source = {
            "basics": {
                "name": "匿名候选人",
                **{key: " \t " for key in (
                    "position", "subtitle", "eyebrow", "tagline", "phone", "email",
                    "location", "status", "details", "photo",
                )},
            },
            "code": [{"label": " ", "icon": " ", "url": " "}, {"icon": "github"}],
            "profile": " \n ",
            "education": [{"school": " ", "major": " ", "date": " ", "detail": " "}, {}],
            "skills": {
                "summary": " ", "items": ["", " \t "],
                "groups": [{"label": " ", "value": " "}, {}],
            },
            "internship_experiences": [{"highlights": [" ", ""]}],
            "project_experiences": [{"highlights": [" ", ""]}],
            "work_experiences": [{"description": " "}],
            "certificates": [{}, " ", {"name": " ", "result": " "}],
            "extra_sections": [{"title": "空补充", "items": [" "]}],
        }
        with tempfile.TemporaryDirectory(prefix="easycv-blank-check-") as temporary:
            root = Path(temporary)
            prefix = Path(temporary) / "minimal"
            generator.generate_resume(source, prefix, formats=("docx",))
            output = prefix.with_suffix(".docx")
            text = docx_text(output)
            for empty_heading in ["求职摘要", "教育背景", "专业技能", "实习经历", "项目经历", "工作经历", "竞赛奖项与证书", "空补充"]:
                with self.subTest(heading=empty_heading):
                    self.assertNotIn(empty_heading, text)
            self.assertEqual("匿名候选人", text.strip())
            with zipfile.ZipFile(output) as archive:
                document = ElementTree.fromstring(archive.read("word/document.xml"))
            namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            # The name header remains; empty sections and contacts add no cards or icons.
            self.assertEqual(1, len(document.findall(".//w:tbl", namespace)))
            self.assertEqual([], document.findall(".//w:hyperlink", namespace))
            self.assertEqual([], document.findall(".//w:drawing", namespace))
            self.assertEqual({output}, set(root.rglob("*")))

    def test_partial_records_do_not_display_labels_for_missing_values(self) -> None:
        source = {
            "basics": {"name": "匿名候选人"},
            "internship_experiences": [{"name": "示例实习标题"}],
            "work_experiences": [{"name": "示例任职标题"}],
            "project_experiences": [{"name": "示例项目标题"}],
            "certificates": [{"name": "示例证书名称"}, {"result": "示例获奖结果"}],
        }
        with tempfile.TemporaryDirectory(prefix="easycv-partial-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            generator.generate_resume(source, prefix, formats=("docx",))
            output = prefix.with_suffix(".docx")
            text = docx_text(output)
            for visible in [
                "示例实习标题", "示例任职标题", "示例项目标题",
                "示例证书名称", "示例获奖结果", "竞赛奖项与证书",
            ]:
                with self.subTest(visible=visible):
                    self.assertEqual(1, text.count(visible))
            for missing_label in ["技术栈", "项目简述", "实习内容", "工作内容"]:
                with self.subTest(label=missing_label):
                    self.assertNotIn(missing_label, text)
            self.assertEqual({output}, set(root.rglob("*")))

    def test_optional_null_sections_and_null_records_are_hidden(self) -> None:
        null_sections = {
            "basics": {"name": "匿名候选人", "details": None, "photo": None},
            **{key: None for key in (
                "code", "profile", "education", "skills", "internship_experiences",
                "work_experiences", "project_experiences", "certificates",
                "extra_sections", "section_titles", "theme",
            )},
        }
        null_records = {
            "basics": {"name": "匿名候选人"},
            "code": [None, {}],
            "education": [None, {}],
            "skills": {"summary": None, "items": None, "groups": [None, {}]},
            "internship_experiences": [None, {"highlights": None}],
            "work_experiences": [None, {}],
            "project_experiences": [None, {"description": None}],
            "certificates": [None, {}, " "],
            "extra_sections": [None, {"title": "空补充", "items": None}],
            "theme": {"margin_mm": None},
        }
        for label, source in (("sections", null_sections), ("records", null_records)):
            with self.subTest(case=label):
                with tempfile.TemporaryDirectory(prefix="easycv-null-check-") as temporary:
                    root = Path(temporary)
                    prefix = root / "anonymous"
                    generator.generate_resume(source, prefix, formats=("docx",))
                    output = prefix.with_suffix(".docx")
                    self.assertEqual("匿名候选人", docx_text(output).strip())
                    with zipfile.ZipFile(output) as archive:
                        document = ElementTree.fromstring(archive.read("word/document.xml"))
                    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    self.assertEqual(1, len(document.findall(".//w:tbl", ns)))
                    self.assertEqual({output}, set(root.rglob("*")))

    def test_partial_values_do_not_leave_empty_card_paragraphs(self) -> None:
        source = {
            "basics": {"name": "匿名候选人"},
            "education": [{"detail": "示例补充教育详情"}],
            "skills": {"groups": [{"value": "示例技能内容"}]},
            "internship_experiences": [{"description": "示例实习职责"}],
            "work_experiences": [{"role": "示例任职职责"}],
            "project_experiences": [{"tech_stack": "示例技术内容"}],
            "certificates": [{"name": "示例证书名称"}, {"result": "示例获奖结果"}],
        }
        with tempfile.TemporaryDirectory(prefix="easycv-partial-value-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            generator.generate_resume(source, prefix, formats=("docx",))
            output = prefix.with_suffix(".docx")
            text = docx_text(output)
            contents = [
                "示例补充教育详情", "示例技能内容", "示例实习职责", "示例任职职责",
                "示例技术内容", "示例证书名称", "示例获奖结果",
            ]
            for content in contents:
                self.assertIn(content, text)
            self.assertNotIn("项目简述", text)
            self.assertNotIn("工作内容", text)
            with zipfile.ZipFile(output) as archive:
                document = ElementTree.fromstring(archive.read("word/document.xml"))
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            def visible_text(element):
                return "".join(node.text or "" for node in element.findall(".//w:t", ns)).strip()

            tables = document.findall(".//w:tbl", ns)
            for table in tables[1:]:
                self.assertTrue(visible_text(table), "Optional fields must not create an empty card")
            for content in ("示例技能内容", "示例实习职责", "示例任职职责", "示例技术内容"):
                matching = [table for table in tables if content in visible_text(table)]
                self.assertEqual(1, len(matching))
                content_cells = [
                    cell for cell in matching[0].findall("w:tr/w:tc", ns)
                    if content in visible_text(cell)
                ]
                self.assertEqual(1, len(content_cells))
                for paragraph in content_cells[0].findall(".//w:p", ns):
                    self.assertTrue(visible_text(paragraph), f"Blank paragraph before {content}")
            for row in document.findall(".//w:tr", ns):
                if visible_text(row) in {"示例证书名称", "示例获奖结果"}:
                    self.assertEqual(1, len(row.findall("w:tc", ns)))
            self.assertEqual({output}, set(root.rglob("*")))

    def test_embedded_photo_is_optional_and_does_not_mutate_input(self) -> None:
        image = io.BytesIO()
        Image.new("RGB", (6, 8), color=(80, 130, 150)).save(image, format="PNG")
        photo = "data:image/png;base64," + base64.b64encode(image.getvalue()).decode("ascii")
        source = {"basics": {"name": "匿名候选人", "photo": photo}}
        untouched = copy.deepcopy(source)
        with tempfile.TemporaryDirectory(prefix="easycv-photo-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            generator.generate_resume(source, prefix, formats=("docx",))
            output = prefix.with_suffix(".docx")
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            with zipfile.ZipFile(output) as archive:
                document = ElementTree.fromstring(archive.read("word/document.xml"))
                media = [name for name in archive.namelist() if name.startswith("word/media/")]
            hero = document.find("w:body/w:tbl", ns)
            self.assertIsNotNone(hero)
            self.assertEqual(3, len(hero.findall("w:tr/w:tc", ns)))
            self.assertEqual(1, len(media))
            self.assertEqual(1, len(document.findall(".//w:drawing", ns)))
            self.assertNotIn("data:image", docx_text(output))
            self.assertEqual(untouched, source)

            without_photo = copy.deepcopy(source)
            without_photo["basics"]["photo"] = ""
            generator.generate_resume(without_photo, prefix, formats=("docx",), force=True)
            with zipfile.ZipFile(output) as archive:
                document = ElementTree.fromstring(archive.read("word/document.xml"))
                self.assertFalse(any(name.startswith("word/media/") for name in archive.namelist()))
            hero = document.find("w:body/w:tbl", ns)
            self.assertIsNotNone(hero)
            self.assertEqual(2, len(hero.findall("w:tr/w:tc", ns)))
            self.assertEqual([], document.findall(".//w:drawing", ns))
            self.assertEqual(untouched, source)
            self.assertEqual({output}, set(root.rglob("*")))

    def test_photo_rejects_paths_urls_and_invalid_image_data(self) -> None:
        image = io.BytesIO()
        Image.new("RGB", (6, 8)).save(image, format="PNG")
        bad_crc = bytearray(image.getvalue())
        chunk = bad_crc.index(b"IDAT")
        length = int.from_bytes(bad_crc[chunk - 4:chunk], "big")
        bad_crc[chunk + 4 + length] ^= 1
        for label, value in [
            ("local-path", r"C:\example\avatar.png"),
            ("remote-url", "https://example.com/avatar.png"),
            ("malformed-base64", "data:image/png;base64,!!!!"),
            ("not-an-image", "data:image/png;base64," + base64.b64encode(b"not an image").decode("ascii")),
            ("unsupported-svg", "data:image/svg+xml;base64,PHN2Zy8+"),
            ("invalid-png-checksum", "data:image/png;base64," + base64.b64encode(bad_crc).decode("ascii")),
        ]:
            with self.subTest(case=label), self.assertRaises(generator.ResumeError):
                generator.normalize_resume({"basics": {"name": "匿名候选人", "photo": value}})

    def test_dotted_output_prefix_is_preserved(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-prefix-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous.v2"
            result = generator.generate_resume(sample_resume(), prefix, formats=("docx",))
            expected = root / "anonymous.v2.docx"
            self.assertEqual(expected, Path(result["docx"]))
            self.assertIn("匿名键盘手", docx_text(expected))
            self.assertEqual({expected}, set(root.rglob("*")))

    def test_unsafe_urls_nonfinite_theme_and_invalid_types_are_rejected(self) -> None:
        cases = []
        unsafe_link = sample_resume()
        unsafe_link["code"][0]["url"] = "javascript:alert('unsafe')"
        cases.append(("unsafe-link", unsafe_link))
        unsafe_project = sample_resume()
        unsafe_project["project_experiences"][0]["url"] = "file:///private/example"
        cases.append(("unsafe-project", unsafe_project))
        invalid_size = sample_resume()
        invalid_size["theme"] = {"body_size": float("nan")}
        cases.append(("nan-font", invalid_size))
        invalid_margin = sample_resume()
        invalid_margin["theme"] = {"margin_mm": {"top": float("inf")}}
        cases.append(("infinite-margin", invalid_margin))
        invalid_line_height = sample_resume()
        invalid_line_height["theme"] = {"body_size": 18, "line_height": 9}
        cases.append(("clipping-line-height", invalid_line_height))
        invalid_skills = sample_resume()
        invalid_skills["skills"]["groups"] = "not a list"
        cases.append(("skill-type", invalid_skills))
        invalid_details = sample_resume()
        invalid_details["basics"]["details"] = ["not plain text"]
        cases.append(("details-type", invalid_details))
        for label, focus in [
            ("focus-too-short", [0.5]),
            ("focus-too-long", [0.0, 0.5, 1.0]),
            ("focus-not-finite", [float("nan"), 0.5]),
            ("focus-below-range", [-0.1, 0.5]),
            ("focus-above-range", [0.5, 1.1]),
            ("focus-not-a-list", "center"),
            ("focus-bool", [True, 0.5]),
        ]:
            invalid_focus = sample_resume()
            invalid_focus["basics"]["photo_focus"] = focus
            cases.append((label, invalid_focus))
        invalid_internships = sample_resume()
        invalid_internships["internship_experiences"] = "not a list"
        cases.append(("internship-list-type", invalid_internships))
        invalid_internship_highlights = sample_resume()
        invalid_internship_highlights["internship_experiences"] = [{"highlights": [123]}]
        cases.append(("internship-highlight-type", invalid_internship_highlights))
        invalid_highlights = sample_resume()
        invalid_highlights["project_experiences"][0]["highlights"] = [123]
        cases.append(("highlight-type", invalid_highlights))
        invalid_paper = sample_resume()
        invalid_paper["theme"] = {"paper": "Poster"}
        cases.append(("unknown-paper", invalid_paper))
        for label, data in cases:
            with self.subTest(case=label), self.assertRaises(generator.ResumeError):
                generator.normalize_resume(data)

    def test_existing_output_is_not_overwritten_by_default(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-collision-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            existing = prefix.with_suffix(".docx")
            original = b"existing document must remain byte-for-byte intact"
            existing.write_bytes(original)
            with self.assertRaises(generator.ResumeError):
                generator.generate_resume(sample_resume(), prefix, formats=("docx",))
            self.assertEqual(original, existing.read_bytes())
            self.assertEqual({existing}, set(root.rglob("*")))

    def test_all_targets_are_checked_before_any_conversion_or_publish(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-preflight-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            existing = prefix.with_suffix(".pdf")
            original = b"existing PDF"
            existing.write_bytes(original)
            with patch.object(generator, "_convert_pdf") as convert:
                with self.assertRaises(generator.ResumeError):
                    generator.generate_resume(sample_resume(), prefix, formats=("docx", "pdf"))
            convert.assert_not_called()
            self.assertEqual(original, existing.read_bytes())
            self.assertFalse(prefix.with_suffix(".docx").exists())
            self.assertEqual({existing}, set(root.rglob("*")))

    def test_pdf_conversion_failure_does_not_publish_or_leave_staging(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-failure-check-") as temporary:
            root = Path(temporary)
            prefix = root / "exports" / "anonymous"
            with patch.object(
                generator, "_convert_pdf", side_effect=generator.ResumeError("模拟转换失败")
            ) as convert:
                with self.assertRaises(generator.ResumeError):
                    generator.generate_resume(sample_resume(), prefix, formats=("docx", "pdf"))
            convert.assert_called_once()
            self.assert_no_staging_left(root, prefix)

    def test_forced_export_conversion_failure_preserves_existing_outputs(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-force-check-") as temporary:
            root = Path(temporary)
            prefix = root / "anonymous"
            originals = {
                prefix.with_suffix(".docx"): b"previous editable document",
                prefix.with_suffix(".pdf"): b"previous PDF document",
            }
            for path, content in originals.items():
                path.write_bytes(content)
            with patch.object(
                generator, "_convert_pdf", side_effect=generator.ResumeError("模拟转换失败")
            ) as convert:
                with self.assertRaises(generator.ResumeError):
                    generator.generate_resume(
                        sample_resume(), prefix, formats=("docx", "pdf"), force=True
                    )
            convert.assert_called_once()
            for path, content in originals.items():
                self.assertEqual(content, path.read_bytes())
            self.assertEqual(set(originals), set(root.rglob("*")))

    def test_mid_publication_failure_and_interrupt_restore_forced_outputs(self) -> None:
        def one_page_pdf(docx, pdf, **kwargs):
            writer = PdfWriter()
            writer.add_blank_page(width=595.28, height=841.89)
            with Path(pdf).open("wb") as stream:
                writer.write(stream)

        real_replace = generator.os.replace
        for failure in (OSError("模拟发布失败"), KeyboardInterrupt()):
            with self.subTest(failure=type(failure).__name__):
                with tempfile.TemporaryDirectory(prefix="easycv-rollback-check-") as temporary:
                    root = Path(temporary)
                    prefix = root / "anonymous"
                    originals = {
                        prefix.with_suffix(".docx"): b"previous editable document",
                        prefix.with_suffix(".pdf"): b"previous PDF document",
                    }
                    for path, content in originals.items():
                        path.write_bytes(content)

                    def fail_second_publish(source, target):
                        if Path(source).name == "resume.pdf":
                            if isinstance(failure, KeyboardInterrupt):
                                # A signal can arrive after the OS has replaced the
                                # file but before the caller resumes bookkeeping.
                                real_replace(source, target)
                            raise failure
                        return real_replace(source, target)

                    expected_error = (
                        KeyboardInterrupt if isinstance(failure, KeyboardInterrupt)
                        else generator.ResumeError
                    )
                    with patch.object(generator, "_convert_pdf", side_effect=one_page_pdf) as convert:
                        with patch.object(generator.os, "replace", side_effect=fail_second_publish) as replace:
                            with self.assertRaises(expected_error):
                                generator.generate_resume(
                                    sample_resume(), prefix, formats=("docx", "pdf"), force=True
                                )
                    convert.assert_called_once()
                    self.assertTrue(any(
                        Path(call.args[0]).name == "resume.pdf"
                        for call in replace.call_args_list
                    ), "The failure must occur while publishing the second artifact")
                    for path, content in originals.items():
                        self.assertEqual(content, path.read_bytes())
                    self.assertEqual(set(originals), set(root.rglob("*")))

    def test_cli_rejects_invalid_parameters_before_writing_outputs(self) -> None:
        invalid_arguments = [
            ["--timeout", "nan"],
            ["--timeout", "inf"],
            ["--timeout", "0"],
            ["--max-pages", "0"],
            ["--max-pages", "1"],
            ["--formats", "html"],
        ]
        with tempfile.TemporaryDirectory(prefix="easycv-cli-check-") as temporary:
            root = Path(temporary)
            source = root / "source.json"
            source.write_text(json.dumps(sample_resume(), ensure_ascii=False), encoding="utf-8")
            for arguments in invalid_arguments:
                with self.subTest(arguments=arguments):
                    stdout, stderr = io.StringIO(), io.StringIO()
                    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                        with patch.object(generator, "_convert_pdf") as convert:
                            try:
                                status = generator.main([
                                    "--source", str(source), "--output", str(root / "out" / "resume"),
                                    *arguments,
                                ])
                            except SystemExit as exc:
                                status = exc.code
                    self.assertEqual(2, status)
                    self.assertTrue(stderr.getvalue().strip())
                    convert.assert_not_called()
                    self.assertEqual({source}, set(root.rglob("*")))

    def test_cli_validate_only_reads_configuration_without_rendering(self) -> None:
        with tempfile.TemporaryDirectory(prefix="easycv-validate-check-") as temporary:
            root = Path(temporary)
            source = root / "source.json"
            source.write_text(json.dumps(sample_resume(), ensure_ascii=False), encoding="utf-8")
            stdout, stderr = io.StringIO(), io.StringIO()
            with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
                with patch.object(generator, "generate_resume") as generate:
                    status = generator.main([
                        "--source", str(source), "--output", str(root / "out" / "resume"),
                        "--formats", "pdf", "--pdf-backend", "word", "--validate-only",
                    ])
            self.assertEqual(0, status)
            self.assertIn("配置检查通过", stdout.getvalue())
            self.assertEqual("", stderr.getvalue())
            generate.assert_not_called()
            self.assertEqual({source}, set(root.rglob("*")))

    def test_page_limit_failure_does_not_publish_or_leave_staging(self) -> None:
        def two_page_pdf(docx, pdf, backend="auto", office_path=None, timeout=120, workdir=None):
            writer = PdfWriter()
            writer.add_blank_page(width=595.28, height=841.89)
            writer.add_blank_page(width=595.28, height=841.89)
            with Path(pdf).open("wb") as stream:
                writer.write(stream)

        with tempfile.TemporaryDirectory(prefix="easycv-pages-check-") as temporary:
            root = Path(temporary)
            prefix = root / "exports" / "anonymous"
            with patch.object(generator, "_convert_pdf", side_effect=two_page_pdf) as convert:
                with self.assertRaises(generator.ResumeError):
                    generator.generate_resume(
                        sample_resume(), prefix, formats=("docx", "pdf"), max_pages=1
                    )
            convert.assert_called_once()
            self.assert_no_staging_left(root, prefix)


if __name__ == "__main__":
    unittest.main()
