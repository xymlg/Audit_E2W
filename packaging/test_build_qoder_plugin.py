from __future__ import annotations

import os
from pathlib import Path
import stat
import subprocess
import tempfile
import unittest
import zipfile

import build_agent_skills as skills
import build_qoder_plugin as build


class BuildPackagingTests(unittest.TestCase):
    def setUp(self) -> None:
        build.TMP.mkdir(exist_ok=True)
        self.work = Path(tempfile.mkdtemp(prefix="test_build_", dir=build.TMP))
        self.original_dist = build.DIST
        build.DIST = self.work

    def tearDown(self) -> None:
        build.DIST = self.original_dist
        build.shutil.rmtree(self.work, ignore_errors=True)

    def test_actual_plugin_source_matches_explicit_whitelist(self) -> None:
        files = build.validate_tree(build.PLUGIN_ROOT, set(build.PLUGIN_SOURCE_FILES))
        self.assertEqual(len(files), 21)

    def test_qoder_skill_sync_does_not_copy_caches(self) -> None:
        original_root, original_plugin = skills.SKILL_ROOT, skills.PLUGIN_SKILL
        skills.SKILL_ROOT = self.work / "skill-source"
        skills.PLUGIN_SKILL = self.work / "plugin-skill"
        try:
            cache = skills.SKILL_ROOT / "scripts" / "__pycache__"
            cache.mkdir(parents=True)
            (skills.SKILL_ROOT / "SKILL.md").write_text("skill", encoding="utf-8")
            (cache / "audit_tool.pyc").write_bytes(b"cache")
            skills.sync_qoder_skill()
            self.assertTrue((skills.PLUGIN_SKILL / "SKILL.md").is_file())
            self.assertFalse((skills.PLUGIN_SKILL / "scripts" / "__pycache__").exists())
        finally:
            skills.SKILL_ROOT, skills.PLUGIN_SKILL = original_root, original_plugin

    def test_tree_rejects_unlisted_and_forbidden_content(self) -> None:
        root = self.work / "tree"
        root.mkdir()
        (root / "README.md").write_text("ok", encoding="utf-8")
        (root / "extra.txt").write_text("no", encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "白名单"):
            build.validate_tree(root, {"README.md"})

        cases = (
            "客户.docx", "~$客户.xlsx", ".~lock.报告.docx#", "_backup/报告.bak",
            "nested/.git/config", "__pycache__/x.pyc", "cache.lock", "Thumbs.db",
        )
        for index, relative in enumerate(cases):
            with self.subTest(relative=relative):
                case = self.work / f"forbidden-{index}"
                path = case / relative
                path.parent.mkdir(parents=True)
                path.write_bytes(b"bad")
                with self.assertRaises(RuntimeError):
                    build.validate_tree(case)

    def test_only_bundled_docx_template_is_allowed(self) -> None:
        root = self.work / "runtime"
        template = root / build.RUNTIME_OFFICE_TEMPLATE
        template.parent.mkdir(parents=True)
        template.write_bytes(b"template")
        self.assertEqual(build.validate_tree(root), [template])
        (root / "客户.xlsx").write_bytes(b"customer")
        with self.assertRaisesRegex(RuntimeError, "Office"):
            build.validate_tree(root)

    def test_tree_rejects_junction(self) -> None:
        target = self.work / "target"
        target.mkdir()
        junction = self.work / "junction"
        result = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(junction), str(target)],
            capture_output=True,
        )
        if result.returncode != 0:
            self.skipTest("当前系统不允许创建 junction")
        with self.assertRaisesRegex(RuntimeError, "junction"):
            build.validate_tree(self.work)

    def test_tree_rejects_symlink(self) -> None:
        target = self.work / "target"
        target.mkdir()
        try:
            os.symlink(target / "missing", self.work / "symlink")
        except OSError:
            self.skipTest("当前系统不允许创建符号链接")
        with self.assertRaisesRegex(RuntimeError, "符号链接"):
            build.validate_tree(self.work)

    def test_zip_checks_exact_whitelist_paths_and_links(self) -> None:
        archive = self.work / "plugin.zip"
        expected = set(build.PLUGIN_SOURCE_FILES) | {
            "engine/审计工具箱引擎.exe", f"engine/{build.RUNTIME_OFFICE_TEMPLATE}",
        }
        with zipfile.ZipFile(archive, "w") as zf:
            for name in sorted(expected):
                zf.writestr(name, b"ok")
        self.assertEqual(build.validate_archive(archive, expected), expected)
        with zipfile.ZipFile(archive, "a") as zf:
            zf.writestr("客户.docx", b"customer")
        with self.assertRaisesRegex(RuntimeError, "Office"):
            build.validate_archive(archive, expected)

        with zipfile.ZipFile(archive, "w") as zf:
            for name in sorted(expected):
                zf.writestr(name, b"ok")
            info = zipfile.ZipInfo("engine/link.exe")
            info.create_system = 3
            info.external_attr = (stat.S_IFLNK | 0o777) << 16
            zf.writestr(info, "target")
        with self.assertRaisesRegex(RuntimeError, "符号链接"):
            build.validate_archive(archive)

    def test_installer_contract_and_portable_top_level(self) -> None:
        build.validate_installer_contract()
        script = self.work / "installer.iss"
        text = (build.ROOT / "packaging" / "installer.iss").read_text(encoding="utf-8")
        text = text.replace("[Files]", '[Files]\nSource: "unexpected"')
        script.write_text(text, encoding="utf-8")
        with self.assertRaisesRegex(RuntimeError, "白名单"):
            build.validate_installer_contract(script)

        version = "test"
        portable = build.DIST / f"审计工具箱_v{version}"
        portable.mkdir()
        for name in build.PORTABLE_TOP_LEVEL - {"_internal"}:
            (portable / name).write_bytes(b"ok")
        template = portable / build.RUNTIME_OFFICE_TEMPLATE
        template.parent.mkdir(parents=True)
        template.write_bytes(b"template")
        self.assertEqual(build.validate_portable_release(version), portable)
        (portable / "客户.xlsx").write_bytes(b"customer")
        with self.assertRaisesRegex(RuntimeError, "Office"):
            build.validate_portable_release(version)


if __name__ == "__main__":
    unittest.main()
