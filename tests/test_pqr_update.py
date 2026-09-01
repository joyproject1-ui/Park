"""프로그램 업데이트 — 담당자 자료를 지우지 않는지가 핵심입니다."""

import os
import shutil
import tempfile
import unittest
import zipfile

from pqr import cli


class CopyProgramTest(unittest.TestCase):

    def setUp(self):
        self.source = tempfile.mkdtemp()
        self.target = tempfile.mkdtemp()

    def tearDown(self):
        for path in (self.source, self.target):
            shutil.rmtree(path, ignore_errors=True)

    def write(self, root, relative, text):
        path = os.path.join(root, relative)
        folder = os.path.dirname(path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(text)
        return path

    def read(self, root, relative):
        with open(os.path.join(root, relative), encoding="utf-8") as handle:
            return handle.read()

    def test_program_files_are_replaced(self):
        self.write(self.source, os.path.join("pqr", "build.py"), "새 버전")
        self.write(self.target, os.path.join("pqr", "build.py"), "옛 버전")
        cli._copy_program(self.source, self.target)
        self.assertEqual(self.read(self.target, os.path.join("pqr", "build.py")), "새 버전")

    def test_input_folder_is_never_touched(self):
        """담당자가 모아 둔 자료가 사라지면 그날 업무가 끝납니다."""
        self.write(self.source, os.path.join("PQR_입력폴더", "읽어보기.txt"), "배포본")
        kept = self.write(self.target, os.path.join("PQR_입력폴더", "QC1-5007", "성적서.pdf"),
                          "내 자료")
        self.write(self.source, os.path.join("pqr", "cli.py"), "새 버전")
        cli._copy_program(self.source, self.target)
        self.assertTrue(os.path.exists(kept))
        self.assertEqual(self.read(self.target, os.path.join("PQR_입력폴더", "QC1-5007",
                                                             "성적서.pdf")), "내 자료")
        # 배포본에 들어 있던 입력폴더 파일이 덮어써서도 안 됩니다.
        self.assertFalse(os.path.exists(
            os.path.join(self.target, "PQR_입력폴더", "읽어보기.txt")))

    def test_out_folder_is_kept(self):
        made = self.write(self.target, os.path.join("out", "reports", "PQR_QC1-5007.md"), "초안")
        self.write(self.source, os.path.join("out", "reports", "PQR_QC1-5007.md"), "빈 배포본")
        cli._copy_program(self.source, self.target)
        self.assertEqual(self.read(self.target, os.path.join("out", "reports",
                                                             "PQR_QC1-5007.md")), "초안")
        self.assertTrue(os.path.exists(made))

    def test_files_only_in_target_survive(self):
        """지우지 않고 덮어쓰기만 합니다 — 담당자가 둔 파일이 섞여 있을 수 있습니다."""
        mine = self.write(self.target, "내 메모.txt", "지우면 안 됩니다")
        self.write(self.source, "README.md", "새 문서")
        cli._copy_program(self.source, self.target)
        self.assertTrue(os.path.exists(mine))

    def test_new_folders_are_created(self):
        self.write(self.source, os.path.join("docs", "pqr", "index.html"), "화면")
        cli._copy_program(self.source, self.target)
        self.assertTrue(os.path.isfile(
            os.path.join(self.target, "docs", "pqr", "index.html")))


class UpdateCommandTest(unittest.TestCase):
    """네트워크 없이 file:// 로 내려받아 전체 흐름을 확인합니다."""

    def setUp(self):
        self.workspace = tempfile.mkdtemp()
        self.target = os.path.join(self.workspace, "program")
        os.makedirs(os.path.join(self.target, "pqr"))
        with open(os.path.join(self.target, "pqr", "build.py"), "w", encoding="utf-8") as handle:
            handle.write("옛 버전")
        self.archive = os.path.join(self.workspace, "update.zip")
        with zipfile.ZipFile(self.archive, "w") as archive:
            archive.writestr("Park-branch/pqr/build.py", "새 버전")
            archive.writestr("Park-branch/README.md", "설명")

    def tearDown(self):
        shutil.rmtree(self.workspace, ignore_errors=True)

    def run_update(self, url):
        class Args(object):
            pass
        args = Args()
        args.url = url
        args.target = self.target
        return cli.cmd_update(args)

    def test_update_replaces_program_files(self):
        code = self.run_update("file://" + self.archive)
        self.assertEqual(code, 0)
        with open(os.path.join(self.target, "pqr", "build.py"), encoding="utf-8") as handle:
            self.assertEqual(handle.read(), "새 버전")
        self.assertTrue(os.path.isfile(os.path.join(self.target, "README.md")))

    def test_network_failure_reports_the_manual_way(self):
        """사내망에서 막히는 일이 흔합니다 — 오류만 뱉고 끝내면 안 됩니다."""
        code = self.run_update("file://" + os.path.join(self.workspace, "없는파일.zip"))
        self.assertEqual(code, 2)
        self.assertTrue(os.path.isdir(self.target))     # 실패해도 프로그램은 그대로


if __name__ == "__main__":
    unittest.main()
