# -*- coding: utf-8 -*-
"""한 함수 안에서 def 로 만든 이름을 뒤에서 대입으로 덮지 않는지 지킨다.

올로원스 2026-09-10: 9항 각주를 고치면서 `bio_text = mass_text = None` 을 넣었는데,
같은 `recipe_ointment.fill` 안 2170줄에 `def bio_text(lot)` 가 이미 있었다. 그때는
`bio_text(lot)` 를 부르는 자리가 모두 대입보다 먼저 실행돼 터지지 않았지만, 각주 뒤에서
9.2 표를 하나라도 더 채우면 `'str' object is not callable` 로 멈춘다 — 각주 자리는 바로
앞서 `notes[-1]` IndexError 로 작성이 통째로 멈췄던 그 자리다.

파이썬은 이런 덮어쓰기를 알려 주지 않으므로 시험으로 막는다.
"""
from __future__ import unicode_literals

import ast
import os
import unittest


PQR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pqr")
SKIP_DIRS = ("__pycache__", "data")


def _sources():
    """pqr 꾸러미의 .py 파일 — (경로, 소스)."""
    for root, dirs, files in os.walk(PQR):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for name in sorted(files):
            if name.endswith(".py"):
                path = os.path.join(root, name)
                with open(path, "rb") as f:
                    yield path, f.read().decode("utf-8")


def _shadowed(tree):
    """(함수 이름, 덮인 이름, def 줄, 대입 줄) — 같은 몸통에서 def 를 대입이 덮은 곳."""
    out = []
    for fn in ast.walk(tree):
        if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        # 같은 몸통(중첩 블록 말고 바로 그 함수의 문장들)에서만 본다 — if/else 로 갈라 쓰는
        # 꼴(둘 중 하나만 실행)까지 잡으면 멀쩡한 코드를 막는다
        defs = {}
        for node in fn.body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                defs.setdefault(node.name, node.lineno)
        if not defs:
            continue
        for node in fn.body:
            if not isinstance(node, ast.Assign):
                continue
            for target in node.targets:
                for n in ast.walk(target):
                    if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Store) and n.id in defs:
                        out.append((fn.name, n.id, defs[n.id], n.lineno))
    return out


class 이름_덮어쓰기(unittest.TestCase):
    def test_def_이름을_대입으로_덮지_않는다(self):
        bad = []
        for path, src in _sources():
            for fn, name, def_line, set_line in _shadowed(ast.parse(src)):
                bad.append("%s  %s(): '%s' — def@%d 를 대입@%d 이 덮음"
                           % (os.path.basename(path), fn, name, def_line, set_line))
        self.assertEqual(bad, [], "함수 이름을 같은 함수 안에서 덮어썼습니다:\n  " + "\n  ".join(bad))

    def test_검사기가_실제로_잡는다(self):
        """검사기 자체가 헛돌지 않는지 — 일부러 겹치게 쓴 코드는 잡혀야 한다."""
        tree = ast.parse("def fill():\n"
                         "    def bio_text(lot):\n"
                         "        return lot\n"
                         "    bio_text = None\n")
        self.assertEqual(_shadowed(tree), [("fill", "bio_text", 2, 4)])

    def test_읽은_파일이_있다(self):
        self.assertTrue(any(True for _ in _sources()))


if __name__ == "__main__":
    unittest.main()
