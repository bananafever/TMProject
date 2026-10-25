"""NAS `iplog` 이름 바꿈 뒤 — 앱 SQL 에 옛 이름(refs·projects·project_grounds, ref_id·project_id)이 남지 않았나.

바꾼 날 호환 뷰 없이 넘어가므로 옛 이름이 하나라도 남으면 그 쿼리는 실패한다.
새 이름: refs→cases, projects→rounds, project_grounds→round_grounds, ref_id→case_id, project_id→round_id.
SQL 처럼 보이는 줄만 본다 — 파이썬 변수 `refs` 같은 것은 건드리지 않는다.
"""

import os
import re
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SKIP_FILES = {"build_exe.py"}
OLD_TABLE = re.compile(r"\b(FROM|JOIN|INTO|UPDATE|TABLE|COLUMNS\s+FROM)\s+`?(refs|projects|project_grounds)\b",
                       re.IGNORECASE)
OLD_COLUMN = re.compile(r"\b(ref_id|project_id)\b")
SQL_WORD = re.compile(r"\b(SELECT|FROM|WHERE|JOIN|INSERT|UPDATE|SET|SUM)\b")


def app_files():
    for name in sorted(os.listdir(HERE)):
        if name.endswith(".py") and not name.startswith("test_") and name not in SKIP_FILES:
            yield name


def old_names(text: str):
    """[(줄 번호, 줄)] — 옛 표 이름이 SQL 자리에 있거나, 옛 칸 이름이 SQL 낱말과 한 줄에 있는 줄."""
    found = []
    for n, line in enumerate(text.splitlines(), 1):
        if OLD_TABLE.search(line) or (OLD_COLUMN.search(line) and SQL_WORD.search(line)):
            found.append((n, line.strip()))
    return found


class OldDbNamesTest(unittest.TestCase):
    def test_app_sql_uses_new_names(self):
        bad = []
        for name in app_files():
            with open(os.path.join(HERE, name), encoding="utf-8") as f:
                bad += [f"{name}:{n}: {line}" for n, line in old_names(f.read())]
        self.assertEqual(bad, [], "옛 iplog 이름이 남았다:\n" + "\n".join(bad))

    def test_scans_app_files(self):
        names = list(app_files())
        self.assertIn("deadline_watch.py", names)
        self.assertIn("db.py", names)
        self.assertNotIn("test_deadline.py", names)

    def test_detector(self):
        self.assertTrue(old_names('cur.execute("SELECT * FROM projects p JOIN refs r ON r.id = p.ref_id")'))
        self.assertTrue(old_names('cur.execute("UPDATE projects SET x = 1")'))
        self.assertTrue(old_names('"SELECT SUM(project_id IS NULL) FROM assignments"'))
        self.assertTrue(old_names('"INSERT INTO project_grounds (a) VALUES (1)"'))
        self.assertFalse(old_names('refs = ref_num_key.findall(subject)'))
        self.assertFalse(old_names('cur.execute("SELECT * FROM rounds p JOIN cases r ON r.id = p.case_id")'))


if __name__ == "__main__":
    unittest.main()
