#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""行测刷题工具 - 本地服务 + SQLite 存储。

启动：
    python3 app.py [端口]
访问：
    http://localhost:8000

数据：
    data/questions.json  题库（只读）
    data/quiz.db         SQLite（答题记录、错题、用户设置的答案），重启保留
"""
import json
import os
import sqlite3
import time
import email
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, unquote

import extract

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")
DB_PATH = os.path.join(DATA_DIR, "quiz.db")
Q_JSON = os.path.join(DATA_DIR, "questions.json")
IMP_JSON = os.path.join(DATA_DIR, "imported.json")   # 用户导入的题库

QUESTIONS = []          # 全部题目（原始 + 导入）
BY_ID = {}              # id -> question
BY_SRC = {}             # source -> question list
PARTS = {}              # source -> set(part)
SECTIONS = {}           # source -> set(section)


def load_questions():
    global QUESTIONS, BY_ID, BY_SRC, PARTS, SECTIONS
    with open(Q_JSON, encoding="utf-8") as f:
        QUESTIONS = json.load(f)
    if os.path.exists(IMP_JSON):
        with open(IMP_JSON, encoding="utf-8") as f:
            QUESTIONS.extend(json.load(f))
    for q in QUESTIONS:
        BY_ID[q["id"]] = q
        BY_SRC.setdefault(q["source"], []).append(q)
        PARTS.setdefault(q["source"], set()).add(q["part"])
        SECTIONS.setdefault(q["source"], set()).add(q.get("section") or "")
    for src in BY_SRC:
        BY_SRC[src].sort(key=lambda q: q["no"])


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db()
    conn.execute("""CREATE TABLE IF NOT EXISTS attempts(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        question_id TEXT NOT NULL,
        selected TEXT,
        correct INTEGER,
        ts INTEGER)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS overrides(
        question_id TEXT PRIMARY KEY,
        answer TEXT,
        explanation TEXT,
        ts INTEGER)""")
    conn.commit()
    conn.close()


def get_answer(q):
    conn = db()
    row = conn.execute("SELECT * FROM overrides WHERE question_id=?", (q["id"],)).fetchone()
    conn.close()
    if row:
        return row["answer"], row["explanation"], True
    return q.get("answer", ""), q.get("explanation", ""), False


def last_attempt(conn, qid):
    return conn.execute(
        "SELECT * FROM attempts WHERE question_id=? ORDER BY ts DESC, id DESC LIMIT 1",
        (qid,)).fetchone()


def qdict(q):
    answer, explanation, is_override = get_answer(q)
    conn = db()
    la = last_attempt(conn, q["id"])
    conn.close()
    return {
        "id": q["id"], "no": q["no"], "source": q["source"],
        "part": q["part"], "section": q.get("section", ""),
        "stem": q["stem"], "options": q.get("options", []),
        "images": q.get("images", []), "material": q.get("material", []),
        "answer": answer, "explanation": explanation,
        "has_answer": bool(answer), "answer_source": "user" if is_override else "doc",
        "last_selected": la["selected"] if la else None,
        "last_correct": la["correct"] if la else None,
        "attempted": la is not None,
    }


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body=b"", ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if body:
            self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def _read_json(self):
        length = int(self.headers.get("Content-Length", 0))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self):
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        qs = parse_qs(parsed.query)
        try:
            if path == "/api/meta":
                self._json({s: {"total": len(v), "parts": sorted(PARTS[s]),
                                "sections": sorted(x for x in SECTIONS[s] if x)}
                            for s, v in BY_SRC.items()})
            elif path == "/api/questions":
                self._questions(qs)
            elif path.startswith("/api/question/"):
                self._question(path.rsplit("/", 1)[1])
            elif path == "/api/stats":
                self._stats()
            elif path.startswith("/api/wrong"):
                self._wrong()
            elif path == "/api/attempts":
                self._attempts(qs)
            elif path == "/static" or path == "/":
                self._static("index.html")
            elif path.startswith("/images/"):
                self._file(os.path.join(DATA_DIR, path.lstrip("/")), "image/png")
            elif path.startswith("/static/"):
                self._static(path[len("/static/"):])
            else:
                self._send(404, b"not found")
        except Exception as e:
            self._json({"error": str(e)}, 500)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/attempt":
                body = self._read_json()
                self._json(self._add_attempt(body))
            elif parsed.path == "/api/answer":
                body = self._read_json()
                self._json(self._set_answer(body))
            elif parsed.path == "/api/reset":
                self._json(self._reset())
            elif parsed.path == "/api/import":
                self._json(self._do_import(self._parse_multipart()))
            elif parsed.path == "/api/delete_source":
                self._json(self._delete_source(self._read_json()))
            else:
                self._send(404, b"not found")
        except Exception as e:
            self._json({"error": str(e)}, 500)

    def _delete_source(self, body):
        global QUESTIONS, BY_ID, BY_SRC, PARTS, SECTIONS
        source = body.get("source")
        if not source:
            return {"error": "缺少 source"}
        qs = BY_SRC.get(source)
        if not qs:
            return {"error": "题库不存在"}
        ids = {q["id"] for q in qs}
        conn = db()
        for qid in ids:
            conn.execute("DELETE FROM attempts WHERE question_id=?", (qid,))
            conn.execute("DELETE FROM overrides WHERE question_id=?", (qid,))
        conn.commit()
        conn.close()
        QUESTIONS = [q for q in QUESTIONS if q["id"] not in ids]
        BY_ID = {k: v for k, v in BY_ID.items() if k not in ids}
        BY_SRC = {k: v for k, v in BY_SRC.items() if k != source}
        PARTS.pop(source, None)
        SECTIONS.pop(source, None)
        # 写回数据文件
        with open(Q_JSON, encoding="utf-8") as f:
            qjson = json.load(f)
        if any(q["source"] == source for q in qjson):
            qjson = [q for q in qjson if q["source"] != source]
            with open(Q_JSON, "w", encoding="utf-8") as f:
                json.dump(qjson, f, ensure_ascii=False, indent=1)
        if os.path.exists(IMP_JSON):
            with open(IMP_JSON, encoding="utf-8") as f:
                imp = json.load(f)
            if any(q["source"] == source for q in imp):
                imp = [q for q in imp if q["source"] != source]
                with open(IMP_JSON, "w", encoding="utf-8") as f:
                    json.dump(imp, f, ensure_ascii=False, indent=1)
        return {"ok": True, "count": len(qs), "source": source}

    def _parse_multipart(self):
        ctype = self.headers.get("Content-Type", "")
        body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
        msg = BytesParser(policy=policy.default).parsebytes(
            b"Content-Type: " + ctype.encode() + b"\r\nMIME-Version: 1.0\r\n\r\n" + body)
        files = []
        for part in msg.iter_parts():
            fn = part.get_filename()
            if fn:
                files.append((os.path.basename(fn), part.get_payload(decode=True) or b""))
        return files

    def _do_import(self, files):
        if not files:
            return {"error": "未收到文件"}
        tmp_dir = os.path.join(DATA_DIR, "tmp_import")
        os.makedirs(tmp_dir, exist_ok=True)
        saved = []
        for name, data in files[:2]:
            p = os.path.join(tmp_dir, name)
            with open(p, "wb") as f:
                f.write(data)
            saved.append(p)
        question_file = saved[0]
        answer_file = saved[1] if len(saved) > 1 else None
        source = os.path.splitext(os.path.basename(question_file))[0]
        qs = extract.import_documents(question_file, answer_file, source)
        if not qs:
            return {"error": "未能从文档中解析出题目，请检查文件格式"}
        max_id = max((int(q["id"][1:]) for q in QUESTIONS), default=0)
        for i, q in enumerate(qs):
            q["id"] = "Q%05d" % (max_id + 1 + i)
        # 导入的题存单独文件，避免重新提取原始题库时被覆盖
        imported = []
        if os.path.exists(IMP_JSON):
            with open(IMP_JSON, encoding="utf-8") as f:
                imported = json.load(f)
        imported.extend(qs)
        with open(IMP_JSON, "w", encoding="utf-8") as f:
            json.dump(imported, f, ensure_ascii=False, indent=1)
        QUESTIONS.extend(qs)
        for q in qs:
            BY_ID[q["id"]] = q
            BY_SRC.setdefault(q["source"], []).append(q)
            PARTS.setdefault(q["source"], set()).add(q["part"])
            SECTIONS.setdefault(q["source"], set()).add(q.get("section") or "")
        return {"ok": True, "count": len(qs),
                "with_answer": sum(1 for q in qs if q["answer"]),
                "source": source}

    def do_DELETE(self):
        parsed = urlparse(self.path)
        try:
            if parsed.path.startswith("/api/attempt/"):
                self._delete_attempt(parsed.path.rsplit("/", 1)[1])
            else:
                self._send(404, b"not found")
        except Exception as e:
            self._json({"error": str(e)}, 500)

    # ---------- 业务 ----------
    def _questions(self, qs):
        src = qs.get("source", [None])[0]
        part = qs.get("part", [None])[0]
        section = qs.get("section", [None])[0]
        mode = qs.get("mode", ["all"])[0]     # all|wrong|noanswer|right|problem
        keyword = (qs.get("kw", [""])[0] or "").strip()
        pool = BY_SRC.get(src, QUESTIONS) if src else QUESTIONS
        if part:
            pool = [q for q in pool if q["part"] == part]
        if section:
            pool = [q for q in pool if (q.get("section") or "") == section]
        if keyword:
            pool = [q for q in pool if keyword in q["stem"] or keyword in q["part"]]
        items = []
        conn = db()
        for q in pool:
            la = last_attempt(conn, q["id"])
            ok = la["correct"] if la else None
            if mode == "wrong" and ok != 0:
                continue
            if mode == "noanswer" and la is not None:
                continue   # 未作答 = 没有作答记录的题
            if mode == "right" and ok != 1:
                continue
            if mode == "problem":
                # 待完善：只保留暂无答案的题
                if get_answer(q)[0]:
                    continue
            items.append(qdict(q))
        conn.close()
        limit = int(qs.get("limit", ["2000"])[0])
        offset = int(qs.get("offset", ["0"])[0])
        page = items[offset:offset + limit]
        self._json({"total": len(items), "offset": offset, "limit": limit,
                    "questions": page})

    def _question(self, qid):
        q = BY_ID.get(qid)
        if not q:
            return self._send(404, b"not found")
        self._json(qdict(q))

    def _add_attempt(self, body):
        qid = body.get("id")
        q = BY_ID.get(qid)
        if not q:
            return {"error": "no such question"}
        selected = body.get("selected", "")
        answer, _, _ = get_answer(q)
        correct = (selected == answer) if answer else None
        conn = db()
        conn.execute("INSERT INTO attempts(question_id, selected, correct, ts) VALUES(?,?,?,?)",
                     (qid, selected, correct, int(time.time())))
        conn.commit()
        conn.close()
        return {"id": qid, "selected": selected, "correct": correct,
                "answer": answer if answer else None,
                "has_answer": bool(answer)}

    def _set_answer(self, body):
        qid = body.get("id")
        q = BY_ID.get(qid)
        if not q:
            return {"error": "no such question"}
        answer = (body.get("answer") or "").strip().upper()
        explanation = (body.get("explanation") or "").strip()
        conn = db()
        conn.execute(
            "INSERT INTO overrides(question_id, answer, explanation, ts) VALUES(?,?,?,?) "
            "ON CONFLICT(question_id) DO UPDATE SET answer=excluded.answer, "
            "explanation=excluded.explanation, ts=excluded.ts",
            (qid, answer, explanation, int(time.time())))
        conn.commit()
        conn.close()
        return {"id": qid, "answer": answer, "explanation": explanation}

    def _stats(self):
        conn = db()
        attempts = conn.execute(
            "SELECT question_id, selected, correct, ts FROM attempts "
            "ORDER BY ts DESC, id DESC").fetchall()
        conn.close()
        seen = {}
        for a in attempts:
            seen.setdefault(a["question_id"], a)
        attempted = len(seen)
        correct = sum(1 for a in seen.values() if a["correct"] == 1)
        wrong_ids = [qid for qid, a in seen.items() if a["correct"] == 0]
        total = len(QUESTIONS)
        has_ans = sum(1 for q in QUESTIONS if get_answer(q)[0])
        by_src = {}
        for src, lst in BY_SRC.items():
            n = len(lst)
            ok = sum(1 for q in lst if (a := seen.get(q["id"])) and a["correct"] == 1)
            att = sum(1 for q in lst if q["id"] in seen)
            ans = sum(1 for q in lst if get_answer(q)[0])
            # 上次作答位置：该试卷最近一次作答的题号
            last_no = None
            latest_ts = -1
            for q in lst:
                a = seen.get(q["id"])
                if a and a["ts"] > latest_ts:
                    latest_ts = a["ts"]
                    last_no = q["no"]
            by_src[src] = {"total": n, "attempted": att, "correct": ok,
                           "has_answer": ans, "last_no": last_no}
        self._json({
            "total": total, "attempted": attempted, "correct": correct,
            "wrong_count": len(wrong_ids), "has_answer": has_ans,
            "by_source": by_src,
        })

    def _wrong(self):
        conn = db()
        all_a = conn.execute("SELECT * FROM attempts ORDER BY ts DESC, id DESC").fetchall()
        conn.close()
        latest = {}
        for a in all_a:
            latest.setdefault(a["question_id"], a)
        ids = [qid for qid, a in latest.items() if a["correct"] == 0]
        items = [qdict(BY_ID[qid]) for qid in ids if qid in BY_ID]
        items.sort(key=lambda q: (q["source"], q["no"]))
        self._json({"total": len(items), "questions": items})

    def _attempts(self, qs):
        qid = qs.get("question_id", [None])[0]
        conn = db()
        if qid:
            rows = conn.execute(
                "SELECT * FROM attempts WHERE question_id=? ORDER BY ts DESC, id DESC",
                (qid,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM attempts ORDER BY ts DESC, id DESC LIMIT 100").fetchall()
        conn.close()
        self._json([dict(r) for r in rows])

    def _delete_attempt(self, qid):
        conn = db()
        conn.execute("DELETE FROM attempts WHERE question_id=?", (qid,))
        conn.commit()
        conn.close()
        self._json({"ok": True})

    def _reset(self):
        conn = db()
        conn.execute("DELETE FROM attempts")
        conn.execute("DELETE FROM overrides")
        conn.commit()
        conn.close()
        return {"ok": True}

    # ---------- 静态 ----------
    def _static(self, rel):
        path = os.path.normpath(os.path.join(STATIC_DIR, rel))
        if not path.startswith(STATIC_DIR) or not os.path.isfile(path):
            return self._send(404, b"not found")
        ctype = {".html": "text/html; charset=utf-8",
                 ".css": "text/css; charset=utf-8",
                 ".js": "application/javascript; charset=utf-8",
                 ".png": "image/png", ".jpg": "image/jpeg"}.get(
                     os.path.splitext(path)[1], "application/octet-stream")
        with open(path, "rb") as f:
            self._send(200, f.read(), ctype)

    def _file(self, path, ctype):
        if not os.path.isfile(path):
            return self._send(404, b"not found")
        with open(path, "rb") as f:
            self._send(200, f.read(), ctype)

    def log_message(self, fmt, *args):
        pass


def main():
    port = int(sys_argv_port())
    load_questions()
    init_db()
    print("题库: %d 题" % len(QUESTIONS))
    print("打开 http://localhost:%d 开始刷题" % port)
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()


def sys_argv_port():
    import sys
    return sys.argv[1] if len(sys.argv) > 1 else "8000"


if __name__ == "__main__":
    main()
