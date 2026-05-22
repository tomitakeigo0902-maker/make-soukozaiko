"""SQLite の接続管理とスキーマ初期化。

DB ファイル(inventory.db)は exe(またはスクリプト)と同じフォルダに作成する。
こうしておくとバックアップは inventory.db をコピーするだけで済む。
"""

import os
import sqlite3
import sys
from contextlib import contextmanager


def base_dir() -> str:
    """exe またはスクリプトが置かれているフォルダ。DB の保存先に使う。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DB_PATH = os.path.join(base_dir(), "inventory.db")


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS materials (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                code          TEXT NOT NULL UNIQUE,
                name          TEXT NOT NULL,
                product_name  TEXT NOT NULL DEFAULT '',
                pack_size     TEXT NOT NULL DEFAULT '',
                unit          TEXT NOT NULL DEFAULT 'kg',
                reorder_point REAL NOT NULL DEFAULT 0,
                supplier      TEXT NOT NULL DEFAULT '',
                location      TEXT NOT NULL DEFAULT '倉庫',
                active        INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE TABLE IF NOT EXISTS transactions (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                material_id INTEGER NOT NULL REFERENCES materials(id) ON DELETE CASCADE,
                type        TEXT NOT NULL CHECK (type IN ('in', 'out')),
                quantity    REAL NOT NULL CHECK (quantity > 0),
                line        TEXT NOT NULL DEFAULT '',
                note        TEXT NOT NULL DEFAULT '',
                import_key  TEXT NOT NULL DEFAULT '',
                created_at  TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
            );

            CREATE INDEX IF NOT EXISTS idx_tx_material ON transactions(material_id);
            CREATE INDEX IF NOT EXISTS idx_tx_created  ON transactions(created_at);
            """
        )
        # 旧バージョンの DB に無い列を追加（移行）
        cols = [r["name"] for r in conn.execute("PRAGMA table_info(materials)")]
        if "location" not in cols:
            conn.execute(
                "ALTER TABLE materials ADD COLUMN location TEXT NOT NULL DEFAULT '倉庫'"
            )
        if "product_name" not in cols:
            conn.execute(
                "ALTER TABLE materials ADD COLUMN product_name TEXT NOT NULL DEFAULT ''"
            )
        if "active" not in cols:
            conn.execute(
                "ALTER TABLE materials ADD COLUMN active INTEGER NOT NULL DEFAULT 1"
            )
        if "pack_size" not in cols:
            conn.execute(
                "ALTER TABLE materials ADD COLUMN pack_size TEXT NOT NULL DEFAULT ''"
            )
        tx_cols = [r["name"] for r in conn.execute("PRAGMA table_info(transactions)")]
        if "import_key" not in tx_cols:
            conn.execute(
                "ALTER TABLE transactions ADD COLUMN import_key TEXT NOT NULL DEFAULT ''"
            )
