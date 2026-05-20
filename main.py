"""倉庫在庫管理アプリ — FastAPI 本体。

工場倉庫の原料の入庫・出庫を記録し、現在庫を自動算出する社内 Web アプリ。
このファイルを直接実行すると 0.0.0.0:8000 でサーバーが起動し、
社内の各 PC からブラウザで http://サーバー名:8000/ にアクセスできる。
"""

import io
import os
import re
import socket
import sqlite3
import sys
from contextlib import asynccontextmanager
from typing import Literal, Optional

import openpyxl
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel, ConfigDict, Field

from database import get_conn, init_db


def resource_dir() -> str:
    """同梱リソース(static)の置き場所。exe では展開先、開発時はスクリプト位置。"""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # type: ignore[attr-defined]
    return os.path.dirname(os.path.abspath(__file__))


STATIC_DIR = os.path.join(resource_dir(), "static")
PORT = 8000


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    yield


app = FastAPI(title="倉庫在庫管理", lifespan=lifespan)


# --- 入力モデル ---------------------------------------------------------------

class MaterialIn(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    code: str = Field(min_length=1, max_length=50)
    name: str = Field(min_length=1, max_length=100)
    product_name: str = Field(default="", max_length=100)
    unit: str = Field(default="kg", max_length=20)
    reorder_point: float = Field(default=0, ge=0)
    supplier: str = Field(default="", max_length=100)
    location: Literal["倉庫", "冷蔵庫", "冷凍庫"] = "倉庫"


class TransactionIn(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    material_id: int
    type: Literal["in", "out"]
    quantity: float = Field(gt=0)
    line: str = Field(default="", max_length=100)
    note: str = Field(default="", max_length=300)


class BatchOutItem(BaseModel):
    material_id: int
    quantity: float = Field(gt=0)


class BatchOutIn(BaseModel):
    model_config = ConfigDict(str_strip_whitespace=True)

    line: str = Field(default="", max_length=100)
    note: str = Field(default="", max_length=300)
    items: list[BatchOutItem] = Field(default_factory=list)


class RegisterIn(BaseModel):
    location: Literal["倉庫", "冷蔵庫", "冷凍庫"] = "倉庫"
    ids: list[int] = Field(default_factory=list)


# --- 在庫計算ヘルパー ---------------------------------------------------------

def _stock(conn: sqlite3.Connection, material_id: int) -> float:
    row = conn.execute(
        "SELECT COALESCE(SUM(CASE type WHEN 'in' THEN quantity ELSE -quantity END), 0) AS s "
        "FROM transactions WHERE material_id = ?",
        (material_id,),
    ).fetchone()
    return row["s"]


def _cell_str(v) -> str:
    """Excel セルの値を文字列に整える（10000.0 → "10000" など）。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _unit_from_pack(pack: str) -> str:
    """入数（例「25kg」）の末尾から単位を取り出す。取れなければ kg。"""
    m = re.search(r"([^\d.\s]+)$", pack)
    return m.group(1) if m else "kg"


# --- 原料マスター API ---------------------------------------------------------

@app.get("/api/materials")
def list_materials():
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT m.*,
                   COALESCE((SELECT SUM(CASE type WHEN 'in' THEN quantity ELSE -quantity END)
                             FROM transactions t WHERE t.material_id = m.id), 0) AS stock
            FROM materials m
            WHERE m.active = 1
            ORDER BY m.code
            """
        ).fetchall()
    result = []
    for r in rows:
        d = dict(r)
        d["low"] = d["reorder_point"] > 0 and d["stock"] <= d["reorder_point"]
        result.append(d)
    return result


@app.get("/api/catalog")
def list_catalog():
    """取り込み済みだが、まだ在庫登録されていない原料（カタログ）の一覧。"""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, code, name, product_name, unit FROM materials "
            "WHERE active = 0 ORDER BY code"
        ).fetchall()
    return [dict(r) for r in rows]


@app.post("/api/materials/register")
def register_materials(reg: RegisterIn):
    """カタログから選んだ原料を、指定の保管場所で在庫登録（active=1）する。"""
    if not reg.ids:
        raise HTTPException(400, "登録する原料が選ばれていません")
    with get_conn() as conn:
        placeholders = ",".join("?" * len(reg.ids))
        cur = conn.execute(
            f"UPDATE materials SET active = 1, location = ? WHERE id IN ({placeholders})",
            [reg.location, *reg.ids],
        )
        return {"registered": cur.rowcount}


@app.post("/api/materials/{material_id}/unregister")
def unregister_material(material_id: int):
    """在庫登録を解除し、カタログに戻す（履歴は残す）。"""
    with get_conn() as conn:
        cur = conn.execute(
            "UPDATE materials SET active = 0 WHERE id = ?", (material_id,)
        )
        if cur.rowcount == 0:
            raise HTTPException(404, "原料が見つかりません")
    return {"ok": True}


@app.post("/api/materials", status_code=201)
def create_material(m: MaterialIn):
    with get_conn() as conn:
        try:
            cur = conn.execute(
                "INSERT INTO materials "
                "(code, name, product_name, unit, reorder_point, supplier, location) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (m.code, m.name, m.product_name, m.unit, m.reorder_point,
                 m.supplier, m.location),
            )
        except sqlite3.IntegrityError:
            raise HTTPException(400, f"原料コード「{m.code}」はすでに登録されています")
        return {"id": cur.lastrowid}


@app.put("/api/materials/{material_id}")
def update_material(material_id: int, m: MaterialIn):
    with get_conn() as conn:
        if conn.execute("SELECT 1 FROM materials WHERE id = ?", (material_id,)).fetchone() is None:
            raise HTTPException(404, "原料が見つかりません")
        try:
            conn.execute(
                "UPDATE materials SET code = ?, name = ?, product_name = ?, unit = ?, "
                "reorder_point = ?, supplier = ?, location = ? WHERE id = ?",
                (m.code, m.name, m.product_name, m.unit, m.reorder_point,
                 m.supplier, m.location, material_id),
            )
        except sqlite3.IntegrityError:
            raise HTTPException(400, f"原料コード「{m.code}」はすでに登録されています")
    return {"ok": True}


@app.delete("/api/materials/{material_id}")
def delete_material(material_id: int):
    with get_conn() as conn:
        cur = conn.execute("DELETE FROM materials WHERE id = ?", (material_id,))
        if cur.rowcount == 0:
            raise HTTPException(404, "原料が見つかりません")
    return {"ok": True}


# --- 入出庫履歴 API -----------------------------------------------------------

@app.get("/api/transactions")
def list_transactions(material_id: Optional[int] = None, limit: int = 200):
    limit = max(1, min(limit, 1000))
    sql = (
        "SELECT t.*, m.code AS material_code, m.name AS material_name, "
        "m.product_name AS material_product, m.unit AS unit "
        "FROM transactions t JOIN materials m ON m.id = t.material_id "
    )
    params: list = []
    if material_id is not None:
        sql += "WHERE t.material_id = ? "
        params.append(material_id)
    sql += "ORDER BY t.created_at DESC, t.id DESC LIMIT ?"
    params.append(limit)
    with get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


@app.post("/api/transactions", status_code=201)
def create_transaction(tx: TransactionIn):
    with get_conn() as conn:
        if conn.execute("SELECT 1 FROM materials WHERE id = ?", (tx.material_id,)).fetchone() is None:
            raise HTTPException(404, "原料が見つかりません")
        cur = conn.execute(
            "INSERT INTO transactions (material_id, type, quantity, line, note) "
            "VALUES (?, ?, ?, ?, ?)",
            (tx.material_id, tx.type, tx.quantity, tx.line, tx.note),
        )
        if _stock(conn, tx.material_id) < 0:
            raise HTTPException(400, "現在庫を超える出庫は登録できません。数量を確認してください")
        return {"id": cur.lastrowid}


@app.put("/api/transactions/{tx_id}")
def update_transaction(tx_id: int, tx: TransactionIn):
    with get_conn() as conn:
        old = conn.execute("SELECT * FROM transactions WHERE id = ?", (tx_id,)).fetchone()
        if old is None:
            raise HTTPException(404, "履歴が見つかりません")
        if conn.execute("SELECT 1 FROM materials WHERE id = ?", (tx.material_id,)).fetchone() is None:
            raise HTTPException(404, "原料が見つかりません")
        conn.execute(
            "UPDATE transactions SET material_id = ?, type = ?, quantity = ?, line = ?, "
            "note = ? WHERE id = ?",
            (tx.material_id, tx.type, tx.quantity, tx.line, tx.note, tx_id),
        )
        for mid in {old["material_id"], tx.material_id}:
            if _stock(conn, mid) < 0:
                raise HTTPException(400, "この修正を行うと在庫がマイナスになります。内容を確認してください")
    return {"ok": True}


@app.delete("/api/transactions/{tx_id}")
def delete_transaction(tx_id: int):
    with get_conn() as conn:
        old = conn.execute("SELECT * FROM transactions WHERE id = ?", (tx_id,)).fetchone()
        if old is None:
            raise HTTPException(404, "履歴が見つかりません")
        conn.execute("DELETE FROM transactions WHERE id = ?", (tx_id,))
        if _stock(conn, old["material_id"]) < 0:
            raise HTTPException(400, "この記録を削除すると在庫がマイナスになります")
    return {"ok": True}


@app.post("/api/transactions/batch-out")
def create_batch_out(batch: BatchOutIn):
    """出庫依頼表のように複数原料をまとめて出庫登録する。

    在庫が足りる明細だけ登録し、不足している明細は登録せず skipped で返す。
    同じ原料が複数行にある場合は登録済み分を反映した在庫で順に判定する。
    """
    if not batch.items:
        raise HTTPException(400, "出庫する明細がありません")
    with get_conn() as conn:
        registered = 0
        skipped = []
        for item in batch.items:
            m = conn.execute(
                "SELECT name, unit FROM materials WHERE id = ?", (item.material_id,)
            ).fetchone()
            if m is None:
                raise HTTPException(404, "原料が見つかりません")
            available = _stock(conn, item.material_id)
            if available < item.quantity:
                skipped.append(
                    {
                        "material_id": item.material_id,
                        "name": m["name"],
                        "unit": m["unit"],
                        "requested": item.quantity,
                        "available": available,
                    }
                )
                continue
            conn.execute(
                "INSERT INTO transactions (material_id, type, quantity, line, note) "
                "VALUES (?, 'out', ?, ?, ?)",
                (item.material_id, item.quantity, batch.line, batch.note),
            )
            registered += 1
        return {"registered": registered, "skipped": skipped}


# --- Excel 取り込み API -------------------------------------------------------

@app.post("/api/import/materials")
async def import_materials(file: UploadFile = File(...)):
    """原料一覧の Excel(.xlsx) を読み込み、カタログへ取り込む。

    見出し行（「ｺｰﾄﾞ」を含む行）を自動で探し、必要な列だけ取り込む。
    新規コードはカタログ（active=0）として追加し、既存コードは
    一般名称・原材料名・単位・仕入先のみ更新する（保管場所・発注点・登録状態は維持）。
    """
    content = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    except Exception:
        raise HTTPException(
            400, "Excel ファイルとして読み込めませんでした。.xlsx 形式か確認してください"
        )
    try:
        rows = [list(r) for r in wb.active.iter_rows(values_only=True)]
    finally:
        wb.close()

    header = []
    header_idx = None
    for i, row in enumerate(rows[:15]):
        cells = [_cell_str(c) for c in row]
        if "ｺｰﾄﾞ" in cells or "コード" in cells:
            header, header_idx = cells, i
            break
    if header_idx is None:
        raise HTTPException(400, "見出し行（「ｺｰﾄﾞ」の列）が見つかりませんでした")

    def find(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    c_code = find("ｺｰﾄﾞ", "コード")
    c_name = find("一般名称")
    c_product = find("原材料名")
    c_pack = find("入数")
    c_supplier = find("購入元1", "購入元")
    if c_code is None:
        raise HTTPException(400, "「ｺｰﾄﾞ」の列が見つかりませんでした")

    def cell(row, idx):
        if idx is None or idx >= len(row):
            return ""
        return _cell_str(row[idx])

    imported = 0
    updated = 0
    with get_conn() as conn:
        existing = {r["code"] for r in conn.execute("SELECT code FROM materials")}
        for row in rows[header_idx + 1:]:
            code = cell(row, c_code)
            if not code:
                continue
            name = cell(row, c_name)
            product = cell(row, c_product)
            unit = _unit_from_pack(cell(row, c_pack))
            supplier = cell(row, c_supplier)
            if code in existing:
                # 既存はExcel由来の項目だけ更新（保管場所・発注点・登録状態は維持）
                conn.execute(
                    "UPDATE materials SET name = ?, product_name = ?, unit = ?, "
                    "supplier = ? WHERE code = ?",
                    (name, product, unit, supplier, code),
                )
                updated += 1
            else:
                # 新規はカタログとして取り込む（active=0、在庫登録は別途）
                conn.execute(
                    "INSERT INTO materials "
                    "(code, name, product_name, unit, reorder_point, supplier, "
                    "location, active) "
                    "VALUES (?, ?, ?, ?, 0, ?, '倉庫', 0)",
                    (code, name, product, unit, supplier),
                )
                existing.add(code)
                imported += 1
    return {"imported": imported, "updated": updated}


# --- 画面 ---------------------------------------------------------------------

@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


if __name__ == "__main__":
    import uvicorn

    host_name = socket.gethostname()
    print("=" * 56)
    print("  倉庫在庫管理アプリ を起動します")
    print(f"  このサーバー上:  http://localhost:{PORT}/")
    print(f"  社内の各 PC から: http://{host_name}:{PORT}/")
    print("  停止するには Ctrl+C を押してください")
    print("=" * 56)
    uvicorn.run(app, host="0.0.0.0", port=PORT)
