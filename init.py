import os
import sqlite3
from contextlib import contextmanager
from dotenv import load_dotenv

load_dotenv()


DATABASE_SCHEMA = """
CREATE TABLE IF NOT EXISTS factorys (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    address TEXT NOT NULL DEFAULT '',
    sort_num INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    factory_id INTEGER NOT NULL REFERENCES factorys(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    image TEXT NOT NULL DEFAULT '',
    sort_num INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS products_skus (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    sku_name TEXT NOT NULL,
    UNIQUE(product_id, sku_name)
);
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    sku_id INTEGER REFERENCES products_skus(id) ON DELETE SET NULL,
    quantity INTEGER NOT NULL DEFAULT 0,
    order_date TEXT NOT NULL DEFAULT (strftime('%Y%m%d', 'now'))
);
CREATE INDEX IF NOT EXISTS idx_factorys_sort_num ON factorys(sort_num);
CREATE INDEX IF NOT EXISTS idx_products_factory_sort ON products(factory_id, sort_num);
CREATE INDEX IF NOT EXISTS idx_orders_product_date ON orders(product_id, order_date);
CREATE INDEX IF NOT EXISTS idx_products_skus_product ON products_skus(product_id, id);
CREATE INDEX IF NOT EXISTS idx_orders_sku_date ON orders(sku_id, order_date);
"""


@contextmanager
def get_db():
    connection = sqlite3.connect(os.getenv('SQLALCHEMY_DATABASE_URI'))
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()



def init_db():
    with get_db() as db:
        db.executescript(DATABASE_SCHEMA)
        migrate_sku_schema(db)
        expected = {
            "factorys": {"id", "name", "address", "sort_num"},
            "products": {"id", "factory_id", "name", "image", "sort_num"},
            "products_skus": {"id", "product_id", "sku_name"},
            "orders": {"id", "product_id", "sku_id", "quantity", "order_date"},
        }
        for table, columns in expected.items():
            actual = {row["name"] for row in db.execute(f"PRAGMA table_info({table})")}
            if actual != columns:
                raise RuntimeError(
                    f"{db.name} 中 {table} 字段不匹配；"
                    f"期望 {sorted(columns)}，实际 {sorted(actual)}"
                )

def migrate_sku_schema(db):
    product_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(products)")
    }
    order_columns = {
        row["name"] for row in db.execute("PRAGMA table_info(orders)")
    }

    sku_ids = {}

    def ensure_sku(product_id, sku_name):
        sku_name = str(sku_name or "").strip()
        if not sku_name:
            return None
        key = (product_id, sku_name)
        if key not in sku_ids:
            db.execute(
                "INSERT OR IGNORE INTO products_skus(product_id, sku_name) VALUES (?, ?)",
                key,
            )
            sku_ids[key] = db.execute(
                "SELECT id FROM products_skus WHERE product_id = ? AND sku_name = ?",
                key,
            ).fetchone()["id"]
        return sku_ids[key]

    if "all_skus" in product_columns:
        for product in db.execute("SELECT id, all_skus FROM products"):
            for sku_name in (product["all_skus"] or "").split(","):
                ensure_sku(product["id"], sku_name)

    if "product_sku" in order_columns:
        if "sku_id" not in order_columns:
            db.execute(
                "ALTER TABLE orders ADD COLUMN sku_id INTEGER REFERENCES products_skus(id) ON DELETE CASCADE"
            )
        for order in db.execute(
            "SELECT id, product_id, product_sku FROM orders"
        ).fetchall():
            sku_id = ensure_sku(order["product_id"], order["product_sku"])
            db.execute("UPDATE orders SET sku_id = ? WHERE id = ?", (sku_id, order["id"]))
        db.execute("ALTER TABLE orders DROP COLUMN product_sku")

    if "all_skus" in product_columns:
        db.execute("ALTER TABLE products DROP COLUMN all_skus")


init_db()