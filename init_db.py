import os
import sqlite3
from pathlib import Path

from dotenv import load_dotenv

from models import database_path


PROJECT_ROOT = Path(__file__).resolve().parent
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
    image TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    order_content TEXT NOT NULL DEFAULT '',
    order_date TEXT NOT NULL DEFAULT (strftime('%Y%m%d', 'now'))
);
CREATE INDEX IF NOT EXISTS idx_factorys_sort_num ON factorys(sort_num);
CREATE INDEX IF NOT EXISTS idx_orders_product_date ON orders(product_id, order_date);
"""


def migrate_schema(connection):
    product_columns = {
        row["name"] for row in connection.execute("PRAGMA table_info(products)")
    }
    order_columns = {
        row["name"] for row in connection.execute("PRAGMA table_info(orders)")
    }
    tables = {
        row["name"]
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table'"
        )
    }

    legacy_order_columns = {"sku_id", "quantity", "product_sku"}
    if "order_content" not in order_columns or order_columns & legacy_order_columns:
        old_orders = connection.execute("SELECT * FROM orders").fetchall()
        sku_names = {}
        if "products_skus" in tables and "sku_id" in order_columns:
            sku_names = {
                row["id"]: row["sku_name"]
                for row in connection.execute("SELECT id, sku_name FROM products_skus")
            }

        connection.execute("DROP TABLE IF EXISTS orders_new")
        connection.execute(
            "CREATE TABLE orders_new ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            "product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE, "
            "order_content TEXT NOT NULL DEFAULT '', "
            "order_date TEXT NOT NULL DEFAULT (strftime('%Y%m%d', 'now')))"
        )
        for order in old_orders:
            keys = order.keys()
            content = str(order["order_content"] or "").strip() if "order_content" in keys else ""
            if not content:
                sku_name = order["product_sku"] if "product_sku" in keys else ""
                if not sku_name and "sku_id" in keys:
                    sku_name = sku_names.get(order["sku_id"], "")
                quantity = order["quantity"] if "quantity" in keys else None
                if quantity is not None:
                    label = str(sku_name or "数量").strip()
                    content = f"{label}x{quantity}"
            connection.execute(
                "INSERT INTO orders_new(id, product_id, order_content, order_date) "
                "VALUES (?, ?, ?, ?)",
                (
                    order["id"],
                    order["product_id"],
                    content,
                    order["order_date"],
                ),
            )
        connection.execute("DROP TABLE orders")
        connection.execute("ALTER TABLE orders_new RENAME TO orders")

    connection.execute("DROP TABLE IF EXISTS products_skus")

    if "sort_num" in product_columns:
        connection.execute("DROP INDEX IF EXISTS idx_products_factory_sort")
        connection.execute("ALTER TABLE products DROP COLUMN sort_num")
    if "all_skus" in product_columns:
        connection.execute("ALTER TABLE products DROP COLUMN all_skus")

    connection.execute("CREATE INDEX IF NOT EXISTS idx_orders_product_date ON orders(product_id, order_date)")


def init_db():
    load_dotenv(PROJECT_ROOT / ".env")
    configured = (
        os.getenv("SQLITE_DATABASE_PATH")
        or os.getenv("DATABASE_URI")
        or "sqlite:///database.db"
    )
    path = database_path(PROJECT_ROOT / "instance", configured)
    if str(path) != ":memory:" and not str(path).startswith("file:"):
        path.parent.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(str(path))
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    try:
        connection.executescript(DATABASE_SCHEMA)
        migrate_schema(connection)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


if __name__ == "__main__":
    init_db()
    print("SQLite database schema initialized.")
