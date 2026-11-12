import sqlite3
from pathlib import Path
from urllib.parse import unquote


class Record(dict):
    """A SQLite result row accessible by both key and attribute."""

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError as error:
            raise AttributeError(name) from error


def database_path(instance_path, configured):
    if not configured:
        return Path(instance_path) / "database.db"

    if configured.startswith("sqlite:///"):
        filename = unquote(configured[len("sqlite:///"):])
        if filename.startswith("/") or (len(filename) > 2 and filename[1:3] in {":\\", ":/"}):
            return Path(filename)
        return Path(instance_path) / filename

    path = Path(configured).expanduser()
    if not path.is_absolute():
        return Path(instance_path) / path
    return path


def connect_db(app):
    path = database_path(app.instance_path, app.config.get("SQLITE_DATABASE_PATH"))
    if str(path) != ":memory:" and not str(path).startswith("file:"):
        path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(str(path))
    connection.row_factory = lambda cursor, row: Record(
        (column[0], row[index]) for index, column in enumerate(cursor.description)
    )
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


class Repository:
    def __init__(self, connection):
        self.db = connection

    def one(self, sql, parameters=()):
        return self.db.execute(sql, parameters).fetchone()

    def all(self, sql, parameters=()):
        return self.db.execute(sql, parameters).fetchall()

    def value(self, sql, parameters=()):
        row = self.one(sql, parameters)
        return next(iter(row.values())) if row else None


class Factorys(Repository):
    def get(self, factory_id):
        return self.one("SELECT id, name, address, sort_num FROM factorys WHERE id = ?", (factory_id,))

    def find_by_name(self, name, order_by_sort=False):
        ordering = "sort_num, id" if order_by_sort else "id"
        return self.one(
            f"SELECT id, name, address, sort_num FROM factorys WHERE name = ? ORDER BY {ordering} LIMIT 1",
            (name,),
        )

    def list(self, query="", missing_address_only=False):
        sql = "SELECT id, name, address, sort_num FROM factorys"
        if missing_address_only:
            sql += " WHERE address IS NULL OR trim(address) = ''"
            parameters = ()
        else:
            sql += " WHERE name LIKE ? OR address LIKE ?"
            parameters = (f"%{query}%", f"%{query}%")
        return self.all(sql + " ORDER BY sort_num, id", parameters)

    def list_for_order_date(self, order_date):
        return self.all(
            "SELECT DISTINCT f.id, f.name, f.address, f.sort_num FROM factorys f "
            "JOIN products p ON p.factory_id = f.id JOIN orders o ON o.product_id = p.id "
            "WHERE o.order_date = ? ORDER BY f.sort_num, f.id",
            (order_date,),
        )

    def count(self):
        return self.value("SELECT COUNT(id) FROM factorys")

    def count_missing_address(self):
        return self.value(
            "SELECT COUNT(id) FROM factorys WHERE address IS NULL OR trim(address) = ''"
        )

    def create(self, name, address=""):
        last_sort = self.value("SELECT MAX(sort_num) FROM factorys")
        return self.db.execute(
            "INSERT INTO factorys(name, address, sort_num) VALUES (?, ?, ?)",
            (name, address, (last_sort if last_sort is not None else -1) + 1),
        ).lastrowid

    def update(self, factory_id, **fields):
        allowed = {"name", "address"}
        if not fields or set(fields) - allowed:
            raise ValueError("Unsupported factory fields")
        assignments = ", ".join(f"{field} = ?" for field in fields)
        self.db.execute(
            f"UPDATE factorys SET {assignments} WHERE id = ?",
            (*fields.values(), factory_id),
        )

    def delete(self, factory_id):
        self.db.execute("DELETE FROM factorys WHERE id = ?", (factory_id,))

    def reorder(self, factory_ids):
        current_ids = {
            row["id"] for row in self.all("SELECT id FROM factorys")
        }
        if len(factory_ids) != len(current_ids) or set(factory_ids) != current_ids:
            return False
        self.db.executemany(
            "UPDATE factorys SET sort_num = ? WHERE id = ?",
            [(position, factory_id) for position, factory_id in enumerate(factory_ids)],
        )
        return True


class Products(Repository):
    def get(self, product_id):
        return self.one("SELECT id, factory_id, name, image FROM products WHERE id = ?", (product_id,))

    def find(self, factory_id, name):
        return self.one(
            "SELECT id, factory_id, name, image FROM products "
            "WHERE factory_id = ? AND name = ? ORDER BY id LIMIT 1",
            (factory_id, name),
        )

    def search(self, query=""):
        sql = (
            "SELECT p.id, p.factory_id, f.name AS factory_name, p.name, p.image "
            "FROM products p JOIN factorys f ON p.factory_id = f.id"
        )
        parameters = ()
        if query:
            sql += " WHERE f.name LIKE ? OR p.name LIKE ?"
            parameters = (f"%{query}%", f"%{query}%")
        return self.all(sql + " ORDER BY f.sort_num, p.id", parameters)

    def search_by_factory_and_product(self, factory_query, product_query):
        return self.all(
            "SELECT p.id, p.factory_id, f.name AS factory_name, p.name, p.image "
            "FROM products p JOIN factorys f ON p.factory_id = f.id "
            "WHERE f.name LIKE ? AND p.name LIKE ? ORDER BY f.sort_num, p.id",
            (f"%{factory_query}%", f"%{product_query}%"),
        )

    def list_choices(self):
        return self.all(
            "SELECT p.id, p.name, f.name AS factory_name FROM products p "
            "JOIN factorys f ON p.factory_id = f.id ORDER BY f.sort_num, p.id"
        )

    def list_orders_for_factory(self, factory_id, order_date):
        return self.all(
            "SELECT o.id AS order_id, o.order_content, o.order_date, "
            "p.id AS product_id, p.name AS product_name, p.image FROM orders o "
            "JOIN products p ON p.id = o.product_id "
            "WHERE p.factory_id = ? AND o.order_date = ? "
            "ORDER BY p.id, o.id",
            (factory_id, order_date),
        )

    def count(self):
        return self.value("SELECT COUNT(id) FROM products")

    def count_missing_image(self):
        return self.value(
            "SELECT COUNT(id) FROM products WHERE image IS NULL OR trim(image) = ''"
        )

    def create(self, factory_id, name, image=""):
        return self.db.execute(
            "INSERT INTO products(factory_id, name, image) VALUES (?, ?, ?)",
            (factory_id, name, image),
        ).lastrowid

    def update(self, product_id, **fields):
        allowed = {"factory_id", "name", "image"}
        if not fields or set(fields) - allowed:
            raise ValueError("Unsupported product fields")
        assignments = ", ".join(f"{field} = ?" for field in fields)
        self.db.execute(
            f"UPDATE products SET {assignments} WHERE id = ?",
            (*fields.values(), product_id),
        )

    def exists_name(self, factory_id, name, excluding_id=None):
        sql = "SELECT id FROM products WHERE factory_id = ? AND name = ?"
        parameters = [factory_id, name]
        if excluding_id is not None:
            sql += " AND id != ?"
            parameters.append(excluding_id)
        return self.value(sql, parameters) is not None

    def delete(self, product_id):
        self.db.execute("DELETE FROM products WHERE id = ?", (product_id,))


class Orders(Repository):
    def get(self, order_id):
        return self.one(
            "SELECT id, product_id, order_content, order_date FROM orders WHERE id = ?",
            (order_id,),
        )

    def list_dashboard(self, start_date, end_date):
        return self.all(
            "SELECT o.order_date, COUNT(DISTINCT p.id) AS products, "
            "COUNT(DISTINCT f.id) AS factories "
            "FROM orders o JOIN products p ON p.id = o.product_id "
            "JOIN factorys f ON f.id = p.factory_id "
            "WHERE o.order_date BETWEEN ? AND ? GROUP BY o.order_date ORDER BY o.order_date",
            (start_date, end_date),
        )

    def list_for_date(self, order_date):
        return self.all(
            "SELECT id, product_id, order_content FROM orders "
            "WHERE order_date = ? ORDER BY id",
            (order_date,),
        )

    def dashboard_totals(self, start_date, end_date):
        return self.one(
            "SELECT COUNT(DISTINCT p.id) AS products, COUNT(DISTINCT f.id) AS factories "
            "FROM orders o JOIN products p ON p.id = o.product_id "
            "JOIN factorys f ON f.id = p.factory_id "
            "WHERE o.order_date BETWEEN ? AND ?",
            (start_date, end_date),
        )

    def create(self, product_id, order_content, order_date):
        return self.db.execute(
            "INSERT INTO orders(product_id, order_content, order_date) VALUES (?, ?, ?)",
            (product_id, order_content, order_date),
        ).lastrowid

    def update(self, order_id, **fields):
        allowed = {"order_content", "order_date"}
        if not fields or set(fields) - allowed:
            raise ValueError("Unsupported order fields")
        assignments = ", ".join(f"{field} = ?" for field in fields)
        self.db.execute(
            f"UPDATE orders SET {assignments} WHERE id = ?",
            (*fields.values(), order_id),
        )


