import csv
import io
import json
import os
import re
import sqlite3
import uuid
import zipfile
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for , session
from dotenv import load_dotenv
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField
from wtforms.validators import DataRequired, Length


# 加载 .env 文件，放在最前面！
load_dotenv()

app = Flask(__name__)
# 从环境变量读取配置
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY")
app.config["DATABASE_URI"] = os.getenv("DATABASE_URI")


# curl.exe -i -X GET http://127.0.0.1:5000/debug?name=TJ 
# curl.exe -i -X POST http://127.0.0.1:5000/debug -H "Content-Type: application/json" -d '{\"name\": \"taojun\"}' 
# curl.exe -i -X PUT http://127.0.0.1:5000/debug -F "name=taojun"
@app.route('/debug', methods=['GET', 'POST'])
def debug():
    # form = MyForm()
    # if form.validate_on_submit():
    #     name = form.name.data
    #     email = form.email.data
    #     return f'Name: {name}, Email: {email}'
    return render_template('debug.html')


@contextmanager
def get_db():
    connection = sqlite3.connect(app.config['DATABASE_URI'])
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




def normalize_order_date(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y%m%d")
    value = str(value or "").strip()
    if not value:
        return ""
    compact = re.fullmatch(r"\d{8}", value)
    if compact:
        return value
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(value[:10], pattern).strftime("%Y%m%d")
        except ValueError:
            continue
    raise ValueError(f"无法识别日期：{value}")


def display_order_date(value):
    value = str(value or "")
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:]}"
    return value


HEADER_ALIASES = {
    "product_code": {"商品编码"},
    "product_sku": {"销售属性"},
    "quantity": {"数量"},
}


def import_rows(filename, content):
    extension = Path(filename).suffix.lower()
    if extension == ".csv":
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("gb18030")
        rows = list(csv.reader(io.StringIO(text)))
    elif extension in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook
        from openpyxl.utils.exceptions import InvalidFileException

        try:
            workbook = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        except (InvalidFileException, OSError, ValueError, zipfile.BadZipFile) as error:
            raise ValueError("Excel 文件损坏或格式无效，请重新选择文件。") from error
        try:
            rows = list(workbook.active.iter_rows(values_only=True))
        finally:
            workbook.close()
    else:
        raise ValueError("仅支持 .xlsx、.xlsm 或 .csv 文件")
    if not rows:
        raise ValueError("表格内容为空")

    columns = {}
    for index, cell in enumerate(rows[0]):
        value = str(cell or "").strip().lower()
        for field, aliases in HEADER_ALIASES.items():
            if value in aliases:
                columns.setdefault(field, index)
    missing_columns = {"product_code", "product_sku", "quantity"} - columns.keys()
    if missing_columns:
        labels = {"product_code": "商品编码", "product_sku": "销售属性", "quantity": "数量"}
        raise ValueError(
            "表格缺少必需列：" + "、".join(labels[column] for column in sorted(missing_columns))
        )

    imported = []
    for row in rows[1:]:
        values = list(row)
        if not any(str(value or "").strip() for value in values):
            continue
        record = {
            field: values[index] if index < len(values) else ""
            for field, index in columns.items()
        }
        product_code = str(record.get("product_code") or "").strip()
        if not product_code:
            continue
        factory, separator, name = product_code.partition("&")
        if not separator:
            factory, separator, name = product_code.partition("＆")
        factory = factory.strip()
        name = name.strip()
        if not separator or not factory or not name:
            raise ValueError(
                f"商品编码“{product_code}”格式无效，应为“厂家名称&商品名称”"
            )
        sku = str(record.get("product_sku") or "").strip()
        raw_quantity = record.get("quantity")
        try:
            numeric_quantity = float(raw_quantity)
            quantity = int(numeric_quantity)
            if quantity < 0 or numeric_quantity != quantity:
                raise ValueError
        except (TypeError, ValueError):
            raise ValueError(f"商品“{name}”的数量必须是非负整数")
        imported.append(
            {
                "name": name,
                "factory": factory,
                "image": "",
                "product_sku": sku,
                "quantity": quantity,
                "order_date": date.today().strftime("%Y%m%d"),
            }
        )
    if not imported:
        raise ValueError("没有找到可导入的商品数据")
    return imported


def search_products(query="", order_date=None):
    clauses = []
    values = []
    if query:
        search_parts = re.split(r"[&＆]", query, maxsplit=1)
        if len(search_parts) == 2:
            factory_query, product_query = search_parts
            clauses.append("(f.name LIKE ? AND p.name LIKE ?)")
            values.extend((f"%{factory_query.strip()}%", f"%{product_query.strip()}%"))
        else:
            clauses.append("(f.name LIKE ? OR p.name LIKE ?)")
            values.extend((f"%{query}%", f"%{query}%"))
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    with get_db() as db:
        products = db.execute(
            f"""SELECT p.*, f.name AS factory_name
                FROM products p JOIN factorys f ON f.id = p.factory_id
                {where} ORDER BY f.sort_num, p.sort_num, p.id""",
            values,
        ).fetchall()
        sku_rows = db.execute(
            """SELECT s.id AS sku_id, s.product_id, s.sku_name
               FROM products_skus s
               ORDER BY s.id"""
        ).fetchall()

    skus_by_product = {}
    for row in sku_rows:
        skus_by_product.setdefault(row["product_id"], []).append(dict(row))
    result = []
    for product in products:
        product_skus = skus_by_product.get(product["id"], [])
        result.append({
            **dict(product),
            "sku_rows": product_skus,
        })
    return result


def ensure_product_sku(db, product_id, sku_name):
    sku_name = str(sku_name or "").strip()
    if not sku_name:
        return None
    db.execute(
        "INSERT OR IGNORE INTO products_skus(product_id, sku_name) VALUES (?, ?)",
        (product_id, sku_name),
    )
    return db.execute(
        "SELECT id FROM products_skus WHERE product_id = ? AND sku_name = ?",
        (product_id, sku_name),
    ).fetchone()["id"]


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        upload = request.files.get("file")
        if not upload or not upload.filename:
            flash("请选择要导入的 Excel 或 CSV 文件。", "danger")
            return redirect(url_for("index"))
        try:
            records = import_rows(upload.filename, upload.read())
            with get_db() as db:
                for record in records:
                    factory_row = db.execute(
                        "SELECT id FROM factorys WHERE name = ? ORDER BY id LIMIT 1",
                        (record["factory"],),
                    ).fetchone()
                    if factory_row:
                        factory_id = factory_row["id"]
                    else:
                        factory_id = db.execute(
                            """INSERT INTO factorys(name, address, sort_num)
                               VALUES (?, ?, (SELECT COALESCE(MAX(sort_num), -1) + 1 FROM factorys))""",
                            (record["factory"], ""),
                        ).lastrowid
                    product_row = db.execute(
                        "SELECT id FROM products WHERE factory_id = ? AND name = ? ORDER BY id LIMIT 1",
                        (factory_id, record["name"]),
                    ).fetchone()
                    if product_row:
                        product_id = product_row["id"]
                    else:
                        product_id = db.execute(
                            """INSERT INTO products(factory_id, name, image, sort_num)
                               VALUES (?, ?, '', ?)""",
                            (
                                factory_id,
                                record["name"],
                                db.execute(
                                    "SELECT COALESCE(MAX(sort_num), -1) + 1 FROM products WHERE factory_id = ?",
                                    (factory_id,),
                                ).fetchone()[0],
                            ),
                        ).lastrowid
                    sku_id = ensure_product_sku(db, product_id, record["product_sku"])
                    db.execute(
                        """INSERT INTO orders(product_id, sku_id, quantity, order_date)
                           VALUES (?, ?, ?, ?)""",
                        (
                            product_id,
                            sku_id,
                            record["quantity"],
                            record["order_date"],
                        ),
                    )
            flash(f"成功导入 {len(records)} 条商品记录。", "success")
            return redirect(url_for("products"))
        except (ValueError, UnicodeDecodeError, csv.Error, OSError, zipfile.BadZipFile) as error:
            flash(str(error), "danger")
        except ImportError:
            app.logger.exception("Excel import dependency is missing")
            flash("Excel 导入组件不可用，请安装 requirements.txt 中的依赖。", "danger")
    with get_db() as db:
        missing_address_count = db.execute(
            "SELECT COUNT(*) FROM factorys WHERE address IS NULL OR TRIM(address) = ''"
        ).fetchone()[0]
        missing_image_count = db.execute(
            "SELECT COUNT(*) FROM products WHERE image IS NULL OR TRIM(image) = ''"
        ).fetchone()[0]
    return render_template(
        "index.html",
        missing_address_count=missing_address_count,
        missing_image_count=missing_image_count,
    )


@app.route("/products")
def products():
    query = request.args.get("name", "").strip()
    with get_db() as db:
        missing_image_count = db.execute(
            "SELECT COUNT(*) FROM products WHERE image IS NULL OR TRIM(image) = ''"
        ).fetchone()[0]
        total_product_count = db.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    has_missing_images = missing_image_count > 0
    displayed_products = (
        search_products(order_date=date.today().strftime("%Y%m%d")) if has_missing_images
        else search_products(query, date.today().strftime("%Y%m%d"))
    )
    if has_missing_images:
        displayed_products = [
            product for product in displayed_products
            if not str(product["image"] or "").strip()
        ]
    return render_template(
        "products.html",
        products=displayed_products,
        query=query,
        has_missing_images=has_missing_images,
        missing_image_count=missing_image_count,
        total_product_count=total_product_count,
    )


@app.route("/factorys")
def factorys():
    query = request.args.get("q", "").strip()
    with get_db() as db:
        missing_address_count = db.execute(
            "SELECT COUNT(*) FROM factorys WHERE address IS NULL OR TRIM(address) = ''"
        ).fetchone()[0]
        total_factory_count = db.execute("SELECT COUNT(*) FROM factorys").fetchone()[0]
        if missing_address_count:
            rows = db.execute(
                """SELECT * FROM factorys
                   WHERE address IS NULL OR TRIM(address) = ''
                   ORDER BY sort_num, id"""
            ).fetchall()
        else:
            rows = db.execute(
                """SELECT * FROM factorys
                   WHERE name LIKE ? OR address LIKE ?
                   ORDER BY sort_num, id""",
                (f"%{query}%", f"%{query}%"),
            ).fetchall()
    
    return render_template(
        "factorys.html",
        factorys=rows,
        query=query,
        has_missing_addresses=missing_address_count > 0,
        missing_address_count=missing_address_count,
        total_factory_count=total_factory_count,
    )


@app.route("/today")
def today():
    day = date.today().strftime("%Y%m%d")
    with get_db() as db:
        product_choices = db.execute(
            """SELECT p.id, p.name, f.name AS factory_name
               FROM products p
               JOIN factorys f ON f.id = p.factory_id
               ORDER BY f.sort_num, p.sort_num, p.id"""
        ).fetchall()
        sku_choices = db.execute(
            """SELECT id, product_id, sku_name FROM products_skus
               ORDER BY product_id, id"""
        ).fetchall()
        skus_by_product = {}
        for sku in sku_choices:
            skus_by_product.setdefault(sku["product_id"], []).append(dict(sku))
        product_choices = [
            {**dict(product), "skus": skus_by_product.get(product["id"], [])}
            for product in product_choices
        ]
        factory_rows = db.execute(
            """SELECT DISTINCT f.id, f.name, f.address, f.sort_num
               FROM factorys f
               JOIN products p ON p.factory_id = f.id
               JOIN orders o ON o.product_id = p.id
               WHERE o.order_date = ?
               ORDER BY f.sort_num, f.id""",
            (day,),
        ).fetchall()
        groups = []
        for factory_row in factory_rows:
            order_rows = db.execute(
                """SELECT o.id AS order_id, o.sku_id, s.sku_name, o.quantity, o.order_date,
                          p.id AS product_id, p.name AS product_name,
                          p.image, p.sort_num AS product_sort_num
                   FROM orders o
                   JOIN products p ON p.id = o.product_id
                   LEFT JOIN products_skus s ON s.id = o.sku_id
                   WHERE p.factory_id = ? AND o.order_date = ?
                   ORDER BY p.sort_num, p.id, o.id""",
                (factory_row["id"], day),
            ).fetchall()
            groups.append({"factory": factory_row, "orders": order_rows})
    return render_template(
        "today.html",
        groups=groups,
        product_choices=product_choices,
        selected_date=display_order_date(day),
        order_date=day,
    )


@app.route("/dashboard")
def dashboard():
    range_days = 30 if request.args.get("days") == "30" else 7
    today = date.today()
    start_day = today - timedelta(days=range_days - 1)
    start_date = start_day.strftime("%Y%m%d")
    end_date = today.strftime("%Y%m%d")
    with get_db() as db:
        daily_rows = db.execute(
            """SELECT o.order_date,
                      COALESCE(SUM(o.quantity), 0) AS pieces,
                      COUNT(DISTINCT p.id) AS products,
                      COUNT(DISTINCT f.id) AS factorys
               FROM orders o
               JOIN products p ON p.id = o.product_id
               JOIN factorys f ON f.id = p.factory_id
               WHERE o.order_date BETWEEN ? AND ?
               GROUP BY o.order_date
               ORDER BY o.order_date""",
            (start_date, end_date),
        ).fetchall()
    stats_by_day = {row["order_date"]: dict(row) for row in daily_rows}
    chart_data = []
    for offset in range(range_days):
        day = start_day + timedelta(days=offset)
        day_key = day.strftime("%Y%m%d")
        daily_stats = stats_by_day.get(
            day_key, {"pieces": 0, "products": 0, "factorys": 0}
        )
        chart_data.append({
            "date": display_order_date(day_key),
            "label": day.strftime("%m-%d"),
            "pieces": daily_stats["pieces"],
            "products": daily_stats["products"],
            "factorys": daily_stats["factorys"],
        })

    max_pieces = max((row["pieces"] for row in chart_data), default=0)
    graph_max = max(max_pieces, 1)
    graph_left = 54
    graph_right = 846
    graph_top = 30
    graph_bottom = 236
    graph_height = graph_bottom - graph_top
    graph_step = (graph_right - graph_left) / max(range_days - 1, 1)
    label_step = max(1, (range_days - 1 + 5) // 6)
    for index, row in enumerate(chart_data):
        row["x"] = round(graph_left + index * graph_step, 2)
        row["y"] = round(graph_bottom - row["pieces"] / graph_max * graph_height, 2)
        row["show_label"] = index % label_step == 0 or index == range_days - 1
    chart_points = " ".join(f"{row['x']},{row['y']}" for row in chart_data)
    chart_ticks = []
    for value in sorted({0, (graph_max + 1) // 2, graph_max}):
        chart_ticks.append({
            "value": value,
            "y": round(graph_bottom - value / graph_max * graph_height, 2),
        })
    total_pieces = sum(row["pieces"] for row in chart_data)
    peak_day = max(chart_data, key=lambda row: row["pieces"])

    return render_template(
        "dashboard.html",
        chart_data=chart_data,
        chart_points=chart_points,
        graph_bottom=graph_bottom,
        chart_ticks=chart_ticks,
        range_days=range_days,
        total_pieces=total_pieces,
        average_pieces=round(total_pieces / range_days, 1),
        peak_pieces=peak_day["pieces"],
        peak_date=peak_day["date"],
    )


def api_error(message, status=400):
    return jsonify({"error": message}), status


def move_row(db, table, row_id, direction, group_column=None, group_id=None):
    if direction not in {"up", "down"}:
        return False
    if table not in {"factorys", "products"}:
        raise ValueError("Unsupported sort table")
    current = db.execute(f"SELECT id, sort_num FROM {table} WHERE id = ?", (row_id,)).fetchone()
    if not current:
        return None
    if group_column:
        if group_column != "factory_id":
            raise ValueError("Unsupported sort group")
        rows = db.execute(
            f"SELECT id FROM {table} WHERE factory_id = ? ORDER BY sort_num, id",
            (group_id,),
        ).fetchall()
    else:
        rows = db.execute(f"SELECT id FROM {table} ORDER BY sort_num, id").fetchall()
    ids = [row["id"] for row in rows]
    position = ids.index(row_id)
    target = position - 1 if direction == "up" else position + 1
    if target < 0 or target >= len(ids):
        return True
    ids[position], ids[target] = ids[target], ids[position]
    for sort_num, item_id in enumerate(ids):
        db.execute(f"UPDATE {table} SET sort_num = ? WHERE id = ?", (sort_num, item_id))
    return True


@app.post("/api/products/<int:product_id>")
def update_product(product_id):
    data = request.get_json(silent=True) or {}
    field = data.get("field")
    value = data.get("value")
    if field != "factory_id":
        return api_error("无效的商品字段")
    try:
        value = int(value)
    except (TypeError, ValueError):
        return api_error("厂家无效")
    with get_db() as db:
        if not db.execute("SELECT 1 FROM factorys WHERE id = ?", (value,)).fetchone():
            return api_error("厂家不存在", 404)
        cursor = db.execute(
            "UPDATE products SET factory_id = ? WHERE id = ?", (value, product_id)
        )
        if cursor.rowcount == 0:
            return api_error("商品不存在", 404)
    return jsonify({"ok": True, "value": value})


@app.post("/api/skus/<int:sku_id>")
def update_sku(sku_id):
    data = request.get_json(silent=True) or {}
    value = data.get("value")
    if not isinstance(value, str) or not value.strip():
        return api_error("SKU 名称不能为空")
    value = value.strip()
    with get_db() as db:
        sku = db.execute(
            "SELECT product_id FROM products_skus WHERE id = ?", (sku_id,)
        ).fetchone()
        if not sku:
            return api_error("SKU 不存在", 404)
        if db.execute(
            "SELECT 1 FROM products_skus WHERE product_id = ? AND sku_name = ? AND id != ?",
            (sku["product_id"], value, sku_id),
        ).fetchone():
            return api_error("该产品已存在同名 SKU", 409)
        db.execute("UPDATE products_skus SET sku_name = ? WHERE id = ?", (value, sku_id))
    return jsonify({"ok": True, "value": value})


@app.post("/api/products/<int:product_id>/skus")
def create_sku(product_id):
    data = request.get_json(silent=True) or {}
    value = data.get("sku_name")
    if not isinstance(value, str) or not value.strip():
        return api_error("SKU 名称不能为空")
    value = value.strip()
    with get_db() as db:
        if not db.execute(
            "SELECT 1 FROM products WHERE id = ?", (product_id,)
        ).fetchone():
            return api_error("商品不存在", 404)
        if db.execute(
            "SELECT 1 FROM products_skus WHERE product_id = ? AND sku_name = ?",
            (product_id, value),
        ).fetchone():
            return api_error("该产品已存在同名 SKU", 409)
        sku_id = db.execute(
            "INSERT INTO products_skus(product_id, sku_name) VALUES (?, ?)",
            (product_id, value),
        ).lastrowid
    return jsonify({"ok": True, "id": sku_id, "sku_name": value}), 201


@app.post("/api/skus/<int:sku_id>/delete")
def delete_sku(sku_id):
    with get_db() as db:
        sku = db.execute(
            "SELECT id FROM products_skus WHERE id = ?", (sku_id,)
        ).fetchone()
        if not sku:
            return api_error("SKU 不存在", 404)
        db.execute("UPDATE orders SET sku_id = NULL WHERE sku_id = ?", (sku_id,))
        cursor = db.execute("DELETE FROM products_skus WHERE id = ?", (sku_id,))
    return jsonify({"ok": True})


@app.post("/api/products/<int:product_id>/orders")
def update_product_orders(product_id):
    data = request.get_json(silent=True) or {}
    try:
        order_date = normalize_order_date(data.get("order_date"))
    except ValueError as error:
        return api_error(str(error))
    entries = data.get("skus")
    if not order_date or not isinstance(entries, list):
        return api_error("排单日期或 SKU 数量格式无效")

    quantities = {}
    for entry in entries:
        if not isinstance(entry, dict):
            return api_error("SKU 数量格式无效")
        try:
            sku_id = int(entry.get("sku_id"))
            quantity = int(entry.get("quantity"))
        except (TypeError, ValueError):
            return api_error("SKU 或件数格式无效")
        if quantity < 0:
            return api_error("件数不能小于零")
        if sku_id in quantities:
            return api_error("排单中包含重复 SKU")
        quantities[sku_id] = quantity

    with get_db() as db:
        if not db.execute("SELECT 1 FROM products WHERE id = ?", (product_id,)).fetchone():
            return api_error("产品不存在", 404)
        valid_skus = {
            row["id"]
            for row in db.execute(
                "SELECT id FROM products_skus WHERE product_id = ?", (product_id,)
            )
        }
        if set(quantities) - valid_skus:
            return api_error("SKU 不属于该产品")
        db.execute(
            "DELETE FROM orders WHERE product_id = ? AND order_date = ? AND sku_id IS NOT NULL",
            (product_id, order_date),
        )
        db.executemany(
            """INSERT INTO orders(product_id, sku_id, quantity, order_date)
               VALUES (?, ?, ?, ?)""",
            [
                (product_id, sku_id, quantity, order_date)
                for sku_id, quantity in quantities.items()
                if quantity > 0
            ],
        )
    return jsonify({"ok": True})


@app.post("/api/products/<int:product_id>/image")
def upload_product_image(product_id):
    upload = request.files.get("image")
    if not upload or not upload.filename:
        return api_error("请选择要上传的产品图片")

    extension = Path(upload.filename).suffix.lower()
    signatures = {
        ".jpg": lambda data: data.startswith(b"\xff\xd8\xff"),
        ".jpeg": lambda data: data.startswith(b"\xff\xd8\xff"),
        ".png": lambda data: data.startswith(b"\x89PNG\r\n\x1a\n"),
        ".gif": lambda data: data.startswith((b"GIF87a", b"GIF89a")),
        ".webp": lambda data: len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP",
    }
    if extension not in signatures:
        return api_error("仅支持 JPG、PNG、GIF 或 WebP 图片")
    content = upload.read()
    if not content or not signatures[extension](content):
        return api_error("图片内容无效或与文件格式不匹配")

    image_name = f"{uuid.uuid4().hex}{extension}"
    image_directory = Path(app.root_path) / "static" / "images"
    image_directory.mkdir(parents=True, exist_ok=True)
    image_path = image_directory / image_name
    image_path.write_bytes(content)
    try:
        with get_db() as db:
            cursor = db.execute(
                "UPDATE products SET image = ? WHERE id = ?", (image_name, product_id)
            )
            if cursor.rowcount == 0:
                image_path.unlink(missing_ok=True)
                return api_error("商品不存在", 404)
    except Exception:
        image_path.unlink(missing_ok=True)
        raise
    return jsonify({"ok": True, "image": image_name})


@app.post("/api/products/<int:product_id>/move")
def move_product(product_id):
    data = request.get_json(silent=True) or {}
    with get_db() as db:
        product = db.execute("SELECT factory_id FROM products WHERE id = ?", (product_id,)).fetchone()
        if not product:
            return api_error("商品不存在", 404)
        result = move_row(
            db, "products", product_id, data.get("direction", ""),
            "factory_id", product["factory_id"],
        )
    if result is False:
        return api_error("无效的排序方向")
    return jsonify({"ok": True})


@app.post("/api/products/<int:product_id>/delete")
def delete_product(product_id):
    with get_db() as db:
        cursor = db.execute("DELETE FROM products WHERE id = ?", (product_id,))
        if cursor.rowcount == 0:
            return api_error("商品不存在", 404)
    return jsonify({"ok": True})


@app.post("/api/orders/<int:order_id>")
def update_order(order_id):
    data = request.get_json(silent=True) or {}
    field = data.get("field")
    value = data.get("value")
    if field not in {"quantity", "order_date"}:
        return api_error("无效的订单字段")
    if field == "quantity":
        try:
            value = int(value)
        except (TypeError, ValueError):
            return api_error("件数必须是整数")
        if value < 0:
            return api_error("件数不能小于零")
    elif field == "order_date":
        try:
            value = normalize_order_date(value)
        except ValueError as error:
            return api_error(str(error))
    with get_db() as db:
        cursor = db.execute(f"UPDATE orders SET {field} = ? WHERE id = ?", (value, order_id))
        if cursor.rowcount == 0:
            return api_error("订单不存在", 404)
    return jsonify({"ok": True, "value": display_order_date(value) if field == "order_date" else value})


@app.post("/api/orders")
def create_order():
    data = request.get_json(silent=True) or {}
    try:
        product_id = int(data.get("product_id"))
        quantity = int(data.get("quantity"))
    except (TypeError, ValueError):
        return api_error("商品或件数格式无效")
    if quantity < 0:
        return api_error("件数不能小于零")
    order_date = date.today().strftime("%Y%m%d")
    raw_sku_id = data.get("sku_id")
    if raw_sku_id in (None, ""):
        sku_id = None
    else:
        try:
            sku_id = int(raw_sku_id)
        except (TypeError, ValueError):
            return api_error("SKU 无效")
    with get_db() as db:
        if not db.execute("SELECT 1 FROM products WHERE id = ?", (product_id,)).fetchone():
            return api_error("商品不存在", 404)
        if sku_id is not None and not db.execute(
            "SELECT 1 FROM products_skus WHERE id = ? AND product_id = ?",
            (sku_id, product_id),
        ).fetchone():
            return api_error("SKU 不属于该商品")
        order_id = db.execute(
            """INSERT INTO orders(product_id, sku_id, quantity, order_date)
               VALUES (?, ?, ?, ?)""",
            (product_id, sku_id, quantity, order_date),
        ).lastrowid
    return jsonify({"ok": True, "id": order_id}), 201


@app.post("/api/orders/<int:order_id>/delete")
def delete_order(order_id):
    with get_db() as db:
        cursor = db.execute("DELETE FROM orders WHERE id = ?", (order_id,))
        if cursor.rowcount == 0:
            return api_error("订单不存在", 404)
    return jsonify({"ok": True})


@app.post("/api/factorys")
def create_factory():
    data = request.get_json(silent=True) or {}
    name = data.get("name", "").strip() if isinstance(data.get("name", ""), str) else ""
    if not name:
        return api_error("厂家名称不能为空")
    address = data.get("address", "").strip() if isinstance(data.get("address", ""), str) else ""
    with get_db() as db:
        db.execute(
            """INSERT INTO factorys(name, address, sort_num)
               VALUES (?, ?, (SELECT COALESCE(MAX(sort_num), -1) + 1 FROM factorys))""",
            (name, address),
        )
    return jsonify({"ok": True})


@app.post("/api/products")
def create_product():
    full_name = request.form.get("name", "")
    separator_match = re.search(r"[&＆]", full_name)
    if not separator_match:
        return api_error("名称格式应为“厂家名称&产品名称”")
    factory_name = full_name[: separator_match.start()]
    product_name = full_name[separator_match.end() :]
    factory_name = factory_name.strip()
    product_name = product_name.strip()
    if not factory_name or not product_name:
        return api_error("名称格式应为“厂家名称&产品名称”")

    try:
        sku_names = json.loads(request.form.get("sku_names", ""))
    except json.JSONDecodeError:
        return api_error("SKU 列表格式无效")
    if not isinstance(sku_names, list) or not sku_names:
        return api_error("至少需要添加一个 SKU")
    if any(not isinstance(sku, str) or not sku.strip() for sku in sku_names):
        return api_error("SKU 名称不能为空")
    sku_names = [sku.strip() for sku in sku_names]
    if len(set(sku_names)) != len(sku_names):
        return api_error("SKU 名称不能重复")

    upload = request.files.get("image")
    if not upload or not upload.filename:
        return api_error("请上传产品图片")
    extension = Path(upload.filename).suffix.lower()
    signatures = {
        ".jpg": lambda content: content.startswith(b"\xff\xd8\xff"),
        ".jpeg": lambda content: content.startswith(b"\xff\xd8\xff"),
        ".png": lambda content: content.startswith(b"\x89PNG\r\n\x1a\n"),
        ".gif": lambda content: content.startswith((b"GIF87a", b"GIF89a")),
        ".webp": lambda content: len(content) >= 12 and content[:4] == b"RIFF" and content[8:12] == b"WEBP",
    }
    if extension not in signatures:
        return api_error("仅支持 JPG、PNG、GIF 或 WebP 图片")
    image_content = upload.read()
    if not image_content or not signatures[extension](image_content):
        return api_error("图片内容无效或与文件格式不匹配")

    image_name = f"{uuid.uuid4().hex}{extension}"
    image_directory = Path(app.root_path) / "static" / "images"
    image_directory.mkdir(parents=True, exist_ok=True)
    image_path = image_directory / image_name
    image_path.write_bytes(image_content)
    try:
        with get_db() as db:
            factory = db.execute(
                "SELECT id FROM factorys WHERE name = ? ORDER BY sort_num, id LIMIT 1",
                (factory_name,),
            ).fetchone()
            factory_created = factory is None
            if factory:
                factory_id = factory["id"]
            else:
                factory_id = db.execute(
                    """INSERT INTO factorys(name, address, sort_num)
                       VALUES (?, ?, (SELECT COALESCE(MAX(sort_num), -1) + 1 FROM factorys))""",
                    (factory_name, ""),
                ).lastrowid

            if db.execute(
                "SELECT 1 FROM products WHERE factory_id = ? AND name = ?",
                (factory_id, product_name),
            ).fetchone():
                image_path.unlink(missing_ok=True)
                return api_error("该厂家下已存在同名产品", 409)
            product_id = db.execute(
                """INSERT INTO products(factory_id, name, image, sort_num)
                   VALUES (?, ?, ?, (SELECT COALESCE(MAX(sort_num), -1) + 1
                                  FROM products WHERE factory_id = ?))""",
                (factory_id, product_name, image_name, factory_id),
            ).lastrowid
            db.executemany(
                "INSERT INTO products_skus(product_id, sku_name) VALUES (?, ?)",
                [(product_id, sku_name) for sku_name in sku_names],
            )
    except Exception:
        image_path.unlink(missing_ok=True)
        raise
    return jsonify({
        "ok": True,
        "id": product_id,
        "factory_created": factory_created,
    }), 201


@app.post("/api/factorys/<int:factory_id>")
def update_factory(factory_id):
    data = request.get_json(silent=True) or {}
    field = data.get("field")
    value = data.get("value")
    if field not in {"name", "address"} or not isinstance(value, str):
        return api_error("无效的厂家字段或内容")
    if field == "name" and not value.strip():
        return api_error("厂家名称不能为空")
    with get_db() as db:
        cursor = db.execute(
            f"UPDATE factorys SET {field} = ? WHERE id = ?", (value.strip(), factory_id)
        )
        if cursor.rowcount == 0:
            return api_error("厂家不存在", 404)
    return jsonify({"ok": True})


@app.post("/api/factorys/<int:factory_id>/move")
def move_factory(factory_id):
    data = request.get_json(silent=True) or {}
    with get_db() as db:
        result = move_row(db, "factorys", factory_id, data.get("direction", ""))
    if result is None:
        return api_error("厂家不存在", 404)
    if result is False:
        return api_error("无效的排序方向")
    return jsonify({"ok": True})


@app.post("/api/factorys/<int:factory_id>/delete")
def delete_factory(factory_id):
    with get_db() as db:
        cursor = db.execute("DELETE FROM factorys WHERE id = ?", (factory_id,))
        if cursor.rowcount == 0:
            return api_error("厂家不存在", 404)
    return jsonify({"ok": True})


@app.errorhandler(413)
def file_too_large(_error):
    if request.endpoint == "upload_product_image":
        return api_error("图片大小不能超过 16 MB。", 413)
    flash("文件大小不能超过 16 MB。", "danger")
    return redirect(url_for("index"))



if __name__ == "__main__":
    app.run(debug=True)
