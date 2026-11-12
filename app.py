import csv
import io
import hmac
import os
import re
import secrets
import uuid
import zipfile
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, session, url_for
from dotenv import load_dotenv
from waitress import serve


from models import Factorys, Orders, Products, connect_db
from extensions import str_to_bool

# 加载 .env 文件，放在最前面！
load_dotenv()

app = Flask(__name__)
app.config["DEBUG"] = os.getenv("DEBUG")
app.config["DATABASE"] = os.getenv("DATABASE")
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY")
app.config["SESSION_COOKIE_SECURE"] = str_to_bool(os.getenv("SESSION_COOKIE_SECURE"))
app.config["SESSION_COOKIE_HTTPONLY"] = str_to_bool(os.getenv("SESSION_COOKIE_HTTPONLY"))
app.config["SESSION_COOKIE_SAMESITE"] = os.getenv("SESSION_COOKIE_SAMESITE")


@app.context_processor
def inject_csrf_token():
    def csrf_token():
        token = session.get("csrf_token")
        if token is None:
            token = secrets.token_urlsafe(32)
            session["csrf_token"] = token
        return token

    return {"csrf_token": csrf_token}


@app.before_request
def protect_unsafe_requests():
    if request.method in {"GET", "HEAD", "OPTIONS", "TRACE"}:
        return None
    expected_token = session.get("csrf_token")
    supplied_token = request.headers.get("X-CSRFToken") or request.form.get("csrf_token", "")
    if expected_token and hmac.compare_digest(expected_token, supplied_token):
        return None
    message = "CSRF token missing or invalid"
    if request.path.startswith("/api/"):
        return api_error(message, 400)
    return message, 400


# curl.exe -i -X GET http://127.0.0.1:5000/debug?name=TJ 
# curl.exe -i -X POST http://127.0.0.1:5000/debug -H "Content-Type: application/json" -d '{\"name\": \"taojun\"}' 
# curl.exe -i -X PUT http://127.0.0.1:5000/debug -F "name=taojun"
@app.route('/debug', methods=['GET', 'POST'])
def debug():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if name:
            address = request.form.get("address", "").strip()
            return f"Name: {name}, Address: {address}"
    return render_template("debug.html")


@contextmanager
def get_db():
    connection = connect_db(app)
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
        rows = rows[1:]
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
        order_content = f"{sku} : {quantity}"
        imported.append(
            {
                "name": name,
                "factory": factory,
                "image": "",
                "order_content": order_content,
                "order_date": date.today().strftime("%Y%m%d"),
            }
        )
    if not imported:
        raise ValueError("没有找到可导入的商品数据")
    return imported


def search_products(query=""):
    with get_db() as db:
        products_repository = Products(db)
        search_parts = re.split(r"[&＆]", query, maxsplit=1)
        if query and len(search_parts) == 2:
            products = products_repository.search_by_factory_and_product(
                search_parts[0].strip(), search_parts[1].strip()
            )
        else:
            products = products_repository.search(query)
        today_orders = Orders(db).list_for_date(date.today().strftime("%Y%m%d"))

    orders_by_product = {}
    for order in today_orders:
        orders_by_product.setdefault(order["product_id"], []).append(dict(order))
    result = [
        {**dict(product), "today_orders": orders_by_product.get(product["id"], [])}
        for product in products
    ]
    return sorted(
        result,
        key=lambda product: (
            not bool(product["today_orders"]),
        ),
    )


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            flash("请选择要导入的 Excel 或 CSV 文件。", "danger")
            return redirect(url_for("index"))
        if Path(upload.filename).suffix.lower() not in {".xlsx", ".xlsm", ".csv"}:
            flash("仅支持 Excel 或 CSV 文件。", "danger")
            return redirect(url_for("index"))
        try:
            records = import_rows(upload.filename, upload.read())
            with get_db() as db:
                factories = Factorys(db)
                products = Products(db)
                orders = Orders(db)
                for record in records:
                    factory = factories.find_by_name(record["factory"])
                    if factory is None:
                        factory_id = factories.create(record["factory"], "")
                    else:
                        factory_id = factory["id"]
                    product = products.find(factory_id, record["name"])
                    if product is None:
                        product_id = products.create(factory_id, record["name"])
                    else:
                        product_id = product["id"]
                    orders.create(product_id, record["order_content"], record["order_date"])
                has_missing_addresses = factories.count_missing_address() > 0
            flash(f"成功导入 {len(records)} 条商品记录。", "success")
            destination = "factorys" if has_missing_addresses else "products"
            return redirect(url_for(destination))
        except (ValueError, UnicodeDecodeError, csv.Error, OSError, zipfile.BadZipFile) as error:
            flash(str(error), "danger")
        except ImportError:
            app.logger.exception("Excel import dependency is missing")
            flash("Excel 导入组件不可用，请安装 requirements.txt 中的依赖。", "danger")
    with get_db() as db:
        missing_address_count = Factorys(db).count_missing_address()
        missing_image_count = Products(db).count_missing_image()
    return render_template(
        "index.html",
        missing_address_count=missing_address_count,
        missing_image_count=missing_image_count,
    )


@app.route("/products")
def products():
    query = request.args.get("name", "").strip()
    with get_db() as db:
        missing_image_count = Products(db).count_missing_image()
        total_product_count = Products(db).count()
    has_missing_images = missing_image_count > 0
    displayed_products = (
        search_products() if has_missing_images
        else search_products(query)
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
    after_address_product_id = request.args.get("after_address_product", type=int)
    with get_db() as db:
        factories = Factorys(db)
        missing_address_count = factories.count_missing_address()
        total_factory_count = factories.count()
        if missing_address_count:
            rows = factories.list(missing_address_only=True)
        else:
            rows = factories.list(query)
        after_address_factory_id = None
        after_address_redirect_url = None
        if after_address_product_id is not None:
            product = Products(db).get(after_address_product_id)
            if product is not None:
                factory = factories.get(product["factory_id"])
                if factory is not None and not str(factory["address"] or "").strip():
                    after_address_factory_id = factory["id"]
                    after_address_redirect_url = url_for(
                        "products",
                        name=f"{factory['name']}&{product['name']}",
                    )

    return render_template(
        "factorys.html",
        factorys=rows,
        query=query,
        has_missing_addresses=missing_address_count > 0,
        missing_address_count=missing_address_count,
        total_factory_count=total_factory_count,
        reorder_enabled=not missing_address_count and not query,
        after_address_factory_id=after_address_factory_id,
        after_address_redirect_url=after_address_redirect_url,
    )


@app.route("/today")
def today():
    day = date.today().strftime("%Y%m%d")
    with get_db() as db:
        factory_rows = Factorys(db).list_for_order_date(day)
        groups = []
        for factory_row in factory_rows:
            order_rows = Products(db).list_orders_for_factory(factory_row["id"], day)
            groups.append({"factory": factory_row, "orders": order_rows})
    return render_template(
        "today.html",
        groups=groups,
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
        orders = Orders(db)
        daily_rows = orders.list_dashboard(start_date, end_date)
        totals = orders.dashboard_totals(start_date, end_date)
    stats_by_day = {row["order_date"]: dict(row) for row in daily_rows}
    chart_data = []
    for offset in range(range_days):
        day = start_day + timedelta(days=offset)
        day_key = day.strftime("%Y%m%d")
        daily_stats = stats_by_day.get(
            day_key, {"products": 0, "factories": 0}
        )
        chart_data.append({
            "date": display_order_date(day_key),
            "label": day.strftime("%m-%d"),
            "products": daily_stats["products"],
            "factories": daily_stats["factories"],
        })

    graph_max = max(
        (max(row["factories"], row["products"]) for row in chart_data),
        default=0,
    )
    graph_max = max(graph_max, 1)
    graph_left = 54
    graph_right = 846
    graph_top = 30
    graph_bottom = 236
    graph_height = graph_bottom - graph_top
    graph_step = (graph_right - graph_left) / max(range_days - 1, 1)
    label_step = max(1, (range_days - 1 + 5) // 6)
    for index, row in enumerate(chart_data):
        row["x"] = round(graph_left + index * graph_step, 2)
        row["factory_y"] = round(
            graph_bottom - row["factories"] / graph_max * graph_height, 2
        )
        row["product_y"] = round(
            graph_bottom - row["products"] / graph_max * graph_height, 2
        )
        row["show_label"] = index % label_step == 0 or index == range_days - 1
    factory_points = " ".join(
        f"{row['x']},{row['factory_y']}" for row in chart_data
    )
    product_points = " ".join(
        f"{row['x']},{row['product_y']}" for row in chart_data
    )
    chart_ticks = []
    for value in sorted({0, (graph_max + 1) // 2, graph_max}):
        chart_ticks.append({
            "value": value,
            "y": round(graph_bottom - value / graph_max * graph_height, 2),
        })
    return render_template(
        "dashboard.html",
        chart_data=chart_data,
        factory_points=factory_points,
        product_points=product_points,
        graph_bottom=graph_bottom,
        chart_ticks=chart_ticks,
        range_days=range_days,
        total_factories=totals["factories"],
        total_products=totals["products"],
    )


def api_error(message, status=400):
    return jsonify({"error": message}), status


def read_image_upload(upload):
    if upload is None or not upload.filename:
        raise ValueError("请选择要上传的产品图片")
    extension = Path(upload.filename).suffix.lower()
    signatures = {
        ".jpg": lambda content: content.startswith(b"\xff\xd8\xff"),
        ".jpeg": lambda content: content.startswith(b"\xff\xd8\xff"),
        ".png": lambda content: content.startswith(b"\x89PNG\r\n\x1a\n"),
        ".gif": lambda content: content.startswith((b"GIF87a", b"GIF89a")),
        ".webp": lambda content: len(content) >= 12 and content[:4] == b"RIFF" and content[8:12] == b"WEBP",
    }
    if extension not in signatures:
        raise ValueError("仅支持 JPG、PNG、GIF 或 WebP 图片")
    content = upload.read()
    if not content or not signatures[extension](content):
        raise ValueError("图片内容无效或与文件格式不匹配")
    return extension, content


def parse_integer(value):
    if isinstance(value, bool) or not isinstance(value, (int, str)):
        raise ValueError
    return int(value)


@app.post("/api/products/<int:product_id>")
def update_product(product_id):
    data = request.get_json(silent=True) or {}
    field = data.get("field")
    value = data.get("value")
    if field != "name" or not isinstance(value, str) or not value.strip():
        return api_error("产品名称不能为空")
    value = value.strip()
    with get_db() as db:
        product = Products(db).get(product_id)
        if product is None:
            return api_error("商品不存在", 404)
        products = Products(db)
        if products.exists_name(product["factory_id"], value, excluding_id=product_id):
            return api_error("该厂家下已存在同名产品", 409)
        products.update(product_id, name=value)
    return jsonify({"ok": True, "value": value})


@app.post("/api/products/<int:product_id>/image")
def upload_product_image(product_id):
    try:
        extension, content = read_image_upload(request.files.get("image"))
    except ValueError as error:
        return api_error(str(error))

    image_name = f"{uuid.uuid4().hex}{extension}"
    image_directory = Path(app.root_path) / "static" / "images"
    image_directory.mkdir(parents=True, exist_ok=True)
    image_path = image_directory / image_name
    image_path.write_bytes(content)
    try:
        with get_db() as db:
            products = Products(db)
            product = products.get(product_id)
            if product is None:
                image_path.unlink(missing_ok=True)
                return api_error("商品不存在", 404)
            products.update(product_id, image=image_name)
    except Exception:
        image_path.unlink(missing_ok=True)
        raise
    return jsonify({"ok": True, "image": image_name})


@app.post("/api/products/<int:product_id>/delete")
def delete_product(product_id):
    with get_db() as db:
        products = Products(db)
        product = products.get(product_id)
        if product is None:
            return api_error("商品不存在", 404)
        products.delete(product_id)
    return jsonify({"ok": True})


@app.post("/api/orders/<int:order_id>")
def update_order(order_id):
    data = request.get_json(silent=True) or {}
    field = data.get("field")
    value = data.get("value")
    if field not in {"order_content", "order_date"}:
        return api_error("无效的订单字段")
    if field == "order_content":
        if not isinstance(value, str) or not value.strip():
            return api_error("排单内容不能为空")
        value = value.strip()
    else:
        try:
            value = normalize_order_date(value)
        except ValueError as error:
            return api_error(str(error))
    with get_db() as db:
        order = Orders(db).get(order_id)
        if order is None:
            return api_error("订单不存在", 404)
        Orders(db).update(order_id, **{field: value})
    return jsonify({"ok": True, "value": display_order_date(value) if field == "order_date" else value})


@app.post("/api/orders")
def create_order():
    data = request.get_json(silent=True) or {}
    if not isinstance(data, dict):
        return api_error("订单数据格式无效")
    try:
        product_id = parse_integer(data.get("product_id"))
    except (TypeError, ValueError):
        return api_error("请选择产品")
    order_content = data.get("order_content")
    if not isinstance(order_content, str) or not order_content.strip():
        return api_error("排单内容不能为空")
    order_content = order_content.strip()
    order_date = date.today().strftime("%Y%m%d")
    with get_db() as db:
        if Products(db).get(product_id) is None:
            return api_error("商品不存在", 404)
        order_id = Orders(db).create(product_id, order_content, order_date)
    return jsonify({"ok": True, "id": order_id}), 201


@app.post("/api/factorys")
def create_factory():
    data = request.get_json(silent=True) or {}
    if not isinstance(data, dict):
        return api_error("厂家数据格式无效")
    name = data.get("name")
    if not isinstance(name, str) or not name.strip():
        return api_error("厂家名称不能为空")
    name = name.strip()
    address = data.get("address", "")
    if not isinstance(address, str):
        return api_error("厂家地址格式无效")
    address = address.strip()
    with get_db() as db:
        Factorys(db).create(name, address)
    return jsonify({"ok": True})


@app.post("/api/products")
def create_product():
    full_name = request.form.get("name", "").strip()
    if not full_name:
        return api_error("产品编码不能为空")
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
        extension, image_content = read_image_upload(request.files.get("image"))
    except ValueError as error:
        return api_error(str(error))

    image_name = f"{uuid.uuid4().hex}{extension}"
    image_directory = Path(app.root_path) / "static" / "images"
    image_directory.mkdir(parents=True, exist_ok=True)
    image_path = image_directory / image_name
    image_path.write_bytes(image_content)
    try:
        with get_db() as db:
            factories = Factorys(db)
            products = Products(db)
            factory = factories.find_by_name(factory_name, order_by_sort=True)
            factory_created = factory is None
            if factory:
                factory_id = factory["id"]
            else:
                factory_id = factories.create(factory_name, "")

            if products.find(factory_id, product_name) is not None:
                image_path.unlink(missing_ok=True)
                return api_error("该厂家下已存在同名产品", 409)
            product_id = products.create(factory_id, product_name, image_name)
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
        factories = Factorys(db)
        factory = factories.get(factory_id)
        if factory is None:
            return api_error("厂家不存在", 404)
        factories.update(factory_id, **{field: value.strip()})
    return jsonify({"ok": True})


@app.post("/api/factorys/reorder")
def reorder_factorys():
    data = request.get_json(silent=True) or {}
    factory_ids = data.get("factory_ids") if isinstance(data, dict) else None
    if not isinstance(factory_ids, list) or any(
        isinstance(factory_id, bool) or not isinstance(factory_id, int) or factory_id < 1
        for factory_id in factory_ids
    ):
        return api_error("厂家排序数据无效")
    with get_db() as db:
        if not Factorys(db).reorder(factory_ids):
            return api_error("厂家列表已变化，请刷新页面后重试", 409)
    return jsonify({"ok": True})


@app.post("/api/factorys/<int:factory_id>/delete")
def delete_factory(factory_id):
    with get_db() as db:
        factories = Factorys(db)
        factory = factories.get(factory_id)
        if factory is None:
            return api_error("厂家不存在", 404)
        factories.delete(factory_id)
    return jsonify({"ok": True})


@app.errorhandler(413)
def file_too_large(_error):
    if request.endpoint == "upload_product_image":
        return api_error("图片大小不能超过 16 MB。", 413)
    flash("文件大小不能超过 16 MB。", "danger")
    return redirect(url_for("index"))



if __name__ == "__main__":
    if(os.getenv('DEBUG')=='True'):
        app.run(debug=True)
    else: 
        host = os.environ.get("APP_HOST")
        port = os.environ.get("APP_PORT")
        threads = os.environ.get("APP_THREADS")
        print('启动生产服务器...', host, port, threads)
        serve(app, host=host, port=port, threads=threads)
        