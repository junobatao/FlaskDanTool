
* **Flask + BootStrap5 + SQLite**
* **dotenv（.env）**
* **Jinja2模板引擎（XSS）**
* **HTML 原生表单（Bootstrap 5 校验、服务端校验与 CSRF 防护）**
* **Python 标准库 sqlite3（参数化 SQL）**
* **openpyxl、csv（Excel、CSV）**
* **uuid（image_name）**
* **SVG 折线图表**
* **@print打印机样式**
* **Toggle切换背景颜色**

数据库使用 Python 标准库 `sqlite3`，不需要 ORM 或迁移扩展。数据库位置由
`.env` 中的 `SQLITE_DATABASE_PATH` 配置；`sqlite:///database.db` 表示
`instance/database.db`。也可使用绝对文件路径。旧配置 `DATABASE_URI` 仍可用。

`init_db.py` 独立于 Flask 应用运行，不导入或启动 `app.py`。首次建库或创建缺失的表和索引时运行：

```powershell
python init_db.py
```

初始化使用 `CREATE TABLE/INDEX IF NOT EXISTS`，不会清空已有记录。启动初始化时会将旧订单中的 SKU 与件数迁移为 `order_content` 文本，并移除旧 SKU 表、订单 SKU/件数字段和产品排序字段。排单内容以文本保存；导入表格时将“销售属性”和“数量”用 ` : ` 拼接。产品名称可在产品列表中直接编辑；每日统计按当天有订单的不同厂家和产品款数计算。厂家排序仍保留。
