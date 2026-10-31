import threading
from contextlib import contextmanager


class MySQLClient:
    """MySQL 数据库访问器（懒连接、参数化查询、事务与自增主键）。"""

    def __init__(
        self,
        host,
        port,
        user,
        password,
        database,
        charset="utf8mb4",
    ):
        try:
            import pymysql

            self.pymysql = pymysql
        except ImportError as exc:
            raise ImportError(
                "检测到未安装 pymysql。请执行 'pip install pymysql'。"
            ) from exc

        self.host = host
        self.port = port
        self.user = user
        self.password = password
        self.database = database
        self.charset = charset

        self.connection = None
        self._lock = threading.RLock()

    # ---------------- 连接管理 ----------------

    def connect(self):
        """懒连接；连接断开时尝试重连。"""
        with self._lock:
            if self.connection is None:
                self.connection = self._create_connection()
            else:
                try:
                    self.connection.ping(reconnect=True)
                except self.pymysql.Error:
                    self.connection = self._create_connection()

        return self.connection

    def _create_connection(self):
        """创建新的数据库连接。"""
        return self.pymysql.connect(
            host=self.host,
            port=self.port,
            user=self.user,
            password=self.password,
            database=self.database,
            charset=self.charset,
        )

    @contextmanager
    def cursor(self, dict_cursor=False):
        """获取游标；dict_cursor=True 时返回字典行。"""
        with self._lock:
            connection = self.connect()

            cursor_cls = (
                self.pymysql.cursors.DictCursor
                if dict_cursor
                else self.pymysql.cursors.Cursor
            )

            with connection.cursor(cursor_cls) as cursor:
                yield cursor

    # ---------------- 读操作 ----------------

    def query(self, sql, params=None):
        """执行查询，返回字典行列表。"""
        with self.cursor(dict_cursor=True) as cursor:
            cursor.execute(sql, params)
            return list(cursor.fetchall())

    def query_one(self, sql, params=None):
        """查询第一行；无结果时返回 None。"""
        with self.cursor(dict_cursor=True) as cursor:
            cursor.execute(sql, params)
            return cursor.fetchone()

    def count(self, sql, params=None):
        """执行 COUNT 查询，返回整数。"""
        row = self.query_one(sql, params)

        if not row:
            return 0

        return int(next(iter(row.values())))

    # ---------------- 写操作 ----------------

    def execute(self, sql, params=None):
        """执行 INSERT、UPDATE、DELETE 等语句，提交并返回影响行数。"""
        with self.cursor() as cursor:
            cursor.execute(sql, params)
            self.connection.commit()
            return cursor.rowcount

    def insert(self, sql, params=None):
        """执行 INSERT，提交并返回自增主键。"""
        with self.cursor() as cursor:
            cursor.execute(sql, params)
            lastrowid = cursor.lastrowid
            self.connection.commit()
            return int(lastrowid or 0)

    # ---------------- 事务 ----------------

    @contextmanager
    def transaction(self):
        """事务上下文；成功提交，异常回滚。"""
        with self._lock:
            connection = self.connect()

            try:
                yield connection
                connection.commit()
            except Exception:
                connection.rollback()
                raise

    # ---------------- 关闭 ----------------

    def close(self):
        """关闭连接并重置连接状态。"""
        with self._lock:
            if self.connection is not None:
                self.connection.close()
                self.connection = None
