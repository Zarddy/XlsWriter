
import pymysql


class DatabaseHelper:

    # --------------------------------------------------①连接数据库和定义游标
    def __init__(self, h='localhost', u='root', p='root', db='database_name'):
        self.db = pymysql.connect(host=h, user=u, password=p, database=db,
                                  cursorclass=pymysql.cursors.DictCursor)
        self.cursor = self.db.cursor()

    # -----------------------------------------------------------②操作
    # 查询操作
    def select(self, sql):
        self.cursor.execute(sql)  # 执行sql语句
        return self.cursor.fetchall()  # 会获取所有数据

    # 增删改操作
    def change(self, sql):
        self.cursor.execute(sql)
        self.db.commit()  # 提交数据
        print('操作成功！')
        return self.cursor.rowcount  # 获取操作的行数

    # ------------------------------------------------③断开连接
    # 自动关闭连接
    def __del__(self):
        self.cursor.close()
        self.db.close()
