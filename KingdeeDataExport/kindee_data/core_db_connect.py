import pymssql
from . import core_other_funcs as funcs


class SQLServer:
    def __init__(self, server, user, password, database):
        # 类的构造函数，初始化DBC连接信息
        self.server = server
        self.user = user
        self.password = password
        self.database = database

    def __GetConnect(self):
        # 得到数据库连接信息，返回conn.cursor()
        if not self.database:
            print(NameError, "没有设置数据库信息")
        try:
            self.conn = pymssql.connect(
                server=self.server, user=self.user, password=self.password, database=self.database)
        except Exception as e:
            print('数据库连接失败！\n错误信息：', e)
            funcs.exit_with_anykey()
        cur = self.conn.cursor()
        if not cur:
            print(NameError, "连接数据库失败！")  # 将DBC信息赋值给cur
        else:
            return cur

    def ExecQuery(self, sql):
        '''
        执行查询语句
        返回一个包含tuple的list，list是元素的记录行，tuple记录每行的字段数值
        '''
        cur = self.__GetConnect()
        try:
            cur.execute(sql)  # 执行查询语句
        except Exception as e:
            print('无法执行SQL语句！\n错误信息：', e)
            funcs.exit_with_anykey()
        result = cur.fetchall()  # fetchall()获取查询结果
        # 查询完毕关闭数据库连接
        self.conn.close()
        return result

    def ExecNonQuery(self, sql):
        try:
            cur = self.__GetConnect()
            cur.execute(sql)
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise Exception
        finally:
            self.conn.close()


def connect_db():
    return SQLServer(server='...', user='',
                     password='', database='master')


def get_all_db_company(msg):
    db_list = msg.ExecQuery('''
        SELECT FDBName, FAcctName FROM [KDAcctDB].[dbo].[t_ad_kdAccount_gl]
    ''')  # 获取账套名与数据库名
    db_name = []
    company_name = []
    for i in range(len(db_list)):
        company_name.append(db_list[i][1])  # 获取公司名
        db_name.append(db_list[i][0])  # 获取数据库名
    return db_name, company_name
