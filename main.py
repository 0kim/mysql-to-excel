import pymysql
import xlsxwriter

mysql_host = '127.0.0.1'
mysql_user = 'root'
mysql_pass = 'mypassword'
mysql_db = 'mydb'
mysql_charset = 'utf8'
mysql_query = 'select * from table;'

xls_filename = 'output.xlsx'
xls_sheetname = 'work1'


class XlsxSheetWriter():
    MAX_ROW = 1048576
    MAX_COL = 16384     #  == 'XFD' @ excel
    DEFAULT_ROW = 0
    DEFAULT_COL = 0

    _cur_col = 0  # A, B, C...
    _cur_row = 0  # 1, 2, 3...
    _workbook = None  # workbook(xls file)
    _worksheet = None # worksheet

    def __init__(self, filename, sheetname):
        self._cur_col = self.DEFAULT_COL
        self._cur_row = self.DEFAULT_ROW
        self._wb = xlsxwriter.Workbook(filename)
        self._ws = self._wb.add_worksheet(sheetname)

    def reset_row(self):
        self._cur_row = self.DEFAULT_ROW

    def reset_col(self):
        self._cur_col = self.DEFAULT_COL

    def inc_row(self):
        if self.MAX_ROW <= self._cur_row:
            raise Exception("Out of range: row " + str(self._cur_row));
        self._cur_row += 1

    def inc_col(self):
        if self.MAX_COL <= self._cur_col:
            raise Exception("Out of range: column " + str(self._cur_col));
        self._cur_col += 1

    def write(self, data):
        self._ws.write(self._cur_row, self._cur_col, data);

    def close(self):
        self._wb.close()


writer = XlsxSheetWriter(filename=xls_filename,
                         sheetname=xls_sheetname)

con = pymysql.connect(host=mysql_host,
                     user=mysql_user,
                     password=mysql_pass,
                     db=mysql_db,
                     charset=mysql_charset)

curs = con.cursor(pymysql.cursors.DictCursor)
curs.execute(mysql_query)
rows = curs.fetchall()

writer.reset_row()
for row in rows:
    writer.reset_col()
    for c in row:
        writer.write(row[c])
        writer.inc_col()
    writer.inc_row()

con.close()
writer.close()

print("Done...")