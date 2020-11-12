import pyodbc


class Connection:
    def __init__(self, server, database, userID, password):
        self.connectionInfo = ''
        self.conn = self.create_connection(server, database, userID, password)

    def create_connection(self, server, database, userID, password):
        try:
            # Download ODBC Driver 13 (https://www.microsoft.com/en-us/download/details.aspx?id=50420)
            cnxn = pyodbc.connect(
                'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' +
                server +
                ';''DATABASE=' +
                database +
                ';UID=' +
                userID +
                ';PWD=' +
                password
            )
        except:
            print('資料庫連線失敗。')
            self.connectionInfo = ('資料庫連線失敗。')
        else:
            print('資料庫連線成功!')
            self.connectionInfo = ('資料庫連線成功!')
            return cnxn


if __name__ == '__main__':
    print('資料庫連線測試')
    server = input('server name: ')
    database = input('database: ')
    userID = input('user ID: ')
    password = input('password: ')
    print('\n')
    Connection(server, database, userID, password).conn
