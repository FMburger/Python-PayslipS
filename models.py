import pandas as pd


class EmployeeListModel:
    def __init__(self, conn):
        self.conn = conn
        self.employeeList = self.create_employeeList()

    def create_employeeList(self):
        query_employeeList = (
            'SELECT DISTINCT TI002, TI004, TI001 ' +
            'FROM PALTI ' +
            'ORDER BY \'TI002\' DESC'
        )
        df_employeeList = pd.read_sql(query_employeeList, self.conn)
        return df_employeeList

    def get_list_payPeriods(self):
        list_payPeriods = self.employeeList[self.employeeList.columns[0]].unique().tolist()
        return list_payPeriods

    def get_list_departments(self, payPeriod):
        df_departments = self.employeeList.query(
            self.employeeList.columns[0] +
            ' == \'' +
            payPeriod +
            '\''
        )
        list_departments = df_departments[self.employeeList.columns[1]].unique().tolist()
        list_departments.sort()
        return list_departments

    def get_list_employees(self, payPeriod, department):
        if department == '所有部門':
            df_employees = self.employeeList.query(
                self.employeeList.columns[0] +
                ' == \'' +
                payPeriod +
                '\''
            )
        else:
            df_employees = self.employeeList.query(
                self.employeeList.columns[0] +
                ' == \'' +
                payPeriod +
                '\' and ' +
                self.employeeList.columns[1] +
                ' == \'' +
                department +
                '\''
            )
        list_employees = df_employees[self.employeeList.columns[2]].unique().tolist()
        list_employees.sort()
        return list_employees


class PayslipModel:
    def __init__(self, conn):
        self.conn = conn

    def create_list_employees(self, payPeriod, department, employee):
        if employee == '所有員工':
            if department == '所有部門': 
                query_employees = (
                    'SELECT DISTINCT TI001 ' +
                    'FROM PALTI ' +
                    'WHERE TI002 = \'' +
                    payPeriod +
                    '\' ' +
                    'ORDER BY \'TI001\' ASC'
                )
            else:
                query_employees = (
                    'SELECT DISTINCT TI001 ' +
                    'FROM PALTI ' +
                    'WHERE TI002 = \'' +
                    payPeriod +
                    '\' and TI004 = \'' +
                    department +
                    '\' ' +
                    'ORDER BY \'TI001\' ASC'
                )
        else:
            query_employees = (
                'SELECT DISTINCT TI001 ' +
                'FROM PALTI ' +
                'WHERE TI001 = \'' +
                employee +
                '\' ' +
                'ORDER BY \'TI001\' ASC'
            )
        df_employees = pd.read_sql(query_employees, self.conn)
        list_employees = df_employees['TI001'].tolist()
        return list_employees

    def get_profile(self, employee):
        query_profile = (
            'SELECT MV001, MV002, MV004, MV008, MV006, MV020, MV009 ' +
            'FROM CMSMV ' +
            'WHERE MV001 = \'' +
            employee +
            '\' ' +
            'ORDER BY \'MV001\' DESC'
            )
        df_profile = pd.read_sql(query_profile, self.conn)
        return df_profile

    def get_palti(self, payPeriod, employee):
        query_palti = (
                'SELECT TI002 as \'發薪年月\', ' +
                'TI004 AS \'部門\', ' +
                'TI001 AS \'員工編號\', ' +
                'TI023 + TI024 AS \'底薪\', ' +
                'TI029 + TI030 AS \'全勤獎金\', ' +
                'TI025 + TI026 AS \'加班費免稅\', ' +
                'TI027 + TI028 AS \'加班費課稅\', ' +
                'TI031 + TI032 AS \'請假扣款\', ' +
                'TI033 AS \'健保費\', ' +
                'TI034 AS \'勞保費\', ' +
                'TI040 + TI041 + TI033 + TI034 AS \'應發金額\', ' +
                'TI040 + TI041 AS \'實發金額\', ' +
                'TI042 + TI043 AS \'課稅金額\', ' +
                'TI027 + TI028 + TI044 + TI045 AS \'津貼合計\', ' +
                'TI054 AS \'公司提繳\', ' +
                'TI015 + TI016 AS \'出勤天數\', ' +
                'TI007 + TI008 + TI009 + TI010 + TI011 + TI012 + TI062 + TI063 + TI064 + TI065 + TI066 + TI067 AS \'加班時數\', ' +
                'TI033 + TI034 AS \'扣款合計\', ' +
                'TI035 + TI036 AS \'所得稅\' ' +
                'FROM PALTI ' +
                'where TI001 = ' +
                '\'' +
                employee +
                '\' ' +
                'and ' +
                'TI002 = ' +
                payPeriod +
                ' ORDER BY \'員工編號\' ASC'
            )
        df_palti = pd.read_sql(query_palti, self.conn)
        return df_palti

    def get_paltj(self, payPeriod, employee):
        query_paltj = (
            'SELECT DISTINCT TJ004 AS \'津貼扣款名稱\', ' +
            'TJ007 AS \'津貼扣款金額\', ' +
            'TJ005 AS \'加扣項\' ' +
            'FROM PALTJ ' +
            'WHERE TJ001 = ' +
            '\'' +
            employee +
            '\' ' +
            'AND TJ002 = ' +
            payPeriod
        )
        df_paltj = pd.read_sql(query_paltj, self.conn)
        return df_paltj

    def get_email(self, employee):
        query_email = (
            'SELECT MV020 ' +
            'FROM CMSMV ' +
            'where MV001 = ' +
            '\'' +
            employee +
            '\' ' +
            'ORDER BY MV001 ASC'
        )
        df_email = pd.read_sql(query_email, self.conn)
        return df_email


class EmailListModel:
    def __init__(self, conn):
        self.conn = conn
        self.emailList = self.get_emailList()

    def get_emailList(self):
        query_emailList = (
            'SELECT DISTINCT MV001, MV002, MV020 ' +
            'FROM CMSMV ' +
            'WHERE MV022 = \'\' ' +
            'ORDER BY \'MV001\' DESC'
        )
        df_emailList = pd.read_sql(query_emailList, self.conn)
        list_emailList = df_emailList.values.tolist()
        return list_emailList
