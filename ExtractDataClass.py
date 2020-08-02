from pandas import read_csv, ExcelFile, to_datetime, DataFrame, ExcelWriter, date_range
from datetime import timedelta


class ExtractData:

    def __init__(self, **kwargs):


        valid_keys = ["faults_file_name", "transaction_file_name", "input_folder", "output_folder",
                  "faults_sheet_name", "employees_sheet_name", "categories_sheet_name", "transation_cols",
                  "faults_from_c", "faults_to_c", "faults_from_r",
                  "emp_from_c", "emp_to_c", "emp_from_r",
                  "cat_from_c", "cat_to_c", "cat_from_r","from_d","to_d"]

        for key in valid_keys:
                setattr(self, key, kwargs.get(key))

    def __read_transaction(self):

        transactions = read_csv(self.transaction_file_name, encoding="ISO-8859-8")
        transactions.columns = transactions.columns.str.lstrip()
        transactions.columns = transactions.columns.str.rstrip()
        df_transactions = transactions[self.transation_cols]
        self.df_transactions = df_transactions
        
        if len(df_transactions) == 0:
            raise ValueError

    def __export_report_file(self):

        xls = ExcelFile(self.faults_file_name)
        self.faults_file = xls

    def __read_df(self,sheet_name, from_c, to_c, skip_r):

        df = self.faults_file.parse(sheet_name,
                                    usecols=range(from_c, to_c),
                                    skiprows=skip_r,
                                    index_col=None,
                                    na_values=['NA'])
        df = df.dropna()
        if len(df) == 0:
            raise ValueError

        return df

    def __read_all_df(self):

        self.df_faults = self.__read_df(self.faults_sheet_name,
                                        self.faults_from_c,
                                        self.faults_to_c,
                                        self.faults_from_r)

        self.df_categories = self.__read_df(self.categories_sheet_name,
                                            self.cat_from_c,
                                            self.cat_to_c,
                                            self.cat_from_r)

        self.df_employees = self.__read_df(self.employees_sheet_name,
                                           self.emp_from_c,
                                           self.emp_to_c,
                                           self.emp_from_r)
        self.df_dates = self.__create_date_dim(self.from_d,
                                               self.to_d)

    def __create_date_dim(self,start='2020-1-1', end='2020-12-31'):


        start_ts = to_datetime(start).date()
        end_ts = to_datetime(end).date()
        first_week = start_ts.isocalendar()[1]

        # record timetsamp is empty for now
        dates = DataFrame(index=date_range(start_ts, end_ts))
        dates.index.name = 'Date'

        days_names = {
            i: name
            for i, name
            in enumerate(['Monday', 'Tuesday', 'Wednesday',
                          'Thursday', 'Friday', 'Saturday',
                          'Sunday'])
        }

        dates['Day'] = dates.index.dayofweek.map(days_names.get)
        dates['Week'] = dates.index.week
        dates['Month'] = dates.index.month
        dates['Quarter'] = dates.index.quarter
        dates['Year_half'] = dates.index.month.map(lambda mth: 1 if mth < 7 else 2)
        dates['Year'] = dates.index.year
        dates.reset_index(inplace=True)
        dates.index.name = 'date_id'

        dates['Week_Project'] = dates['Date'].apply(lambda x: (x + timedelta(days=1)).week)
        first_week = min(dates.Week_Project)
        dates['Week_Project'] -= first_week - 1

        return dates



    def load_data(self):

        with ExcelWriter(self.output_folder +  '//output.xlsx') as writer:

            self.df_employees.to_excel(writer, sheet_name='Employees')
            self.df_faults.to_excel(writer, sheet_name='Faults')
            self.df_categories.to_excel(writer, sheet_name='Categories')
            self.df_transactions.to_excel(writer, sheet_name='Transactions')
            self.df_dates.to_excel(writer, sheet_name='Dates')

    def extract_all(self):

        self.__read_transaction()
        self.__export_report_file()
        self.__read_all_df()
        self.__create_date_dim()



