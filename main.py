import pandas as pd
import numpy as np
import sys
from datetime import datetime as time
from tqdm import tqdm
import random
from sklearn import linear_model


TARGET_DIRECTORY = "docs"
EXTENSION = "xlsx"

class Representer:
 
    instance = None

    def __init__(self):
        self.xs = []
        self.ys = []
        self.model = linear_model.LinearRegression()

    def register(self, n_sheets, n_cols, n_rows, data_percentage, seconds):
        self.xs.append([n_sheets, n_cols, n_rows, data_percentage])
        self.ys.append(seconds.total_seconds())
        print(n_sheets, n_cols, n_rows, data_percentage, " : ", seconds)

    def get_factors(self):
        self.model.fit(self.xs, self.ys)
        a, b, c, d = self.model.coef_
        print(f"Load time = {a:.5f} * n_sheets + {b:.5f} * n_cols + {c:.5f} * n_rows + {d:.5f} * data_percentage")

    @classmethod
    def get_instance(cls):
        if cls.instance is None:
            cls.instance = Representer()
        return cls.instance


def get_name(n_sheets : int, n_cols : int, n_rows : int, data_percentage : float) -> str:
    return f"{TARGET_DIRECTORY}/{n_sheets}-{n_cols}-{n_rows}-{data_percentage}.{EXTENSION}"

def timer(representer):
    def _timer(func):
        def _wrapper(*args, **kwargs):
            a = time.now()
            func(*args, **kwargs)
            b = time.now()
            representer.register(*args, b - a)
        return _wrapper
    return _timer
    

class Workbook:

    def __init__(self, name):
        self.writer = pd.ExcelWriter(name, engine='xlsxwriter')
        self.sheet_count = 0

    def append(self, dataframe: pd.DataFrame):
        self.sheet_count += 1
        dataframe.to_excel(self.writer, sheet_name=f"{self.sheet_count}")


    def close(self):
        self.writer.save()

def create_workbook(name:str, n_sheets:int, n_cols:int, n_rows:int, data_percentage:float):
    columns = list(map(str, range(n_cols)))
    writter = Workbook(name)
        
    for s in range(n_sheets):
        values = np.random.choice([1, np.nan], size=(n_rows, n_cols), p=[data_percentage, 1 - data_percentage])
        df = pd.DataFrame(values, columns=columns, index=range(n_rows))
        writter.append(df)
    writter.close()


def create_file_sequence(a_callable, no_tqdm=True):
    sheet_variants = [1, 5, 10]                         # 3
    cols_count_variants = [10, 50, 100, 400, 700]       # 5
    rows_count_variants = [10, 50, 100, 400, 700]       # 5
    data_percentage_variants = [0.1, 0.5, 0.9]          # 3

    if no_tqdm:
        def pipe(iterable, position):
            return iterable

        tqdmw = pipe
    else:
        tqdmw = tqdm

    for n_sheets in tqdmw(sheet_variants, position=0):
        for n_cols in tqdmw(cols_count_variants, position=1):
            for n_rows in tqdmw(rows_count_variants, position=2):
                for data_percentage in tqdmw(data_percentage_variants,position=3):
                    a_callable(n_sheets, n_cols, n_rows, data_percentage)

def create_test_data():
    create_file_sequence(                   
            lambda a, b, c, d : create_workbook(get_name(a, b, c, d), 
                                    n_sheets=a,
                                    n_cols=b,
                                    n_rows=c,
                                    data_percentage=d), no_tqdm=False)


@timer(Representer.get_instance())
def single_read(a, b, c, d):
    pd.read_excel(get_name(a, b, c, d), sheet_name=None)

def test_all_cases():
    create_file_sequence(lambda a, b, c, d: single_read(a, b, c, d))
    Representer.get_instance().get_factors()

def entrypoint(args):
    if len(args)==2:
        if args[1] == "-c":
            create_test_data()
    test_all_cases()
    return 0

if __name__=="__main__":
    sys.exit(entrypoint(sys.argv))
