import pandas as pd
import numpy as np
import sys
from datetime import datetime as time
from tqdm import tqdm
import random
from sklearn import linear_model


class Representer:
 
    instance = None

    def __init__(self):
        self.xs = []
        self.ys = []
        self.model = linear_model.LinearRegression()

    def register(self, a, b, c, d, seconds):
        self.xs.append([a, b, c, d])
        self.ys.append(seconds.total_seconds())
        print(a, b, c, d, " : ", seconds)

    def get_factors(self):
        self.model.fit(self.xs, self.ys)
        a, b, c, d = self.model.coef_
        print(f"Load time = {a:.5f} * n_sheets + {b:.5f} * n_cols + {c:.5f} * n_rows + {d:.5f} * data_percentage")

    @classmethod
    def get_instance(cls):
        if cls.instance is None:
            cls.instance = Representer()
        return cls.instance

def timer(name, callable, *args):
    global repr
    a = time.now()
    callable(*args)
    b = time.now()
    Representer.get_instance().register(*name, b - a)
    

def create_workbook(name:str, n_sheets:int, n_cols:int, n_rows:int, data_percentage:float):
    columns = list(map(str, range(n_cols)))
    df = pd.DataFrame(np.nan, columns=columns, index=range(n_rows))
    for i in range(n_rows):
        for j in range(n_cols):
            if random.uniform(0,1) < data_percentage:
                df.iloc[i,j] = 1
    df.to_excel(f"{name}.xlsx")


def create_file_sequence(a_callable, no_tqdm=True):
    sheet_variants = [1, 5, 10, 50, 100]           # 6
    cols_count_variants = [10, 50, 100, 500, 1000]  # 6
    rows_count_variants = [10, 50, 100, 500, 1000]  # 6
    data_percentage_variants = [0.1, 0.5, 0.9]          # 5

    if no_tqdm:
        def pipe(iterable, position):
            return iterable

        tqdm = pipe

    for n_sheets in tqdm(sheet_variants, position=0):
        for n_cols in tqdm(cols_count_variants, position=1):
            for n_rows in tqdm(rows_count_variants, position=2):
                for data_percentage in tqdm(data_percentage_variants,position=3):
                    a_callable(n_sheets, n_cols, n_rows, data_percentage)

def create_test_data():
    create_file_sequence(                   
            lambda a, b, c, d : create_workbook(f"{a}-{b}-{c}-{d}", 
                                    n_sheets=a,
                                    n_cols=b,
                                    n_rows=c,
                                    data_percentage=d))


def single_read(a, b, c, d):
    timer((a, b, c, d),
            (lambda aa, bb, cc, dd: lambda: pd.read_excel(f"{aa}-{bb}-{cc}-{dd}.xlsx"))(a, b, c, d))

def test_all_cases():
    global repr
    create_file_sequence(lambda a, b, c, d: single_read(a, b, c, d))
    Representer.get_instance().get_factors()

def entrypoint(args):
    if len(args)>2:
        if args[2] == "-c":
            create_test_data()
    test_all_cases()
    return 0

if __name__=="__main__":
    sys.exit(entrypoint(sys.argv))
