import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 向下填充，依據最前面有的數值
def fill_down_first(df):
    return df.ffill()

# 向下填充，依據最後有的數值
def fill_down_last(df):
    return df.bfill()

# 線性插值填充
def linear_interpolation(df):
    return df.interpolate(method='linear')

# 二次逼近插值填充
def quadratic_interpolation(df):
    return df.apply(lambda col: col.interpolate(method='polynomial', order=2))

# 指數插值填充
def exponential_interpolation(df):
    return df.apply(lambda col: np.exp(np.log(col).interpolate(method='linear')))

# 整數處理
def round_values(df, method):
    if method == 'floor':
        return df.applymap(np.floor)
    elif method == 'ceil':
        return df.applymap(np.ceil)
    elif method == 'round':
        return df.round()
    else:
        raise ValueError("Unknown rounding method")

def plot_data(df, title):
    df.plot(kind='line', marker='.', x="--")
    plt.title(title)
    plt.xlabel('Index')
    plt.ylabel('Values')
    plt.legend(loc='best')
    plt.grid(True)
    plt.show()

def main():
    # 讀取Excel表格
    df = pd.read_excel('data.xlsx')

    # 選擇填充方式
    fill_method = 'exponential'  # 可選 'first', 'last', 'linear', 'exponential'
    rounding_method = 'round'  # 可選 'floor', 'ceil', 'round', 'none'

    if fill_method == 'first':
        df_filled = fill_down_first(df)
    elif fill_method == 'last':
        df_filled = fill_down_last(df)
    elif fill_method == 'linear':
        df_filled = linear_interpolation(df)
    elif fill_method == 'quadratic':
        df_filled = quadratic_interpolation(df)
    elif fill_method == 'exponential':
        df_filled = exponential_interpolation(df)
    else:
        raise ValueError("Unknown fill method")

    # 對數值進行整數處理
    if rounding_method != 'none':
        df_filled = round_values(df_filled, rounding_method)

    # 將結果寫回Excel表格
    df_filled.to_excel('output.xlsx', index=False)

    # 繪製圖表
    plot_data(df_filled, 'Filled and Rounded Data')

if __name__ == '__main__':
    main()
