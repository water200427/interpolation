# 數據插值

此專案提供了一個Python腳本，用於使用各種插值方法填充Excel文件中的缺失值。該腳本支持不同的四捨五入方法，並且可以繪製填充後的數據。

## 要求

- Python 3.11

## 設置

1. **克隆倉庫**（如果適用）：
    ```sh
    git clone https://github.com/water200427/interpolation.git
    cd interpolation
    ```

2. **創建虛擬環境**：
    ```sh
    python -m venv venv
    ```

3. **啟動虛擬環境**：
    - 在Windows上：
        ```sh
        .\venv\Scripts\activate
        ```
    - 在macOS/Linux上：
        ```sh
        source venv/bin/activate
        ```

4. **安裝所需的依賴項**：
    ```sh
    pip install -r requirements.txt
    ```

## 使用方法

1. **將你的輸入Excel文件放置在專案目錄中，並命名為`data.xlsx`**。

2. **運行腳本**：
    ```sh
    python data_interpolation.py
    ```

3. **輸出將保存為`output.xlsx`，並顯示一個圖表**。

## 配置

你可以通過修改`data_interpolation.py`中的[fill_method](http://_vscodecontentref_/0)和[rounding_method](http://_vscodecontentref_/1)變量來配置插值和四捨五入方法。

- **插值方法**：
    - `'first'`：使用第一個可用值向下填充。
    - `'last'`：使用最後一個可用值向下填充。
    - `'linear'`：線性插值。
    - `'exponential'`：指數插值。
    - `'quadratic'`：二次插值。

- **四捨五入方法**：
    - `'floor'`：向下取整。
    - `'ceil'`：向上取整。
    - `'round'`：四捨五入。
    - `'none'`：保留原始插值數據。

## 範例

```python
# data_interpolation.py中的範例配置
fill_method = 'exponential'  # 選項: 'first', 'last', 'linear', 'exponential', 'quadratic'
rounding_method = 'round'  # 選項: 'floor', 'ceil', 'round', 'none'