import pandas as pd

def main():
    # 讀取 CSV 檔案
    df = pd.read_csv('上市公司資料.csv')
    
    # 移除包含缺失值的行
    df1 = df.dropna()
    
    # 重新索引 DataFrame，僅保留特定的列
    df2 = df1.reindex(columns=['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收'])
    
    # 重命名列名稱
    df3 = df2.rename(columns={
        '營業收入-上月營收':'上月營收',
        '營業收入-當月營收':'當月營收'
        })
    
    # 將整理後的 DataFrame 存成 CSV 檔案
    df3.to_csv('上市公司資料整理.csv', encoding='utf-8')
    
    # 將整理後的 DataFrame 存成 Excel 檔案
    df3.to_excel('上市公司資料整理.xlsx')
    
    # 打印完成訊息
    print("存檔完成")

if __name__ == '__main__':
    main()