import pandas as pd

df_buyers = pd.read_excel("./data/sharepoint/sap_buyers.xls", engine="openpyxl")

print(df_buyers.head())