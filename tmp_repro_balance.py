import pandas as pd
from page_product_liquidity import sort_cashflow_rows_preserve_blank_positions

df = pd.DataFrame({
    '日期': [pd.NaT, pd.Timestamp('2024-05-01'), pd.NaT],
    '现金流类型': ['', '申购', ''],
    '关联产品': ['', '', ''],
    '现金流入': [0.0, 100.0, 50.0],
    '现金流出': [0.0, 0.0, 0.0],
    '期末余额': [0.0, 0.0, 0.0]
})
print('before')
print(df)
print('sorted preserve blank positions')
print(sort_cashflow_rows_preserve_blank_positions(df))

# simulate update of blank row amount
edited_df = df.copy()
edited_df.loc[0,'现金流入'] = 500.0
current_balance = 0.0
for row_i in range(len(edited_df)):
    inflow = float(edited_df.iloc[row_i]['现金流入']) if pd.notna(edited_df.iloc[row_i]['现金流入']) else 0
    outflow = float(edited_df.iloc[row_i]['现金流出']) if pd.notna(edited_df.iloc[row_i]['现金流出']) else 0
    current_balance = current_balance + inflow - outflow
    edited_df.iloc[row_i, edited_df.columns.get_loc('期末余额')] = current_balance
print('after')
print(edited_df)
