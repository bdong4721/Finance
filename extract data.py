import pandas as pd




df = pd.read_fwf('NDX_1min.txt', sep=" ", header =None, names = ['Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume']
df_final = df.loc[[687733:982719], :]
""""
with open('NDX_1min.txt') as f:
    lines = f.readline()

stock = pd.DataFrame()

for stock in lines:
    df.columns = ['Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume']
    
    if stock ==0:
        None
    else:
        stock_f = stock.append(lines, sort=Fale)
    print(stock_f)    
"""
