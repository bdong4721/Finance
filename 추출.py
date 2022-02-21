df = pd.read_fwf('NDX_1min.txt', sep=" ", header =None, names = ['Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume']
print(df[687733:982719])
