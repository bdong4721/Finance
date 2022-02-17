#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import datetime
import time
import matplotlib.pyplot as plt



with open('NDX_1min.txt') as f:
    lines = f.readline()

stock = pd.DataFrame()

for i in lines:
    df.columns = ['Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume']
    
    if lines ==0:
        None
    else:
        stock = stock.append(lines, sort=False)
        
        
    

