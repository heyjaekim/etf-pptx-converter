# -*- coding: utf-8 -*-
"""
Created on Mon Jun  7 20:52:24 2021

@author: 가족
"""

import numpy as np
import plotly.express as px
import plotly.io as pio
import pandas as pd

pio.renderers.default='jpeg'


df1 = pd.read_excel(r'C:\ETF\ETF Snapshot_Treemap_V3.xlsx', sheet_name=1)
df2 = pd.read_excel(r'C:\ETF\ETF Snapshot_Treemap_V3.xlsx', sheet_name=3)

fig1 = px.treemap(df1, path=['Main', 'Sector_adj', 'Tickers'], values='Mkt Cap_adj', 
                 color='MTD 수익률(%)', color_continuous_scale='IceFire',
                 color_continuous_midpoint= np.average(df2['MTD 수익률(%)']))


fig2 = px.treemap(df2, path=['Main', 'Sector_adj2', 'Sector_adj1', 'Tickers'], values='Mkt Cap_adj', 
                 color='MTD 수익률(%)', color_continuous_scale='IceFire',
                 color_continuous_midpoint= np.average(df2['MTD 수익률(%)']))

fig1.update_layout(width=2000, height=2000)
fig2.update_layout(width=2000, height=2000)


fig1.show()
fig2.show()

