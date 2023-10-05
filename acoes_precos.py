import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib as plt
from scipy import stats
from scipy.stats import norm, skewnorm
from urllib.request import urlopen, urlretrieve, Request
from urllib.error import URLError, HTTPError
import requests
import json
import time
import plotly.express as px
import plotly.graph_objects as go

#Nome desse app: "De olho nos preços para Dividendos"

# Importing the pywin32 module
# import win32com.client
  
# # Opening Excel software using the win32com
# File = win32com.client.Dispatch("Excel.Application")
  
# # Optional line to show the Excel software
# #File.Visible = 1
  
# # Opening your workbook
# Workbook = File.Workbooks.open(r"E:\Users\JONATHAN LOKO\Documents\Repositorio\DOPD\acoes_que_acompanhamos_margens_precos_teto.xlsx")
  
# # Refeshing all the shests
# Workbook.RefreshAll()
  
# # Saving the Workbook
# Workbook.Save()
  
# # Closing the Excel File
# File.Quit()

# Load the Excel file
# import openpyxl
# workbook = openpyxl.load_workbook("acoes_que_acompanhamos_margens_precos_teto.xlsx")

# # Refresh all data connections in the workbook
# for connection in workbook.connections:
#     connection.refresh()
# Save the updated workbook
# workbook.save("acoes_que_acompanhamos_margens_precos_teto.xlsx")

# import xlwings as xw

# # Open the Excel file
# workbook = xw.Book("acoes_que_acompanhamos_margens_precos_teto.xlsx")

# # Refresh all data connections in the workbook
# for connection in workbook.connections:
#     connection.refresh()

# # Save and close the updated workbook
# workbook.save("acoes_que_acompanhamos_margens_precos_teto.xlsx")
# workbook.close()


page_name = "De olho nos preços para Dividendos"
df = pd.read_excel('acoes_que_acompanhamos_margens_precos_teto.xlsx', skiprows=1)
df = df[['Nome_Empresa','media_valor_div_5anos','Barsi_preco_Teto','Preço Atual','margem_preco_Barsi %','meu preço','margem meu preco %']]
df = df.sort_values(by=['margem meu preco %'], ascending=False)

##############################################################################

#Icons link: https://icons.getbootstrap.com/

    # 1. as sidebar menu
with st.sidebar:
    selected = option_menu(menu_title=None,
                          options=["Home",], #required
                          icons=['house'], #optional
                          menu_icon="cast", default_index=0) #optional
    # 2. Creating the home page   
if selected == "Home":
    st.title(f"Bem vindo/a a {page_name}")
    st.subheader('Verifique melhores preços')
 
    st.dataframe(df)

##############################################################################