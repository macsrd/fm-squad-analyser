import pandas as pd
from tkinter import filedialog
import time

timestr = time.strftime("%Y%m%d-%H%M%S")

#opening export file from Squad View
###filename =  filedialog.askopenfilename(initialdir = r"D:\Dokumenty\Sports Interactive\Football Manager 2021",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
filename = 'Export_Pre_28_29.html'
players = pd.read_html(filename, index_col=None, header=0, encoding='utf-8')
players = pd.concat(players)


#opening file with positional attributes
atr = pd.read_excel('attributes.xlsx')

#Opening file with renamed columns
renamed = pd.read_excel('rename.xlsx')

#removing nan and sigames hyperlink from dataframe (cleaning input data)
players.dropna(inplace=True)
players = players.drop(players[(players['Nazwisko'] == 'NaN') | (players['Nazwisko'] == 'http://www.sigames.com/')].index)

#renaming attributes columns to long name
players.rename(columns=renamed.set_index('old_name')['new_name'], inplace=True)


#creating class for calculating skills
class Calculate():
    """Kalkulowanie umiejętności w zależności od pozycji"""

    def __init__(self, sk, sk_1, sk_2, sk_3):
        """Inicjalizacja atrybutów umiejętności (skill)"""
        self.sk = sk
        self.sk_1 = sk_1
        self.sk_2 = sk_2
        self.sk_3 = sk_3

    def calculate_skill(self):
        """ Kalkulowanie umiejętności w zależności od pozycji"""
        sk_1_col = self.sk_1
        sk_2_col = self.sk_2
        sk_3_col = self.sk_3
        
        self.sk_1 = (atr[self.sk_1].values.tolist())
        self.sk_1 = [x for x in self.sk_1 if pd.isnull(x) == False]
        
        self.sk_2 = (atr[self.sk_2].values.tolist())
        self.sk_2 = [x for x in self.sk_2 if pd.isnull(x) == False]
        
        self.sk_3 = (atr[self.sk_3].values.tolist())
        self.sk_3 = [x for x in self.sk_3 if pd.isnull(x) == False]

        players[sk_1_col] = (players[self.sk_1].sum(axis=1))/(len(self.sk_1)*20)*100
        players[sk_2_col] = (players[self.sk_2].sum(axis=1))/(len(self.sk_2)*20)*100
        players[sk_3_col] = (players[self.sk_3].sum(axis=1))/(len(self.sk_3)*20)*100
        players[self.sk] = ((players[sk_1_col] + players[sk_2_col]*0.8 + players[sk_3_col]*0.6)/3).round(decimals=2)

#defining position skill calculation functions                                               
def gk_analysis():
    GK = Calculate('GK', 'GK_1', 'GK_2', 'GK_3')
    GK.calculate_skill()

def bpd_analysis():
    BPD = Calculate('BPD', 'BPD_1', 'BPD_2', 'BPD_3')
    BPD.calculate_skill()

def iwb_analysis():
    IWB = Calculate('IWB', 'IWB_1', 'IWB_2', 'IWB_3')
    IWB.calculate_skill()

def dm_analysis():
    DM = Calculate('DM', 'DM_1', 'DM_2', 'DM_3')
    DM.calculate_skill()

def w_analysis():
    W = Calculate('W', 'W_1', 'W_2', 'W_3')
    W.calculate_skill()

def cm_analysis():
    CM = Calculate('CM', 'CM_1', 'CM_2', 'CM_3')
    CM.calculate_skill()

def ss_analysis():
    SS = Calculate('SS', 'SS_1', 'SS_2', 'SS_3')
    SS.calculate_skill()
    
def highlight_max(s):
    is_max = s == s.max()
    return ['background: green' if cell else '' for cell in is_max]

#Executing position analysis functions in one function
def analysis_all():
    gk_analysis()
    bpd_analysis()
    iwb_analysis()
    dm_analysis()
    w_analysis()
    cm_analysis()
    ss_analysis()

analysis_all()
    
selected_columns = players[['Nazwisko', 'GK', 'BPD', 'IWB', 'DM', 'W', 'CM', 'SS', 'Preferowana noga']]
GK_columns = players[['Nazwisko', 'GK']]
BPD_columns = players[['Nazwisko', 'BPD']]
IWB_columns = players[['Nazwisko', 'IWB', 'Preferowana noga']]
IWB_L_columns = IWB_columns.drop(IWB_columns[(IWB_columns['Preferowana noga'] == 'Tylko lewa') | (IWB_columns['Preferowana noga'] == 'Lewa')].index)
IWB_R_columns = IWB_columns.drop(IWB_columns[(IWB_columns['Preferowana noga'] == 'Tylko prawa') | (IWB_columns['Preferowana noga'] == 'Prawa')].index)
DM_columns = players[['Nazwisko', 'DM']]
W_columns = players[['Nazwisko', 'W', 'Preferowana noga']]
W_R_columns = W_columns.drop(W_columns[(W_columns['Preferowana noga'] == 'Tylko lewa') | (W_columns['Preferowana noga'] == 'Lewa')].index)
W_L_columns = W_columns.drop(W_columns[(W_columns['Preferowana noga'] == 'Tylko prawa') | (W_columns['Preferowana noga'] == 'Prawa')].index)
CM_columns = players[['Nazwisko', 'CM']]
SS_columns = players[['Nazwisko', 'SS']]

#creating new dataframes for each player position
analysis = selected_columns.copy()
analysis = analysis.style.highlight_max(color = 'green', axis = 0)

GK = GK_columns.copy()
GK.sort_values(by='GK', inplace=True, ascending=False)

BPD = BPD_columns.copy()
BPD.sort_values(by='BPD', inplace=True, ascending=False)

IWB = IWB_columns.copy()
IWB.sort_values(by='IWB', inplace=True, ascending=False)

IWB_R = IWB_R_columns.copy()
IWB_R.sort_values(by='IWB', inplace=True, ascending=False)

IWB_L = IWB_L_columns.copy()
IWB_L.sort_values(by='IWB', inplace=True, ascending=False)

DM = DM_columns.copy()
DM.sort_values(by='DM', inplace=True, ascending=False)

W = W_columns.copy()
W.sort_values(by='W', inplace=True, ascending=False)

W_R = W_R_columns.copy()
W_R.sort_values(by='W', inplace=True, ascending=False)

W_L = W_L_columns.copy()
W_L.sort_values(by='W', inplace=True, ascending=False)

CM = CM_columns.copy()
CM.sort_values(by='CM', inplace=True, ascending=False)

SS = SS_columns.copy()
SS.sort_values(by='SS', inplace=True, ascending=False)

#saving dataframes into new excel file
ExcelWriter = pd.ExcelWriter
with ExcelWriter('Squad Analysis_' + timestr + '.xlsx') as writer:
    analysis.to_excel(writer, sheet_name='Analysis', index=False)
    GK.to_excel(writer, sheet_name='Goalkeepers', index=False)
    BPD.to_excel(writer, sheet_name='Ball Playing Defenders', index=False)
    IWB.to_excel(writer, sheet_name='Inverted Wing Backs', index=False)
    IWB_R.to_excel(writer, sheet_name='Right IWBs', index=False)
    IWB_L.to_excel(writer, sheet_name='Left IWBs', index=False)
    DM.to_excel(writer, sheet_name='Defensive Midfielders', index=False)
    W.to_excel(writer, sheet_name='Wingers', index=False)
    W_R.to_excel(writer, sheet_name='Right Wingers', index=False)
    W_L.to_excel(writer, sheet_name='Left Wingers', index=False)
    CM.to_excel(writer, sheet_name='Central Midfilders', index=False)
    SS.to_excel(writer, sheet_name='Shadow Strikers', index=False)
