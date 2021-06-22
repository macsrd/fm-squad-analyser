import pandas as pd
from tkinter import filedialog
import time

#filename =  filedialog.askopenfilename(initialdir = r"D:\Dokumenty\Sports Interactive\Football Manager 2021",title = "Select file",filetypes = (("xlsx files","*.xlsx"),("all files","*.*")))
players = pd.read_excel('Export 2.xlsx', index_col=None, header=0)
atr = pd.read_excel('attributes.xlsx')

nan_value = float("NaN")
players.replace("", nan_value, inplace=True)
players.dropna(inplace=True)

#drop_columns = players[(players['Nazwisko'] == 'nan') | (players['Nazwisko'] == 'http://www.sigames.com/')].index

players = players.drop(players[(players['Nazwisko'] == 'NaN') | (players['Nazwisko'] == 'http://www.sigames.com/')].index)
print(players)

#export column names rename to match skills from list

players.rename(columns={'Agr': 'Agresja', 'Bły': 'Błyskotliwość', 'Chw': 'Chwytanie', 'Dec': 'Decyzje', 'Det': 'Determinacja', 
                        'D Wrz': 'Długie wyrzuty', 'Dśr': 'Dośrodkowania', 'Drb': 'Drybling', 'Eks': 'Ekscentryczność', 
                        'GbP': 'Gra bez piłki', 'Głw': 'Gra głową', 'GnP': 'Gra na przedpolu', '1v1': 'Jeden na jednego', 
                        'Kom': 'Komunikacja', 'Knc': 'Koncentracja', 'Kry': 'Krycie', 'Odb': 'Odbiór piłki', 'Opa': 'Opanowanie', 
                        'Pst': 'Piąstkowanie (Tendencja)', 'Pod': 'Podania', 'Pra': 'Pracowitość', 'Przg': 'Przegląd sytuacji', 
                        'Pwd': 'Przewidywanie', 'PPł': 'Przyjęcie piłki', 'Psp': 'Przyspieszenie', 'Prz': 'Przywództwo', 'Ref': 'Refleks', 
                        'Rów': 'Równowaga', 'RzK': 'Rzuty karne', 'RzR': 'Rzuty rożne', 'RzW': 'Rzuty wolne', 'Sił': 'Siła', 
                        'Sko': 'Skoczność', 'NSp': 'Sprawność', 'Str': 'Strzały z dystansu', 'Szb': 'Szybkość', 'Tec': 'Technika', 
                        'Ust': 'Ustawianie się', 'Wal': 'Waleczność', 'Wsp': 'Współpraca', 'WPK': 'Wychodzi poza pole karne (Tendencja)', 
                        'Wyk': 'Wykańczanie akcji', 'Wkp': 'Wykopy', 'Rzu': 'Wyrzuty', 'Wyt': 'Wytrzymałość', 'ZWy': 'Zasięg wyskoku', 'Zwi': 'Zwinność'}, inplace=True)

## Skills calculation functions 
def gk_analysis():
    """Calculation of GK skills basing on attributes lists saved in attributes.xlsx file"""
    GK_1 = (atr['GK_1'].to_list())
    GK_1 = [str(x) for x in GK_1]
    GK_1 = [x for x in GK_1 if x != 'nan']

    GK_2 = (atr['GK_2'].to_list())
    GK_2 = [str(x) for x in GK_2]
    GK_2 = [x for x in GK_2 if x != 'nan']

    GK_3 = (atr['GK_3'].to_list())
    GK_3 = [str(x) for x in GK_3]
    GK_3 = [x for x in GK_3 if x != 'nan']

    players['GK_1'] = (players[GK_1].sum(axis=1))/(len(GK_1)*20)*100
    players['GK_2'] = (players[GK_2].sum(axis=1))/(len(GK_2)*20)*100
    players['GK_3'] = (players[GK_3].sum(axis=1))/(len(GK_3)*20)*100
    players['GK'] = ((players['GK_1'] + players['GK_2']*0.8 + players['GK_3']*0.6)/3).round(decimals=2)

# Calculating Ball Playing Defender skills
def bpd_analysis():
    """Calculation of BPD skills basing on attributes lists saved in attributes.xlsx file"""
    BPD_1 = (atr['BPD_1'].to_list())
    BPD_1 = [str(x) for x in BPD_1]
    BPD_1 = [x for x in BPD_1 if x != 'nan']

    BPD_2 = (atr['BPD_2'].to_list())
    BPD_2 = [str(x) for x in BPD_2]
    BPD_2 = [x for x in BPD_2 if x != 'nan']

    BPD_3 = (atr['BPD_3'].to_list())
    BPD_3 = [str(x) for x in BPD_3]
    BPD_3 = [x for x in BPD_3 if x != 'nan']

    players['BPD_1'] = (players[BPD_1].sum(axis=1))/(len(BPD_1)*20)*100
    players['BPD_2'] = (players[BPD_2].sum(axis=1))/(len(BPD_2)*20)*100
    players['BPD_3'] = (players[BPD_3].sum(axis=1))/(len(BPD_3)*20)*100
    players['BPD'] = ((players['BPD_1'] + players['BPD_2']*0.8 + players['BPD_3']*0.6)/3).round(decimals=2)

def iwb_analysis():
    """Calculation of IWB skills basing on attributes lists saved in attributes.xlsx file"""
    IWB_1 = (atr['IWB_1'].to_list())
    IWB_1 = [str(x) for x in IWB_1]
    IWB_1 = [x for x in IWB_1 if x != 'nan']

    IWB_2 = (atr['IWB_2'].to_list())
    IWB_2 = [str(x) for x in IWB_2]
    IWB_2 = [x for x in IWB_2 if x != 'nan']

    IWB_3 = (atr['IWB_3'].to_list())
    IWB_3 = [str(x) for x in IWB_3]
    IWB_3 = [x for x in IWB_3 if x != 'nan']

    players['IWB_1'] = (players[IWB_1].sum(axis=1))/(len(IWB_1)*20)*100
    players['IWB_2'] = (players[IWB_2].sum(axis=1))/(len(IWB_2)*20)*100
    players['IWB_3'] = (players[IWB_3].sum(axis=1))/(len(IWB_3)*20)*100
    players['IWB'] = ((players['IWB_1'] + players['IWB_2']*0.8 + players['IWB_3']*0.6)/3).round(decimals=2)

# Calculating Defensive Midfielder skills
def dm_analysis():
    """Calculation of DM skills basing on attributes lists saved in attributes.xlsx file"""
    DM_1 = (atr['DM_1'].to_list())
    DM_1 = [str(x) for x in DM_1]
    DM_1 = [x for x in DM_1 if x != 'nan']

    DM_2 = (atr['DM_2'].to_list())
    DM_2 = [str(x) for x in DM_2]
    DM_2 = [x for x in DM_2 if x != 'nan']

    DM_3 = (atr['DM_3'].to_list())
    DM_3 = [str(x) for x in DM_3]
    DM_3 = [x for x in DM_3 if x != 'nan']

    players['DM_1'] = (players[DM_1].sum(axis=1))/(len(DM_1)*20)*100
    players['DM_2'] = (players[DM_2].sum(axis=1))/(len(DM_2)*20)*100
    players['DM_3'] = (players[DM_3].sum(axis=1))/(len(DM_3)*20)*100
    players['DM'] = ((players['DM_1'] + players['DM_2']*0.8 + players['DM_3']*0.6)/3).round(decimals=2)

def w_analysis():
    """Calculation of Wingers skills basing on attributes lists saved in attributes.xlsx file"""
    W_1 = (atr['W_1'].to_list())
    W_1 = [str(x) for x in W_1]
    W_1 = [x for x in W_1 if x != 'nan']

    W_2 = (atr['W_2'].to_list())
    W_2 = [str(x) for x in W_2]
    W_2 = [x for x in W_2 if x != 'nan']

    W_3 = (atr['W_3'].to_list())
    W_3 = [str(x) for x in W_3]
    W_3 = [x for x in W_3 if x != 'nan']

    players['W_1'] = (players[W_1].sum(axis=1))/(len(W_1)*20)*100
    players['W_2'] = (players[W_2].sum(axis=1))/(len(W_2)*20)*100
    players['W_3'] = (players[W_3].sum(axis=1))/(len(W_3)*20)*100
    players['W'] = ((players['W_1'] + players['W_2']*0.8 + players['W_3']*0.6)/3).round(decimals=2)

# Calculating Central Midfielder skills
def cm_analysis():
    """Calculation of CM skills basing on attributes lists saved in attributes.xlsx file"""
    CM_1 = (atr['CM_1'].to_list())
    CM_1 = [str(x) for x in CM_1]
    CM_1 = [x for x in CM_1 if x != 'nan']

    CM_2 = (atr['CM_2'].to_list())
    CM_2 = [str(x) for x in CM_2]
    CM_2 = [x for x in CM_2 if x != 'nan']

    CM_3 = (atr['CM_3'].to_list())
    CM_3 = [str(x) for x in CM_3]
    CM_3 = [x for x in CM_3 if x != 'nan']

    players['CM_1'] = (players[CM_1].sum(axis=1))/(len(CM_1)*20)*100
    players['CM_2'] = (players[CM_2].sum(axis=1))/(len(CM_2)*20)*100
    players['CM_3'] = (players[CM_3].sum(axis=1))/(len(CM_3)*20)*100
    players['CM'] = ((players['CM_1'] + players['CM_2']*0.8 + players['CM_3']*0.6)/3).round(decimals=2)

def ss_analysis():
    """Calculation of SS skills basing on attributes lists saved in attributes.xlsx file"""
    SS_1 = (atr['SS_1'].to_list())
    SS_1 = [str(x) for x in SS_1]
    SS_1 = [x for x in SS_1 if x != 'nan']

    SS_2 = (atr['SS_2'].to_list())
    SS_2 = [str(x) for x in SS_2]
    SS_2 = [x for x in SS_2 if x != 'nan']

    SS_3 = (atr['SS_3'].to_list())
    SS_3 = [str(x) for x in SS_3]
    SS_3 = [x for x in SS_3 if x != 'nan']

    players['SS_1'] = (players[SS_1].sum(axis=1))/(len(SS_1)*20)*100
    players['SS_2'] = (players[SS_2].sum(axis=1))/(len(SS_2)*20)*100
    players['SS_3'] = (players[SS_3].sum(axis=1))/(len(SS_3)*20)*100
    players['SS'] = ((players['SS_1'] + players['SS_2']*0.8 + players['SS_3']*0.6)/3).round(decimals=2)

#Executing functions
gk_analysis()
bpd_analysis()
iwb_analysis()
dm_analysis()
w_analysis()
cm_analysis()
ss_analysis()
    

## Tutaj objekt/klasę można zrobić
selected_columns = players[['Nazwisko', 'GK', 'BPD', 'IWB', 'DM', 'W', 'CM', 'SS', 'Preferowana noga']]
GK_columns = players[['Nazwisko', 'GK']]
BPD_columns = players[['Nazwisko', 'BPD']]
IWB_columns = players[['Nazwisko', 'IWB', 'Preferowana noga']]
DM_columns = players[['Nazwisko', 'DM']]
W_columns = players[['Nazwisko', 'W', 'Preferowana noga']]
CM_columns = players[['Nazwisko', 'CM']]
SS_columns = players[['Nazwisko', 'SS']]

#creating new dataframes for each player position
analysis = selected_columns.copy()

GK = GK_columns.copy()
GK.sort_values(by='GK', inplace=True, ascending=False)

BPD = BPD_columns.copy()
BPD.sort_values(by='BPD', inplace=True, ascending=False)

IWB = IWB_columns.copy()
IWB.sort_values(by='IWB', inplace=True, ascending=False)

DM = DM_columns.copy()
DM.sort_values(by='DM', inplace=True, ascending=False)

W = W_columns.copy()
W.sort_values(by='W', inplace=True, ascending=False)

CM = CM_columns.copy()
CM.sort_values(by='CM', inplace=True, ascending=False)

SS = SS_columns.copy()
SS.sort_values(by='SS', inplace=True, ascending=False)

#saving dataframes into new excel file
ExcelWriter = pd.ExcelWriter
with ExcelWriter('Squad Analysis.xlsx') as writer:
    analysis.to_excel(writer, sheet_name='Analysis', index=False)
    GK.to_excel(writer, sheet_name='Goalkeepers', index=False)
    BPD.to_excel(writer, sheet_name='Ball Playing Defenders', index=False)
    IWB.to_excel(writer, sheet_name='Inverted Wing Backs', index=False)
    DM.to_excel(writer, sheet_name='Defensive Midfielders', index=False)
    W.to_excel(writer, sheet_name='Wingers', index=False)
    CM.to_excel(writer, sheet_name='Central Midfilders', index=False)
    SS.to_excel(writer, sheet_name='Shadow Strikers', index=False)