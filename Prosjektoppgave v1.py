# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 15:05:45 2025

@author: Karl Peter Fjørtoft
"""


# Del a) Skriv et program som leser inn filen ‘support_uke_24.xlsx’ og lagrer data fra 
# kolonne 1 i en array med variablenavn ‘u_dag’, dataen i 
# kolonne 2 lagres i arrayen ‘kl_slett’, data i 
# kolonne 3 lagres i arrayen ‘varighet’ og dataen i 
# kolonne 4 lagres i arrayen ‘score’. 
# Merk: filen ‘support_uke_24.xlsx’ må ligge i samme mappe som Python-programmet ditt.


import pandas as pd

support = pd.read_excel("support_uke_24.xlsx")

u_dag = support['Ukedag'].values

# print (u_dag)

# data_u_dag = pd.DataFrame(u_dag)
# data_u_dag.to_excel("u_dag.xlsx")

kl_slett = support['Klokkeslett'].values

# print (kl_slett)

varighet = support['Varighet'].values


score = support['Tilfredshet'].values

# print (score)
data_score = pd.DataFrame(score)
data_score.to_excel("score.xlsx")




# data = pd.read_excel("norges-befolkning.xlsx")

# year = data['Aarstall'].values
# antall = data['Innbyggere'].values