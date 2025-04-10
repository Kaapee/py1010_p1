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

# Laster inn pakken pandas som behandler data, og kan behandle excel-filer
import pandas as pd
import matplotlib.pyplot as plt
import re

# Laster inn excelfil og kaller den support
support = pd.read_excel("support_uke_24.xlsx")

# # Lagrer dataen i kolonnen Ukedag i excelarket til arrayen u_dag 
u_dag = support['Ukedag'].values

# # Lagrer dataen i kolonnen Klokkeslett i excelarket til arrayen kl_slett 
kl_slett = support['Klokkeslett'].values

# # Lagrer dataen i kolonnen Varighet i excelarket til arrayen varighet 
varighet = support['Varighet'].values

# # Lagrer dataen i kolonnen Tilfredshet i excelarket til arrayen score
score = support['Tilfredshet'].values

# En test for å sjekke at vi laster inn filen og skriver ut fornuftige verdier
# Gir kolonnene i excelarket navn data = support[["Ukedag", "Klokkeslett", "Varighet", "Tilfredshet"]]
# data = support[["Ukedag", "Klokkeslett", "Varighet", "Tilfredshet"]]
# data.to_excel("Sammensatt.xlsx")

##################################################################################################################
# Del b) Skriv et program som finner antall henvendelser for hver de 5 ukedagene. 
# Resultatet visualiseres ved bruk av et søylediagram (stolpediagram).

# Laster inn ukedager og teller opp antall henvendelser for hver ukedag
ukedager = support["Ukedag"].value_counts()

# En test for å se hvordan tabellen ser ut
# print(ukedager)

#Lager fast rekkefølge for ukedager for å presentere disse i riktig rekkefølge
ukedager_rekkefolge = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]

#Sorterer ukedagene etter riktig rekkefølge 
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.loc.html
ukedager_sortert = ukedager.loc[ukedager_rekkefolge]

# Test for å se at verdiene blir riktige
# print (ukedager_sortert)

# Lager et søylediagram med de sorterte antall henvendelser per ukedag
ukedager_sortert.plot(kind="bar", title="Henvendelser per ukedag")


##################################################################################################################

# Del c) Skriv et program som finner minste og lengste samtaletid som er loggført for uke 24. 
# Svaret skrives til skjerm med informativ tekst.

# Fant flere enkle funksjoner for pandas her: https://pandas.pydata.org/docs/user_guide/basics.html

samtale_korteste = varighet.min()
print("Den korteste samtalen denne uken varte i: ",samtale_korteste, "\n")

samtale_lengste = varighet.max()
print("Den lengste samtalen denne uken varte i: ",samtale_lengste, "\n")


##################################################################################################################

#      Del d) KREVENDE: Skriv et program som regner ut gjennomsnittlig samtaletid basert på alle henvendelser i uke 24.

# Her fungerer det ikkje å bruke varighet.mean() fordi samtaletiden er lagret som string med formatet hh:mm:ss
# Vi kan formatere om dette formatet ved å bruke to.timedelta funksjonen i pandas
# https://pandas.pydata.org/docs/reference/api/pandas.to_timedelta.html

# Eg starter ved å formattere om til numeriske samtaleverdier:
samtaletid_numerisk = pd.to_timedelta(support['Varighet'])

# Test for å sjekke at samtaletiden ser riktig ut
# print (samtaletid_numerisk)

# Eg finner snitt-samtaletid ved å bruke .mean funksjonen på den numeriske arrayen for samtalelengder:
samtaletid_snitt = samtaletid_numerisk.mean()
print("Snitt for samtaletid denne uken er: ", samtaletid_snitt, "\n")


##################################################################################################################

#     Del e) Supportvaktene i MORSE er delt inn i 2-timers bolker: kl 08-10, kl 10-12, kl 12-14 og kl 14-16. 
#     Skriv et program som finner det totale antall henvendelser supportavdelingen mottok for hver av tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24. 
#     Resultatet visualiseres ved bruk av et sektordiagram (kakediagram).


# https://docs.python.org/3/library/re.html#regular-expression-examples
# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.Series.str.match.html

# Her går en for løkke gjennom tekstverdien til kl_slett og øker verdien med 1 for hver treff på tekst som starter på:
# r : raw string, unngår potensielle problemer med søket, "0[89]:" her søker vi mellom tekst-verdiene 0 og enten 8 eller 9 etterfulgt av :
# for eksempel klokkeslett som starter på "08:" og "09:"
 


kl_08_10 = 0

for i in kl_slett:
    if re.match(r"0[89]:", i):
        kl_08_10 +=1
        
print("Antall telefoner mellom 08-10: ", kl_08_10)

kl_10_12 = 0

for i in kl_slett:
    if re.match(r"1[01]:", i):
        kl_10_12 +=1
        
print("Antall telefoner mellom 10-12: ", kl_10_12)


kl_12_14 = 0

for i in kl_slett:
    if re.match(r"1[23]:", i):
        kl_12_14 +=1
        
print("Antall telefoner mellom 12-14: ", kl_12_14)


kl_14_16 = 0

for i in kl_slett:
    if re.match(r"1[45]:", i):
        kl_14_16 +=1
        
print("Antall telefoner mellom 14-16: ", kl_14_16)

totalt_antall_samtaler = (kl_08_10 + kl_10_12 + kl_12_14 + kl_14_16)
print ("Totalt antall samtaler denne uken er :", totalt_antall_samtaler)

