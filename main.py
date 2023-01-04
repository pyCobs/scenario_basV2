import openpyxl as xl
import time
import json

from load import load_pft, dico_scenario
from write import use_scenario

# chargement des 3 feuilles utiles pour ce script
start = time.time()
wb, sheet_pft, sheet_scenario = load_pft()
end = time.time()

print(f"Temps de chargement : {end-start} sec")

# chargement du tableau avec les années de décalage
start = time.time()
scenario, dico_nd = dico_scenario(sheet_scenario)

# création du scénario
use_scenario(sheet_pft, scenario, dico_nd)
end = time.time()

print(f"Temps de traitement : {end-start} sec")

# sauvegarde
start = time.time()
wb.save(filename="PFM1 2023 bas.xlsm")
end = time.time()
print(f"Temps de sauvegarde : {end-start} sec")

wb.close()
