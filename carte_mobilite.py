# V2 sur un fichier de villes/localisation d'une autre source
#%% Imports
# import folium
# from folium import plugins
# from folium.plugins.marker_cluster import MarkerCluster
import json
# import requests
import numpy as np
import pandas as pd
from OSMPythonTools.nominatim import Nominatim
from time import sleep
from glob import glob
from os import path

# Création de l'objet de recherche des adresses
nominatim = Nominatim()

#%% Lecture des fichiers

arExcel = glob("*.xlsx")
# arExcel = [ 
#     {'fic': "./2021 - KA120-SCH Accréditation Erasmus dans l'enseignement scolaire.xlsx",
#             'feuil':'Worksheet'},
#     {'fic': "2022 - KA122-SCH Projets de mobilité de courte durée pour les élèves et le personnel de l'enseignement scolaire.xlsx",
#             'feuil': "Worksheet"},
#     {'fic': "2022 KA121-SCH Projets de mobilité accrédités pour les élèves et le personnel de l'enseignement scolaire.xlsx",
#             'feuil': "Worksheet"}
#         ]

#%% Modificiations des fichiers, ajout des infos code + GPS
cpt_total = 0
cpt_ko = 0
cpt_adr_ko = 0
for xl in arExcel:
    fic = xl
    feuil = 'Worksheet' # xl['feuil']

    # On passe les fichiers déjà traité et "_modifs" 
    if (fic.endswith("_modif.xlsx") or path.exists(fic.replace('.xlsx', '_modif.xlsx'))):
        print(f"{fic} déjà traité")
        continue

    df = pd.read_excel(fic, sheet_name=feuil, header=0)

    cpt_total = cpt_total + df.shape[0]

    # Requêtes des adresses et positions GPS
    df_out = df.copy()
    df_out['CodePostal'] = ""
    df_out['lat'] = ""
    df_out['lon'] = ""
    df_out['note'] = ""

    print("Fichier "+fic)
    print("Analyse des "+str(df.shape[0])+" lignes")

    for index, row in df.iterrows():
        row = np.asarray(row)
        print(row[-2]+', '+row[-1].replace(',', ''))
        try:
            # Test requete 1 : ville + adresse
            ville = row[-2].replace(" CEDEX","").replace(" Cédex","").replace(" cedex","").replace(" Cedex","").lower()
            osm_res = nominatim.query(ville + ', '+row[-1].replace(',', '').lower(), addressdetails = True)

            if osm_res.displayName() is None:
                # Réessaye juste la ville, sans adresse
                sleep(2)
                osm_res_2 = nominatim.query(ville, addressdetails = True)
                if osm_res_2.displayName() is None:
                    # Non trouvé
                    df_out.at[index, 'CodePostal'] = "Non trouvé"
                    cpt_ko = cpt_ko + 1
                else:
                    # Ville seule
                    d_res = osm_res_2.toJSON()
                    ar_add = osm_res_2.displayName().split(',')
                    df_out.at[index, 'CodePostal'] = ar_add[-2]
                    df_out.at[index, 'lat'] = d_res[0]['lat']
                    df_out.at[index, 'lon'] = d_res[0]['lon'] 
                    df_out['note'] = "Ville sans adresse"
                    cpt_adr_ko = cpt_adr_ko + 1

            else:
                # Ville + adresse Ok
                d_res = osm_res.toJSON()
                ar_add = osm_res.displayName().split(',')
                df_out.at[index, 'CodePostal'] = ar_add[-2]
                df_out.at[index, 'lat'] = d_res[0]['lat']
                df_out.at[index, 'lon'] = d_res[0]['lon']  

        except:
            df_out.at[index, 'CodePostal'] = "Pb connexion"
        
        print('CP:'+df_out.at[index, 'CodePostal']+', lat lon: '+
                df_out.at[index, 'lat']+', '+df_out.at[index, 'lon'])
        sleep(2)

    df_out.to_excel(fic.replace('.xlsx', '_modif.xlsx'), sheet_name=feuil)

# Info 
print(f"{cpt_total} lignes lues : villes trouvées sans les adresses {cpt_adr_ko}, villes+adresses non trouvées {cpt_ko}")


#%% ['Civilite', 'Nom', 'Prenom', 'Mandat', 'Circonscription', 'Departement', 'Candidat', 'DatePublication']
# ['Code projet', 'Organisme candidat', 'Résultat de la sélection', 'Montant accordé', 'Ville', 'Adresse', 'CodePostal', 'lat', 'lon']

# Recherche des labels de couleurs 
arExcel = glob("*_modif.xlsx")

l_codes_proj = []

def format_code_projet(str_code):
    v_codes = str_code.split('-')
    code_new = v_codes[0] + " " + v_codes[3] + " "+ v_codes[4]
    return code_new

for xl in arExcel:
    fic = xl # .replace('.xlsx', '_modif.xlsx')
    feuil = "Worksheet" # xl['feuil']
    
    df = pd.read_excel(fic, sheet_name=feuil, header=0)
    n_lig_orig = len(df)
    
    # Supprime les lignes sans CodePostal
    df = df.dropna(subset=['CodePostal'])
    df = df[df.CodePostal!='']

    for code in df['Code projet'].to_list():
        code_trim = format_code_projet(code)
        l_codes_proj.append(code_trim)

    print(f"{fic} : {len(df)} lignes conservées / {n_lig_orig}")

# Couleurs des projets par type
from collections import Counter
print(Counter(l_codes_proj))

l_codes_proj = list(set(l_codes_proj))

v_couleurs = ["darkblue","red","green","lightblue","lightred",
"purple","orange", "magenta", "blueviolet"]

d_couleurs = dict(zip(l_codes_proj, v_couleurs))

#%% Appariemment des données bis
arExcel = glob("*_modif.xlsx")

data_sites = []
n_ko = 0

for xl in arExcel:
    fic = xl # .replace('.xlsx', '_modif.xlsx')
    feuil = "Worksheet" # xl['feuil']
    
    df = pd.read_excel(fic, sheet_name=feuil, header=0)
    
    # Supprime les lignes sans CodePostal
    df = df.dropna(subset=['CodePostal'])
    df = df[df.CodePostal!='']
    print(f"{fic} : {len(df)} lignes")

    for idx, row in df.iterrows():
        code_proj = format_code_projet(row["Code projet"])

        # Test des différents cas
        if row['CodePostal'] == "Non trouvé":
            # Français de l'étranger : au milieu de l'atlantique
            print(f"!! coords KO (->Paris) : {row['Organisme candidat']}, {row['Ville']}, {row['Adresse']}")
            coords = [48.856729,2.291466]
            str_cp = 'CP non trouvé'
            n_ko = n_ko + 1
        else:
            coords = [float(row['lat']), float(row['lon'])]
            str_cp = row['CodePostal']


        info = f"{row['Organisme candidat']}<br>{row['Code projet']}<br>{row['Ville']}({str_cp})<br/>{row['Adresse']}"
        
        # Vérif coords : si NaN
        if  any(np.isnan(coords)):
            print(f"!! coords KO (->Paris) : {', '.join(row)}")
            coords = [48.856729,2.291466]
        


        item = {"code": code_proj,
        "organisme": row['Organisme candidat'],
        "coordinates" : coords,
        "couleur": d_couleurs[code_proj] if code_proj in d_couleurs.keys() else "lightgray",
        "infos": f"{info}"
        }
        data_sites.append(item)

    

print(f" # KO : non trouvés {n_ko}")


#%% Sauvegarde en JSON 
with open('carto_programmes.json', 'w') as fp:
    json.dump(data_sites, fp)

# -> Fichier utilisé dynamiquement par index.html

#%% Empty cell to run all above