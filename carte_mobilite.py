# V2 sur un fichier de villes/localisation d'une autre source
#%% Imports
import folium
from folium import plugins
from folium.plugins.marker_cluster import MarkerCluster
import json
import requests
import numpy as np
import pandas as pd
from random import random
import unicodedata
import math


#%% Fonctions utilitaires pour supprimer les accents
# Sur un dataframe
prep_string_df = lambda x: x.str.normalize('NFKD').str.replace("œ","oe").str.encode('ascii', errors='ignore').str.decode('utf-8').str.replace('-',' ').str.replace("'",' ').str.upper()
# Sur une chaine
prep_string = lambda x: unicodedata.normalize('NFKD',x).replace("œ","oe").encode('ascii', errors='ignore').decode('utf-8').replace('-',' ').replace("'",' ').upper()


#%% Lecture des données

#  Fichier eucircos
df_regions = pd.read_csv('eucircos_regions_departements_circonscriptions_communes_gps.csv', encoding='utf-8', sep=";", 
dtype={'nom_département': str, 'codes_postaux':str, 'latitude':float, 'longitude':float})
#['EU_circo', 'code_région', 'nom_région', 'chef-lieu_région','numéro_département', 'nom_département', 'préfecture', 
# 'numéro_circonscription', 'nom_commune', 'codes_postaux', 'code_insee', 'latitude', 'longitude', 'éloignement']
df_regions['COM'] = prep_string_df(df_regions['nom_commune'])
df_regions['DEP'] = prep_string_df(df_regions['nom_département'])
df_regions['REG'] = prep_string_df(df_regions['nom_région'])



#%% Fichier INSEE
df_insee = pd.read_csv('codesinseecommunesgeolocalisees.csv', encoding='utf-8', sep=",")
df_insee["longitude"] = df_insee["longitude_radian"] * 180 / math.pi
df_insee["latitude"] = df_insee["latitude_radian"] * 180 / math.pi
df_insee.drop(['latitude_radian', 'longitude_radian'], axis=1, inplace=True)
df_insee['COM'] = prep_string_df(df_insee['Nom'])
#['Insee', 'Nom', 'Altitude', 'code_postal', 'longitude_radian',
# 'latitude_radian', 'pop99', 'surface', 'longitude', 'latitude']

#%% Fichier la poste
df_poste = pd.read_csv('laposte_hexasmal.csv', encoding='utf-8', sep=";", 
dtype={'Code_postal':str, 'latitude':float, 'longitude':float})

df_poste['numéro_département'] = df_poste['Code_commune_INSEE'].apply(lambda x : x[0:3] if x.startswith('97') or x.startswith('98') else x[0:2])

# Code_commune_INSEE;Nom_commune;Code_postal;Ligne_5;Libellé_d_acheminement;latitude;longitude
df_poste.drop(['Ligne_5', 'Libellé_d_acheminement'], axis=1, inplace=True)
df_poste.drop_duplicates(inplace=True)

# On ajoute les départements dans df_poste
df_poste = pd.merge(df_poste,
             df_regions[['numéro_département','nom_département']].drop_duplicates(), 
             how='left', on='numéro_département')

# Colonne pour recherche
df_poste['COM'] = prep_string_df(df_poste['Nom_commune'])
df_poste['DEP'] = prep_string_df(df_poste['nom_département'])

#%% Chargement fichiers Excel à renseigner 
from openpyxl import load_workbook
from OSMPythonTools.nominatim import Nominatim
from time import sleep
import re
nominatim = Nominatim()

#%% Lecture des fichiers

                        
arExcel = [ 
    {'fic': "./2021 - KA120-SCH Accréditation Erasmus dans l'enseignement scolaire.xlsx",
            'feuil':'Worksheet'},
    {'fic': "2022 - KA122-SCH Projets de mobilité de courte durée pour les élèves et le personnel de l'enseignement scolaire.xlsx",
            'feuil': "Worksheet"},
    {'fic': "2022 KA121-SCH Projets de mobilité accrédités pour les élèves et le personnel de l'enseignement scolaire.xlsx",
            'feuil': "Worksheet"}
        ]

#%% Modificiations des fichiers, ajout des infos code + GPS
cpt_total = 0
cpt_ko = 0
cpt_adr_ko = 0
for xl in arExcel[2:]:
    fic = xl['fic']
    feuil = xl['feuil']
    
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

# Couleurs des projets par type
l_couleurs = {"Type 1" : "darkblue",
"Type 2" : "red",
"Type 3" : "green",
"Type 4" : "lightblue",
"Type 5" : "lightred",
"Type 6" : "purple",
"Type 7" : "orange"}


#%% Appariemment des données bis

data_sites = []
n_ko = 0

for xl in arExcel:
    fic = xl['fic'].replace('.xlsx', '_modif.xlsx')
    feuil = xl['feuil']
    print(f"{fic} : {len(df)} lignes")

    df = pd.read_excel(fic, sheet_name=feuil, header=0)

    for idx, row in df.iterrows():
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


        info = f"{row['Code projet']}<br>{row['Ville']}({str_cp})<br/>{row['Adresse']}"
        
        # Vérif coords : si NaN
        if  any(np.isnan(coords)):
            print(f"!! coords KO (->Paris) : {', '.join(row)}")
            coords = [48.856729,2.291466]
        


        item = {"status": "etab",
        "organisme": row['Organisme candidat'],
        "coordinates" : coords,
        "couleur": "lightgray" if row['CodePostal'] == "Non trouvé" else "lightred",
        "infos": f"{info}"
        }
        data_sites.append(item)

    

print(f" # KO : non trouvés {n_ko}")


#%% Sauvegarde en JSON 
with open('carto_programmes.json', 'w') as fp:
    json.dump(data_sites, fp)

# -> Fichier utilisé dynamiquement par index.html

#%% Empty cell to run all above