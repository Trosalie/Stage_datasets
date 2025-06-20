# %%
# Importer les bibliothèques nécessaires
import os
import re
import csv
import pandas as pd
import load_file as lf
import uml_class as uml
import json

#%%
# Préparation des variables
# Dictionnaire de référence pour les attributs
dictionnaireReference = json.load(open('dic_hierarchise.json', 'r', encoding='utf-8'))
# dictionnaireReference = pd.DataFrame.from_dict(dictionnaireReference, orient='index')
# dictionnaireReference = dictionnaireReference.transpose()
listeChampsSpatiauxHierarchises = ["pays", "region", "departement", "epci", "quartier", "commune", "iris", "geopoint"]
listeChampsTemporelsHierarchises = ['annee', 'trimestre', 'mois', 'semaine', 'date']

regexAnnee = r'(19\d{2}|20\d{2})$'
regexMois = r'^(0[1-9]|1[0-2])$'
regexJour = r'(0[1-9]|[12]\d|3[01])'
regexDate = r'(' + regexAnnee[:-1] + r'[-/]?' + regexMois[1:-1] + r'[-/]?' + regexJour + r')|(' + regexJour + r'[-/]' + regexMois[1:-1] + r'[-/]' + regexAnnee[:-1] + r')|('+ regexMois[1:-1] + r'[-/]' + regexAnnee[:-1] + r')'
regexHeure = r'(([10]\d)|(2[0-3]))[:h]([0-5]\d)([:h]([0-5]\d))?'
regexTrim = r'(19\d{2}|20\d{2})_[a-zA-Z]{1}[1-3]'

listeRegexTemporel = [[regexDate, 'date'], [regexAnnee, 'annee'], [regexTrim, 'trimestre']]
# %%
# Récupérer la liste des datasets
def getFiles(origine='Opendata'):
    fichiers = []

    for dossier in os.walk(origine):
        for fichier in dossier[2]:
            if fichier.endswith('.csv') or fichier.endswith('.xlsx'):
                fichiers.append(os.path.join(dossier[0], fichier))
    
    return sorted(fichiers)

listeDatasets = getFiles()

# %%
# Fonctions utiles

def remplacerAccents(listeValeurs):
    accents_remplacement = [
        (r'[éèêë]', 'e'),
        (r'[àâä]', 'a'),
        (r'[ôö]', 'o'),
        (r'[îï]', 'i'),
        (r'[ùûü]', 'u')
    ]

    for accents, lettre in accents_remplacement:
        listeValeurs = [re.sub(accents, lettre, x) for x in listeValeurs]

    return listeValeurs

def obtenirDataframeXLSX(df):
    # Vérifier si le DataFrame est vide ou ne contient pas de données valides
    if df.empty:
        raise ValueError("Le DataFrame est vide ou ne contient pas de données valides.")

    # Verifier si la dernière colonne du dataframe est vide
    # Parcours des colonnes du DataFrame par la fin
    for colonne in range(len(df.columns)-1):
        # Si la colonne est vide ou trop peu de valeur non nulle, on la supprime
        nbValeurs = df.iloc[:, -1].notnull().sum()
        if nbValeurs < 2:
            df.drop(df.columns[-1], axis=1, inplace=True)
        else:
            # Si la colonne n'est pas vide, on arrête le parcours
            break

    # Parcourir chaque ligne du DataFrame
    for i in range(df.shape[0]):
        trouve = False
        # Compter le nombre de valeurs nulles dans la ligne
        nbNull = df.iloc[i].isnull().sum()
        # Si il n'y a aucune valeur nulle, c'est probablement la ligne des entêtes
        if nbNull < 1:
            # On vérifie si la ligne suivante peut être la ligne des entêtes
            i += 1
            ligneSuivante = df.iloc[i].tolist()
            
            # Verifier que la ligne ne contient que des chaines de caractères
            if all(isinstance(cell, str) for cell in ligneSuivante):
                # Post-traitement de la ligne des entêtes
                # Remplacer tous les accents par la lettre respective sans accent
                ligneSuivante = remplacerAccents(ligneSuivante)

                # Verifier la présence des champs spatiaux et temporels dans la ligne
                for cell in ligneSuivante:
                    # Si la cellule contient un champ spatial ou temporel c'est probablement la ligne des entêtes
                    if any(champ in cell.lower() for champ in listeChampsSpatiauxHierarchises + listeChampsTemporelsHierarchises):
                        trouve = True
                        break

            # Si la ligne ne contient pas de champs, on considère la ligne initiale
            if not trouve:
                i -= 1
                if any(x in df.iloc[i].tolist() for x in ["Variables : ", "Variables", "Variable", "Attributs", "Attribut"]):
                    trouve = False
                else:
                    # Vérifier s'il y a trop de valeurs numériques dans la ligne en excluant les années
                    nbNumeriques = 0
                    for elt in df.iloc[i].tolist():
                        if (isinstance(elt, int) or isinstance(elt, float)) and not re.match(regexAnnee, str(elt)):
                            nbNumeriques += 1
                    if nbNumeriques > 5:
                        i -= 1
                    
                    trouve = True

        if trouve:
            break

    # On retourne le DataFrame formaté à partir de la ligne des entêtes
    df.columns = df.iloc[i].tolist()
    df = df.iloc[i+1:]
    df.reset_index(drop=True, inplace=True)

    return df

def obtenirDataframeCSV(dataset):
    # Lire la première ligne du fichier CSV pour déterminer le séparateur
    with open(dataset, 'r', encoding='utf-8') as f:
        ligne = f.readline()

        # Compter le nombre de virgules et de points-virgules dans la ligne
        nbVirgules = ligne.count(',')
        nbPointsVirgules = ligne.count(';')

        # Déterminer le séparateur en fonction du nombre de virgules et de points-virgules
        separateur = ',' if nbVirgules > nbPointsVirgules else ';'
    
    # Lire le fichier CSV avec le séparateur déterminé
    try:
        df = pd.read_csv(dataset, sep=separateur, encoding='utf-8', low_memory=False, on_bad_lines='skip')
    except UnicodeDecodeError:
        # Si une erreur d'encodage se produit, essayer avec un autre encodage
        df = pd.read_csv(dataset, sep=separateur, encoding='latin1', low_memory=False, on_bad_lines='skip')
    
    return df

def chercherAttribtusDansEntetes(listeEntetes, listeAttributsRetenus=None):
    if listeAttributsRetenus is None:
        listeAttributsRetenus = []

    # Dictionnaire de correspondance entête -> granularité
    nomEnteteParGranularite = {
        "geopoint": ["geopoint", "lat-long", "latlon", "lattitude", "longitude", "adresse"],
        "iris": ["iris", "code iris", "iris code", "iris code insee"],
        "commune": ["code postal", "code postaux", "postal", "code postal insee", "insee code postal", "commune"],
        "quartier": ["quartier prioritaire", "quartiers prioritaires", "code qp", "qp code", "quartier", "quartiers", "code quartier", "quartier code", "quartier code insee"],
        "epci": ["epci", "epci code", "code epci", "epci code insee"],
        "departement": ["departement", "departements", "code departement", "departement code", "departement code insee"],
        "region": ["region", "regions", "code region", "region code", "region code insee"],
        "pays": ["pays", "pays de résidence", "pays de naissance", "pays d'origine"],
        "annee": ["annee", "annees", "year", "years"],
        "trimestre": ["trimestre", "trimestres", "quarter", "quarters", "semester", "semestres"],
        "mois": ["mois", "months", "month"],
        "semaine": ["semaine", "semaines", "week", "weeks"],
        "date": ["date", "dates", "datetime", "timestamp"]
    }

    # Parcours des entetes de colonnes pour identifier la granularité
    for entete in listeEntetes:
        entete = entete.lower()
        granularite = None
        for key, values in nomEnteteParGranularite.items():
            if any(champ in entete for champ in values):
                granularite = key
                break

        if granularite:
            if granularite in listeChampsSpatiauxHierarchises:
                typeDonnee = "spatial"
            elif granularite in listeChampsTemporelsHierarchises:
                typeDonnee = "temporel"
            else:
                typeDonnee = None

            listeAttributsRetenus.append({
                "nom_attribut": entete,
                "granularite": granularite,
                "type_donnee": typeDonnee
            })

    return listeAttributsRetenus

def estGeopoint(cell):
    cell = str(cell).strip()
    match = re.match(r'^(-?\d+(?:\.\d+)?)[,; ]\s*(-?\d+(?:\.\d+)?)$', cell)
    
    if match:
        try:
            lat, lon = float(match.group(1)), float(match.group(2))
            if -90 <= lat <= 90 and -180 <= lon <= 180:
                return [True, 'geopoint']
        except Exception:
            pass
    return [False, None]

def estSpatial(cell):
    cell = str(cell).lower()
    infoGeopoint = estGeopoint(cell)
    if infoGeopoint[0]:
        return infoGeopoint
    
    for champ, valeurs in dictionnaireReference.items():
        if cell in valeurs:
            return [True, champ]
        
    return [False, None]

def estTemporel(cell):
    if not isinstance(cell, str):
        cell = str(cell)
    for regex, label in listeRegexTemporel:
        if re.match(regex, cell):
            return [True, label]
    if cell.lower() in ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre']:
        return [True, 'mois']
    return [False, None]

def recupererAttributsQualitatifs(headers, df, listeEntetesRetenues, listeAttributsRetenus):
    """
    Identifie les colonnes contenant des attributs qualitatifs (catégoriels),
    incluant les index numériques et les jours du mois.
    
    Args:
        headers: Liste des en-têtes de colonnes
        df: DataFrame pandas à analyser
        liste_attributs_spatiaux: Dictionnaire des attributs spatiaux identifiés
        liste_attributs_temporels: Dictionnaire des attributs temporels identifiés
        
    Returns:
        Une liste des attributs qualitatifs identifiés
    """
    n_rows = min(100, df.shape[0])
    liste_attributs_qualitatifs = []
    
    for header in headers:
        # Ignorer les colonnes déjà identifiées comme spatiales ou temporelles
        if header in listeEntetesRetenues:
            continue
            
        # Récupérer les valeurs de la colonne
        col_values = df[header].head(n_rows)
        
        # Convertir en string, retirer les valeurs 'nan' et valeurs vides
        str_values = col_values.astype(str)
        clean_values = str_values[~str_values.isin(['', 'nan', 'NaN', 'None', 'NONE', 'null', 'Null', 'NULL'])]
        
        if len(clean_values) == 0:
            continue
            
        # Nombre de valeurs uniques
        unique_values = clean_values.nunique()
        ratio_unique = unique_values / len(clean_values) if len(clean_values) > 0 else 0
        
        # CAS 1: Texte avec peu de valeurs uniques - clairement catégoriel
        if 0 < ratio_unique < 0.5:
            try:
                # Vérifier si les valeurs sont numériques
                numeric_values = pd.to_numeric(clean_values)
                
                # Vérifier si c'est un jour du mois (1-31)
                if numeric_values.min() >= 1 and numeric_values.max() <= 31:
                    liste_attributs_qualitatifs.append(header)
                    continue
                    
                # Si les valeurs sont numériques mais en petit nombre (< 20 valeurs distinctes)
                # c'est probablement une variable catégorielle
                if unique_values < 20:
                    liste_attributs_qualitatifs.append(header)
            except:
                # Si la conversion échoue, c'est probablement du texte catégoriel
                liste_attributs_qualitatifs.append(header)
                
        # CAS 2: Texte court mais pas trop de valeurs uniques
        elif col_values.astype(str).str.len().mean() < 30 and unique_values < 50:
            liste_attributs_qualitatifs.append(header)
            
        # CAS 3: Séquences d'entiers potentiellement ID/index
        else:
            try:
                numeric_values = pd.to_numeric(clean_values)
                
                # Vérifier si les valeurs sont principalement des entiers
                if (numeric_values % 1 == 0).mean() > 0.95:
                    sorted_values = sorted(numeric_values.dropna().unique())
                    
                    if len(sorted_values) > 1:
                        # Calculer les écarts
                        gaps = [sorted_values[i+1] - sorted_values[i] for i in range(len(sorted_values)-1)]
                        avg_gap = sum(gaps) / len(gaps)
                        
                        # Un index a souvent des écarts très proches de 1 ou constants
                        if 0.8 < avg_gap < 1.2 or (max(gaps) - min(gaps) < 2):
                            liste_attributs_qualitatifs.append(header)
                            continue
                            
                        # Autre critère: si le nom contient des indices d'ID
                        if any(x in header.lower() for x in ['id', 'code', 'key', 'index', 'numero', 'num', 'identifiant']):
                            liste_attributs_qualitatifs.append(header)
            except:
                pass
                
    return liste_attributs_qualitatifs

def recupererAttributsQuantitatifs(headers, df, listeEntetesRetenues, listeAttributsRetenus):
    """
    Identifie les colonnes contenant des attributs quantitatifs (numériques continus).
    Distingue entre variables quantitatives et identifiants/index numériques.
    
    Args:
        headers: Liste des en-têtes de colonnes
        df: DataFrame pandas à analyser
        listeEntetesRetenues: Liste des en-têtes retenues
        listeAttributsRetenus: Liste des attributs retenus

    Returns:
        Une version complétée de listeAttributsRetenus et listeEntetesRetenues avec les attributs quantitatifs identifiés.
    """
    n_rows = min(100, len(df))  # Analyse un échantillon représentatif
    liste_attributs_quantitatifs = []
    
    for header in headers:
        # Ignorer les colonnes déjà identifiées comme spatiales, temporelles ou qualitatives
        if (header in listeEntetesRetenues):
            continue

        # Récupérer les valeurs de la colonne
        col_values = df[header].head(n_rows)
        
        # Essayer de convertir en numérique
        try:
            numeric_values = pd.to_numeric(col_values)
            
            # Ignorer les colonnes avec trop de valeurs manquantes
            if numeric_values.isna().mean() > 0.5:
                continue
                
            # Vérifier s'il s'agit de données numériques continues
            unique_values = numeric_values.nunique()
            non_na_count = len(numeric_values.dropna())
            ratio_unique = unique_values / non_na_count if non_na_count > 0 else 0
            
            # Détection d'identifiants ou index numériques
            # 1. Vérifier si les valeurs sont majoritairement entières
            is_mostly_integer = (numeric_values % 1 == 0).mean() > 0.95
            
            # 2. Vérifier la distribution des écarts entre valeurs consécutives
            if is_mostly_integer and unique_values >= 5:
                sorted_values = sorted(numeric_values.dropna().unique().tolist())
                gaps = [sorted_values[i+1] - sorted_values[i] for i in range(len(sorted_values)-1)]
                
                # Si les écarts sont variables mais pas trop grands
                if max(gaps) / (sum(gaps)/len(gaps) if gaps else 1) < 10:
                    # Si ce sont des nombres entiers avec des écarts irréguliers mais relativement petits
                    # C'est probablement une variable quantitative ordinale
                    liste_attributs_quantitatifs.append(header)
                # Si trop d'écarts très grands par rapport à la moyenne, c'est probablement un identifiant
                else:
                    continue
                    
            # Pour les nombres à virgule ou avec beaucoup de valeurs uniques -> quantitatif
            elif (not is_mostly_integer) or ratio_unique > 0.5 or unique_values > 20:
                liste_attributs_quantitatifs.append(header)
            
        except:
            # Si la conversion échoue, ce n'est pas une variable numérique
            continue
    
    return liste_attributs_quantitatifs

def traiterDataframe(df):

    """
    etapes a suivre :
        - Chercher les granularités dans les entêtes de colonne
        - Les enregistrer dans une liste [[nom_colonne, granularité, type]]
        - Parcourir les valeurs et classer en conséquence dans la liste (Temporel, Spatial, Qualitatif, Quantitatif)
        - Enregistrer avec la classe UML
    """

    # Parcours des entêtes de colonnes
    listeEntetes = df.columns.tolist()
    listeAttributsRetenus = chercherAttribtusDansEntetes(remplacerAccents(listeEntetes))
    listeEntetesRetenues = []

    # Parcours des cellules du DataFrame
    df_sample = df.sample(n=min(100, len(df)), random_state=42)  # Prendre un échantillon de 100 lignes ou moins si le DataFrame est plus petit
    
    for index, row in df_sample.iterrows():
        for col in row.index:
            cell = row[col]

            # Conversion de la cellule en chaîne de caractères
            if pd.isna(cell):
                continue
            cell = str(cell).strip().lower()

            # Vérification de la granularité spatiale
            infoSpatial = estSpatial(cell)
            if infoSpatial[0]:
                # Enregistrer l'attribut dans la liste
                listeEntetesRetenues.append(col)
                listeAttributsRetenus.append({
                    "nom_attribut": col,
                    "granularite": infoSpatial[1],
                    "type_donnee": "spatial"
                })
                continue

            # Vérification de la granularité temporelle
            infoTemporel = estTemporel(cell)
            if infoTemporel[0]:
                # Enregistrer l'attribut dans la liste
                listeEntetesRetenues.append(col)
                listeAttributsRetenus.append({
                    "nom_attribut": col,
                    "granularite": infoTemporel[1],
                    "type_donnee": "temporel"
                })
                continue
    
    # Verifier les colonnes de type qualitatives et quantitatives
    
    return

# %%
# Boucle de lecture des datasets
for dataset in listeDatasets:
    # Récupérer le nom du fichier et l'extension
    fichier = dataset.split('/')[-1]
    nomFichier, extension = fichier.split('.')

    # Lire le fichier en fonction de son extension
    print(f"Analyse du fichier : {nomFichier[:15]} ({extension})")

    # Variables utiles
    listeDataframe = []

    try :
        match extension:
            case 'csv':
                df = obtenirDataframeCSV(dataset)
                listeDataframe.append([df, nomFichier])
                continue

            case 'xlsx':
                try :
                    for indexFeuille, nomFeuille in enumerate(pd.ExcelFile(dataset).sheet_names):
                        if any(x in nomFeuille.lower() for x in ['variable', 'documentation', 'description', 'presentation', 'présentation']):
                            continue
                        df = pd.read_excel(dataset, sheet_name=nomFeuille, engine='openpyxl', header=None)
                        df = obtenirDataframeXLSX(df)

                        listeDataframe.append([df, f"{nomFichier}_{indexFeuille}"])
                except Exception as e:
                    print(f"Erreur lors de la lecture du fichier Excel {fichier[:10]} | {nomFeuille} : {e}")
                    continue

            case _:
                print(f"Format de fichier non supporté : {extension}")
                continue
        
        # Analyser les DataFrames
        for df, nomDf in listeDataframe:
            try:
                traiterDataframe(df)
            except Exception as e:
                print(f"Erreur lors du traitement du DataFrame {nomDf[:20]} : {e}")
                continue

    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {fichier} : {e}")
        continue
# %%
