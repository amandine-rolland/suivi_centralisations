import os
import pandas as pd
import polars as pl
from xlsxwriter import Workbook
from datetime import timedelta

from sqlalchemy import create_engine
from sqlalchemy.engine import URL

#import traceback # permet d'afficher les messages d'erreur du try except

import warnings
warnings.filterwarnings('ignore')

# Defining the connection string
driver = "SQL Server"
server = "172.16.22.227"
user = "usrSupplyChain"
pwd =  "u575u99yCha1n$"
BDD_name = "SupplyChain"

# conn = pyodbc.connect(f'DRIVER={driver}; Server={server}; UID={user}; PWD={pwd}; DataBase={BDD_name}')
connection_string = f'DRIVER={driver}; Server={server}; UID={user}; PWD={pwd}; DataBase={BDD_name}'
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
engine = create_engine(connection_url)

con = engine.connect()

## paramètres ##

param = pd.read_excel(r"\\BUREAU2022\ACHATSChaussures\LOGISTIQUE\Rapports BDD Supply Chain\data_prep\SuiviCentra\Parametres_centra_kpi.xlsx",
                      sheet_name="IDCatalogue")

save_folder = r"\\BUREAU2022\ACHATSChaussures\LOGISTIQUE\Rapports BDD Supply Chain\data_prep\SuiviCentra"


# ## fonctions ##################################################################


def get_deblocage(marque, saison, prix_achat):
    data = param[(param["Marque"] == marque) & (param["Saison"] == saison)].reset_index(drop=True)
    deblocage_folder = data["DossierSuiviDesDeblocages"].iloc[0]
    deblocage_file = data["Fichier"].iloc[0]
    deblocage_sheet = data["Onglet"].iloc[0]

    #if not pd.isna(deblocage_folder) and not pd.isna(deblocage_file) and not pd.isna(deblocage_sheet):
    try:
        deblocage_path = os.path.join(deblocage_folder, deblocage_file)
        print("fichier déblocage :",deblocage_path)
        df = pd.read_excel(deblocage_path, sheet_name=deblocage_sheet)
        
        # type des données
        df["DateDeblocage"] = pd.to_datetime(df["DateDeblocage"]).dt.date
        df["IDCommande"] = df["IDCommande"].astype(str)
        
        # ajout de la valo débloquée
        #prix_achat = get_PrixAchat_RefFourCoul(achat, RefFourCouleur)
        df = pd.merge(df, prix_achat,
                      left_on=["IDCommande", "ReferenceFournisseurCouleur"],
                     right_on=["IDCommandeAchat", "ReferenceFournisseurCouleur"],
                     how="left")
        df["MontantDeblocage"] = df["QuantiteDeblocagePiece"] * df["PrixAchatPiece"]
        
    #else:
    except Exception:
        df = pd.DataFrame()
        print("Nom de dossier, fichier et/ou onglet manquant.")
            
    return df



def get_annulations(marque, saison, prix_achat):
    """ quantité et montant annulés par la marque à ajouter au tableau de synthèse"""
    try :
        data = param[(param["Marque"] == marque) & (param["Saison"] == saison)].reset_index(drop=True)
        folder = data["DossierSuiviDesDeblocages"].iloc[0]
        file = data["Fichier"].iloc[0]
        sheet = data["OngletAnnulations"].iloc[0]

        df = pd.read_excel(os.path.join(folder,file), sheet_name=sheet)
        df["IDCommande"] = df["IDCommande"].astype(str)
        df = pd.merge(df, prix_achat,
                      left_on=["IDCommande", "ReferenceFournisseurCouleur"],
                      right_on=["IDCommandeAchat", "ReferenceFournisseurCouleur"],
                      how="left"
                     )
        df["MontantAnnulation"] = df["QuantiteAnnulationPiece"] * df["PrixAchatPiece"]
    
    except Exception as e:
        df = pd.DataFrame()
        print(e)
    
    return df


def int_to_sql(liste_int):

    return "(" + ",".join([str(x) for x in liste_int]) + ")"


def str_to_sql(liste_str):

    return "('" + "','".join(liste_str) + "')"


def get_achat(idcat_sql):

    q = f"""
    SELECT 
    ca.IDCommandeAchat
    , ca.IDCatalogue
    , ca.IDArticle
    , DateLivraisonAttendue
    , (ca.QuantiteCommandee * a.QteParPack) as QuantiteCommandeePiece
    , (ca.PrixAchat / a.QteParPack) as PrixAchatPiece
    , (ca.QuantiteCommandee * ca.PrixAchat) as MontantAchat 
    FROM CommandeAchat as ca
    LEFT JOIN Article as a ON a.IDArticle = ca.IDArticle
    WHERE IDCatalogue IN {idcat_sql}
    AND Statut <> 99
    
    """

    return pd.read_sql(q, con)


def get_idcmd_achat(achat):

    return achat["IDCommandeAchat"].unique()


def get_reception(liste_idcmd_achat_sql):

    q = f"""  
    SELECT
    m.IDCommande
    , m.IDArticle
    , m.DateMouvement as DateReception
    , (m.QuantiteMouvement * a.QteParPack) as QuantiteReceptionPiece
    , (m.QuantiteMouvement * m.CAMVMvt) as MontantReception
    FROM Mouvement as m
    LEFT JOIN Article as a ON a.IDArticle = m.IDArticle
    WHERE IDTypeMouvement = 25
    AND IDCommande IN {liste_idcmd_achat_sql}
    """
    df = pd.read_sql(q, con)
    # transformer datetime en date
    if not df.empty:
        df["DateReception"] = pd.to_datetime(df["DateReception"])
        df["DateReception"] = df["DateReception"].dt.date
        
        # grouper par date de réception
        df = (df
          .groupby(["IDCommande", "IDArticle", "DateReception"])[["QuantiteReceptionPiece", "MontantReception"]]
          .sum()
          .reset_index()
         )
    
    return df



def ref_four_couleur():
    q = """
    WITH ArtRefCouleur as (
    -- Requete qui renvoi le code article, unite produit, ref fournisseur et couleur (pour les pièces)    
        SELECT IDArticle, Conditionnement, CategorieFEDAS, UniteProduit, ReferenceFournisseur, CodeCouleur, QteParPack
        FROM Article
        LEFT JOIN Couleur ON IDArticle = CodeM3
        WHERE IDArticle NOT IN ('AJUSTE', 'PERTE', 'TVA20', 'TVA55', 'ZGESCOM', 'ZGESPORT', 'ZPUB', 'ZTRADE')
    ),

    ArticleReferenceFournisseur as (
    -- requete qui démultiplie les lignes de pack et renvoi la ref fournisseur de chaque code article contenu dans les pack
        SELECT a.IDArticle,
                a.CategorieFEDAS,
                a.ReferenceFournisseur,
                a.UniteProduit,
                a.QteParPack,
                ISNULL(a.CodeCouleur, '') as CodeCouleur,
                ap.CodeArticle,
                a2.ReferenceFournisseur as RefFourPiece,
                a2.CodeCouleur as CodeCouleurPiece,
        CASE
            WHEN a.UniteProduit = 'PIECE' THEN a.ReferenceFournisseur
            ELSE a2.ReferenceFournisseur
        END AS ReferenceFournisseurClean,

        CASE
            WHEN a.UniteProduit = 'PIECE' THEN ISNULL(a.CodeCouleur, '')
            ELSE ISNULL(a2.CodeCouleur, '')
        END AS CodeCouleurClean   

        FROM ArtRefCouleur AS a
        LEFT JOIN ArticlePack as ap ON a.IDArticle = ap.CodePack
        LEFT JOIN ArtRefCouleur AS a2 ON ap.CodeArticle = a2.IDArticle
        ),

    ArticleDistinct AS (
        SELECT DISTINCT IDArticle, UniteProduit, CategorieFEDAS, QteParPack,
                ReferenceFournisseurClean as ReferenceFournisseur,
                CodeCouleurClean as CodeCouleur
        FROM ArticleReferenceFournisseur
    ),

    last as (
        SELECT IDArticle, UniteProduit, ReferenceFournisseur,
                STRING_AGG(CodeCouleur, '') as CodeCouleur, QteParPack
        FROM ArticleDistinct
        GROUP BY IDArticle, UniteProduit, CategorieFEDAS,
                    ReferenceFournisseur, QteParPack
                )
    SELECT *,
    CASE
         WHEN CodeCouleur IS NOT NULL AND CodeCouleur <> '' THEN ReferenceFournisseur+'-'+CodeCouleur
         ELSE ReferenceFournisseur
    END AS ReferenceFournisseurCouleur

    FROM last

    """ 
    return pd.read_sql(q, con)

def get_PrixAchat_RefFourCoul(achat, RefFourCouleur):
    """
    renvoi le prix d'achat par ReferenceFournisseurCouleur 
    """
       
    px_achat_u = pd.merge(achat[["IDCommandeAchat", "IDArticle", "PrixAchatPiece"]], 
                      RefFourCouleur[["IDArticle", "ReferenceFournisseurCouleur"]], 
                      on="IDArticle", 
                      how='left'
                     )
    del px_achat_u["IDArticle"]
    px_achat_u = px_achat_u.drop_duplicates(subset=["IDCommandeAchat", "ReferenceFournisseurCouleur", "PrixAchatPiece"])
    px_achat_u = px_achat_u[["IDCommandeAchat", "ReferenceFournisseurCouleur", "PrixAchatPiece"]]
    
    # Y a t_il des doublons de prix d'achat à la refFourCouleur
    print("Nombre de doublons de prix :",
          px_achat_u[["ReferenceFournisseurCouleur", "PrixAchatPiece"]].drop_duplicates().duplicated().sum())
    
    return px_achat_u


def get_vente(idcat_sql):
    """
    récupération des IDCommandeVente grace aux IDCatalogue
    puis des éxpéditions vers les magasins via les IDCommandeVente
    """
    
    q = f"""
    SELECT IDCommandeVente
    , IDCatalogue
    , IDArticle
    , SUM(QuantiteCommandee) AS QuantiteCommandee
    FROM CommandeVente
    WHERE IDCatalogue IN {idcat_sql}
    AND Statut <> 99
    GROUP BY IDCommandeVente
    , IDCatalogue
    , IDArticle
    """
    
    return pd.read_sql(q, con)

def get_idcmd_vente(vente):

    return vente["IDCommandeVente"].unique()

def get_expedition(liste_idcmd_vente_sql):
    """
    renvoi les expéditions des commandes de vente
    basée sur les IDCommandeVente
    """
    
    q = f"""
    SELECT m.IDCommande
    , m.IDArticle
    , m.DateMouvement as DateExpedition
    , (- m.QuantiteMouvement * a.QteParPack) as QuantiteExpeditionPiece
    , (- m.QuantiteMouvement * m.CAMVMvt) as MontantExpedition
    FROM Mouvement as m
    LEFT JOIN Article as a ON a.IDArticle = m.IDArticle
    WHERE IDTypeMouvement = 31
    AND IDCommande IN {liste_idcmd_vente_sql}
    """
    df = pd.read_sql(q, con)
    # transformer datetime en date
    if not df.empty:
        df["DateExpedition"] = pd.to_datetime(df["DateExpedition"])
        df["DateExpedition"] = df["DateExpedition"].dt.date
        
    return df



def calcul_pct_df(df, denom_qte, denom_valo):

    if not df.empty:
        # colonnes quantité cumulé à utiliser
        for col in [col for col in df.columns if col.startswith("Quantite")]:
            new_col = "pct_" + col
            df[new_col] = df[col] / denom_qte
        for col in [col for col in df.columns if col.startswith("Montant")]:
            new_col = "pct_" + col
            df[new_col] = df[col] / denom_valo

    return df


def calcul_pct_cum_df(df, denom_qte, denom_valo):

    if not df.empty:
        # colonnes quantité cumulé à utiliser
        for col in [
                col for col in df.columns
                if col.startswith("Quantite") and col.endswith("_cum")
        ]:
            new_col = "pct_" + col
            df[new_col] = df[col] / denom_qte
        for col in [
                col for col in df.columns
                if col.startswith("Montant") and col.endswith("_cum")
        ]:
            new_col = "pct_" + col
            df[new_col] = df[col] / denom_valo
    return df


def calc_synthese(achat, reception, expedition, deblocage, annulation):  # ensuite ajouter déblocage

    qte_achat = achat["QuantiteCommandeePiece"].sum()
    qte_recep = reception["QuantiteReceptionPiece"].sum()
    qte_expe = expedition["QuantiteExpeditionPiece"].sum()

    valo_achat = achat["MontantAchat"].sum()
    valo_recep = reception["MontantReception"].sum()
    valo_expe = expedition["MontantExpedition"].sum()

    pct_recep = round(qte_recep / qte_achat, 3)
    pct_expe = round(qte_expe / qte_achat, 3)

    pct_valo_recep = round(valo_recep / valo_achat, 3)
    pct_valo_expe = round(valo_expe / valo_achat, 3)

    if not deblocage.empty:
        qte_debl = deblocage["QuantiteDeblocagePiece"].sum()
        valo_debl = deblocage["MontantDeblocage"].sum()
        pct_debl = round(qte_debl / qte_achat, 3)
        pct_valo_debl = round(valo_debl / valo_achat, 3)

        data = {
            "Unité": ["QuantitePiece", "Valo €"],
            "Commande": [qte_achat, valo_achat],
            "Deblocage": [qte_debl, valo_debl],
            "Reception": [qte_recep, valo_recep],
            "Expedition": [qte_expe, valo_expe],
            "% Deblocage": [pct_debl, pct_valo_debl],
            "% Reception": [pct_recep, pct_valo_recep],
            "% Expedition": [pct_expe, pct_valo_expe]
        }
    else:
        data = {
            "Unité": ["QuantitePiece", "Valo €"],
            "Commande": [qte_achat, valo_achat],
            "Reception": [qte_recep, valo_recep],
            "Expedition": [qte_expe, valo_expe],
            "% Reception": [pct_recep, pct_valo_recep],
            "% Expedition": [pct_expe, pct_valo_expe]
        }
        
    if not annulation.empty:
        qte_annul = annulation["QuantiteAnnulationPiece"].sum()
        valo_annul = annulation["MontantAnnulation"].sum()
        pct_annul = round(qte_annul / qte_achat, 3)
        pct_valo_annul = round(valo_annul / valo_achat, 3)
        
        data["Annulation"] = [qte_annul, valo_annul]
        data["% Annulation"] = [pct_annul, pct_valo_annul]
    
    df = pd.DataFrame(data=data)
    #df.index = ["QuantitePiece", "Valo €"]
    return df


def add_idcatalogue(reception, achat, expedition, vente):

    reception = pd.merge(reception,
                         achat[["IDCatalogue",
                                "IDCommandeAchat"]].drop_duplicates(),
                         left_on="IDCommande",
                         right_on="IDCommandeAchat",
                         how="left")
    expedition = pd.merge(expedition,
                          vente[["IDCatalogue",
                                 "IDCommandeVente"]].drop_duplicates(),
                          left_on="IDCommande",
                          right_on="IDCommandeVente",
                          how="left")
    return reception, expedition


def calc_retard_livraisons(achat, reception, RefFourCouleur):
    
    if not reception.empty:
        df = pd.merge(achat, reception, left_on=["IDCommandeAchat", "IDArticle"], right_on=["IDCommande", "IDArticle"], how="left")
        df["DateReception"] = pd.to_datetime(df["DateReception"])
        df["DateLivraisonAttendue"] = pd.to_datetime(df["DateLivraisonAttendue"])
        df["Reception-attendu"] = df["DateReception"] - df["DateLivraisonAttendue"] - timedelta(hours = 24)
        # SI Reception-attendu > 0 alors il ya un retard de livraison
        df["RetardLivraison"] = df["Reception-attendu"].apply(lambda x: x > timedelta(days=0))
        # calcul du retard de livraison pondéré par la quantité réceptionnée
        df["Reception*Qte/1000"] = df["Reception-attendu"] * df["QuantiteReceptionPiece"] /1000 # diviser par 1000 sinon ça fait un int too big
        ## Focus sur les retards
        retard = df[df["RetardLivraison"]]

        ## en ref fournisseur Couleur

        retard = pd.merge(retard, RefFourCouleur, on="IDArticle", how="left")
        retard_g = (retard
                    .groupby(["ReferenceFournisseurCouleur",
                              "IDCommandeAchat", "IDCatalogue",
                              "DateLivraisonAttendue","QuantiteCommandeePiece", "PrixAchatPiece", "MontantAchat"
                             ])
                    .agg(DateReceptionMin = ("DateReception", min), 
                         DateReceptionMax = ("DateReception", max),
                         MontantReception = ("MontantReception", sum),
                         ReceptionQteDiv1000 = ("Reception*Qte/1000", sum), 
                         QuantiteReceptionPiece = ("QuantiteReceptionPiece", sum)
                        )
                    .reset_index()
                   )
        retard_g["RetardMoyenPondéré"] = retard_g["ReceptionQteDiv1000"] / retard_g["QuantiteReceptionPiece"] * 1000

    else:
        retard_g=pd.DataFrame()
    
    return retard_g


def set_glossaire():

    data = {
        "Remarque": [
            "les % sont calculés en divisant par les quantités ou montants des commandes d'achat dans M3",
            "les quantités sont exprimées en pièces",
            """délai pondéré = délai de livraison pondéré par les quantités réceptionnées.\n
            Un délai de 24h est soustrait, ce qui correspond au temps laissé à l'entrepôt pour réaliser les réceptions""",
            "délai livraison = DateRéception - DateDéblocage - 24h",
            "RetardLivraison = DateReception - DateLivraisonAttendueCmdAchat - 24h"
            "RetardMoyenPondéré = retard moyen en jours pondéré par la qté réceptionnée"
        ]
    }

    return pl.DataFrame(data)


def get_warnings(deblocage):
    if not deblocage.empty:
        # doublons de déblocage
        doublons = deblocage[deblocage[[
            "IDCommande", "ReferenceFournisseurCouleur"
        ]].duplicated()][["IDCommande", "ReferenceFournisseurCouleur"]]
        df = pd.merge(deblocage, doublons, on=["IDCommande", "ReferenceFournisseurCouleur"], how="inner")
    else:
        df = pd.DataFrame()

    return df


def groupby_ref_couleur_date(df, RefFourCouleur):
    if not df.empty:
        # ajouter la référence fournisseur couleur
        if "ReferenceFournisseurCouleur" not in df.columns:
            df = pd.merge(df, RefFourCouleur[["IDArticle", "ReferenceFournisseurCouleur"]], on="IDArticle", how="left")

        # grouper par ref four couleur et date
        col_gp = ["ReferenceFournisseurCouleur"] + [col for col in df.columns if col.startswith("Date")]
        col_agg = [col for col in df.columns if (col.startswith("Quantite") and col.endswith("Piece")) or col.startswith("Montant")]

        df = df.groupby(col_gp)[col_agg].sum().reset_index()
    
    return df


def get_ref_date(achat_ref, deblocage_ref_date, reception_ref_date, expedition_ref_date):
    """
    renvoi un dataframe avec toutes les referencesfournisseurs couleurs possibles et toutes les dates de mouvement
    """
    
    # récupération de toutes les réf couleur possibles pour les df non vides
    ref_all = pd.concat([
                        df["ReferenceFournisseurCouleur"]
                        for df in [achat_ref, deblocage_ref_date, reception_ref_date, expedition_ref_date]
                        if not df.empty
                        ]).drop_duplicates()

    RefFourCoul_all = pd.DataFrame({"ReferenceFournisseurCouleur":ref_all})

    # récupération de toutes les dates de mouvement
    DateMvt = pd.DataFrame()
    for df in [deblocage_ref_date, reception_ref_date, expedition_ref_date]:
        if not df.empty:    
            col = [col for col in df.columns if col.startswith("Date")]
            date_mvt = df[col]
            date_mvt = date_mvt.rename(columns = {col[0]: "Date"})
            DateMvt = pd.concat([DateMvt, date_mvt]).drop_duplicates()

    # produit cartésien pour avoir pour chaque ref une ligne pour chaque date
    ref_date_mvt = pd.merge(RefFourCoul_all, DateMvt, how="cross").sort_values(by=["ReferenceFournisseurCouleur", "Date"])


    return ref_date_mvt



def merge_df_by_ref_four(deblocage_grouped, reception_grouped, expedition_grouped, ref_date_mvt):
    """
    renvoi les données des déblocages, expédition et / ou réception compilées
    """
    # savoir quels sont les df non vides
    not_empty_df = [
        df
        for df in [deblocage_grouped, reception_grouped, expedition_grouped]
        if not df.empty
    ]

    if len(not_empty_df) == 1:
        df = not_empty_df[0]
        # ajouter toutes les ref four couleur présentes dans les commandes d'achat
        col_date = [col for col in df if col.startswith("Date")][0]
        df = pd.merge(ref_date_mvt,
                      df,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date],
                      how="left")

    elif len(not_empty_df) == 2:
        # merge des 2 df sur les colonnes de date
        df1 = not_empty_df[0]
        df2 = not_empty_df[1]
        col_date1 = [col for col in df1 if col.startswith("Date")][0]
        col_date2 = [col for col in df2 if col.startswith("Date")][0]
        # ajouter toutes les ref sur toutes les dates
        df = pd.merge(ref_date_mvt,
                      df1,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date1],
                      how="left"
                     )
        df = pd.merge(df,
                      df2,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date2],
                      how="left")
    elif len(not_empty_df) == 3:  # 3 dataframes à fusionner
        df1 = not_empty_df[0]
        df2 = not_empty_df[1]
        df3 = not_empty_df[2]
        col_date1 = [col for col in df1 if col.startswith("Date")][0]
        col_date2 = [col for col in df2 if col.startswith("Date")][0]
        col_date3 = [col for col in df3 if col.startswith("Date")][0]

        df = pd.merge(ref_date_mvt,
                      df1,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date1],
                      how="left"
                     )
        df = pd.merge(df,
                      df2,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date2],
                      how="left")
        df = pd.merge(df,
                      df3,
                      left_on=["ReferenceFournisseurCouleur", "Date"],
                      right_on=["ReferenceFournisseurCouleur", col_date3],
                      how="left")

    else:
        df = pd.DataFrame()

    return df


def cumsum_by_ref_date(df):
    """
    renvoi un tableau des données de déblocage, réception, expédition en date et les sommes cumulées
    """
    if not df.empty:
        # grouper par date
        # Colonne unique de date
        # on n'a pas forcément les 3 colonnes de date, on peut en avoir que 2 (si c'est une centra sans déblocage)
        # identifier les colonnes de date par leur nom
        colonnes_dates = [col for col in df.columns if col.startswith("Date")]

        # on rempli horizontalement (axis=1) avec la première date non vide (vers la gauche puis vers la droite)
        # puis on prend la 1ère colonne
        df["Date"] = df[colonnes_dates].bfill(axis=1).ffill(axis=1).iloc[:, 0]

        # grouper par Date
        other_col = [col for col in df.columns if not col.startswith("Date")]
        other_col.remove("ReferenceFournisseurCouleur")
        df1 = df.groupby(["ReferenceFournisseurCouleur", "Date"])[other_col].sum().reset_index()

        # calculer les cumsum
        df1 = calcul_cumsum_by_ref(df1)
    else:
        df1 = pd.DataFrame()

    return df1


def calcul_cumsum_by_ref(df):
    # colonne non date
    other_col = [col for col in df.columns if not col.startswith("Date") and not col.startswith("Reference")]
    for col in other_col:
        new_col = col + "_cum"
        df[new_col] = df.groupby(["ReferenceFournisseurCouleur"])[col].cumsum()

    return df


def sum_by_ref_date_deblocage(df):
    """
    renvoi un tableau des données de déblocage, réception, expédition en date de déblocage
    """
    # quand il y a des NaN dans les dates de déblocages (=> il y a des réceptions ou expé mais pas de déblocage)
    # alors on rempli les NaN avec la prochaine date de déblocage
    # pour les réceptions/expéditions de fin de saison, il n'y a pas de déblocage à venir (car les déblocages sont finis)
    # donc on met la date max entre la réception et l'expédition
    if not df.empty:
        # identifier les colonnes de date
        #col_date = [col for col in df.columns if col.startswith("Date")]

        if "DateDeblocage" in df.columns:
            max_dt = df["Date"].max() # permet de mettre une date pour les lorsque les déblocages sont finis
                                        # et qu'il reste des réceptions expéditions à venir
            df["DateDeblocage"] = df.groupby("ReferenceFournisseurCouleur")["DateDeblocage"].bfill()      

            df["DateDeblocage"] = df["DateDeblocage"].fillna(max_dt)

            col_non_date = [
                col for col in df.columns if not col.startswith("Date") and not col.startswith("Reference")
            ]
            df1 = df.groupby(["ReferenceFournisseurCouleur", "DateDeblocage"])[col_non_date].sum().reset_index()

            # récupérer les date déblocage
            dt_debl = pd.DataFrame({"DateDeblocage":df["DateDeblocage"].drop_duplicates()})
            # récupérer toutes les réf fournisseur
            ref_debl = pd.DataFrame({"ReferenceFournisseurCouleur":df["ReferenceFournisseurCouleur"].drop_duplicates()})
            # tableau de toutes les réf et toutes les dates de déblocage
            ref_dt_debl = pd.merge(ref_debl, dt_debl, how="cross").sort_values(by=["ReferenceFournisseurCouleur", "DateDeblocage"])

            df2 = pd.merge(ref_dt_debl, df1, on=["ReferenceFournisseurCouleur","DateDeblocage"], how="left")
            df2 = df2.fillna(0)
            # calculer les cumsum
            df3 = calcul_cumsum_by_ref(df2)
        else:
            df3 = pd.DataFrame()
    else:
        df3 = pd.DataFrame()
    return df3


def calc_delai_livraison_by_ref(deblocage, reception, RefFourCouleur):
    """ le délai théorique entre la livraison et la réception est de 24h pour Logs
        pondéré par la qté receptionnée    
    """
    if not deblocage.empty and not reception.empty:
        # passer les réceptions en par ref fournisseur
        reception_ref = pd.merge(reception, RefFourCouleur[["IDArticle", "ReferenceFournisseurCouleur"]], on="IDArticle", how="left")
        # puis grouper par ref four couleur
        reception_ref = (reception_ref
                         .groupby(["IDCommande", "ReferenceFournisseurCouleur", "DateReception"])[["QuantiteReceptionPiece", "MontantReception"]]
                         .sum()
                         .reset_index()
                        )

        delai = pd.merge(deblocage, reception_ref, on=["IDCommande", "ReferenceFournisseurCouleur"], how="left")
        # supprimer les lignes ou la date de réception est antérieure à la date de déblocage
        delai = delai[delai["DateReception"] > delai["DateDeblocage"]]
        # calculer le delta entre le déblocage et la réception
        delai["delai j"] = delai["DateReception"] - delai["DateDeblocage"] + pd.Timedelta(days=-1)
        # supprimer les lignes où la date de réception est antérieure à la date de déblocage
        delai = delai[delai["delai j"] > timedelta(days=0)]

        # convertir en secondes
        delai["delai j"] = delai["delai j"].apply(lambda x : x.total_seconds())

        delai["delai*qte"] = delai["delai j"] * delai["QuantiteReceptionPiece"] / (60*60*24) # division par (60*60*24) pour l'avoir en jours


        # grouper par date de déblocage - aggreger par moyenne de délai
        delai_g = delai.groupby(["ReferenceFournisseurCouleur","DateDeblocage"], dropna=False).agg({"delai j":"mean", "QuantiteReceptionPiece": sum, "delai*qte":sum}).reset_index()
        # calculer delai moyen pondéré
        delai_g["delai pondere j"] = delai_g["delai*qte"] / delai_g["QuantiteReceptionPiece"]
        delai_g.drop(columns=["QuantiteReceptionPiece", "delai*qte"], inplace=True)

    
    else:
        delai_g = pd.DataFrame()
    
    return delai_g
    
    return delai_g


def save_by_ref(marque, saison, parametrage, synthese, reception, recep, expedition,
         expe, achat, vente, deblocage, debl, data_by_date,
         data_by_date_deblocage, prix_achat, delai_livraison, warning, retard, annulation):
    save_name = "_".join(["suivi_par_ref", marque, saison]) + ".xlsx"
    with Workbook(os.path.join(save_folder, save_name)) as wb:
        glossaire = set_glossaire()
        glossaire.write_excel(workbook=wb, worksheet="glossaire", autofit=True)
        # enregistrer les paramètres utilisés
        pl.from_pandas(parametrage).write_excel(workbook=wb,
                                                worksheet="parametrage",
                                                autofit=True)
        
        # enregistrer les alertes (doublons)
        if not warning.empty:
            ws = wb.add_worksheet("alerte")
            ws.write("A1", "Doublons de déblocage") # permet d'ajouter une ligne avec du texte avant le tableau
            pl.from_pandas(warning).write_excel(workbook=wb, worksheet=ws, position="A2", autofit=True)        
        
        # enregistrer les data
        pl.from_pandas(synthese).write_excel(workbook=wb,
                                             worksheet="synthese",
                                             autofit=True)
        pl.from_pandas(data_by_date).write_excel(workbook=wb,
                                                 worksheet="par date",
                                                 autofit=True)
        if not data_by_date_deblocage.empty:
            pl.from_pandas(data_by_date_deblocage).write_excel(
                workbook=wb, worksheet="par date deblocage", autofit=True)
        if not delai_livraison.empty:
            pl.from_pandas(delai_livraison).write_excel(
                workbook=wb, worksheet="delai livraison", autofit=True)
        if not debl.empty:
            pl.from_pandas(debl).write_excel(workbook=wb,
                                             worksheet="deblocages",
                                             autofit=True)
        pl.from_pandas(recep).write_excel(workbook=wb,
                                          worksheet="reception",
                                          autofit=True)
        pl.from_pandas(expe).write_excel(workbook=wb,
                                         worksheet="expedition",
                                         autofit=True)
        if not deblocage.empty:
            pl.from_pandas(deblocage).write_excel(workbook=wb,
                                                  worksheet="detail_deblocage",
                                                  autofit=True)
        if not annulation.empty:
            pl.from_pandas(annulation).write_excel(workbook=wb,
                                                  worksheet="detail_annulation",
                                                  autofit=True)
        pl.from_pandas(reception).write_excel(workbook=wb,
                                              worksheet="detail_recep",
                                              autofit=True)
        if not retard.empty:
            pl.from_pandas(retard).write_excel(workbook=wb,
                                               worksheet="retardLivraison",
                                               autofit=True
                                              )
        pl.from_pandas(expedition).write_excel(workbook=wb,
                                               worksheet="detail_expe",
                                               autofit=True)
        pl.from_pandas(achat).write_excel(workbook=wb,
                                          worksheet="commande_achat",
                                          autofit=True)
        pl.from_pandas(vente).write_excel(workbook=wb,
                                          worksheet="commande_vente",
                                          autofit=True)
        pl.from_pandas(prix_achat).write_excel(workbook=wb,
                                               worksheet="prix_achat",
                                               autofit=True)

    return


###############################################################################

def main():
    RefFourCouleur = ref_four_couleur()
    
    for marque in param["Marque"].unique():
        for saison in param[param["Marque"] == marque]["Saison"].unique(): # permet de recupérer les données sur plusieurs saisons
            parametrage = param[(param["Marque"] == marque) & (param["Saison"] == saison)]
            idcat = parametrage["IDCatalogue"].unique()
            print(marque, ":" ,idcat)
    
            idcat_sql = int_to_sql(idcat)
    
            # 1- réceptions
            achat = get_achat(idcat_sql)
    
            idcmd_achat = get_idcmd_achat(achat)
            idcmd_achat_sql = str_to_sql(idcmd_achat)
    
            reception = get_reception(idcmd_achat_sql)
            
            # 2 - Prix d'achat par RefFour Coul - il faut l'enregistrer dans le fichier excel car besoin ensuite dans le PwBI
            # et dans get_deblocage
            prix_achat = get_PrixAchat_RefFourCoul(achat, RefFourCouleur)
    
            # 3 - expéditions    
            vente = get_vente(idcat_sql)
    
            idcmd_vente = get_idcmd_vente(vente)
            idcmd_vente_sql = str_to_sql(idcmd_vente)
    
            expedition = get_expedition(idcmd_vente_sql)
    
            # 4 - Totaux commandés pour le calucl des % vs qté & montant commandée dans M3
            qte_achat_tot = achat["QuantiteCommandeePiece"].sum()
            montant_achat_tot = achat["MontantAchat"].sum()
    
            
            # 5 - déblocages et annulations
            deblocage = get_deblocage(marque, saison, prix_achat)
            annulation = get_annulations(marque, saison, prix_achat)
          
            # 6 - synthese
            synthese = calc_synthese(achat, reception, expedition, deblocage, annulation)
            
            # 7 - ajouter la ref fournisseur
            achat_ref = groupby_ref_couleur_date(achat, RefFourCouleur) # pour pouvoir récupérer la liste des ref four couleur
            deblocage_ref_date = groupby_ref_couleur_date(deblocage, RefFourCouleur)
            reception_ref_date = groupby_ref_couleur_date(reception, RefFourCouleur)
            expedition_ref_date = groupby_ref_couleur_date(expedition, RefFourCouleur)
            
            # 8 - compiler les données des 3 df
            # pour pouvoir faire des cumsum utilisables dans PowerBi en gardant le détail de la réf couleur,
            # il va falloir une ligne pour chaque date et chaque ReferenceFournisseurCouleur
            ref_date_mvt = get_ref_date(achat_ref, deblocage_ref_date, reception_ref_date, expedition_ref_date)
            
            data =  merge_df_by_ref_four(deblocage_ref_date, reception_ref_date, expedition_ref_date, ref_date_mvt)
            
            data_by_date = cumsum_by_ref_date(data)
            
            # 9 - compiler par date de déblocage
            data_by_date_deblocage = sum_by_ref_date_deblocage(data)
            
            
                            ### CALCULER LES %###
                
            # 10 - calucl des % cumulé vs qté & montant commandée dans M3
            data_by_date = calcul_pct_cum_df(data_by_date, qte_achat_tot, montant_achat_tot)
            data_by_date_deblocage = calcul_pct_cum_df(data_by_date_deblocage, qte_achat_tot, montant_achat_tot)
            
            # 10 - Calcul des % non cumulés vs qté & montant commandée dans M3
            deblocage_ref_date = calcul_pct_df(deblocage_ref_date, qte_achat_tot, montant_achat_tot)
            reception_ref_date = calcul_pct_df(reception_ref_date, qte_achat_tot, montant_achat_tot)
            expedition_ref_date = calcul_pct_df(expedition_ref_date, qte_achat_tot, montant_achat_tot)
            
            # 11 - calcul du délai de livraison moyen
            delai_livraison = calc_delai_livraison_by_ref(deblocage, reception, RefFourCouleur)
            
            # 12 - calcul retard de livraison
            retard = calc_retard_livraisons(achat, reception,RefFourCouleur)
                        
            # ajout des IDCatalogue à réception et expédition
            reception, expedition = add_idcatalogue(reception, achat, expedition, vente)
            
            # alertes
            alertes = get_warnings(deblocage)
            
            # Enregistrement en polars pour l'autoajustement des colonnes
            save_by_ref(marque, saison, parametrage, synthese ,reception , reception_ref_date , expedition ,expedition_ref_date ,
                 achat_ref, vente, deblocage ,deblocage_ref_date , data_by_date, data_by_date_deblocage, prix_achat,
                 delai_livraison, alertes, retard, annulation
                )


if __name__ == "__main__":
    main()