from Sheets_Manip import *
import datetime
import pandas as pd
from pandas.tseries.offsets import DateOffset
from dateutil import parser

import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")


class Transpa():
    
    def __init__(self):
        self.errors2 = []
        self.df_test = pd.DataFrame()
        
    

    def col_df2(self, df):
        """
        This method initializes self.df_test with a predefined structure:
        - Row labels represent different validation checks.
        - The first column is 'Modif/Tests', filled with the row labels.
        - Other columns correspond to unique file identifiers found in the last column of df.
        """

        # Define the row labels
        row_labels = [
            'Somme VM (col.24) = NAV Annoncée (col.5)', 'Tests maturité produits de taux (col. 39)', 'Test fréquence produits de taux (col. 38): vide', 'Test fréquence produits de taux (col. 38): texte',
            'Test validité redemtion rates produits de taux (col. 41): Vide','Test validité redemtion rates produits de taux (col. 41): Texte','Test validité redemtion rates produits de taux (col. 41): Négatif', 
            'Taux de coupon (col. 33): vide', 'Taux de coupon (col. 33): négatif','Taux de coupon (col. 33): élevé', 'Taux de coupon (col. 33): très élevé','Test notation crédit (col. 59)', 'Test delta dans [-1,1] (col. 93)',
            'Test Covered (col.55)', 'Test code NACE (col.54)', 'Test code CIC (col.12)', 'Test pays renseigné (col. 52)', 'Ratio VM (col.24) /Nominal (col. 19)', 'Test info type produit (col. 131)','Test info type infrastructure (col.132) : Vide ou Texte',
            'Test info type infrastructure (col.132) : monétaire', 'Test info type infrastructure (col.132) : obligataire','Test info type infrastructure (col.132) : action'
        ]

        # Extract unique values from the last column, excluding unknown security types
        unique_groups = [
            val for val in df.iloc[:, -1].unique().tolist()
            if not str(val).startswith("UNKNOWN security type_")
        ]

        # Ensure 'Modif/Tests' is the first column
        columns = ['Modif/Tests'] + unique_groups

        # Create the DataFrame with empty string values
        self.df_test = pd.DataFrame('', index=row_labels, columns=columns)

        # Fill the 'Modif/Tests' column with the row labels
        self.df_test['Modif/Tests'] = row_labels

        return self.df_test


    def VM(self, df, ecart):
       """
       This function checks whether the Monetary Value (VM) of each fund matches the sum of the VMs of its components.
       If the relative difference exceeds the threshold, the 'net_asset' value is adjusted to match the calculated sum.
       """

       print('Test des VM...')

       # Sécurisation des colonnes numériques
       df[df.columns[25]] = pd.to_numeric(df[df.columns[25]], errors='coerce')
       df['5_Net asset valuation of the portfolio or the share class in portfolio currency'] = pd.to_numeric(df['5_Net asset valuation of the portfolio or the share class in portfolio currency'], errors='coerce')

       df_grouped = df.groupby(df.columns[0])
       wrong_VM_founds = []

       for group_name, group in df_grouped:
           sum_VM = group[df.columns[25]].sum()
           net_asset = group['5_Net asset valuation of the portfolio or the share class in portfolio currency'].iloc[0]

           if pd.isna(sum_VM) or pd.isna(net_asset) or sum_VM == 0:
               continue
            
           relative_diff = abs(sum_VM - net_asset) / sum_VM

           if relative_diff > ecart:
               df.loc[group.index, '5_Net asset valuation of the portfolio or the share class in portfolio currency'] = sum_VM
               col_name = group['Agregated_name_product'].iloc[0]
               self.df_test.loc['Somme VM (col.24) = NAV Annoncée (col.5)', col_name] = 'X'
               wrong_VM_founds.append((group_name, col_name))

       if wrong_VM_founds:
           self.errors2.append(f"( * ) Les fonds suivants ont été ajustés car la somme des VM différait de plus de {ecart * 100} % :")
           self.errors2.append('')
           for item in wrong_VM_founds:
               self.errors2.append(" - " + item[0] + ' situé dans ' + item[1])
           self.errors2.append('')
       else:
           self.errors2.append("( * ) Pas de problème de somme de VM.")
           self.errors2.append('')

       return df



    def maturity(self, df, y):
        """
        This function checks that each maturity date is after December 31st of year y.
        If a maturity date is before the limit, one year is added and the user is notified.
        After that, any missing or invalid maturity date is replaced with January 1st, 1900.
        An 'X' is also added in self.df_test at the row 'Tests maturité produits de taux (col. 39)' and the column
        corresponding to the last column of each affected row.
        """
        
        print('Test des maturités...')
        
        col_date = '39_Maturity date'
        df_modif = df.copy()

        # Step 1: Parse dates without replacing missing values
        def try_parse(val):
            try:
                if isinstance(val, (int, float)):
                    return pd.to_datetime("1899-12-30") + pd.to_timedelta(val, unit="D")
                return parser.parse(str(val), dayfirst=True)
            except Exception:
                return pd.NaT

        df_modif[col_date] = df_modif[col_date].apply(try_parse)

        # Step 2: Correct dates that are too early
        limit = pd.Timestamp(f"{y}-12-31")
        need_to_correct = df_modif[col_date].notna() & (df_modif[col_date] < limit)

        modified_cells = []
        for _, row in df_modif[need_to_correct].iterrows():
            name = row.iloc[-2]
            column_name = row.iloc[-1]
            modified_cells.append(f"- {name} situé dans le fichier {column_name}")
            if 'Tests maturité produits de taux (col. 39)' in self.df_test.index and column_name in self.df_test.columns:
                self.df_test.loc['Tests maturité produits de taux (col. 39)', column_name] = 'X'

        df_modif.loc[need_to_correct, col_date] += pd.DateOffset(years=1)

        # Step 3: Replace remaining missing or invalid dates with 1900-01-01
        df_modif[col_date] = df_modif[col_date].fillna(pd.Timestamp("1900-01-01"))

        # Step 4: User messages
        if modified_cells:
            self.errors2.append("( * ) Les actifs suivants ont eu leur maturité modifiée :")
            self.errors2.append('')
            self.errors2.extend(modified_cells)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Pas de problème de date de maturité.")
            self.errors2.append('')
            self.errors2.append('')

        return df_modif


    def coupon_freq(self, df):
        """
        This method identifies all bonds in the DataFrame using their CIC code and checks
        whether they have a coupon frequency. If not, it sets the frequency to 1,
        reports them to the user, and marks the issue in self.df_test under 'Fréquence de coupon'.
        """
        
        print('Test des fréquences de coupon des obligations...')
        
        df_modif = df.copy()
        CIC = df_modif['12_CIC code of the instrument']
        no_coup = []
        text_coup = []

        for i in range(len(CIC)):
            code = str(CIC.iloc[i])
            # Check if the instrument is a bond based on CIC code
            if len(code) > 2 and (code[2] == '1' or code[2] == '2'):
                freq = df_modif['38_Coupon payment frequency'].iloc[i]
                # Check if coupon frequency is missing (empty string or NaN)
                if pd.isna(freq) or freq == '':
                    # Set missing frequency to 1
                    df_modif.at[i, '38_Coupon payment frequency'] = 1

                    row = df_modif.iloc[i]
                    file_name = row.iloc[-1]
                    no_coup.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test fréquence produits de taux (col. 38): vide', file_name] = 'X'
                    
                elif type(freq) == str:
                    
                    row = df_modif.iloc[i]
                    file_name = row.iloc[-1]
                    text_coup.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test fréquence produits de taux (col. 38): texte', file_name] = 'X'
                    
                     

        # Report missing coupon frequencies
        if no_coup:
            self.errors2.append("( * ) Les obligations suivantes n'avaient pas de fréquence de coupon (valeur 1 attribuée) :")
            self.errors2.append('')
            self.errors2.extend(no_coup)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Toutes les obligations ont une fréquence de coupon.")
            self.errors2.append('')
            self.errors2.append('')
            
        
        if text_coup:
            self.errors2.append("( * ) Les obligations suivantes ont une fréquence de coupon au format texte:")
            self.errors2.append('')
            self.errors2.extend(text_coup)
            self.errors2.append('')
            self.errors2.append('')

        return df_modif

    
    def storage(self, df):
        """
        This method analyzes each instrument in the DataFrame and ensures that,
        for identical instruments, the CIC and NACE codes are consistent.

        It also checks that NACE codes starting with 64, 65, or 66 begin with 'K'.

        If inconsistencies are found, it notifies the user and marks the issues in self.df_test.
        """

        print('Test des codes CIC et codes NACE...')
        
        grouped = df.groupby('17_Instrument name')
        wrong_nace = []
        wrong_cic = []

        for group_name, group in grouped:
            # Check for CIC code consistency using column header
            cic_col = '12_CIC code of the instrument'
            if group[cic_col].nunique() > 1:
                correct_cic = group[cic_col].mode()[0]
                wrong_rows = group[group[cic_col] != correct_cic]
                for _, row in wrong_rows.iterrows():
                    file_name = row.iloc[-1]
                    wrong_cic.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test code CIC (col.12)', file_name] = 'X'

            # Check for NACE code consistency (column 56)
            nace_col_index = 55
            if group.iloc[:, nace_col_index].nunique() > 1:
                correct_nace = group.iloc[:, nace_col_index].mode()[0]
                wrong_rows = group[group.iloc[:, nace_col_index] != correct_nace]
                for _, row in wrong_rows.iterrows():
                    file_name = row.iloc[-1]
                    wrong_nace.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Code NACE', file_name] = 'X'

            # Check if NACE code starts with 64, 65, or 66, it must begin with 'K'
            for _, row in group.iterrows():
                nace_code = str(row.iloc[nace_col_index]).strip()
                if nace_code.startswith(('64', '65', '66')) and not nace_code.startswith('K'):
                    file_name = row.iloc[-1]
                    wrong_nace.append(f"- {row.iloc[-2]} situé dans le fichier {file_name} a un code NACE commençant par {nace_code[:2]} mais ne commence pas par 'K'.")
                    self.df_test.at['Test code NACE (col.54)', file_name] = 'X'

        # Report NACE code issues
        if wrong_nace:
            self.errors2.append("( * ) Les actifs suivants ont un problème au niveau du code NACE :")
            self.errors2.append('')
            self.errors2.extend(wrong_nace)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Le fichier ne contient pas d'erreur de regroupement des produits au niveau du code NACE.")
            self.errors2.append('')
            self.errors2.append('')

        # Report CIC code issues
        if wrong_cic:
            self.errors2.append("( * ) Les actifs suivants ont un problème au niveau du code CIC :")
            self.errors2.append('')
            self.errors2.extend(wrong_cic)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Le fichier ne contient pas d'erreur de regroupement des produits au niveau du code CIC.")
            self.errors2.append('')
            self.errors2.append('')

        return None

    

    def redemp_rate(self, df):
        """
        This function takes a data frame as an entry and analyses the column entitled '41_Redemption rate'.
        
        It makes sure that:
            - None of its cells is empty. If so, we replace the empty cell with 1.
            - None of its cells is in a text format
            - None of its cells is negative or equal to zero
            - If the value is greater than 100, it is divided by 100
        """
        
        print('Test des redemption rates...')

        empt_red = []
        text_red = []
        neg_red = []
        
        for idx, row in df.iterrows():
            if str(row['12_CIC code of the instrument'])[2] == '1' or str(row['12_CIC code of the instrument'])[2] == '2':
                redemption_rate = row['41_Redemption rate']

                # Appliquer la logique demandée
                if pd.isna(redemption_rate) or redemption_rate == '':
                    df.at[idx, '41_Redemption rate'] = 1
                    file_name = row.iloc[-1]
                    empt_red.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test validité redemtion rates produits de taux (col. 41): Vide', file_name] = 'X'
                else:
                    if isinstance(redemption_rate, str):
                        file_name = row.iloc[-1]
                        text_red.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                        self.df_test.at['Test validité redemtion rates produits de taux (col. 41): Text', file_name] = 'X'
                    else:
                        if redemption_rate <= 0:
                            file_name = row.iloc[-1]
                            neg_red.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                            self.df_test.at['Test validité redemtion rates produits de taux (col. 41): Négatif', file_name] = 'X'
                        elif redemption_rate > 100:
                            df.at[idx, '41_Redemption rate'] = redemption_rate / 100

        if text_red:
            self.errors2.append("( * ) Les actifs suivants ont un taux de remboursement du nominal au format texte : ")
            self.errors2.append('')
            self.errors2.extend(text_red)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Le fichier ne contient pas de taux de remboursement du nominal au format texte.")
            self.errors2.append('')
            self.errors2.append('')
            
        if neg_red:
            self.errors2.append("( * ) Les actifs suivants ont un taux de remboursement du nominal négatif : ")
            self.errors2.append('')
            self.errors2.extend(neg_red)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Le fichier ne contient pas de taux de remboursement du nominal négatif.")
            self.errors2.append('')
            self.errors2.append('')
        
        return df


    
    def modified_duration(self, df):
        """
        This function ensures that no row in the column '90_Modified Duration to maturity date' is empty or negative.
        """
        print('Test des durations modifiées')

        target_col = '90_Modified Duration to maturity date'
        penultimate_col = df.columns[-2]
        last_col = df.columns[-1]

        collected_values = []

        # Convert to numeric, coercing errors to NaN
        df[target_col] = pd.to_numeric(df[target_col], errors='coerce')

        # Identify NaN or negative rows
        mask = df[target_col].isna() | (df[target_col] < 0)
        collected_values = df.loc[mask, penultimate_col].tolist()

        for idx in df[mask].index:
            col_name = df.at[idx, last_col]
            if col_name in self.df_test.columns:
                self.df_test.at['Test duration modifiée (col. 90)', col_name] = 'X'

        # Replace NaN by 0 and take absolute value
        df[target_col] = df[target_col].fillna(0).abs()

        if collected_values:
            self.errors2.append("( * ) Les actifs suivants ont une duration modifiée négative ou vide : ")
            self.errors2.append('')
            self.errors2.extend(collected_values)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Le fichier ne contient pas de duration modifiée vide ou négative.")
            self.errors2.append('')
            self.errors2.append('')

        return df


    
    def coupon_rate(self, df, taux_tres_haut, taux_haut):
        """
        This method analyzes every nominal rate and ensures that:
            - No rate is strictly under 0
            - No rate is strictly superior to 15
            - The user can easily identify the products with high rates (between 9 and 15)

        It notifies the user and marks issues in self.df_test under 'Taux de coupon'.
        """
        
        print('Test des taux de coupon...')

        neg_rate = []
        high_rate = []
        very_high_rate = []
        empt_rate = []

        df_cleaned = df.copy()
        index_to_del = []

        for index, row in df.iterrows():
            if str(row['12_CIC code of the instrument'])[2] == '1' or str(row['12_CIC code of the instrument'])[2] == '2':
                value = row.iloc[34]

                if pd.isna(value) or value == "":
                    file_name = row.iloc[-1]
                    empt_rate.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Taux de coupon (col. 33): vide', file_name] = 'X'
                
                else:
                    
                    if pd.isna(value) or type(value) == str:
                        continue
                    
                    file_name = row.iloc[-1]

                    if value < 0:
                        neg_rate.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                        self.df_test.at['Taux de coupon (col. 33): négatif', file_name] = 'X'

                    if value > taux_tres_haut:
                        very_high_rate.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                        self.df_test.at['Taux de coupon (col. 33): très élevé', file_name] = 'X'

                    elif value > taux_haut:
                        high_rate.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                        self.df_test.at['Taux de coupon (col. 33): élevé', file_name] = 'X'

        # Remove rows with negative rates
        df_cleaned.drop(index=index_to_del, inplace=True)
        df_cleaned.reset_index(drop=True, inplace=True)

        # Report empty coupon rates
        if empt_rate:
            self.errors2.append("( * ) Les obligations suivantes ont un taux de coupon vide:")
            self.errors2.append('')
            self.errors2.extend(empt_rate)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Les obligations ont toutes un taux de coupon.")
            self.errors2.append('')
            self.errors2.append('')

        # Report high coupon rates
        if high_rate:
            self.errors2.append(f"( * ) Les actifs suivants ont un taux de coupon assez élevé (> {taux_haut}) et sont à vérifier :")
            self.errors2.append('')
            self.errors2.extend(high_rate)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append(f"( * ) Aucun taux de coupon n'est supérieur à {taux_haut}.")
            self.errors2.append('')
            self.errors2.append('')

        # Report negative coupon rates
        if neg_rate:
            self.errors2.append("( * ) Les actifs suivants ont un taux d'intérêt négatif ou nul et ont été supprimés :")
            self.errors2.append('')
            self.errors2.extend(neg_rate)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Aucun taux d'intérêt nominal n'est négatif ou nul.")
            self.errors2.append('')
            self.errors2.append('')

        # Report very high coupon rates
        if very_high_rate:
            self.errors2.append(f"( * ) Les actifs suivants ont un taux d'intérêt très élevé (>= {taux_tres_haut}) et ont été supprimés :")
            self.errors2.append('')
            self.errors2.extend(very_high_rate)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append(f"( * ) Aucun taux d'intérêt nominal n'est supérieur à {taux_tres_haut}.")
            self.errors2.append('')
            self.errors2.append('')

        return df_cleaned

    def not_cred(self,df):
        """

        This function takes a data frame as an input and anayses the column entitled '59_Credit quality step'.
        
        It makes sure that no cell of this column is in a text format.
        """
        
        print('Test de la colonne 59...')
        
        cred_text = []
        
        for _, row in df.iterrows():
            if type(row['59_Credit quality step']) == str:
                file_name = row.iloc[-1]
                cred_text.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                self.df_test.at['Test notation crédit (col. 59)', file_name] = 'X'
                
        if cred_text:
            self.errors2.append("( * ) Les actifs suivants ont une Notation de crédit au format Texte :")
            self.errors2.append('')
            self.errors2.extend(cred_text)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Aucune notation de crédit n'est au format texte.")
            self.errors2.append('')
            self.errors2.append('')
        
        return None
                
                
    
    def delt(self, df):
        """
        This method analyzes each delta value and ensures it is within the range [-1, 1].

        It notifies the user about any anomalies without modifying the DataFrame.
        """

        print('Test des deltas...')
        
        wrong_deltas = []
        df_cleaned = df

        for _, row in df.iterrows():
            value = row.iloc[94]
            if pd.notna(value) and type(value) != str:
                if value > 1 or value < -1:
                    wrong_deltas.append(f"- {row.iloc[-2]} situé dans le fichier {row.iloc[-1]}")
                    file_name = row.iloc[-1]
                    self.df_test.at['Test delta dans [-1,1] (col. 93)', file_name] = 'X'

        if wrong_deltas:
            self.errors2.append("( * ) Les actifs suivants ont un delta anormal et sont à vérifier:")
            self.errors2.append('')
            self.errors2.extend(wrong_deltas)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append("( * ) Aucun delta n'est supérieur à 1 ou inférieur à -1")
            self.errors2.append('')
            self.errors2.append('')

        return df_cleaned

    
    def ratio_nom_vm(self, df, rat):
        """
        This method checks for high VM/Nominal ratios (>= 5), reports them to the user,
        and marks the issue in self.df_test at the row 'Ratio VM (col.24) /Nominal (col. 19)' and the column
        specified in the last column of each affected row.
        """
        
        print('Test des ratio VM/Nominal...')

        high_ratio = []
        grouped = df.groupby(df.columns[0])  # Group by the first column (e.g., instrument ID)

        for _, group in grouped:
            for _, row in group.iterrows():
                col20 = row[df.columns[20]]  # Nominal
                col25 = row[df.columns[25]]  # VM

                if pd.notnull(col20) and col20 != 0:
                    ratio = col25 / col20
                    if ratio >= rat:
                        file_name = row.iloc[-1]
                        high_ratio.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")

                        # Mark 'X' in self.df_test at the row 'Ratio VM/Nominal' and column file_name
                        if 'Ratio VM (col.24) /Nominal (col. 19)' in self.df_test.index and file_name in self.df_test.columns:
                            self.df_test.loc['Ratio VM (col.24) /Nominal (col. 19)', file_name] = 'X'

        # Report high ratios
        if high_ratio:
            self.errors2.append(f"( * ) Les actifs suivants ont un ratio VM/Nominal élevé (> {rat}) et sont à vérifier :")
            self.errors2.append('')
            self.errors2.extend(high_ratio)
            self.errors2.append('')
            self.errors2.append('')
        else:
            self.errors2.append(f"( * ) Aucun ratio VM/Nominal n'est supérieur à {rat}.")
            self.errors2.append('')
            self.errors2.append('')

        return None

    def covered(self, df):
        """
        This function takes a data frame as input and returns a modified version
        where every state's bond and every loan are coming with a C or NC value.
        """

        print('Test Covered...')

        cov_list = []

        for index, row in df.iterrows():
            cic_code = str(row['12_CIC code of the instrument'])

            # Check CIC code pattern to determine if 'NC' should be added
            if len(cic_code) > 3 and cic_code[2] == '1' and cic_code[3] == '1':
                if pd.isna(row['55_Covered / not covered']) or row['55_Covered / not covered'] == '':
                    df.at[index, '55_Covered / not covered'] = 'NC'
                    file_name = row.iloc[-1]
                    cov_list.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test Covered (col.55)', file_name] = 'X'

            elif len(cic_code) > 2 and cic_code[2] == '8':
                if pd.isna(row['55_Covered / not covered']) or row['55_Covered / not covered'] == '':
                    df.at[index, '55_Covered / not covered'] = 'NC'
                    file_name = row.iloc[-1]
                    cov_list.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test Covered (col.55)', file_name] = 'X'

        # Final step: replace all remaining empty cells in column 55 with 'NC'
        df['55_Covered / not covered'] = df['55_Covered / not covered'].replace('', pd.NA)
        df['55_Covered / not covered'] = df['55_Covered / not covered'].fillna('NC')

        if cov_list:
            self.errors2.append("( * ) Les actifs suivants ont subi un rajout de 'NC' dans la colonne 55:")
            self.errors2.append('')
            self.errors2.extend(cov_list)
            self.errors2.append('')
            self.errors2.append('')

        return df


    
    def empt_country(self,df):
        """

        This function takes a data frame as input and makes sure that every bond is associated with a country.
        """
        
        print('Test des pays...')
        
        empt_list = []
        
        for _, row in df.iterrows():
            if str(row['12_CIC code of the instrument'])[2] == '1' or str(row['12_CIC code of the instrument'])[2] == '2':
                if pd.isna(row['52_Issuer country']) or row['52_Issuer country'] == '':
                    file_name = row.iloc[-1]
                    empt_list.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test Pays Renseigné (col. 52)', file_name] = 'X'
        
        if empt_list:
            self.errors2.append("( * ) Les obligations suivantes n'ont pas de pays renseigné à la colonne 52, veuillez en ajouter:")
            self.errors2.append('')
            self.errors2.extend(empt_list)
            self.errors2.append('')
            self.errors2.append('')
        
        return None

   
    def empt_under(self,df):
        """

        This function takes a data frame as input and makes sure that every bond is associated with a country.
        """
        
        print('Test des Underlying assets...')
        
        empt_131 = []
        
        for _, row in df.iterrows():
            if pd.isna(row['131_Underlying Asset Category']) or row['131_Underlying Asset Category'] == '':
                file_name = row.iloc[-1]
                empt_131.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                self.df_test.at['Test info type produit (col. 131)', file_name] = 'X'
        
        if empt_131:
            self.errors2.append("( * ) Les actifs suivants n'ont pas de catégorie d'asset renseignée à la colonne 131, veuillez en ajouter:")
            self.errors2.append('')
            self.errors2.extend(empt_131)
            self.errors2.append('')
            self.errors2.append('')
        
        return None
    
    def infra(self,df):
        """
        
        This function is making sure that both columns 131 and 132 are working together and that no errors are noticed between them.
        
        """
        
        print('Test des info type infra...')
        
        col131 = '131_Underlying Asset Category'
        col132 = '132_Infrastructure_investment'
        
        cash = []
        bond = []
        action = []
        
        for _,row in df.iterrows():
            
            if pd.isna(row[col132]) or row[col132] == '' or type(row[col132]) == str:
                row[col132] = 0
                file_name = row.iloc[-1]
                self.df_test.at['Test info type infrastructure (col.132) : Vide ou Texte', file_name] = 'X'
            
            #If a title is registered as cash but was put as an infra, we notice and correct it
            if row[col131] == 7 or row[col131] == '7':
                if row[col132] in [1,2,3,4,'1','2','3','4']:
                    file_name = row.iloc[-1]
                    cash.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test info type infrastructure (col.132) : monétaire', file_name] = 'X'
                    row[col132] = 5
         
            
            #We check if no bond is registered as as action as type product
            if row[col132] == 1 or row[col132] == '1' or row[col132] == 3 or row[col132] == '3':
                if row[col131] == 'A':
                    file_name = row.iloc[-1]
                    bond.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test info type infrastructure (col.132) : obligataire', file_name] = 'X'
            
                    
            #We check if no action is registered as a bond as type product
            if row[col132] == 2 or row[col132] == '2' or row[col132] == 4 or row[col132] == '4':
                if row[col131] == 'A':
                    file_name = row.iloc[-1]
                    action.append(f"- {row.iloc[-2]} situé dans le fichier {file_name}")
                    self.df_test.at['Test info type infrastructure (col.132) : action', file_name] = 'X'
                    
        if cash:
            self.errors2.append("( * ) Les actifs suivants étaient inscrits en tant que monnaie mais renseignés infra :")
            self.errors2.append('')
            self.errors2.extend(cash)
            self.errors2.append('')
            self.errors2.append('')
            
        
        if bond:
            self.errors2.append("( * ) Les obligations suivantes étaient enregistrés en tant que produit de type action :")
            self.errors2.append('')
            self.errors2.extend(bond)
            self.errors2.append('')
            self.errors2.append('')
            
            
        if action:
            self.errors2.append("( * ) Les actions suivantes étaient enregistrées en tant que produit de type obligation : ")
            self.errors2.append('')
            self.errors2.extend(action)
            self.errors2.append('')
            self.errors2.append('')
                    
            
        
        return None
    
    def empty_port_id(self,df):
       """
       Identifie les lignes dont la colonne '1_Portfolio identifying data' est vide,
       les regroupe par la dernière colonne, retourne la liste des groupes,
       puis supprime ces lignes du DataFrame et retourne aussi le DataFrame modifié.
       """
       print('Test de la colonne 1...')
       # Étape 1 : filtrer les lignes avec identifiant de portefeuille vide
       mask_empty = df['1_Portfolio identifying data'].isna() | (df['1_Portfolio identifying data'] == '')
       df_empty = df[mask_empty]

       # Étape 2 : identifier la dernière colonne
       last_col = df.columns[-1]
   
       # Étape 3 : regrouper par la dernière colonne et collecter les noms de groupes
       group_names = df_empty[last_col].dropna().unique().tolist()

       # Étape 4 : supprimer les lignes avec identifiant vide du DataFrame original
       df_cleaned = df[~mask_empty].copy()
       
       if group_names:
            self.errors2.append("( * ) IL FAUT VERIFIER LES IDENTIFIANTS DANS LES TPT SUIVANTS (VIDE):")
            self.errors2.append('')
            self.errors2.extend(group_names)
            self.errors2.append('')
            self.errors2.append('')

       return df_cleaned

          
            
    