import pandas as pd
import warnings
import os
import zipfile


pd.set_option('future.no_silent_downcasting', True)

import os

class SheetsManip:
    
    def __init__(self):
        self.errors = []
        self.df_col = pd.DataFrame({'Modif/Tests': ['Colonne(s) manquante(s)']})

    
    def export_txt(self,list, folder_path, name):
        os.makedirs(folder_path, exist_ok = True)
        
        file_path = os.path.join(folder_path, name)
        
        # Export TXT
        with open(file_path, "w") as f:
            for e in list:
                f.write(e + "\n")
    
        return None

    def conv_from_excel(self, path):
        """
        This function loads an Excel file and returns a cleaned DataFrame.
        It automatically removes any empty first row or column that could interfere with processing.
        It uses the appropriate engine for .xlsx and .xls files.
        Only one valid sheet is selected based on column count and data density.
        """

        # List of the normalized name of the columns
        L = ['1_Portfolio identifying data', '2_Type of identification code for the fund share or portfolio', '3_Portfolio name', '4_Portfolio currency ( B )',     
        '5_Net asset valuation of the portfolio or the share class in portfolio currency', '6_Valuation date', '7_Reporting date', '8_Share price', 
        '8b_Total number of shares', '9_' +'%'+' cash', '10_Portfolio Modified Duration', '11_Complete SCR Delivery', '12_CIC code of the instrument', 
        '13_Economic zone of the quotation place', '14_Identification code of the financial instrument', '15_Type of identification code for the instrument', 
        '16_Grouping code for multiple leg instruments', '17_Instrument name', '17b_Asset / Liability', '18_Quantity', '19_Nominal amount','20_Contract size for derivatives', 
        '21_Quotation currency (A)', '22_Market valuation in quotation currency  ( A )', '23_Clean market valuation in quotation currency (A)',
        '24_Market valuation in portfolio currency  (B)', '25_Clean market valuation in portfolio currency (B)','26_Valuation weight','27_Market exposure amount in quotation currency (A)',
        '28_Market exposure amount in portfolio currency (B)','29_Market exposure amount for the 3rd currency in quotation currency of the underlying asset ( C )',
        '30_Market Exposure in weight','31_Market exposure for the 3rd currency in weight over NAV', '32_Interest rate type', '33_Coupon rate',
        '34_Interest rate reference identification','35_Identification type for interest rate index', '36_Interest rate index name', '37_Interest rate Margin',
        '38_Coupon payment frequency','39_Maturity date','40_Redemption type','41_Redemption rate','42_Callable / putable',	'43_Call / put date',
        '44_Issuer / bearer option exercise','45_Strike price for embedded (call/put) options',	'46_Issuer name','47_Issuer identification code','48_Type of identification code for issuer',
        '49_Name of the group of the issuer','50_Identification of the group','51_Type of identification code for issuer group','52_Issuer country',
        '53_Issuer economic area','54_Economic sector','55_Covered / not covered','56_Securitisation','57_Explicit guarantee by the country of issue',
        '58_Subordinated debt','58b_Nature of the TRANCHE','59_Credit quality step','60_Call / Put / Cap / Floor','61_Strike price',
        '62_Conversion factor (convertibles)/ concordance factor  / parity (options)', '63_Effective Date of Instrument','64_Exercise type',
        '65_Hedging Rolling','67_CIC code of the underlying asset','68_Identification code of the underlying asset','69_Type of identification code for the underlying asset',
        '70_Name of the underlying asset','71_Quotation currency of the underlying asset ( C )','72_Last valuation price of the underlying asset',
        '73_Country of quotation of the underlying asset','74_Economic Area of quotation of the underlying asset',
        '75_Coupon rate of the underlying asset','76_Coupon payment frequency of the underlying asset',	'77_Maturity date of the underlying asset',
        '78_Redemption profile of the underlying asset','79_Redemption rate of the underlying asse','80_Issuer name of the underlying asset',
        '81_Issuer identification code of the underlying asset','82_Type of issuer identification code of the underlying asset',
        '83_Name of the group of the issuer of the underlying asset','84_Identification of the group of the underlying asset',
        '85_Type of the group identification code of the underlying asset','86_Issuer country of the underlying asset',
        '87_Issuer economic area of the underlying asset','88_Explicit guarantee by the country of issue of the underlying asset','89_Credit quality step of the underlying asset',
        '90_Modified Duration to maturity date','91_Modified duration to next option exercise date','92_Credit sensitivity',
        '93_Sensitivity to underlying asset price (delta)','94_Convexity / gamma for derivatives','94b_Vega',
        '95_Identification of the original portfolio for positions embedded in a fund','97_SCR_Mrkt_IR_up weight over NAV','98_SCR_Mrkt_IR_down weight over NAV','99_SCR_Mrkt_Eq_type1 weight over NAV',
        '100_SCR_Mrkt_Eq_type2 weight over NAV','101_SCR_Mrkt_Prop weight over NAV','102_SCR_Mrkt_Spread_bonds weight over NAV',
        '103_SCR_Mrkt_Spread_structured weight over NAV','104_SCR_Mrkt_Spread_derivatives_up weight over NAV',
        '105_SCR_Mrkt_Spread_derivatives_down weight over NAV','105a_SCR_Mrkt_FX_up weight over NAV', '105b_SCR_Mrkt_FX_down weight over NAV',
        '106_Asset pledged as collateral','107_Place of deposit','108_Participation','110_Valorisation method','111_Value of acquisition',
        '112_Credit rating','113_Rating agency','114_Issuer economic area','115_Fund Issuer Code','116_Fund Issuer Code Type',
        '117_Fund Issuer Name','118_Fund Issuer Sector','119_Fund Issuer Group Code','120_Fund Issuer Group Code Type',
        '121_Fund Issuer Group name','122_Fund Issuer Country','123_Fund CIC code', '123a_Fund Custodian Country','124_Duration',
        '125_Accrued Income (Security Denominated Currency)','126_Accrued Income (Portfolio Denominated Currency)',
        '127_Bond Floor (convertible instrument only)','128_Option premium (convertible instrument only)','129_Valuation Yield',
        '130_Valuation Z-spread','131_Underlying Asset Category','132_Infrastructure_investment','133_custodian_name',	
        '134_type1_private_equity_portfolio_eligibility','135_type1_private_equity_issuer_beta','137_counterparty_sector',	
        '138_Collateral_eligibility','139_Collateral_Market_valuation_in_portfolio_currency','TPTVersion']  # Replace with actual list of column names

        print('Traitement de ' + path)

        # Detect file extension and select appropriate engine
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            raise ValueError("Le fichier que vous tentez de convertir est au format csv, veuillez le convertir en .xls ou .xlsx")
        elif ext == ".xlsb":
            raise ValueError("Le fichier est au format .xlsb. Veuillez le convertir au format .xls avant de le traiter.")
        elif ext == ".xlsx":
            engine = "openpyxl"
        elif ext == ".xls":
            try:
                with zipfile.ZipFile(path, 'r') as z:
                    engine = "openpyxl"  # .xlsx with a wrong format
            except zipfile.BadZipFile:
                try:
                    import xlrd
                    engine = "xlrd"  # True .xls
                except ImportError:
                    raise ImportError("Le fichier est au format .xls mais le module 'xlrd' n'est pas installé.")
        else:
            raise ValueError("Format non pris en charge, veuillez utiliser un format .csv, .xlsx ou .xls")

        # Read all sheets
        sheets = pd.read_excel(path, sheet_name=None, header=None, engine=engine)

        # Select the best sheet based on column count and data density
        best_sheet = None
        best_score = 0

        for name, content in sheets.items():
            if name.lower() == "disclamer" or content.empty:
                continue

            # Remove empty first row and column if necessary
            if content.shape[0] > 0 and content.iloc[0].isna().all():
                content = content.iloc[1:].reset_index(drop=True)
            if content.shape[1] > 0 and content.iloc[:, 0].isna().all():
                content = content.iloc[:, 1:].reset_index(drop=True)

            # Check if sheet has enough columns
            if content.shape[1] > 120:
                # Score based on non-empty cells in second row (first data row)
                non_empty_ratio = content.iloc[1].notna().sum() / content.shape[1]
                score = non_empty_ratio

                if score > best_score:
                    best_score = score
                    best_sheet = (name, content)

        if best_sheet is None:
            raise ValueError(f"Le fichier {path} ne contient pas de feuille respectant le format du TPT (au moins 120 colonnes).")

        # Process the selected sheet
        sheet_name, df_sheet = best_sheet
        df_sheet = df_sheet.reset_index(drop=True)
        df_sheet.columns = df_sheet.iloc[0]
        df_sheet = df_sheet[1:].reset_index(drop=True)
        df_sheet = self.normalized(path)[0]
        df_sheet.columns = L
        df_sheet = self.agregated_name(df_sheet)
        df_sheet.insert(loc=len(df_sheet.columns), column="File's name", value=[path] * len(df_sheet))

        print('Traitement terminé')
        return df_sheet




    def concat(self,list_path):
        """
        This function can be used in order to concatenate all the TPTs into one single sheet.

        list_path: object = list, contains the path of all the files that you want to add in your TPT.
        """
        df = self.conv_from_excel(list_path[0])
        
        if len(list_path) == 1:
            return df
        
        for i in range(1,len(list_path)):
            dfb = self.conv_from_excel(list_path[i])

            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=FutureWarning)
                df = pd.concat([df, dfb], ignore_index=True)

        
        return df
    
    def agregated_name(self, df):
        """
        This function will:
            - Add the column 'Agregated_name_product' if it does not exist yet.
            - Complete this column if it already exists.

        It uses the columns:
            - '1_Portfolio identifying data'
            - '14_Identification code of the financial instrument'
        to generate a unique identifier.
        """

        col_portfolio = df['1_Portfolio identifying data'].fillna('').infer_objects(copy=False).astype(str)
        col_identifier = df['14_Identification code of the financial instrument'].fillna('').infer_objects(copy=False).astype(str)
        fusion = col_identifier + "_" + col_portfolio

        if 'Agregated_name_product' in df.columns:
            masque_vides = df['Agregated_name_product'].isna() | (df['Agregated_name_product'] == '')
            df.loc[masque_vides, 'Agregated_name_product'] = fusion[masque_vides]
        else:
            df['Agregated_name_product'] = fusion

        self.dfg = df
        return df


    def normalized(self, path):
        """
        Normalize a TPT DataFrame by:
        - Detecting the first sheet with at least 120 columns.
        - Automatically detecting the correct header row.
        - Removing empty unnamed columns.
        - Renaming partially filled unnamed columns.
        - Ensuring all required columns (based on expected prefixes) are present.
        """

        Miss = False

        # Étape 0 : Détection de la première feuille avec au moins 120 colonnes
        try:
            xls = pd.ExcelFile(path, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"Erreur lors de l'ouverture du fichier Excel : {e}")

        target_sheet = None
        header_row_index = None

        for sheet_name in xls.sheet_names:
            try:
                df_check = pd.read_excel(path, sheet_name=sheet_name, header=None, engine='openpyxl')
                if df_check.shape[1] >= 120:
                    for i, row in df_check.iterrows():
                        if any(str(cell).startswith(('1_', '1.', '1')) for cell in row if isinstance(cell, str)):
                            target_sheet = sheet_name
                            header_row_index = i
                            break
                    if target_sheet:
                        break
            except Exception:
                continue

        if target_sheet is None or header_row_index is None:
            raise ValueError("Impossible de détecter automatiquement la feuille et la ligne d'en-tête.")

        try:
            df_ref = pd.read_excel(path, sheet_name=target_sheet, header=header_row_index, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"Erreur lors de la lecture de la feuille de référence : {e}")

        ref_columns = [str(col).strip() for col in df_ref.columns]

        # Étape 1.1 : Concaténation des feuilles avec les mêmes en-têtes
        dfs_to_concat = []
        for sheet_name in xls.sheet_names:
            try:
                df_check = pd.read_excel(path, sheet_name=sheet_name, header=None, engine='openpyxl')
                if df_check.shape[0] <= header_row_index:
                    continue
                df_candidate = pd.read_excel(path, sheet_name=sheet_name, header=header_row_index, engine='openpyxl')
                candidate_columns = [str(col).strip() for col in df_candidate.columns]
                if candidate_columns == ref_columns:
                    dfs_to_concat.append(df_candidate)
            except Exception:
                continue

        if not dfs_to_concat:
            raise ValueError("Aucune feuille avec des en-têtes compatibles n'a été trouvée.")

        df = pd.concat(dfs_to_concat, ignore_index=True)

        # Étape 1.5 : Gestion des en-têtes numériques et suffixes 'Portfolio'
        first_header = df.columns[0]
        if isinstance(first_header, (int, float)):
            next_row = df.iloc[0]
            if 'Portfolio' in str(next_row.iloc[0]) or 'portfolio' in str(next_row.iloc[0]):
                new_columns = []
                for col_idx, col in enumerate(df.columns):
                    header_str = str(col)
                    suffix = str(next_row.iloc[col_idx]).strip()
                    new_columns.append(f"{header_str}_{suffix}" if suffix else header_str)
                df.columns = new_columns
                df = df.iloc[1:]
            else:
                df.columns = [str(col) for col in df.columns]
        else:
            df.columns = [str(col) for col in df.columns]

        # Remplacement des préfixes avec points
        df.columns = [
            col.replace('8.b', '8b')
            .replace('17.b', '17b')
            .replace('58.b', '58b')
            .replace('94.b', '94b')
            .replace('105.a', '105a')
            .replace('105. a', '105a')
            .replace('105.b', '105b')
            .replace('123.a', '123a')
            .replace('25_Clean market valuation in quotation currency (A)','23_Clean market valuation in quotation currency (A)')
            if isinstance(col, str) else col
            for col in df.columns
        ]

        # Étape 2 : Nettoyage des colonnes Unnamed
        cleaned_columns = []
        for i, col in enumerate(df.columns):
            if col in ["NUM_DATA", "Num_Data"]:
                df.drop(columns=col, inplace=True)
                continue
            if isinstance(col, str) and col.startswith("Unnamed"):
                col_data = df.iloc[:, i]
                if col_data.replace('', pd.NA).isna().all():
                    df.drop(columns=col, inplace=True)
                    continue
                else:
                    cleaned_columns.append(f"{i}_Recovered")
            else:
                cleaned_columns.append(col)

        if len(cleaned_columns) != df.shape[1]:
            raise ValueError("Mismatch entre le nombre de colonnes et les noms nettoyés.")
        df.columns = cleaned_columns

        # Étape 3 : Vérification des colonnes sans nom
        for i, col in enumerate(df.columns):
            if col == '' or pd.isna(col):
                col_data = df.iloc[:, i]
                if col_data.replace('', pd.NA).isna().all():
                    df.drop(columns=col, inplace=True)
                else:
                    prev_index = i - 1
                    raise ValueError(
                        f"Le fichier {path} contient une colonne sans nom à l'indice {i} (après la colonne {prev_index}) "
                        "mais avec des valeurs non vides. Veuillez corriger cette colonne et relancer la transparisation."
                    )

        # Étape 4 : Normalisation des colonnes attendues
        expected_prefixes = [
            '1', '2', '3', '4', '5', '6', '7', '8', '8b', '9', '10', '11', '12', '13', '14', '15', '16',
            '17', '17b', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
            '31', '32', '33', '34', '35', '36', '37', '38', '39', '40', '41', '42', '43', '44', '45',
            '46', '47', '48', '49', '50', '51', '52', '53', '54', '55', '56', '57', '58', '58b', '59',
            '60', '61', '63', '64', '65', '67', '68', '69', '70', '71', '72', '73', '74', '75',
            '76', '77', '78', '79', '80', '81', '82', '83', '84', '85', '86', '87', '88', '89', '90',
            '91', '92','93', '94', '94b', '95', '97', '98', '99', '100', '101', '102', '103', '104',
            '105', '105a', '105b', '106', '107', '108', '110', '111', '112', '113', '114', '115',
            '116', '117', '118', '119', '120', '121', '122', '123', '123a', '124', '125', '126',
            '127', '128', '129', '130', '131', '132', '133', '134', '135', '137', '138', '139',
            'TPTVersion'
        ]

        special_names = ['3L', '3X', 0, '0', 'A', 'B', 'C', 'D', 'E', 'F', 'L']
        df.rename(columns={col: '131_Underlying Asset Category' for col in df.columns if col in special_names}, inplace=True)

        actual_prefixes = []
        col_names = list(df.columns)

        for col in col_names:
            if type(col) != str:
                col = str(int(col)) if isinstance(col, float) else str(col)

            if col in ['TPTVersion', '1000_TPTVersion', '1000_TPT_Version', '1000_TPT Version', '1000' '1000_TPT_Version	', 1000] or 'version' in col.lower():
                actual_prefixes.append('TPTVersion')
            else:
                prefix = ''
                started = False
                for char in col:
                    if not started:
                        if char != ' ':
                            started = True
                            if char not in ['_', ' ']:
                                prefix += char
                    else:
                        if char in ['_', ' ']:
                            break
                        prefix += char
                actual_prefixes.append(prefix)

        missing_columns = []

        for i, expected_prefix in enumerate(expected_prefixes):
            if expected_prefix not in actual_prefixes:
                col_name = 'TPTVersion' if expected_prefix == 'TPTVersion' else f'{expected_prefix}_Missing'
                df.insert(loc=i+1, column=col_name, value='')
                missing_columns.append(col_name)

        if missing_columns:
            if '( * ) Les colonnes suivantes ont été ajoutées dans {path} :' not in self.errors:
                self.errors.append(f'( * ) Les colonnes suivantes ont été ajoutées dans {path} :')
                for col in missing_columns:
                    self.errors.append(f'   - {col}')
                self.errors.append('')
                self.errors.append('')
                self.df_col[path] = 'X'
        else:
            self.df_col[path] = ''

        return df, Miss


    def convert_to_excel(self,df,path,name):
        """
        This function converts the data frame df into an excel file registered is a file located by path.
        
        The format of the excel will be a,xlsx.
        
        The name of the file will be the input name.
        """
        if not name.endswith('.xlsx'): #If the name you've put as an input does not finish with .xlsx, we add it
            name += '.xlsx'
        
        full_path = os.path.join(path, name)
        
        if not os.path.exists(path):
            raise FileNotFoundError(f"Le dossier spécifié n'existe pas: {path}")
        
        df.to_excel(full_path, index= False)

        print(f"Fichier Excel sauvegardé à : {full_path}")
        return None
    

    def add_and_reject(self, df, path):
        """
        Compares two portfolios (previous and current) to identify:
        - Assets that were removed (present before, missing now)
        - Assets that were added (missing before, present now)

        Args:
            df (DataFrame): Data from the previous year's portfolio (transparisation)
            path (str): Path to the Excel file containing the current portfolio

        Returns:
            tuple: (df_add, df_del)
                - df_add: newly added assets
                - df_del: removed assets
        """

        # Read the current portfolio from the Excel file (starting from row 14)
        dfb = pd.read_excel(path, header=13)

        # Group both datasets by their asset identifiers
        df_grouped = df.groupby('1_Portfolio identifying data')
        dfb_grouped = dfb.groupby('Code Valeur')

        # Initialize output DataFrames for added and removed assets
        df_add = pd.DataFrame(columns=['Code Valeur', 'Libellé Valeur', 'Encours'])
        df_del = pd.DataFrame(columns=['Code Valeur', 'Libellé Valeur', 'Encours'])

        print('Analyse des données à retirer...')
        # Identify removed assets: present in previous year but not in current stock
        for group_name, group in df_grouped:
            if group_name not in dfb['Code Valeur'].values:
                df_del.loc[len(df_del)] = {
                    'Code Valeur': group_name,
                    'Libellé Valeur': group['3_Portfolio name'].iloc[0],
                    'Encours': group['5_Net asset valuation of the portfolio or the share class in portfolio currency'].iloc[0]
                }

        print('Analyse des données à rajouter...')
        # Identify added assets: present in current stock but not in previous year
        for group_name, group in dfb_grouped:
            if group_name not in df['1_Portfolio identifying data'].values and group['Valeur de réalisation'].sum() > 1000 and 'PR' not in group['Portefeuille']:
                df_add.loc[len(df_add)] = {
                    'Code Valeur': group_name,
                    'Libellé Valeur': group['Libellé Valeur'].iloc[0],
                    'Encours': group['Valeur de réalisation'].iloc[0]
                }


        return df_add, df_del


    def transpa_type(self,df):
        # Référentiels codés en dur
        ref_A_C = { "A": "monétaire","B": "monétaire","C": "monétaire","D": "monétaire","E": "monétaire","F": "monétaire",
            "5": "alternatifs", "6": "alternatifs", "7": "monétaire", "8": "alternatifs", "9": "immobilier"
        }
        ref_G_I = {
            "41": "alternatifs", "42": "alternatifs","43": "monétaire", "44": "alternatifs", "45": "immobilier", "46": "alternatifs",
            "47": "alternatifs", "48": "alternatifs", "49": "alternatifs"
        }
        ref_K_M = {
            "0": "alternatifs", "1": "actions", "2": "actions", "3": "alternatifs"
        }
        ref_P1_Q1 = {
            "fixed": "3", "floating": "4", "variable": "4", "inflation_linked": "4"
        }
        ref_P2_Q2 = {
            "fixed": "1", "floating": "2", "variable": "2", "inflation_linked": "2"
        }
        
        result = {}

        for _, row in df.iterrows():
            key = row.get('Agregated_name_product')
            cic_code = str(row.get('12_CIC code of the instrument', ''))
            econ_zone_raw = str(row.get('13_Economic zone of the quotation place', '')).lower()
            rate_type = str(row.get('32_Interest rate type', '')).lower()
            isin = str(row.get('14_Identification code of the financial instrument', '')).lower()

            third_char = cic_code[2] if len(cic_code) >= 3 else ''
            last_two = cic_code[-2:] if len(cic_code) >= 2 else ''

            # Nettoyage de la zone économique
            try:
                econ_zone = str(int(float(econ_zone_raw)))
            except ValueError:
                econ_zone = None

            try:
                if third_char == "8":
                    value = 1
                elif third_char > "4":
                    value = ref_A_C.get(third_char)
                elif third_char == "4":
                    value = ref_G_I.get(last_two)
                elif third_char == "3":
                    value = ref_K_M.get(econ_zone)
                elif last_two == "22":
                    value = ref_P1_Q1.get(rate_type)
                else:
                    value = ref_P2_Q2.get(rate_type)
            except Exception:
                value = None

            result[key] = value

        return result


    def input_SAS(self, df):
        
        print("Création de l'input SAS...")
        
        # Step 1: Create a copy of the input DataFrame
        df_copy = df.copy()

        # Step 2: Define the output DataFrame with the required columns
        columns = [
            "Agregated_ISIN", "ID_Titre", "Lib_Titre", "ID_Fund", "NAV_Fund",
            "Value_Market", "Nominal", "Maturity_Date", "Redemption_Rate",
            "Coupon_Rate", "Frequency_Asset", "Modified_Duration", "Devise",
            "Delta", "Transpa_Type", "Covered", "Govies", "Infrastructure_Type",
            "IssuerGroup_Code", "Code_CIC"
        ]
        output_df = pd.DataFrame(columns=columns)

        # Step 3: Populate the output DataFrame with mapped or default values
        output_df["Agregated_ISIN"] = df_copy.get("Agregated_name_product", "")
        output_df["ID_Titre"] = df_copy.get("14_Identification code of the financial instrument", "")
        output_df["Lib_Titre"] = df_copy.get("17_Instrument name", "")
        output_df["ID_Fund"] = df_copy.get("1_Portfolio identifying data", "")
        output_df["NAV_Fund"] = df_copy.get("5_Net asset valuation of the portfolio or the share class in portfolio currency", "")

        # Fill missing maturity dates with a default value
        maturity_series = df_copy.get("39_Maturity date", "")
        output_df["Maturity_Date"] = maturity_series.fillna("01/01/1900")

        # Convert numeric columns with error handling
        output_df["Redemption_Rate"] = pd.to_numeric(df_copy.get("41_Redemption rate", ""), errors='coerce')
        output_df["Coupon_Rate"] = pd.to_numeric(df_copy.get("33_Coupon rate", ""), errors='coerce')

        # Handle frequency: replace 0 with 1 if applicable
        freq_series = df_copy.get("38_Coupon payment frequency", "")
        freq_cleaned = freq_series.apply(lambda x: 1 if pd.notnull(x) and str(x).strip() != '' and float(x) == 0 else x)
        output_df["Frequency_Asset"] = pd.to_numeric(freq_cleaned, errors='coerce')

        output_df["Modified_Duration"] = df_copy.get("90_Modified Duration to maturity date", "")
        output_df["Devise"] = df_copy.get("21_Quotation currency (A)", "")

        # Replace missing or empty Delta values with 0
        delta_series = df_copy.get("93_Sensitivity to underlying asset price (delta)", "")
        output_df["Delta"] = delta_series.apply(lambda x: 0 if pd.isna(x) or str(x).strip() == '' else x)

        output_df["Covered"] = df_copy.get("55_Covered / not covered", "")
        output_df["Govies"] = 0

        # Convert infrastructure type to integer, replacing missing values with 0
        infra_series = df_copy.get("132_Infrastructure_investment", 0).fillna(0)
        output_df["Infrastructure_Type"] = pd.to_numeric(infra_series, errors='coerce').fillna(0).astype('int64')

        # Clean issuer group code: replace '0' with empty string
        issuer_series = df_copy.get("47_Issuer identification code", "")
        output_df["IssuerGroup_Code"] = issuer_series.apply(lambda x: "" if pd.notnull(x) and str(x).strip() == '0' else x)

        output_df["Code_CIC"] = df_copy.get('12_CIC code of the instrument', "")

        # Map Transpa_Type using a dictionary
        transpa_dict = self.transpa_type(df)
        output_df["Transpa_Type"] = output_df["Agregated_ISIN"].map(transpa_dict)

        # Convert object columns to string while preserving blanks
        for col in ["Agregated_ISIN", "ID_Titre", "Lib_Titre", "ID_Fund", "Maturity_Date",
                    "Devise", "Transpa_Type", "Covered", "IssuerGroup_Code", "Code_CIC"]:
            output_df[col] = output_df[col].apply(lambda x: str(x) if pd.notnull(x) else "")

        # Step 4: Directly assign 'Nominal' and 'Value_Market' from df_copy
        output_df["Nominal"] = pd.to_numeric(df_copy.get("19_Nominal amount", 0), errors='coerce').fillna(0)
        output_df["Value_Market"] = pd.to_numeric(df_copy.get("24_Market valuation in portfolio currency  (B)", 0), errors='coerce').fillna(0)

        # Step 5: Final cleaning
        output_df["Covered"] = output_df["Covered"].replace("", "NC")
        output_df = output_df[output_df["Value_Market"].notna()]
        output_df["ID_Titre"] = output_df["ID_Titre"].apply(lambda x: "TITRE A COMPLETER" if pd.isna(x) or str(x).strip() == "" else x)
        output_df["Devise"] = output_df["Devise"].apply(lambda x: "EUR" if pd.isna(x) or str(x).strip() == "" else x)
        output_df = output_df.drop_duplicates(subset="Agregated_ISIN")


        return output_df


    def conv_date(self,df):
       """
       Convertit les colonnes de dates au format texte 'dd/mm/yyyy' et remplace les '01/01/1900' dans la colonne de maturité.
       Colonnes ciblées : '6_Valuation date', '7_Reporting date', '39_Maturity date'
       """
       date_cols = ['6_Valuation date', '7_Reporting date', '39_Maturity date']

       for col in date_cols:
           if col in df.columns:
               df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')

        # Remplacement rapide des '01/01/1900' dans la colonne de maturité
       if '39_Maturity date' in df.columns:
        df.loc[df['39_Maturity date'] == '01/01/1900', '39_Maturity date'] = '01/01/1900'

        return df

        
    
    def list_of_files(self,folder_path):
        """
        This function takes as an input the path to the folder that contains every files that you want to want to transparize and converts all the elements of the folder into a list of files.
        You can decide to work without this function but if the number of files that you want to analyze is huge, it will be easier for you to use this function.
        """
        return os.listdir(folder_path)
    