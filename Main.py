"""
This file will be used to call the different functions. You can refer to the notice if you want a better explanation on how this folder works.  
"""

from Sheets_Manip import *
from Transpa_Tests import *

import time

#CAREFUL: when you use the path of the file, you need to watch for any backslash \ because you will have to double it \\

SM = SheetsManip()
TR = Transpa()

def transparisation(L,y,exit,name):
    """
    This function automatically calls every steps of the transparisation process at the right time:
    
    L : List of all the TPT you want to analyze.
    y : Year of the transparisation.
    exit : Path to the folder where you want to save the Transparisation when it is finished. 
    name: Name of the Transparisation file that the user wants
    """
    
    print('Début de la transparisation.')
    
    #We make sure that the lists of the modifications and errors are resseted
    SM.errors = []
    TR.errors2 = []
    
    #We concatenate and normalize the TPTs
    df = SM.concat(L)
    
    #We make sure that the data frame used to sum up every errors and modifications is built correctly
    TR.df_test = TR.col_df2(df)
    
    #We point out the assets where | sum_VM - net_asset | > 500 and delete them from the data frame.
    df = TR.VM(df)
    
    #We make sure that, for two same instruments, the CIC code and the NACE code are the same.
    TR.storage(df)
    
    #We verify that no maturity is younger than the 31/12/y.
    df = TR.maturity(df,y)
    
    #We make sure that every obligations has a coupon frequency.
    TR.coupon_freq(df)
    
    #We point out every unusual nominal rate for the obligations.
    df = TR.redemp_rate(df)
    
    #Same for the level of the coupon rates
    df = TR.coupon_rate(df)
    
    #Same for the deltas
    df = TR.delt(df)
    
    #Same for the ratio VM/Nominal
    TR.ratio_nom_vm(df)
    
    #Makes sure that no notation credit is in a text format
    TR.not_cred(df)
    
    #Makes sure that every state's bonds and every loan is associated with a C or NC value
    df = TR.covered(df)
    
    #Makes sure that every bond have a country
    TR.empt_country(df)
    
    #Makes sure that every asset have an Underlying asset category
    TR.empt_under(df)
    
    #We make many tests on infrastructure type product
    TR.infra(df)
    
    #Because the final transparisation file does not contain the path of the files, we delete this column
    df = df.drop("File's name", axis = 1)
    
    #We then clean every rows that aren't filled with something
    df = df.dropna(how = 'all')
    
    #We concatenate both data frames that we created to summarize the modifications and errors
    df_sum = pd.concat([SM.df_col,TR.df_test], ignore_index=True)
    df_sum = df_sum.T
    df_sum = df_sum.iloc[:, :-1]
    
    for _,row in df_sum.iterrows():
        row[df_sum.columns[0]].replace(nom_du_dossier,"")

    #To make sure that we'll have no empty rows, we use the following instruction
    df = df[df.count(axis=1) >= 10]
    
    #Now that we verified everthing, we export the data frames into the folder mentionned in the input
    """
    #Previous method used to convert data frames into a single file: we now convert all of it into a file with many sheets
    SM.convert_to_excel(df,exit,name)
    SM.convert_to_excel(df_sum,exit,'Résumé des modifications et erreurs')
    """
    
    #We are also using the state of the stock to see which one are new and which one need to be deleted
    
    full_path = os.path.join(exit, name)
    
    with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
        if not name.endswith('.xlsx'): #If the name you've put as an input does not finish with .xlsx, we add it
            name += '.xlsx'
            
        df.to_excel(writer, sheet_name = 'Transpa_OPCVM')
        df_sum.to_excel(writer, sheet_name = 'Résumé erreurs')

    
    #We also create the log to help the user notice the errors and modifications made by our program
    joinedlist = SM.errors + TR.errors2
    SM.export_txt(joinedlist, exit,'Modifications et Erreurs.txt')
    print('Transparisation terminée')
    return None


stk = 'Test_2024\\Etat_Stock_OPCVM_02122024.xlsx'

#ONLY PUT the sheet of the previous transparisation entitled 'Transpa_OPCVM
pr = 'Test_2024\\03_Transpa_R3S_Niv1_VF_202312.xlsm'

def pre_transp(stock_path,prev_path,ex,name):
    
    print('Début de la pré-transparisation')
    
    df_prev = pd.read_excel(prev_path, sheet_name='Transpa_OPCVM')
    df_add, df_del = SM.add_and_reject(df_prev, stock_path)
    
    full_path = os.path.join(ex, name)
    
    with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
        if not name.endswith('.xlsx'): #If the name you've put as an input does not finish with .xlsx, we add it
            name += '.xlsx'
            
        df_add.to_excel(writer, sheet_name = 'Fonds à rajouter')
        df_del.to_excel(writer, sheet_name= 'Fonds retirés')
        
    print('Pré-transparisation terminée')
        
    return None
"""

start = time.time()
pre_transp(stk,pr,'Results','Pré_Transpa.xlsx')
end = time.time()
duree = end - start
minutes = int(duree // 60)
secondes = int(duree % 60)
print(f"Temps d'exécution : {minutes} min {secondes} s")


#Follow the indicates modifications to use the following commands
start = time.time()
nom_du_dossier = 'Test_2024\\TPT_2024' #Replace by the name of the folder that contains your TPT
inter = SM.list_of_files(nom_du_dossier) 
files = [nom_du_dossier + '\\' + inter[i] for i in range(len(inter))]
#Adapt the year, the folder where you want your results, the name of the result you want and the path to the file that contains your OPCVM stock.
transparisation(files,2024,'Results','TranspaTest.xlsx')
end = time.time()

duree = end - start
minutes = int(duree // 60)
secondes = int(duree % 60)
print(f"Temps d'exécution : {minutes} min {secondes} s")
"""
nom_du_dossier = ""

path = 'Fichiers Bug à traiter\\I21 - FR0010540716 part IC - Ampere - 2024-12-31.xlsx'
df = pd.read_excel(path)
print(df)
for i in range(len(list(df.columns))):
    print(list(df.columns)[i])