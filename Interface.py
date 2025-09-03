import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import time
import os
import sys

# Importation des modules personnalisés si disponibles
from Sheets_Manip import SheetsManip
from Transpa_Tests import Transpa
import win32com.client as win32


# Valeurs par défaut pour l'envoi de mails
DEFAULT_OBJET = "[Ne pas répondre] Test Mail Automatique"
DEFAULT_MESSAGE_TEMPLATE = """Bonjour {prenom},

Dans le cadre de la transparisation Solvabilité 2, pouvez-vous nous faire parvenir les TPT au 31/12/2025 pour les fonds suivants s’il vous plaît :

{liste_portefeuilles}

Cordialement,"""

# Redirection de la console vers l'interface
class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

# Nettoyage des chemins
def clean_path(path):
    return path.replace('\\\\', '\\')

# Fonctions de transparisation
def pre_transp(stock_path, prev_path, ex, name):
    df_prev = pd.read_excel(prev_path, sheet_name='Transpa_OPCVM')
    df_add, df_del = SheetsManip().add_and_reject(df_prev, stock_path)
    full_path = os.path.join(ex, name if name.endswith('.xlsx') else name + '.xlsx')
    with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
        df_add.to_excel(writer, sheet_name='Fonds à rajouter')
        df_del.to_excel(writer, sheet_name='Fonds à retirer')

def transparisation(L, y, exit, name, a, b, c, d, fich_suivi):
    
    print('Début de la transparisation.')
    
    SM = SheetsManip()
    TR = Transpa()
    df = SM.concat(L)
    
    print('Concaténation terminée.')
    
    TR.df_test = TR.col_df2(df)
    df = TR.empty_port_id(df)
    df = TR.VM(df,a)
    TR.storage(df)
    df = TR.maturity(df, y)
    TR.coupon_freq(df)
    df = TR.redemp_rate(df)
    df = TR.modified_duration(df)
    df = TR.coupon_rate(df,b,c)
    df = TR.delt(df)
    TR.ratio_nom_vm(df,d)
    TR.not_cred(df)
    df = TR.covered(df)
    TR.empt_country(df)
    TR.empt_under(df)
    TR.infra(df)
    
    df = SM.conv_date(df)
    
    """
    df_suivi = SM.input_SAS(df)
    """
    print("Phase d'enregistrement...")
    
    if "File's name" in df.columns:
        df = df.drop("File's name", axis=1)
        
    df = df.dropna(how='all')
    
    SM.df_col.columns = [os.path.basename(col) for col in SM.df_col.columns]
    TR.df_test.columns = [os.path.basename(col) for col in TR.df_test.columns]
    
    df_sum = pd.concat([SM.df_col, TR.df_test], ignore_index=True).T
    df_sum = df_sum.iloc[:, :-1]
    df = df[df.count(axis=1) >= 10]
    
    
    full_path = os.path.join(exit, name if name.endswith('.xlsx') else name + '.xlsx')
    with pd.ExcelWriter(full_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Transpa_OPCVM')
        df_sum.to_excel(writer, sheet_name='Résumé erreurs')
    
    """
    full_path2 = os.path.join(exit, "Input_SAS.xlsx")
    df_suivi.to_excel(full_path2, index=False, sheet_name='Transpa')
    """
    
    joinedlist = SM.errors + TR.errors2
    SM.export_txt(joinedlist, exit, 'Modifications et Erreurs.txt')
    
    print('Fichiers en enregistrés.')
    
    print('Transparisation terminée !')

# Application principale
class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Interface Transparisation et Envoi de Mails")
        self.root.geometry("1000x700")
        self.root.configure(bg="#e6f0ff")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        self.create_pretransp_tab()
        self.create_transp_tab()
        self.create_mail_tab()
        self.create_console()
        
        self.lock = threading.Lock()

        sys.stdout = ConsoleRedirector(self.console)
        sys.stderr = ConsoleRedirector(self.console)
        
        
        self.suivi_console = scrolledtext.ScrolledText(self.tab3, height=10)
        self.suivi_console.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")
        
        
        self.tab3.grid_rowconfigure(5, weight=1)
        self.tab3.grid_columnconfigure(1, weight=1)


    def create_pretransp_tab(self):
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Pré-Transparisation")

        self.stk_path = tk.StringVar()
        self.pr_path = tk.StringVar()
        self.save_path1 = tk.StringVar()
        self.filename1 = tk.StringVar(value="Pré_Transpa.xlsx")

        self.create_file_selector(self.tab1, "Stock OPCVM", self.stk_path, 0)
        self.create_file_selector(self.tab1, "Précédente Transpa", self.pr_path, 1)
        self.create_folder_selector(self.tab1, "Dossier de sauvegarde", self.save_path1, 2)
        self.create_entry(self.tab1, "Nom du fichier", self.filename1, 3)

        tk.Button(self.tab1, text="Lancer Pré-Transparisation", bg="#4d94ff", fg="white",
                  command=self.run_pretransp).grid(row=4, column=0, columnspan=3, pady=10)

    def create_transp_tab(self):
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="Transparisation")

        self.folder_path = tk.StringVar()
        self.year = tk.StringVar(value="2024")
        self.save_path2 = tk.StringVar()
        self.filename2 = tk.StringVar(value="TranspaTest.xlsx")

        self.create_folder_selector(self.tab2, "Dossier TPT", self.folder_path, 0)
        self.chem_fich_suivi = tk.StringVar()
        tk.Label(self.tab2, text="Fichier de suivi", bg="#e6f0ff").grid(row=1, column=0, sticky='w', padx=10, pady=5)
        tk.Entry(self.tab2, textvariable=self.chem_fich_suivi, width=60).grid(row=1, column=1, padx=5)
        tk.Button(self.tab2, text="Parcourir", command=lambda: self.chem_fich_suivi.set(
            filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xlsm")])
        )).grid(row=1, column=2)
        self.chem_fich_corresp = tk.StringVar()

        self.create_entry(self.tab2, "Année", self.year, 2)
        self.create_folder_selector(self.tab2, "Dossier de sauvegarde", self.save_path2, 3)
        self.create_entry(self.tab2, "Nom du fichier", self.filename2, 4)

        frame_params = ttk.LabelFrame(self.tab2, text="Paramétrage de la transparisation", padding=(10, 10))
        frame_params.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        self.use_ecart = tk.BooleanVar(value=False)
        self.ecart_value = tk.StringVar(value="0.001")
        self.use_palliers = tk.BooleanVar(value=False)
        self.pallier_tres_eleve = tk.StringVar(value="15")
        self.pallier_eleve = tk.StringVar(value="9")
        self.use_ratio = tk.BooleanVar(value=False)
        self.ratio_value = tk.StringVar(value="5")

        tk.Checkbutton(frame_params, text="Personnaliser écart VM-NAV (%)", variable=self.use_ecart).grid(row=0, column=0, sticky="w", pady=2)
        tk.Entry(frame_params, textvariable=self.ecart_value, width=10).grid(row=0, column=1, pady=2)
        tk.Checkbutton(frame_params, text="Personnaliser Palliers Coupon", variable=self.use_palliers).grid(row=1, column=0, sticky="w", pady=2)
        tk.Label(frame_params, text="Très élevé : ").grid(row=1, column=1, sticky="e")
        tk.Entry(frame_params, textvariable=self.pallier_tres_eleve, width=5).grid(row=1, column=2, pady=2)
        tk.Label(frame_params, text="Élevé : ").grid(row=1, column=3, sticky="e")
        tk.Entry(frame_params, textvariable=self.pallier_eleve, width=5).grid(row=1, column=4, pady=2)
        tk.Checkbutton(frame_params, text="Personnaliser ratio VM", variable=self.use_ratio).grid(row=2, column=0, sticky="w", pady=2)
        tk.Entry(frame_params, textvariable=self.ratio_value, width=10).grid(row=2, column=1, pady=2)

        # Bouton déplacé ici, après le cadre de paramétrage
        tk.Button(self.tab2, text="Lancer Transparisation", bg="#4d94ff", fg="white",
                command=self.run_transp).grid(row=6, column=0, columnspan=3, pady=10)


    def create_mail_tab(self):
        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="Envoi de Mails")

        self.objet_personnalise = tk.BooleanVar(value=False)
        self.message_personnalise = tk.BooleanVar(value=False)

        self.entry_objet = tk.Entry(self.tab3)
        self.entry_objet.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        tk.Label(self.tab3, text="Objet du mail:", bg="#e6f0ff").grid(row=0, column=0, sticky="e")
        self.bouton_objet = tk.Button(self.tab3, text="Mode personnalisé", command=self.toggle_objet_mode)
        self.bouton_objet.grid(row=0, column=2, padx=5)
        self.toggle_objet_mode()

        self.text_message = tk.Text(self.tab3, height=10)
        self.text_message.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        tk.Label(self.tab3, text="Message:", bg="#e6f0ff").grid(row=1, column=0, sticky="ne")
        self.bouton_message = tk.Button(self.tab3, text="Mode personnalisé", command=self.toggle_message_mode)
        self.bouton_message.grid(row=1, column=2, padx=5, sticky="n")
        self.toggle_message_mode()

        tk.Label(self.tab3, text="Adresse mail manuelle:", bg="#e6f0ff").grid(row=2, column=0, sticky="e")
        self.entry_email = tk.Entry(self.tab3)
        self.entry_email.grid(row=2, column=1, padx=5, pady=5, sticky="we")

        tk.Label(self.tab3, text="Fichier Excel:", bg="#e6f0ff").grid(row=3, column=0, sticky="e")
        self.entry_fichier = tk.Entry(self.tab3)
        self.entry_fichier.grid(row=3, column=1, sticky="we", padx=5, pady=5)
        tk.Button(self.tab3, text="Parcourir", command=self.selectionner_fichier).grid(row=3, column=2, sticky="we", padx=5, pady=5)

        tk.Button(self.tab3, text="Envoyer les mails", command=self.envoyer_mails).grid(row=4, column=1, pady=10)

        self.tab3.grid_rowconfigure(1, weight=1)
        self.tab3.grid_columnconfigure(1, weight=1)

    def create_console(self):
        self.console = scrolledtext.ScrolledText(self.root, height=10, state='disabled', bg="#f0f8ff")
        self.console.pack(fill='x', padx=10, pady=5)

    def create_file_selector(self, parent, label, var, row):
        tk.Label(parent, text=label, bg="#e6f0ff").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        tk.Entry(parent, textvariable=var, width=60).grid(row=row, column=1, padx=5)
        tk.Button(parent, text="Parcourir", command=lambda: var.set(filedialog.askopenfilename())).grid(row=row, column=2)

    def create_folder_selector(self, parent, label, var, row):
        tk.Label(parent, text=label, bg="#e6f0ff").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        tk.Entry(parent, textvariable=var, width=60).grid(row=row, column=1, padx=5)
        tk.Button(parent, text="Parcourir", command=lambda: var.set(filedialog.askdirectory())).grid(row=row, column=2)

    def create_entry(self, parent, label, var, row):
        tk.Label(parent, text=label, bg="#e6f0ff").grid(row=row, column=0, sticky='w', padx=10, pady=5)
        tk.Entry(parent, textvariable=var, width=60).grid(row=row, column=1, columnspan=2, padx=5)

    def toggle_objet_mode(self):
        if self.objet_personnalise.get():
            self.entry_objet.config(state="normal", bg="white")
            self.bouton_objet.config(text="Mode par défaut")
        else:
            self.entry_objet.delete(0, tk.END)
            self.entry_objet.insert(0, DEFAULT_OBJET)
            self.entry_objet.config(state="disabled", bg="#d0e7ff")
            self.bouton_objet.config(text="Mode personnalisé")

    def toggle_message_mode(self):
        if self.message_personnalise.get():
            self.text_message.config(state="normal", bg="white")
            self.bouton_message.config(text="Mode par défaut")
        else:
            self.text_message.delete("1.0", tk.END)
            self.text_message.insert("1.0", DEFAULT_MESSAGE_TEMPLATE.format(prenom="{Prénom}", liste_portefeuilles="•             IDF001 ;\n•             IDF002 ;"))
            self.text_message.config(state="disabled", bg="#d0e7ff")
            self.bouton_message.config(text="Mode personnalisé")

    def selectionner_fichier(self):
        fichier = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])
        self.entry_fichier.delete(0, tk.END)
        self.entry_fichier.insert(0, fichier)

    def run_pretransp(self):
        threading.Thread(target=self._run_pretransp).start()

    def _run_pretransp(self):
        
        with self.lock:    
            try:
                print('Début de la pré-transparisation.')
                start = time.time()
                stock = clean_path(self.stk_path.get())
                prev = clean_path(self.pr_path.get())
                folder = clean_path(self.save_path1.get())
                name = self.filename1.get()
                pre_transp(stock, prev, folder, name)
                print('Pré-transparisation terminée !')
                self.print_duration(start)
            except Exception as e:
                print(f"Erreur : {e}")

    def run_transp(self):
        threading.Thread(target=self._run_transp).start()

    def _run_transp(self):
        try:
            a = float(self.ecart_value.get())
            b = float(self.pallier_tres_eleve.get())
            c = float(self.pallier_eleve.get())
            d = float(self.ratio_value.get())

            start = time.time()
            folder = clean_path(self.folder_path.get())
            year = int(self.year.get())
            save = clean_path(self.save_path2.get())
            name = self.filename2.get()

            files = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith('.xlsx') or f.endswith('.xlsm')]

            if not files:
                print("Aucun fichier Excel trouvé dans le dossier sélectionné.")
                return

            
            transparisation(files, year, save, name, a, b, c, d, self.chem_fich_suivi.get())

            self.print_duration(start)

        except Exception as e:
            print(f"Erreur : {e}")


    def detect_header_row(self, file_path):
        try:
            df_raw = pd.read_excel(file_path, engine='openpyxl', header=None)
            for i, row in df_raw.iterrows():
                if str(row.iloc[0]).strip() == 'Nom du fond':
                    return i
            return None
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel : {e}")
            return None

    def envoyer_mails(self):
        fichier_excel = self.entry_fichier.get()
        message_custom = self.text_message.get("1.0", tk.END).strip() if self.message_personnalise.get() else None
        adresse_manuelle = self.entry_email.get()

        if not fichier_excel or not os.path.exists(fichier_excel):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier Excel valide.")
            return

        header_row = self.detect_header_row(fichier_excel)
        if header_row is None:
            messagebox.showerror("Erreur", "Impossible d'identifier la ligne d'en-tête. "
                                        "Assurez-vous que le fichier contient une colonne 'Nom du fond'.")
            return

        try:
            df = pd.read_excel(fichier_excel, engine='openpyxl', header=header_row)
            df_original = df.copy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel : {e}")
            return

        outlook = win32.Dispatch('outlook.application')
        df = df.dropna(subset=["Adresse mail", "Identifiant Portefeuille"])
        grouped = df.groupby("Adresse mail")

        for email, group in grouped:
            etats = group["Etat du reçu"].dropna().astype(int)
            if all(etat == 2 for etat in etats):
                self.suivi_console.insert(tk.END, f"Aucun email envoyé à {email} car TPT déjà reçu.\n")
                self.suivi_console.see(tk.END)
                continue

            relance = any(etat == 1 for etat in etats)
            objet_base = self.entry_objet.get() if self.objet_personnalise.get() else DEFAULT_OBJET
            objet_final = f"[RELANCE] {objet_base}" if relance else objet_base

            prenom = group["Prénom du contact"].dropna().astype(str).values[0] if "Prénom du contact" in group else ""
            identifiants = group["Identifiant Portefeuille"].dropna().astype(str).tolist()
            liste_portefeuilles = "\n".join([f"•             {idf} ;" for idf in identifiants])

            message_final = message_custom if message_custom else DEFAULT_MESSAGE_TEMPLATE.format(
                prenom=prenom if prenom else "",
                liste_portefeuilles=liste_portefeuilles
            )

            try:
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = objet_final
                mail.Body = message_final
                mail.Send()
                self.suivi_console.insert(tk.END, f"Email envoyé à {prenom} ({email})\n")
                self.suivi_console.see(tk.END)

                indices_a_mettre_a_jour = df_original[
                    (df_original["Adresse mail"] == email) & (df_original["Etat du reçu"] == 0)
                ].index
                df_original.loc[indices_a_mettre_a_jour, "Etat du reçu"] = 1

            except Exception as e:
                self.suivi_console.insert(tk.END, f"Erreur lors de l'envoi à {email} : {e}\n")
                self.suivi_console.see(tk.END)

        if adresse_manuelle:
            objet_base = self.entry_objet.get() if self.objet_personnalise.get() else DEFAULT_OBJET
            objet_final = objet_base
            try:
                mail = outlook.CreateItem(0)
                mail.To = adresse_manuelle
                mail.Subject = objet_final
                mail.Body = message_custom if message_custom else DEFAULT_MESSAGE_TEMPLATE.format(
                    prenom="",
                    liste_portefeuilles="•             [Identifiant Portefeuille]"
                )
                mail.Send()
                self.suivi_console.insert(tk.END, f"Email envoyé à {adresse_manuelle}\n")
                self.suivi_console.see(tk.END)
            except Exception as e:
                self.suivi_console.insert(tk.END, f"Erreur lors de l'envoi à {adresse_manuelle} : {e}\n")
                self.suivi_console.see(tk.END)

        try:
            df_original.to_excel(fichier_excel, index=False)
            self.suivi_console.insert(tk.END, "Fichier Excel mis à jour avec les nouveaux états.\n")
        except Exception as e:
            self.suivi_console.insert(tk.END, f"Erreur lors de la sauvegarde du fichier Excel : {e}\n")

        self.suivi_console.insert(tk.END, "L'envoi automatique est terminé.\n")
        self.suivi_console.see(tk.END)
        
    def print_duration(self, start):
        end = time.time()
        duration = end - start
        minutes = int(duration // 60)
        seconds = int(duration % 60)
        print(f"Temps d'exécution : {minutes} min {seconds} s")



# Lancement de l'application
if __name__ == "__main__":
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

