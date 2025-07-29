# Fichier : apptest.py

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl
# import pandas as pd # Note: pandas est importé mais non utilisé dans le code montré

class WeeklySettings:
    def __init__(self, window):
        self.window = window
        window.title("Weekly Settings")

        # --- Variables de données ---
        self.donnees_excel = None
        self.config_colonnes = {
            "COREXT": 8,
            "Transco By Source": 8,
            "Transco By Ref": 6,
            "Missing Data": 4
        }
        self.servers_map = {
            "EUROPE": "wlypasdev000078LITES_GLORYV4EU_ASS",
            "ASIA": "wlypasdev000078LITES_GLORYV4AS_INT",
            "NY": "wlypasdev000078LITES_GLORYV4NA_INT",
        }

        # --- Création des widgets ---
        self.create_widgets()

    def create_widgets(self):
        """Crée tous les widgets de l'interface"""
        
        # -- Conteneur pour la connexion
        login_frame = ttk.Frame(self.window)
        login_frame.pack(pady=5, padx=5, fill="x")

        # -- Sélection du serveur
        server_label = tk.Label(login_frame, text="Serveur:")  # Nom plus clair
        server_label.pack(anchor="w")
        
        options = list(self.servers_map.keys())
        self.selectServer = tk.StringVar(value=options[0])  # Initialise avec "EUROPE"
        
        self.menu = tk.OptionMenu(login_frame, self.selectServer, *options)
        self.menu.pack(pady=5, fill="x")

        # -- Identifiant
        username_label = tk.Label(login_frame, text="Username:")
        username_label.pack(anchor="w")
        self.username_var = tk.StringVar()
        username_entry = ttk.Entry(login_frame, textvariable=self.username_var)
        username_entry.pack(fill="x")

        # -- Mot de passe
        password_label = ttk.Label(login_frame, text="Password:")
        password_label.pack(anchor="w", pady=(10, 0))  # Plus d'espacement
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(login_frame, textvariable=self.password_var, show="*")
        password_entry.pack(fill="x")

        # -- Étiquette pour afficher les résultats/statuts
        self.etiquette_resultat = tk.Label(
            self.window, 
            text="Veuillez sélectionner un fichier à traiter.", 
            font=("Helvetica", 12),
            wraplength=600  # Permet le retour à la ligne pour les longs messages
        )
        self.etiquette_resultat.pack(pady=15)

        # -- Conteneur pour les boutons d'action
        button_frame = ttk.Frame(self.window)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # -- Définition des boutons
        filedialog_btn = ttk.Button(
            button_frame, 
            text="Select a file", 
            command=self.open_files
        )
        filedialog_btn.pack(side=tk.LEFT, padx=5)

        connect_btn = ttk.Button(
            button_frame, 
            text="Connexion", 
            command=self.on_connect_click
        )
        connect_btn.pack(side=tk.LEFT, padx=5)

        cancel_btn = ttk.Button(
            button_frame, 
            text="Cancel", 
            command=self.window.destroy
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)

    def open_files(self):
        """Ouvre l'explorateur de fichiers pour sélectionner un fichier Excel."""
        self.etiquette_resultat.config(text="Sélection du fichier en cours...")
        
        filepath = filedialog.askopenfilename(
            title="Ouvrir un fichier excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
        )
        
        if not filepath:
            self.etiquette_resultat.config(text="Aucun fichier sélectionné.")
            return

        self.load_excel_data(filepath)

    def load_excel_data(self, filepath):
        """Charge les données du fichier Excel."""
        sheets_to_process = self.config_colonnes.keys()
        self.donnees_excel = {}

        try:
            workbook = openpyxl.load_workbook(filepath)
            filename = filepath.split('/')[-1]  # Nom du fichier seulement
            
            sheets_found = []
            sheets_missing = []
            
            # Vérifier quelles feuilles sont disponibles
            available_sheets = [sheet.title for sheet in workbook.worksheets]
            
            for sheet_name in sheets_to_process:
                if sheet_name in available_sheets:
                    sheets_found.append(sheet_name)
                else:
                    sheets_missing.append(sheet_name)
            
            if not sheets_found:
                self.etiquette_resultat.config(
                    text=f"Aucune feuille requise trouvée dans '{filename}'"
                )
                return
            
            # Traiter les feuilles trouvées
            for sheet in workbook.worksheets:
                if sheet.title in sheets_found:
                    nb_colonnes = self.config_colonnes[sheet.title]
                    
                    # Lecture des données
                    all_rows = list(sheet.values)
                    if not all_rows:
                        print(f"Feuille '{sheet.title}' est vide, ignorée.")
                        continue
                        
                    headers = all_rows[0][:nb_colonnes]
                    
                    # Vérifier que les en-têtes ne sont pas vides
                    if not any(headers):
                        print(f"Feuille '{sheet.title}' n'a pas d'en-têtes valides.")
                        continue
                    
                    data = []
                    for row_values in all_rows[1:]:
                        if any(row_values):  # Ignorer les lignes complètement vides
                            # Créer un dictionnaire pour chaque ligne
                            row_data = dict(zip(headers, row_values[:nb_colonnes]))
                            data.append(row_data)

                    # Stocker les données de la feuille
                    self.donnees_excel[sheet.title] = data
                    print(f"Feuille '{sheet.title}' traitée: {len(data)} lignes de données.")

            # Message de statut
            status_msg = f"Fichier '{filename}' chargé.\n"
            status_msg += f"Feuilles traitées: {', '.join(sheets_found)}"
            
            if sheets_missing:
                status_msg += f"\nFeuilles manquantes: {', '.join(sheets_missing)}"
            
            self.etiquette_resultat.config(text=status_msg)

        except Exception as e:
            print(f"Erreur lors du chargement: {e}")
            self.etiquette_resultat.config(
                text=f"Erreur lors de la lecture du fichier: {str(e)}"
            )
            self.donnees_excel = None

    def on_connect_click(self):
        """Vérifie si un fichier a été chargé et lance la connexion/traitement."""
        if self.donnees_excel is None:
            self.etiquette_resultat.config(
                text="Veuillez d'abord sélectionner et charger un fichier!"
            )
            return

        if not self.donnees_excel:  # Dictionnaire vide
            self.etiquette_resultat.config(
                text="Aucune donnée valide trouvée dans le fichier!"
            )
            return

        # Récupérer les informations de connexion
        serveur = self.selectServer.get()
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        
        if not username or not password:
            self.etiquette_resultat.config(
                text="Le nom d'utilisateur et le mot de passe sont requis."
            )
            return

        database = self.servers_map.get(serveur)
        if not database:
            self.etiquette_resultat.config(
                text=f"Erreur: Serveur '{serveur}' non trouvé."
            )
            return
             
        self.process_connection(serveur, database, username, password)

    def process_connection(self, serveur, database, username, password):
        """Traite la connexion et les données."""
        try:
            self.etiquette_resultat.config(text=f"Connexion à {serveur}...")
            
            # Ici, la logique de connexion à la base de données (ex: pyodbc)
            # conn = pyodbc.connect(...)
            # cursor = conn.cursor()
            
            print(f"Connexion à {serveur}")
            print(f"DATABASE: {database}")
            print(f"UID: {username}")
            print(f"PWD: {'*' * len(password)}")

            # Traitement des données par feuille
            total_processed = 0
            for sheet_name, sheet_data in self.donnees_excel.items():
                print(f"\nTraitement de la feuille: {sheet_name} ({len(sheet_data)} lignes)")
                
                if sheet_name == 'COREXT':
                    result = self.process_corext(sheet_data)
                elif sheet_name == 'Transco By Source':
                    result = self.process_transco_source(sheet_data)
                elif sheet_name == 'Transco By Ref':
                    result = self.process_transco_ref(sheet_data)
                elif sheet_name == 'Missing Data':
                    result = self.process_missing_data(sheet_data)
                else:
                    print(f"Feuille non reconnue: {sheet_name}")
                    continue
                
                total_processed += len(sheet_data)
            
            self.etiquette_resultat.config(
                text=f"Traitement terminé! {total_processed} lignes traitées."
            )
        
        except Exception as e:
            print(f"Exception lors du traitement: {e}")
            self.etiquette_resultat.config(text=f"Erreur de connexion: {str(e)}")
        
        finally:
            print("Fin du traitement")
            # if 'conn' in locals():
            #     conn.close()

    # Méthodes de traitement (à implémenter selon vos besoins)
    def process_corext(self, data):
        """Traite les données COREXT."""
        print(f"Traitement COREXT: {len(data)} lignes")
        # Votre logique de traitement ici
        return len(data)

    def process_transco_source(self, data):
        """Traite les données Transco By Source."""
        print(f"Traitement Transco By Source: {len(data)} lignes")
        # Votre logique de traitement ici
        return len(data)

    def process_transco_ref(self, data):
        """Traite les données Transco By Ref."""
        print(f"Traitement Transco By Ref: {len(data)} lignes")
        # Votre logique de traitement ici
        return len(data)

    def process_missing_data(self, data):
        """Traite les données Missing Data."""
        print(f"Traitement Missing Data: {len(data)} lignes")
        # Votre logique de traitement ici
        return len(data)


def main():
    """Fonction principale pour lancer l'application."""
    window = tk.Tk()
    window.geometry("700x400")  # Taille par défaut
    window.resizable(True, True)  # Permettre le redimensionnement
    
    app = WeeklySettings(window)
    window.mainloop()


if __name__ == "__main__":
    main()