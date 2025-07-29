# Fichier : apptest.py (Option 2 avec gestion des actions C/D)

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl

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
            "Missing Data": 6  # Mise à jour pour inclure toutes les colonnes
        }
        
        # Mapping des colonnes Excel vers les noms de colonnes en base
        self.column_mapping = {
            "Missing Data": {
                "Product": "product_name",
                "Category": "category_type", 
                "FO System": "front_office_system",
                "Attribute": "attribute_name",
                "Comment": "comment_text",
                # Actions (C,D) n'est PAS mappée car elle ne va pas en base
                # created_on sera ajoutée automatiquement avec GETDATE()
            },
            "COREXT": {
                "ID": "corext_id",
                "Name": "corext_name", 
                "Type": "corext_type",
                "Status": "status_value",
                # Actions (C,D) n'est PAS mappée
                # created_on sera ajoutée automatiquement
            },
            "Transco By Source": {
                "Source": "source_name",
                "Target": "target_name",
                "Priority": "priority_level",
                # created_on sera ajoutée automatiquement
            },
            "Transco By Ref": {
                "Reference": "reference_id",
                "Value": "reference_value",
                "Description": "ref_description",
                # created_on sera ajoutée automatiquement
            }
        }
        
        # Colonnes qui définissent l'action (ne pas inclure dans les WHERE/INSERT)
        self.action_columns = ["Actions (C,D)", "action_type"]
        
        # Colonnes automatiques par feuille (différentes pour chaque table)
        self.auto_columns_by_sheet = {
            "Missing Data": {
                "created_on": "GETDATE()",
                "created_by": "username_from_login"  # Utilise le username de connexion
            },
            "COREXT": {
                "last_update": "GETDATE()",
                "updated_by": "username_from_login",
                "status": "ACTIVE"
            },
            "Transco By Source": {
                "import_date": "GETDATE()",
                "source_system": "WEEKLY_IMPORT"
            },
            "Transco By Ref": {
                "created_date": "GETDATE()",
                "validation_status": "PENDING"
            }
        }
        
        # Configuration spéciale pour COREXT (SELECT puis INSERT/DELETE sur 2 tables)
        self.corext_config = {
            "select_columns": ["ID", "Name", "Type"],  # Colonnes pour la requête SELECT
            "target_tables": ["corext_main_table", "corext_details_table"],  # 2 tables cibles
            "select_table": "corext_source_table"  # Table source pour le SELECT
        }
        
        self.servers_map = {
            "EUROPE": "wlypasdev000078LITES_GLORYV4EU_ASS",
            "ASIA": "wlypasdev000078LITES_GLORYV4AS_INT", 
            "NY": "wlypasdev000078LITES_GLORYV4NA_INT",
        }

        self.create_widgets()

    def create_widgets(self):
        """Crée tous les widgets de l'interface"""
        
        # -- Conteneur pour la connexion
        login_frame = ttk.Frame(self.window)
        login_frame.pack(pady=5, padx=5, fill="x")

        # -- Sélection du serveur
        server_label = tk.Label(login_frame, text="Serveur:")
        server_label.pack(anchor="w")
        
        options = list(self.servers_map.keys())
        self.selectServer = tk.StringVar(value=options[0])
        
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
        password_label.pack(anchor="w", pady=(10, 0))
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(login_frame, textvariable=self.password_var, show="*")
        password_entry.pack(fill="x")

        # -- Étiquette pour afficher les résultats/statuts
        self.etiquette_resultat = tk.Label(
            self.window, 
            text="Veuillez sélectionner un fichier à traiter.", 
            font=("Helvetica", 12),
            wraplength=600
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
        """
        Charge les données Excel avec format dictionnaire et mapping
        Format: {'sheet_name': [{'col1': 'val1', 'col2': 'val2'}, ...]}
        """
        sheets_to_process = self.config_colonnes.keys()
        self.donnees_excel = {}

        try:
            workbook = openpyxl.load_workbook(filepath)
            filename = filepath.split('/')[-1]
            
            sheets_found = []
            
            for sheet in workbook.worksheets:
                if sheet.title in sheets_to_process:
                    nb_colonnes = self.config_colonnes[sheet.title]
                    
                    # Lecture des données
                    all_rows = list(sheet.values)
                    if not all_rows:
                        continue
                        
                    headers = all_rows[0][:nb_colonnes]
                    data = []
                    
                    for row_values in all_rows[1:]:
                        if any(row_values):
                            # Créer dictionnaire avec les headers Excel
                            row_data = dict(zip(headers, row_values[:nb_colonnes]))
                            data.append(row_data)

                    self.donnees_excel[sheet.title] = data
                    sheets_found.append(sheet.title)
                    
                    print(f"\n=== Feuille '{sheet.title}' chargée ===")
                    print(f"Headers Excel: {headers}")
                    if sheet.title in self.column_mapping:
                        db_headers = [self.column_mapping[sheet.title].get(h, h) for h in headers]
                        print(f"Headers DB: {db_headers}")
                    print(f"Données: {len(data)} lignes")
                    
                    # Afficher les actions pour debug
                    if data and "Actions (C,D)" in headers:
                        actions_summary = {}
                        for row in data:
                            action = row.get("Actions (C,D)", "Unknown")
                            actions_summary[action] = actions_summary.get(action, 0) + 1
                        print(f"Actions trouvées: {actions_summary}")

            status_msg = f"Fichier '{filename}' chargé.\n"
            status_msg += f"Feuilles traitées: {', '.join(sheets_found)}"
            self.etiquette_resultat.config(text=status_msg)

        except Exception as e:
            print(f"Erreur lors du chargement: {e}")
            self.etiquette_resultat.config(text=f"Erreur: {str(e)}")
            self.donnees_excel = None

    def on_connect_click(self):
        """Vérifie si un fichier a été chargé et lance la connexion/traitement."""
        if self.donnees_excel is None:
            self.etiquette_resultat.config(text="Veuillez d'abord sélectionner un fichier!")
            return

        serveur = self.selectServer.get()
        username = self.username_var.get().strip()
        password = self.password_var.get().strip()
        
        if not username or not password:
            self.etiquette_resultat.config(text="Username et password requis.")
            return

        database = self.servers_map.get(serveur)
        if not database:
            self.etiquette_resultat.config(text=f"Serveur '{serveur}' non trouvé.")
            return
             
        self.process_connection(serveur, database, username, password)

    def process_connection(self, serveur, database, username, password):
        """Traite la connexion et les données."""
        try:
            self.etiquette_resultat.config(text=f"Connexion à {serveur}...")
            
            # Ici, la logique de connexion à la base de données (ex: pyodbc)
            # conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={database};UID={username};PWD={password}")
            # cursor = conn.cursor()
            
            print(f"\n=== Connexion à {serveur} ===")
            print(f"DATABASE: {database}")
            print(f"UID: {username}")

            total_processed = 0
            for sheet_name, sheet_data in self.donnees_excel.items():
                print(f"\n=== Traitement de la feuille: {sheet_name} ===")
                
                # Dispatcher selon le nom de la feuille
                if sheet_name == 'Missing Data':
                    processed = self.process_missing_data(sheet_data, sheet_name)
                elif sheet_name == 'COREXT':
                    processed = self.process_corext(sheet_data, sheet_name)
                elif sheet_name == 'Transco By Source':
                    processed = self.process_transco_source(sheet_data, sheet_name)
                elif sheet_name == 'Transco By Ref':
                    processed = self.process_transco_ref(sheet_data, sheet_name)
                else:
                    print(f"Feuille non reconnue: {sheet_name}")
                    continue
                
                total_processed += processed
            
            self.etiquette_resultat.config(
                text=f"Traitement terminé! {total_processed} opérations exécutées."
            )
        
        except Exception as e:
            print(f"Erreur: {e}")
            self.etiquette_resultat.config(text=f"Erreur: {str(e)}")
        
        finally:
            print("Fin du traitement")
            # if 'conn' in locals():
            #     conn.close()

    def process_missing_data(self, sheet_data, sheet_name):
        """
        Traite les données Missing Data avec gestion des actions C/D
        DELETE en premier, puis INSERT
        """
        print(f"\n--- Traitement Missing Data: {len(sheet_data)} lignes ---")
        
        # Séparer les actions par type
        delete_rows = []
        insert_rows = []
        other_rows = []
        
        for row_dict in sheet_data:
            action = row_dict.get("Actions (C,D)", "").strip().upper()
            if action == "D":
                delete_rows.append(row_dict)
            elif action == "C":
                insert_rows.append(row_dict)
            else:
                other_rows.append(row_dict)
                print(f"Action non reconnue: '{action}' pour la ligne: {row_dict}")
        
        print(f"Actions détectées: {len(delete_rows)} DELETE, {len(insert_rows)} INSERT, {len(other_rows)} autres")
        
        total_processed = 0
        
        # 1. TRAITER LES DELETE EN PREMIER
        if delete_rows:
            print(f"\n=== PHASE 1: Exécution des {len(delete_rows)} DELETE Missing Data ===")
            for i, row_dict in enumerate(delete_rows):
                try:
                    sql, params = self.build_delete_query_for_sheet("Missing Data", row_dict)
                    print(f"DELETE {i+1}/{len(delete_rows)}: {sql}")
                    print(f"  Params: {params}")
                    
                    # cursor.execute(sql, params)
                    # affected_rows = cursor.rowcount
                    # print(f"  Lignes supprimées: {affected_rows}")
                    
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur DELETE ligne {i+1}: {e}")
        
        # 2. TRAITER LES INSERT ENSUITE
        if insert_rows:
            print(f"\n=== PHASE 2: Exécution des {len(insert_rows)} INSERT Missing Data ===")
            for i, row_dict in enumerate(insert_rows):
                try:
                    sql, params = self.build_insert_query_for_sheet("Missing Data", row_dict)
                    print(f"INSERT {i+1}/{len(insert_rows)}: {sql}")
                    print(f"  Params: {params}")
                    
                    # cursor.execute(sql, params)
                    # print(f"  Ligne insérée avec succès")
                    
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur INSERT ligne {i+1}: {e}")
        
        # conn.commit()  # Valider toutes les transactions
        print(f"\nMissing Data terminé: {total_processed} opérations")
        return total_processed

    def build_delete_query(self, sheet_name, row_dict):
        """Construit une requête DELETE avec les conditions WHERE."""
        table_name = self.get_table_name(sheet_name)
        
        conditions = []
        params = []
        
        for excel_col, value in row_dict.items():
            # Ignorer la colonne action et les valeurs nulles
            if excel_col not in self.action_columns and value is not None and str(value).strip():
                db_col = self.get_db_column_name(sheet_name, excel_col)
                conditions.append(f"{db_col} = ?")
                params.append(str(value).strip())
        
        if not conditions:
            raise ValueError("Aucune condition WHERE valide pour le DELETE")
        
        sql = f"DELETE FROM {table_name} WHERE {' AND '.join(conditions)}"
        return sql, params

    def build_insert_query_for_sheet(self, sheet_name, row_dict, table_name=None):
        """Construit une requête INSERT spécifique à chaque feuille."""
        if table_name is None:
            table_name = self.get_table_name(sheet_name)
        
        columns = []
        values_placeholders = []
        params = []
        
        # Ajouter les colonnes du fichier Excel (sauf Actions)
        for excel_col, value in row_dict.items():
            if excel_col not in self.action_columns:
                db_col = self.get_db_column_name(sheet_name, excel_col)
                columns.append(db_col)
                values_placeholders.append('?')
                params.append(str(value).strip() if value is not None else None)
        
        # Ajouter les colonnes automatiques spécifiques à cette feuille
        if sheet_name in self.auto_columns_by_sheet:
            for auto_col, auto_value in self.auto_columns_by_sheet[sheet_name].items():
                columns.append(auto_col)
                if auto_value == "GETDATE()":
                    values_placeholders.append("GETDATE()")  # Fonction SQL directe
                elif auto_value == "username_from_login":
                    values_placeholders.append("?")
                    params.append(self.username_var.get() or "SYSTEM")
                else:
                    values_placeholders.append("?")
                    params.append(auto_value)
        
        if not columns:
            raise ValueError(f"Aucune colonne valide pour l'INSERT {sheet_name}")
        
        sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join(values_placeholders)})"
        return sql, params

    def build_delete_query_for_sheet(self, sheet_name, row_dict, table_name=None):
        """Construit une requête DELETE spécifique à chaque feuille."""
        if table_name is None:
            table_name = self.get_table_name(sheet_name)
        
        conditions = []
        params = []
        
        for excel_col, value in row_dict.items():
            # Ignorer la colonne action et les valeurs nulles
            if excel_col not in self.action_columns and value is not None and str(value).strip():
                db_col = self.get_db_column_name(sheet_name, excel_col)
                conditions.append(f"{db_col} = ?")
                params.append(str(value).strip())
        
        if not conditions:
            raise ValueError(f"Aucune condition WHERE valide pour le DELETE {sheet_name}")
        
        sql = f"DELETE FROM {table_name} WHERE {' AND '.join(conditions)}"
        return sql, params

    def get_table_name(self, sheet_name):
        """Retourne le nom de table en base de données."""
        table_mapping = {
            "Missing Data": "missing_data_table",
            "COREXT": "corext_table", 
            "Transco By Source": "transco_source_table",
            "Transco By Ref": "transco_ref_table"
        }
        return table_mapping.get(sheet_name, sheet_name.lower().replace(' ', '_') + '_table')

    def get_db_column_name(self, sheet_name, excel_column):
        """Retourne le nom de colonne en base de données."""
        if sheet_name in self.column_mapping:
            return self.column_mapping[sheet_name].get(excel_column, excel_column.lower().replace(' ', '_'))
        return excel_column.lower().replace(' ', '_')

    # Méthodes pour les autres feuilles (avec gestion des actions si nécessaire)
    def process_corext(self, sheet_data, sheet_name):
        """
        Traite les données COREXT: SELECT puis INSERT/DELETE sur 2 tables
        """
        print(f"\n--- Traitement COREXT: {len(sheet_data)} lignes ---")
        
        # Séparer les actions par type
        delete_rows = []
        insert_rows = []
        
        for row_dict in sheet_data:
            action = row_dict.get("Actions (C,D)", "").strip().upper()
            if action == "D":
                delete_rows.append(row_dict)
            elif action == "C":
                insert_rows.append(row_dict)
        
        print(f"COREXT Actions: {len(delete_rows)} DELETE, {len(insert_rows)} INSERT")
        total_processed = 0
        
        # 1. TRAITER LES DELETE EN PREMIER
        if delete_rows:
            print(f"\n=== COREXT PHASE 1: {len(delete_rows)} DELETE sur 2 tables ===")
            for i, row_dict in enumerate(delete_rows):
                try:
                    # D'abord faire le SELECT pour récupérer les données complètes
                    select_data = self.execute_corext_select(row_dict)
                    
                    if select_data:
                        # DELETE sur les 2 tables cibles avec les données récupérées
                        for table in self.corext_config["target_tables"]:
                            sql, params = self.build_delete_query_for_sheet("COREXT", select_data, table)
                            print(f"DELETE {i+1} sur {table}: {sql}")
                            print(f"  Params: {params}")
                            # cursor.execute(sql, params)
                        total_processed += 1
                    else:
                        print(f"  Aucune donnée trouvée pour DELETE ligne {i+1}")
                        
                except Exception as e:
                    print(f"  Erreur DELETE COREXT ligne {i+1}: {e}")
        
        # 2. TRAITER LES INSERT ENSUITE
        if insert_rows:
            print(f"\n=== COREXT PHASE 2: {len(insert_rows)} INSERT sur 2 tables ===")
            for i, row_dict in enumerate(insert_rows):
                try:
                    # D'abord faire le SELECT pour récupérer les données complètes
                    select_data = self.execute_corext_select(row_dict)
                    
                    if select_data:
                        # INSERT sur les 2 tables cibles avec les données récupérées
                        for table in self.corext_config["target_tables"]:
                            sql, params = self.build_insert_query_for_sheet("COREXT", select_data, table)
                            print(f"INSERT {i+1} sur {table}: {sql}")
                            print(f"  Params: {params}")
                            # cursor.execute(sql, params)
                        total_processed += 1
                    else:
                        print(f"  Aucune donnée trouvée pour INSERT ligne {i+1}")
                        
                except Exception as e:
                    print(f"  Erreur INSERT COREXT ligne {i+1}: {e}")
        
        print(f"\nCOREXT terminé: {total_processed} opérations")
        return total_processed

    def execute_corext_select(self, row_dict):
        """
        Exécute le SELECT pour récupérer les données complètes depuis la table source
        """
        source_table = self.corext_config["select_table"]
        select_columns = self.corext_config["select_columns"]
        
        # Construire les conditions WHERE avec les colonnes de sélection
        conditions = []
        params = []
        
        for excel_col in select_columns:
            if excel_col in row_dict and row_dict[excel_col] is not None:
                db_col = self.get_db_column_name("COREXT", excel_col)
                conditions.append(f"{db_col} = ?")
                params.append(str(row_dict[excel_col]).strip())
        
        if not conditions:
            print("  Aucune condition WHERE valide pour le SELECT COREXT")
            return None
        
        sql = f"SELECT * FROM {source_table} WHERE {' AND '.join(conditions)}"
        print(f"  SELECT COREXT: {sql}")
        print(f"  SELECT Params: {params}")
        
        # cursor.execute(sql, params)
        # result = cursor.fetchone()  # ou fetchall() si plusieurs résultats attendus
        
        # SIMULATION: retourner des données fictives pour test
        simulation_data = {
            "corext_id": row_dict.get("ID", "TEST_ID"),
            "corext_name": row_dict.get("Name", "TEST_NAME"),
            "corext_type": row_dict.get("Type", "TEST_TYPE"),
            "additional_field": "SELECTED_VALUE"  # Données supplémentaires du SELECT
        }
        
        return simulation_data

    def build_corext_delete_query(self, table_name, select_data):
        """Construit une requête DELETE pour COREXT avec les données du SELECT."""
        return self.build_delete_query_for_sheet("COREXT", select_data, table_name)

    def build_corext_insert_query(self, table_name, select_data):
        """Construit une requête INSERT pour COREXT avec les données du SELECT."""
        return self.build_insert_query_for_sheet("COREXT", select_data, table_name)

    # Méthodes pour les autres feuilles avec leur logique spécifique
    def process_transco_source(self, sheet_data, sheet_name):
        """Traite les données Transco By Source avec sa logique spécifique."""
        print(f"\n--- Traitement Transco By Source: {len(sheet_data)} lignes ---")
        
        # Séparer les actions par type
        delete_rows = []
        insert_rows = []
        
        for row_dict in sheet_data:
            action = row_dict.get("Actions (C,D)", "").strip().upper()
            if action == "D":
                delete_rows.append(row_dict)
            elif action == "C":
                insert_rows.append(row_dict)
        
        print(f"Transco Source Actions: {len(delete_rows)} DELETE, {len(insert_rows)} INSERT")
        total_processed = 0
        
        # 1. DELETE en premier
        if delete_rows:
            print(f"\n=== PHASE 1: {len(delete_rows)} DELETE Transco Source ===")
            for i, row_dict in enumerate(delete_rows):
                try:
                    sql, params = self.build_delete_query_for_sheet("Transco By Source", row_dict)
                    print(f"DELETE {i+1}: {sql}")
                    print(f"  Params: {params}")
                    # cursor.execute(sql, params)
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur DELETE ligne {i+1}: {e}")
        
        # 2. INSERT ensuite
        if insert_rows:
            print(f"\n=== PHASE 2: {len(insert_rows)} INSERT Transco Source ===")
            for i, row_dict in enumerate(insert_rows):
                try:
                    sql, params = self.build_insert_query_for_sheet("Transco By Source", row_dict)
                    print(f"INSERT {i+1}: {sql}")
                    print(f"  Params: {params}")
                    # cursor.execute(sql, params)
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur INSERT ligne {i+1}: {e}")
        
        print(f"\nTransco Source terminé: {total_processed} opérations")
        return total_processed

    def process_transco_ref(self, sheet_data, sheet_name):
        """Traite les données Transco By Ref avec sa logique spécifique."""
        print(f"\n--- Traitement Transco By Ref: {len(sheet_data)} lignes ---")
        
        # Séparer les actions par type
        delete_rows = []
        insert_rows = []
        
        for row_dict in sheet_data:
            action = row_dict.get("Actions (C,D)", "").strip().upper()
            if action == "D":
                delete_rows.append(row_dict)
            elif action == "C":
                insert_rows.append(row_dict)
        
        print(f"Transco Ref Actions: {len(delete_rows)} DELETE, {len(insert_rows)} INSERT")
        total_processed = 0
        
        # 1. DELETE en premier
        if delete_rows:
            print(f"\n=== PHASE 1: {len(delete_rows)} DELETE Transco Ref ===")
            for i, row_dict in enumerate(delete_rows):
                try:
                    sql, params = self.build_delete_query_for_sheet("Transco By Ref", row_dict)
                    print(f"DELETE {i+1}: {sql}")
                    print(f"  Params: {params}")
                    # cursor.execute(sql, params)
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur DELETE ligne {i+1}: {e}")
        
        # 2. INSERT ensuite
        if insert_rows:
            print(f"\n=== PHASE 2: {len(insert_rows)} INSERT Transco Ref ===")
            for i, row_dict in enumerate(insert_rows):
                try:
                    sql, params = self.build_insert_query_for_sheet("Transco By Ref", row_dict)
                    print(f"INSERT {i+1}: {sql}")
                    print(f"  Params: {params}")
                    # cursor.execute(sql, params)
                    total_processed += 1
                except Exception as e:
                    print(f"  Erreur INSERT ligne {i+1}: {e}")
        
        print(f"\nTransco Ref terminé: {total_processed} opérations")
        return total_processed

    # Méthode utilitaire pour tester les requêtes par feuille
    def test_queries(self):
        """Méthode de test pour vérifier la génération des requêtes SQL par feuille."""
        print("\n=== TEST DE GÉNÉRATION DE REQUÊTES PAR FEUILLE ===")
        
        # Test Missing Data
        missing_data_test = {
            'Product': 'Repo',
            'Category': 'TRAN', 
            'FO System': 'Murex V3',
            'Attribute': 'Spread',
            'Actions (C,D)': 'D',
            'Comment': 'Test comment'
        }
        
        print("\n--- TEST MISSING DATA ---")
        try:
            sql, params = self.build_delete_query_for_sheet("Missing Data", missing_data_test)
            print(f"DELETE SQL: {sql}")
            print(f"DELETE Params: {params}")
            
            missing_data_test['Actions (C,D)'] = 'C'
            sql, params = self.build_insert_query_for_sheet("Missing Data", missing_data_test)
            print(f"INSERT SQL: {sql}")
            print(f"INSERT Params: {params}")
        except Exception as e:
            print(f"Erreur Missing Data: {e}")
        
        # Test COREXT
        corext_test = {
            'corext_id': 'TEST_ID',
            'corext_name': 'TEST_NAME',
            'corext_type': 'TEST_TYPE',
            'additional_field': 'SELECTED_VALUE'
        }
        
        print("\n--- TEST COREXT ---")
        try:
            sql, params = self.build_insert_query_for_sheet("COREXT", corext_test, "corext_main_table")
            print(f"COREXT INSERT SQL: {sql}")
            print(f"COREXT INSERT Params: {params}")
        except Exception as e:
            print(f"Erreur COREXT: {e}")
        
        # Test Transco By Source
        transco_source_test = {
            'Source': 'SRC_TEST',
            'Target': 'TGT_TEST',
            'Priority': '1',
            'Actions (C,D)': 'C'
        }
        
        print("\n--- TEST TRANSCO BY SOURCE ---")
        try:
            sql, params = self.build_insert_query_for_sheet("Transco By Source", transco_source_test)
            print(f"TRANSCO SOURCE INSERT SQL: {sql}")
            print(f"TRANSCO SOURCE INSERT Params: {params}")
        except Exception as e:
            print(f"Erreur Transco Source: {e}")


def main():
    """Fonction principale pour lancer l'application."""
    window = tk.Tk()
    window.geometry("700x400")
    window.resizable(True, True)
    
    app = WeeklySettings(window)
    
    # Pour tester la génération de requêtes (optionnel)
    # app.test_queries()
    
    window.mainloop()


if __name__ == "__main__":
    main()