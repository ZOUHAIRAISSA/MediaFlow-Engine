#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gestionnaire pour la feuille "DriveFolders" de Google Sheets
Fonctionnalités :
- Lecture complète de la feuille avec structure préservée
- Ajout de nouvelles lignes
- Visualisation et modification des lignes existantes
- Préservation des dropdowns et formats existants
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import threading
from typing import Dict, List, Optional, Any
import os
import pickle

# Google APIs
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False

# Constantes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

class DriveFoldersManager:
    """Gestionnaire pour la feuille DriveFolders"""
    
    def __init__(self, spreadsheet_name: str = "automatisated kyopa insetion 2", 
                 worksheet_name: str = "DriveFolders"):
        self.spreadsheet_name = spreadsheet_name
        self.worksheet_name = worksheet_name
        self.service = None
        self.spreadsheet_id = None
        self.worksheet_data = []
        self.headers = []
        self.creds_manager = GoogleCredentialsManager()
    
    def authenticate(self):
        """Authentification Google"""
        if not GOOGLE_AVAILABLE:
            raise ImportError("Librairies Google non disponibles")
        
        creds = self.creds_manager.authenticate()
        self.service = build('sheets', 'v4', credentials=creds)
        return True
    
    def find_spreadsheet(self) -> Optional[str]:
        """Trouver le spreadsheet par nom"""
        try:
            drive_service = build('drive', 'v3', credentials=self.creds_manager.creds)
            results = drive_service.files().list(
                q=f"name='{self.spreadsheet_name}' and mimeType='application/vnd.google-apps.spreadsheet'",
                fields="files(id, name)"
            ).execute()
            
            files = results.get('files', [])
            if files:
                return files[0]['id']
            return None
        except Exception as e:
            raise Exception(f"Erreur lors de la recherche du spreadsheet: {e}")
    
    def load_worksheet(self) -> List[List]:
        """Charger les données de la feuille DriveFolders"""
        if not self.spreadsheet_id:
            self.spreadsheet_id = self.find_spreadsheet()
            if not self.spreadsheet_id:
                raise Exception(f"Spreadsheet '{self.spreadsheet_name}' non trouvé")
        
        try:
            # Charger toutes les données de la feuille (A à AB = 28 colonnes)
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{self.worksheet_name}!A:AB"  # Colonnes A à AB
            ).execute()
            
            self.worksheet_data = result.get('values', [])
            
            if self.worksheet_data:
                self.headers = self.worksheet_data[0]
                # S'assurer qu'on a bien toutes les colonnes A à AB (28 colonnes)
                while len(self.headers) < 28:
                    self.headers.append(f"Colonne_{len(self.headers) + 1}")
            else:
                # Créer les en-têtes par défaut si la feuille est vide
                self.headers = [f"Colonne_{chr(65 + i)}" for i in range(28)]  # A à AB
            
            return self.worksheet_data
        except Exception as e:
            raise Exception(f"Erreur lors de la lecture de la feuille: {e}")
    
    def get_dropdown_options(self, column_index: int) -> List[str]:
        """Récupérer les options de dropdown pour une colonne"""
        try:
            # Obtenir les informations détaillées de la feuille
            result = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id,
                ranges=[f"{self.worksheet_name}!A:AB"],
                includeGridData=True
            ).execute()
            
            sheet = result['sheets'][0]
            if 'data' in sheet and sheet['data']:
                grid_data = sheet['data'][0]
                if 'rowData' in grid_data and grid_data['rowData']:
                    # Chercher les données de validation dans la première ligne (en-têtes)
                    first_row = grid_data['rowData'][0]
                    if 'values' in first_row and column_index < len(first_row['values']):
                        cell = first_row['values'][column_index]
                        if 'dataValidation' in cell:
                            validation = cell['dataValidation']
                            if 'condition' in validation:
                                condition = validation['condition']
                                if condition['type'] == 'ONE_OF_LIST':
                                    return condition.get('values', [])
            
            return []
        except Exception as e:
            print(f"Erreur lors de la récupération des options dropdown: {e}")
            return []
    
    def get_row_data(self, row_index: int) -> Dict[str, Any]:
        """Récupérer les données d'une ligne sous forme de dictionnaire"""
        if row_index < 1 or row_index >= len(self.worksheet_data):
            return {}
        
        row = self.worksheet_data[row_index]
        row_data = {}
        
        for i, header in enumerate(self.headers):
            value = row[i] if i < len(row) else ""
            row_data[header] = value
        
        return row_data
    
    def add_row(self, row_data: Dict[str, Any]) -> bool:
        """Ajouter une nouvelle ligne"""
        try:
            # Préparer les valeurs dans l'ordre des colonnes
            values = []
            for header in self.headers:
                value = row_data.get(header, "")
                values.append([value])
            
            # Ajouter la ligne
            body = {
                'values': [values]
            }
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"{self.worksheet_name}!A:Z",
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            # Recharger les données
            self.load_worksheet()
            return True
            
        except Exception as e:
            raise Exception(f"Erreur lors de l'ajout de la ligne: {e}")
    
    def update_row(self, row_index: int, row_data: Dict[str, Any]) -> bool:
        """Mettre à jour une ligne existante"""
        try:
            if row_index < 1 or row_index >= len(self.worksheet_data):
                raise Exception("Index de ligne invalide")
            
            # Préparer les valeurs à mettre à jour
            updates = []
            for i, header in enumerate(self.headers):
                if header in row_data:
                    cell_range = f"{self.worksheet_name}!{chr(65 + i)}{row_index + 1}"
                    updates.append({
                        'range': cell_range,
                        'values': [[row_data[header]]]
                    })
            
            # Effectuer les mises à jour
            body = {
                'valueInputOption': 'RAW',
                'data': updates
            }
            
            self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body=body
            ).execute()
            
            # Recharger les données
            self.load_worksheet()
            return True
            
        except Exception as e:
            raise Exception(f"Erreur lors de la mise à jour de la ligne: {e}")
    
    def get_worksheet_info(self) -> Dict[str, Any]:
        """Obtenir les informations sur la structure de la feuille"""
        try:
            result = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id,
                ranges=[self.worksheet_name],
                includeGridData=True
            ).execute()
            
            sheet = result['sheets'][0]
            properties = sheet['properties']
            
            return {
                'title': properties['title'],
                'row_count': properties['gridProperties']['rowCount'],
                'column_count': properties['gridProperties']['columnCount'],
                'headers': self.headers,
                'data_rows': len(self.worksheet_data) - 1 if len(self.worksheet_data) > 1 else 0
            }
        except Exception as e:
            raise Exception(f"Erreur lors de la récupération des informations: {e}")


class GoogleCredentialsManager:
    """Gestionnaire des identifiants Google (réutilisé depuis Manager.py)"""
    
    def __init__(self, credentials_file='credentials.json', token_file='token.pickle'):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.creds = None
    
    def authenticate(self):
        """Authentification Google"""
        if not GOOGLE_AVAILABLE:
            raise ImportError("Librairies Google non disponibles")
        
        # Charger les tokens existants
        if os.path.exists(self.token_file):
            with open(self.token_file, 'rb') as token:
                self.creds = pickle.load(token)
        
        # Renouveler si nécessaire
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_file):
                    raise FileNotFoundError(f"Fichier {self.credentials_file} non trouvé")
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            # Sauvegarder les tokens
            with open(self.token_file, 'wb') as token:
                pickle.dump(self.creds, token)
        
        return self.creds


class DriveFoldersGUI:
    """Interface graphique pour le gestionnaire DriveFolders"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Gestionnaire DriveFolders")
        self.root.geometry("1200x800")
        self.root.configure(bg='#2b2b2b')
        
        # Manager
        self.manager = DriveFoldersManager()
        
        # Variables
        self.current_row_index = None
        self.row_vars = {}
        
        # Style moderne
        self.setup_modern_style()
        
        # Interface
        self.create_widgets()
        
        # Connexion automatique
        self.after(100, self.auto_connect)
    
    def setup_modern_style(self):
        """Configuration du style moderne"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Couleurs modernes
        style.configure('Modern.TFrame', background='#2b2b2b')
        style.configure('Modern.TLabel', background='#2b2b2b', foreground='#ffffff')
        style.configure('Modern.TButton', background='#4a90e2', foreground='#ffffff')
        style.map('Modern.TButton', background=[('active', '#357abd')])
        
        style.configure('Success.TButton', background='#28a745', foreground='#ffffff')
        style.map('Success.TButton', background=[('active', '#218838')])
        
        style.configure('Danger.TButton', background='#dc3545', foreground='#ffffff')
        style.map('Danger.TButton', background=[('active', '#c82333')])
    
    def create_widgets(self):
        """Créer l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self.root, style='Modern.TFrame')
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Titre
        title = ttk.Label(main_frame, text="Gestionnaire DriveFolders", 
                         font=('Arial', 16, 'bold'), style='Modern.TLabel')
        title.pack(pady=(0, 20))
        
        # Barre d'outils
        toolbar = ttk.Frame(main_frame, style='Modern.TFrame')
        toolbar.pack(fill='x', pady=(0, 10))
        
        ttk.Button(toolbar, text="🔄 Actualiser", command=self.refresh_data, 
                  style='Modern.TButton').pack(side='left', padx=5)
        ttk.Button(toolbar, text="➕ Ajouter ligne", command=self.add_row_dialog, 
                  style='Success.TButton').pack(side='left', padx=5)
        ttk.Button(toolbar, text="✏️ Modifier ligne", command=self.edit_row_dialog, 
                  style='Modern.TButton').pack(side='left', padx=5)
        
        # Informations sur la feuille
        info_frame = ttk.LabelFrame(main_frame, text="Informations sur la feuille")
        info_frame.pack(fill='x', pady=(0, 10))
        
        self.info_label = ttk.Label(info_frame, text="Non connecté", style='Modern.TLabel')
        self.info_label.pack(pady=5)
        
        # Liste des lignes
        list_frame = ttk.LabelFrame(main_frame, text="Lignes de données")
        list_frame.pack(fill='both', expand=True)
        
        # Treeview pour afficher les données (A à AB = 28 colonnes)
        columns = ['Row'] + [f'Col{i+1}' for i in range(28)]  # Colonnes A à AB
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack
        self.tree.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')
        
        # Bind double-click
        self.tree.bind('<Double-1>', self.on_row_double_click)
        
        # Barre de statut
        self.status_bar = ttk.Label(main_frame, text="Prêt", 
                                   style='Modern.TLabel', relief='sunken')
        self.status_bar.pack(side='bottom', fill='x', pady=(10, 0))
    
    def after(self, ms, func):
        """Wrapper pour after"""
        self.root.after(ms, func)
    
    def update_status(self, message):
        """Mettre à jour la barre de statut"""
        self.status_bar.config(text=message)
        self.root.update_idletasks()
    
    def auto_connect(self):
        """Connexion automatique"""
        def connect_worker():
            try:
                self.update_status("Connexion à Google...")
                self.manager.authenticate()
                self.refresh_data()
                self.update_status("Connecté - Données chargées")
            except Exception as e:
                self.update_status(f"Erreur: {e}")
                messagebox.showerror("Erreur de connexion", f"Impossible de se connecter:\n{e}")
        
        threading.Thread(target=connect_worker, daemon=True).start()
    
    def refresh_data(self):
        """Actualiser les données"""
        def refresh_worker():
            try:
                self.update_status("Chargement des données...")
                self.manager.load_worksheet()
                self.populate_tree()
                self.update_info()
                self.update_status("Données actualisées")
            except Exception as e:
                self.update_status(f"Erreur: {e}")
                messagebox.showerror("Erreur", f"Impossible de charger les données:\n{e}")
        
        threading.Thread(target=refresh_worker, daemon=True).start()
    
    def populate_tree(self):
        """Peupler le treeview avec les données"""
        # Nettoyer le tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.manager.headers:
            return
        
        # Configurer les colonnes (A à AB = 28 colonnes)
        columns = ['Row'] + [chr(65 + i) for i in range(28)]  # A, B, C, ..., AB
        self.tree['columns'] = columns
        
        # Configurer les en-têtes
        for i, col in enumerate(columns):
            if i == 0:
                self.tree.heading(col, text="Row")
            else:
                header_name = self.manager.headers[i-1] if i-1 < len(self.manager.headers) else f"Col{i}"
                self.tree.heading(col, text=f"{col} - {header_name}")
            self.tree.column(col, width=120, minwidth=80)
        
        # Ajouter les données
        for i, row in enumerate(self.manager.worksheet_data[1:], 1):  # Skip header
            # S'assurer qu'on a 28 colonnes (A à AB)
            row_values = []
            for j in range(28):
                if j < len(row):
                    row_values.append(row[j])
                else:
                    row_values.append("")
            values = [str(i)] + row_values
            self.tree.insert('', 'end', values=values, tags=(i,))
    
    def update_info(self):
        """Mettre à jour les informations"""
        try:
            info = self.manager.get_worksheet_info()
            info_text = f"Feuille: {info['title']} | Colonnes: {info['column_count']} | Lignes de données: {info['data_rows']}"
            self.info_label.config(text=info_text)
        except Exception as e:
            self.info_label.config(text=f"Erreur info: {e}")
    
    def on_row_double_click(self, event):
        """Double-clic sur une ligne"""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            row_index = int(item['values'][0])
            self.edit_row_dialog(row_index)
    
    def add_row_dialog(self):
        """Dialogue pour ajouter une nouvelle ligne"""
        if not self.manager.headers:
            messagebox.showwarning("Attention", "Aucune donnée chargée")
            return
        
        dialog = RowEditDialog(self.root, self.manager.headers, {}, "Ajouter une nouvelle ligne", self.manager)
        if dialog.result:
            def add_worker():
                try:
                    self.update_status("Ajout de la ligne...")
                    self.manager.add_row(dialog.result)
                    self.refresh_data()
                    self.update_status("Ligne ajoutée avec succès")
                    messagebox.showinfo("Succès", "Ligne ajoutée avec succès")
                except Exception as e:
                    self.update_status(f"Erreur: {e}")
                    messagebox.showerror("Erreur", f"Impossible d'ajouter la ligne:\n{e}")
            
            threading.Thread(target=add_worker, daemon=True).start()
    
    def edit_row_dialog(self, row_index=None):
        """Dialogue pour modifier une ligne"""
        if not self.manager.headers:
            messagebox.showwarning("Attention", "Aucune donnée chargée")
            return
        
        if row_index is None:
            selection = self.tree.selection()
            if not selection:
                messagebox.showwarning("Attention", "Veuillez sélectionner une ligne")
                return
            item = self.tree.item(selection[0])
            row_index = int(item['values'][0])
        
        # Récupérer les données de la ligne
        row_data = self.manager.get_row_data(row_index)
        
        dialog = RowEditDialog(self.root, self.manager.headers, row_data, f"Modifier la ligne {row_index}", self.manager)
        if dialog.result:
            def update_worker():
                try:
                    self.update_status("Mise à jour de la ligne...")
                    self.manager.update_row(row_index, dialog.result)
                    self.refresh_data()
                    self.update_status("Ligne mise à jour avec succès")
                    messagebox.showinfo("Succès", "Ligne mise à jour avec succès")
                except Exception as e:
                    self.update_status(f"Erreur: {e}")
                    messagebox.showerror("Erreur", f"Impossible de mettre à jour la ligne:\n{e}")
            
            threading.Thread(target=update_worker, daemon=True).start()
    
    def run(self):
        """Lancer l'application"""
        self.root.mainloop()


class RowEditDialog:
    """Dialogue pour éditer une ligne"""
    
    def __init__(self, parent, headers, row_data, title, manager=None):
        self.result = None
        self.headers = headers
        self.row_data = row_data
        self.manager = manager
        
        # Créer la fenêtre
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("1000x700")
        self.dialog.configure(bg='#2b2b2b')
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variables pour les champs
        self.vars = {}
        self.dropdown_options = {}
        
        self.create_widgets()
        
        # Centrer la fenêtre
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (1000 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"1000x700+{x}+{y}")
    
    def create_widgets(self):
        """Créer les widgets du dialogue"""
        # Frame principal avec scrollbar
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Canvas pour le scroll
        canvas = tk.Canvas(main_frame, bg='#2b2b2b')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Créer les champs pour toutes les colonnes A à AB (28 colonnes)
        for i in range(28):
            header = self.headers[i] if i < len(self.headers) else f"Colonne_{chr(65 + i)}"
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill='x', pady=2)
            
            label = ttk.Label(frame, text=f"{chr(65 + i)} - {header}:", width=25)
            label.pack(side='left', padx=(0, 10))
            
            # Récupérer les options dropdown si disponibles
            dropdown_options = []
            if self.manager:
                try:
                    dropdown_options = self.manager.get_dropdown_options(i)
                except:
                    pass
            
            # Déterminer la valeur actuelle
            current_value = self.row_data.get(header, "") if header in self.row_data else ""
            
            # Créer le widget approprié
            if header.lower() in ['select', 'selected', 'checkbox', 'check']:
                # Checkbox pour la colonne "select"
                var = tk.BooleanVar()
                # Convertir la valeur en booléen
                if str(current_value).lower() in ['true', '1', 'yes', 'oui', '✓', '✔', 'x']:
                    var.set(True)
                else:
                    var.set(False)
                self.vars[header] = var
                
                checkbox = ttk.Checkbutton(frame, variable=var)
                checkbox.pack(side='left', fill='x', expand=True)
                
            elif dropdown_options:
                # Combobox pour les colonnes avec dropdown
                var = tk.StringVar(value=current_value)
                self.vars[header] = var
                
                combobox = ttk.Combobox(frame, textvariable=var, values=dropdown_options, width=47)
                combobox.pack(side='left', fill='x', expand=True)
                
            else:
                # Entry normal pour les autres colonnes
                var = tk.StringVar(value=current_value)
                self.vars[header] = var
                
                entry = ttk.Entry(frame, textvariable=var, width=50)
                entry.pack(side='left', fill='x', expand=True)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Boutons
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(button_frame, text="Annuler", command=self.dialog.destroy).pack(side='right', padx=5)
        ttk.Button(button_frame, text="OK", command=self.ok_clicked).pack(side='right', padx=5)
    
    def ok_clicked(self):
        """Bouton OK cliqué"""
        self.result = {}
        for header, var in self.vars.items():
            if isinstance(var, tk.BooleanVar):
                # Pour les checkboxes, convertir en texte approprié
                self.result[header] = "TRUE" if var.get() else "FALSE"
            else:
                self.result[header] = var.get()
        self.dialog.destroy()


def main():
    """Fonction principale"""
    if not GOOGLE_AVAILABLE:
        messagebox.showerror("Erreur", 
                           "Librairies Google non disponibles.\n"
                           "Installez avec: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return
    
    app = DriveFoldersGUI()
    app.run()


if __name__ == "__main__":
    main()
