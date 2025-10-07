#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface moderne pour gérer Google Sheets et Google Drive
Fonctionnalités :
- Visualisation de la feuille "Etsy Listing Template"
- Upload de plusieurs sous-dossiers vers Google Drive
- Mise à jour automatique des liens dans Google Sheets
- Filtrage par statut (erreur)
- Téléchargement de dossiers par liens
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from threading import Thread
import os
import zipfile
import tempfile
from pathlib import Path
from typing import Dict, List, Optional
import json
import webbrowser
import subprocess
import sys

# Google APIs
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    import pickle
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False

# Constantes
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
SPREADSHEET_NAME = "automatisated kyopa insetion 2"
WORKSHEET_NAME = "Etsy Listing Template"

class GoogleCredentialsManager:
    """Gestionnaire des identifiants Google"""
    
    def __init__(self, credentials_file='credentials.json', token_file='token.pickle'):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.creds = None
    
    def authenticate(self):
        """Authentification Google"""
        if not GOOGLE_AVAILABLE:
            raise ImportError("Librairies Google non disponibles. Installez: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        
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
                    raise FileNotFoundError(f"Fichier {self.credentials_file} non trouvé. Téléchargez-le depuis Google Cloud Console.")
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES)
                self.creds = flow.run_local_server(port=0)
            
            # Sauvegarder les tokens
            with open(self.token_file, 'wb') as token:
                pickle.dump(self.creds, token)
        
        return self.creds

class GoogleSheetsManager:
    """Gestionnaire Google Sheets"""
    
    def __init__(self, credentials_manager):
        self.creds_manager = credentials_manager
        self.service = None
        self.spreadsheet_id = None
    
    def connect(self):
        """Connexion au service Google Sheets"""
        creds = self.creds_manager.authenticate()
        self.service = build('sheets', 'v4', credentials=creds)
        return True
    
    def find_spreadsheet(self, name: str) -> Optional[str]:
        """Trouver un spreadsheet par nom"""
        try:
            # Utiliser Google Drive API pour chercher
            drive_service = build('drive', 'v3', credentials=self.creds_manager.creds)
            results = drive_service.files().list(
                q=f"name='{name}' and mimeType='application/vnd.google-apps.spreadsheet'",
                fields="files(id, name)"
            ).execute()
            
            files = results.get('files', [])
            if files:
                return files[0]['id']
            return None
        except Exception as e:
            raise Exception(f"Erreur lors de la recherche du spreadsheet: {e}")
    
    def get_worksheet_data(self, worksheet_name: str) -> List[List]:
        """Récupérer les données d'une feuille"""
        if not self.spreadsheet_id:
            self.spreadsheet_id = self.find_spreadsheet(SPREADSHEET_NAME)
            if not self.spreadsheet_id:
                raise Exception(f"Spreadsheet '{SPREADSHEET_NAME}' non trouvé")
        
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{worksheet_name}!A:Z"
            ).execute()
            
            return result.get('values', [])
        except Exception as e:
            raise Exception(f"Erreur lors de la lecture de la feuille: {e}")
    
    def update_cell(self, worksheet_name: str, cell_range: str, value: str):
        """Mettre à jour une cellule"""
        try:
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=f"{worksheet_name}!{cell_range}",
                valueInputOption='RAW',
                body={'values': [[value]]}
            ).execute()
        except Exception as e:
            raise Exception(f"Erreur lors de la mise à jour: {e}")

class GoogleDriveManager:
    """Gestionnaire Google Drive"""
    
    def __init__(self, credentials_manager):
        self.creds_manager = credentials_manager
        self.service = None
        self.etsy_folder_id = "1YbCxswBnYswOAx-o09rn-TLMe5GgedrK"  # ID du dossier Photos Etsy Kyopadeco Shop
    
    def connect(self):
        """Connexion au service Google Drive"""
        creds = self.creds_manager.authenticate()
        self.service = build('drive', 'v3', credentials=creds)
        return True
    
    def list_folders(self) -> List[Dict]:
        """Lister les dossiers dans Google Drive"""
        try:
            results = self.service.files().list(
                q="mimeType='application/vnd.google-apps.folder'",
                fields="files(id, name, parents)"
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            raise Exception(f"Erreur lors de la liste des dossiers: {e}")
    
    def list_etsy_subfolders(self) -> List[Dict]:
        """Lister seulement les sous-dossiers du dossier Photos Etsy Kyopadeco Shop"""
        try:
            results = self.service.files().list(
                q=f"'{self.etsy_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'",
                fields="files(id, name, parents)"
            ).execute()
            
            return results.get('files', [])
        except Exception as e:
            raise Exception(f"Erreur lors de la liste des sous-dossiers Etsy: {e}")
    
    def create_folder(self, name: str, parent_id: str = None) -> str:
        """Créer un dossier"""
        try:
            folder_metadata = {
                'name': name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            
            if parent_id:
                folder_metadata['parents'] = [parent_id]
            
            folder = self.service.files().create(
                body=folder_metadata,
                fields='id'
            ).execute()
            
            return folder.get('id')
        except Exception as e:
            raise Exception(f"Erreur lors de la création du dossier: {e}")
    
    def upload_folder(self, local_path: Path, parent_id: str, progress_callback=None) -> str:
        """Uploader un dossier local vers Google Drive"""
        try:
            # Créer le dossier principal
            folder_id = self.create_folder(local_path.name, parent_id)
            
            # Uploader récursivement
            self._upload_folder_recursive(local_path, folder_id, progress_callback)
            
            return folder_id
        except Exception as e:
            raise Exception(f"Erreur lors de l'upload: {e}")
    
    def upload_subfolders_only(self, parent_local_path: Path, parent_id: str, progress_callback=None) -> List[Dict]:
        """Uploader seulement les sous-dossiers d'un dossier parent vers Google Drive"""
        uploaded_folders = []
        try:
            for item in parent_local_path.iterdir():
                if item.is_dir():  # Seulement les dossiers
                    if progress_callback:
                        progress_callback(f"Upload du sous-dossier: {item.name}")
                    
                    # Créer le dossier dans Google Drive
                    folder_id = self.create_folder(item.name, parent_id)
                    
                    # Uploader le contenu du sous-dossier
                    self._upload_folder_recursive(item, folder_id, progress_callback)
                    
                    # Ajouter à la liste des dossiers uploadés
                    folder_url = self.get_folder_url(folder_id)
                    uploaded_folders.append({
                        'name': item.name,
                        'id': folder_id,
                        'url': folder_url
                    })
                    
                    if progress_callback:
                        progress_callback(f"✅ {item.name} uploadé avec succès")
            
            return uploaded_folders
        except Exception as e:
            raise Exception(f"Erreur lors de l'upload des sous-dossiers: {e}")
    
    def _upload_folder_recursive(self, local_path: Path, parent_id: str, progress_callback=None):
        """Upload récursif"""
        for item in local_path.iterdir():
            if progress_callback:
                progress_callback(f"Upload: {item.name}")
            
            if item.is_file():
                self._upload_file(item, parent_id)
            elif item.is_dir():
                sub_folder_id = self.create_folder(item.name, parent_id)
                self._upload_folder_recursive(item, sub_folder_id, progress_callback)
    
    def _upload_file(self, file_path: Path, parent_id: str) -> str:
        """Uploader un fichier"""
        try:
            file_metadata = {
                'name': file_path.name,
                'parents': [parent_id]
            }
            
            media = MediaFileUpload(str(file_path), resumable=True)
            
            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            return file.get('id')
        except Exception as e:
            raise Exception(f"Erreur upload fichier {file_path.name}: {e}")
    
    def get_folder_url(self, folder_id: str) -> str:
        """Obtenir l'URL d'un dossier"""
        return f"https://drive.google.com/drive/folders/{folder_id}"
    
    def download_folder_by_url(self, folder_url: str, download_path: Path, progress_callback=None) -> bool:
        """Télécharger un dossier par son URL"""
        try:
            # Extraire l'ID du dossier de l'URL
            folder_id = self._extract_folder_id_from_url(folder_url)
            if not folder_id:
                raise Exception("ID de dossier non trouvé dans l'URL")
            
            # Télécharger récursivement
            folder_info = self.service.files().get(fileId=folder_id).execute()
            folder_name = folder_info['name']
            
            download_folder = download_path / folder_name
            download_folder.mkdir(parents=True, exist_ok=True)
            
            self._download_folder_recursive(folder_id, download_folder, progress_callback)
            
            return True
        except Exception as e:
            raise Exception(f"Erreur téléchargement: {e}")
    
    def _extract_folder_id_from_url(self, url: str) -> Optional[str]:
        """Extraire l'ID du dossier depuis l'URL"""
        if "/folders/" in url:
            return url.split("/folders/")[1].split("?")[0]
        elif "/drive/u/0/folders/" in url:
            return url.split("/drive/u/0/folders/")[1].split("?")[0]
        return None
    
    def _download_folder_recursive(self, folder_id: str, local_path: Path, progress_callback=None):
        """Téléchargement récursif"""
        try:
            results = self.service.files().list(
                q=f"'{folder_id}' in parents",
                fields="files(id, name, mimeType)"
            ).execute()
            
            files = results.get('files', [])
            
            for file_info in files:
                if progress_callback:
                    progress_callback(f"Téléchargement: {file_info['name']}")
                
                if file_info['mimeType'] == 'application/vnd.google-apps.folder':
                    # Dossier
                    sub_folder = local_path / file_info['name']
                    sub_folder.mkdir(exist_ok=True)
                    self._download_folder_recursive(file_info['id'], sub_folder, progress_callback)
                else:
                    # Fichier
                    self._download_file(file_info['id'], local_path / file_info['name'])
        except Exception as e:
            raise Exception(f"Erreur téléchargement récursif: {e}")
    
    def _download_file(self, file_id: str, local_path: Path):
        """Télécharger un fichier"""
        try:
            request = self.service.files().get_media(fileId=file_id)
            
            with open(local_path, 'wb') as f:
                downloader = request.execute()
                f.write(downloader)
        except Exception as e:
            raise Exception(f"Erreur téléchargement fichier: {e}")

class ModernGoogleDriveApp(tk.Tk):
    """Interface moderne pour Google Drive Manager"""
    
    def __init__(self):
        super().__init__()
        
        # Configuration de base
        self.title("Google Drive & Sheets Manager")
        self.geometry("1200x800")
        self.configure(bg='#2b2b2b')
        
        # Style moderne
        self.setup_modern_style()
        
        # Managers
        self.creds_manager = GoogleCredentialsManager()
        self.sheets_manager = GoogleSheetsManager(self.creds_manager)
        self.drive_manager = GoogleDriveManager(self.creds_manager)
        
        # Variables
        self.worksheet_data = []
        self.filtered_data = []
        self.selected_drive_folder = None
        
        # Interface utilisateur
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
        style.map('Modern.TButton',
                  background=[('active', '#357abd')])
        
        style.configure('Success.TButton', background='#28a745', foreground='#ffffff')
        style.map('Success.TButton',
                  background=[('active', '#218838')])
        
        style.configure('Warning.TButton', background='#ffc107', foreground='#000000')
        style.map('Warning.TButton',
                  background=[('active', '#e0a800')])
        
        style.configure('Danger.TButton', background='#dc3545', foreground='#ffffff')
        style.map('Danger.TButton',
                  background=[('active', '#c82333')])
    
    def create_widgets(self):
        """Créer l'interface utilisateur"""
        # Frame principal
        main_frame = ttk.Frame(self, style='Modern.TFrame')
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Titre
        title = ttk.Label(main_frame, text="Google Drive & Sheets Manager", 
                         font=('Arial', 16, 'bold'), style='Modern.TLabel')
        title.pack(pady=(0, 20))
        
        # Notebook pour les onglets
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Onglet 1: Visualisation des données
        self.create_data_tab()

        # Onglet Dashboard (module séparé)
        try:
            from dashboard_tab import DashboardTab
            self.dashboard_tab = DashboardTab(
                self.notebook,
                sheets_manager=self.sheets_manager,
                update_status_callback=self.update_status,
                stock_worksheet_name="stock Etsy Listing",
                drive_manager=self.drive_manager,
            )
            self.notebook.add(self.dashboard_tab, text="📈 Dashboard")
        except Exception as e:
            placeholder = ttk.Frame(self.notebook, style='Modern.TFrame')
            self.notebook.add(placeholder, text="📈 Dashboard")
            ttk.Label(placeholder, text=f"Erreur chargement Dashboard: {e}", style='Modern.TLabel').pack(padx=12, pady=12)
        
        # Onglet 2: Upload des dossiers
        self.create_upload_tab()
        
        # Onglet 3: Téléchargements
        self.create_download_tab()
        
        # Onglet 4: Traitement Vidéos
        self.create_video_processing_tab()
        
        # Onglet 5: Traitement des Images
        self.create_image_processing_tab()
        
        # Onglet 6: Configuration
        self.create_config_tab()
        
        # Barre de statut
        self.status_bar = ttk.Label(main_frame, text="Prêt", 
                                   style='Modern.TLabel', relief='sunken')
        self.status_bar.pack(side='bottom', fill='x', pady=(10, 0))
    
    def create_data_tab(self):
        """Créer l'onglet de visualisation des données"""
        data_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(data_frame, text="📊 Données Etsy")
        
        # Contrôles
        controls_frame = ttk.Frame(data_frame, style='Modern.TFrame')
        controls_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(controls_frame, text="🔄 Actualiser", 
                  command=self.refresh_data, style='Modern.TButton').pack(side='left', padx=5)
        
        ttk.Button(controls_frame, text="❌ Filtrer Erreurs", 
                  command=self.filter_errors, style='Warning.TButton').pack(side='left', padx=5)
        
        ttk.Button(controls_frame, text="👁️ Tout Afficher", 
                  command=self.show_all, style='Modern.TButton').pack(side='left', padx=5)
        
        # Recherche
        search_frame = ttk.Frame(controls_frame, style='Modern.TFrame')
        search_frame.pack(side='right')
        
        ttk.Label(search_frame, text="Recherche:", style='Modern.TLabel').pack(side='left')
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        search_entry.pack(side='left', padx=5)
        search_entry.bind('<KeyRelease>', self.on_search)
        
        # Tableau des données
        self.create_data_table(data_frame)
    
    def create_data_table(self, parent):
        """Créer le tableau des données"""
        # Frame pour le tableau
        table_frame = ttk.Frame(parent, style='Modern.TFrame')
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Treeview avec scrollbars
        self.tree = ttk.Treeview(table_frame, height=20)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Placement
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        # Bind double-click pour édition
        self.tree.bind('<Double-1>', self.on_row_double_click)
    
    def create_upload_tab(self):
        """Créer l'onglet d'upload"""
        upload_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(upload_frame, text="📤 Upload Dossiers")
        
        # Sélection du dossier Drive de destination
        dest_frame = ttk.LabelFrame(upload_frame, text="Dossier de destination Google Drive")
        dest_frame.pack(fill='x', padx=10, pady=10)
        
        self.drive_folder_var = tk.StringVar(value="Aucun dossier sélectionné")
        ttk.Label(dest_frame, textvariable=self.drive_folder_var, 
                 style='Modern.TLabel').pack(side='left', padx=10, pady=5)
        
        ttk.Button(dest_frame, text="Choisir Dossier Drive", 
                  command=self.select_drive_folder, style='Modern.TButton').pack(side='right', padx=10, pady=5)
        
        # Sélection du dossier parent local
        local_frame = ttk.LabelFrame(upload_frame, text="Dossier parent contenant les sous-dossiers à uploader")
        local_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Sélection du dossier parent
        parent_selection_frame = ttk.Frame(local_frame, style='Modern.TFrame')
        parent_selection_frame.pack(fill='x', padx=10, pady=10)
        
        self.parent_folder_var = tk.StringVar(value="Aucun dossier parent sélectionné")
        ttk.Label(parent_selection_frame, textvariable=self.parent_folder_var, 
                 style='Modern.TLabel').pack(side='left', padx=10, pady=5)
        
        ttk.Button(parent_selection_frame, text="Choisir Dossier Parent", 
                  command=self.select_parent_folder, style='Success.TButton').pack(side='right', padx=10, pady=5)
        
        # Liste des sous-dossiers détectés
        subfolders_label = ttk.Label(local_frame, text="Sous-dossiers détectés:", style='Modern.TLabel')
        subfolders_label.pack(anchor='w', padx=10, pady=(10, 5))
        
        list_frame = ttk.Frame(local_frame, style='Modern.TFrame')
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.subfolders_listbox = tk.Listbox(list_frame, height=8)
        subfolders_scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.subfolders_listbox.yview)
        self.subfolders_listbox.configure(yscrollcommand=subfolders_scrollbar.set)
        
        self.subfolders_listbox.pack(side='left', fill='both', expand=True)
        subfolders_scrollbar.pack(side='right', fill='y')
        
        # Contrôles
        control_frame = ttk.Frame(local_frame, style='Modern.TFrame')
        control_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(control_frame, text="🔄 Actualiser Sous-dossiers", 
                  command=self.refresh_subfolders, style='Modern.TButton').pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="✏️ Modifier Sous-dossiers", 
                  command=self.modify_subfolders, style='Warning.TButton').pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="🚀 Commencer Upload", 
                  command=self.start_subfolder_upload, style='Modern.TButton').pack(side='right', padx=5)
        
        # Progress bar
        self.upload_progress = ttk.Progressbar(upload_frame, mode='indeterminate')
        self.upload_progress.pack(fill='x', padx=10, pady=5)
        
        # Log d'upload
        log_frame = ttk.LabelFrame(upload_frame, text="Journal d'upload")
        log_frame.pack(fill='x', padx=10, pady=10)
        
        self.upload_log = tk.Text(log_frame, height=8, state='disabled')
        upload_log_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.upload_log.yview)
        self.upload_log.configure(yscrollcommand=upload_log_scroll.set)
        
        self.upload_log.pack(side='left', fill='both', expand=True)
        upload_log_scroll.pack(side='right', fill='y')
        
        # Variables
        self.selected_parent_folder = None
        self.detected_subfolders = []
    
    def create_download_tab(self):
        """Créer l'onglet de téléchargement"""
        download_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(download_frame, text="📥 Téléchargements")
        
        # Dossier de destination
        dest_frame = ttk.LabelFrame(download_frame, text="Dossier de téléchargement")
        dest_frame.pack(fill='x', padx=10, pady=10)
        
        self.download_path_var = tk.StringVar()
        ttk.Entry(dest_frame, textvariable=self.download_path_var, 
                 state='readonly').pack(side='left', fill='x', expand=True, padx=10, pady=5)
        
        ttk.Button(dest_frame, text="Choisir Dossier", 
                  command=self.select_download_folder, style='Modern.TButton').pack(side='right', padx=10, pady=5)
        
        # Actions de téléchargement
        actions_frame = ttk.LabelFrame(download_frame, text="Actions")
        actions_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(actions_frame, text="📥 Télécharger Dossiers avec Erreurs", 
                  command=self.download_error_folders, style='Warning.TButton').pack(pady=5)
        
        ttk.Button(actions_frame, text="📥 Télécharger par Drive Folder URL", 
                  command=self.download_by_drive_folder_url, style='Modern.TButton').pack(pady=5)
        
        # Progress et log
        self.download_progress = ttk.Progressbar(download_frame, mode='indeterminate')
        self.download_progress.pack(fill='x', padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(download_frame, text="Journal de téléchargement")
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.download_log = tk.Text(log_frame, state='disabled')
        download_log_scroll = ttk.Scrollbar(log_frame, orient='vertical', command=self.download_log.yview)
        self.download_log.configure(yscrollcommand=download_log_scroll.set)
        
        self.download_log.pack(side='left', fill='both', expand=True)
        download_log_scroll.pack(side='right', fill='y')
    
    def create_config_tab(self):
        """Créer l'onglet de configuration"""
        config_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(config_frame, text="⚙️ Configuration")
        
        # Informations de connexion
        conn_frame = ttk.LabelFrame(config_frame, text="Connexion Google")
        conn_frame.pack(fill='x', padx=10, pady=10)
        
        self.connection_status = ttk.Label(conn_frame, text="Non connecté", 
                                          style='Modern.TLabel')
        self.connection_status.pack(pady=5)
        
        ttk.Button(conn_frame, text="🔐 Reconnecter", 
                  command=self.reconnect, style='Modern.TButton').pack(pady=5)
        
        # Configuration des fichiers
        files_frame = ttk.LabelFrame(config_frame, text="Configuration des fichiers")
        files_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(files_frame, text=f"Spreadsheet: {SPREADSHEET_NAME}", 
                 style='Modern.TLabel').pack(anchor='w', padx=10, pady=2)
        ttk.Label(files_frame, text=f"Feuille: {WORKSHEET_NAME}", 
                 style='Modern.TLabel').pack(anchor='w', padx=10, pady=2)
        
        # Informations sur les colonnes
        columns_frame = ttk.LabelFrame(config_frame, text="Colonnes importantes")
        columns_frame.pack(fill='x', padx=10, pady=10)
        
        columns_text = """
        • sku kyopa: Identifiant unique du produit
        • Drive Folder URL Kyopa: URL du dossier Google Drive
        • sku original: Lien vers le dossier original
        • Status: Statut du produit (erreur, OK, etc.)
        """
        
        ttk.Label(columns_frame, text=columns_text, 
                 style='Modern.TLabel', justify='left').pack(anchor='w', padx=10, pady=5)
        
        # Boutons utiles
        utils_frame = ttk.LabelFrame(config_frame, text="Utilitaires")
        utils_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(utils_frame, text="🌐 Ouvrir Google Drive", 
                  command=lambda: webbrowser.open('https://drive.google.com'), 
                  style='Modern.TButton').pack(side='left', padx=5, pady=5)
        
        ttk.Button(utils_frame, text="📋 Ouvrir Google Sheets", 
                  command=self.open_sheets, style='Modern.TButton').pack(side='left', padx=5, pady=5)
    
    def create_video_processing_tab(self):
        """Créer l'onglet de lancement du traitement vidéos"""
        video_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(video_frame, text="🎬 Traitement Vidéos")

        # Intégrer directement l'interface BatchProcessorFrame
        from batchprocessor import BatchProcessorFrame
        embedded = BatchProcessorFrame(video_frame)
        embedded.pack(fill='both', expand=True)

    def create_image_processing_tab(self):
        """Créer l'onglet de traitement des images"""
        image_frame = ttk.Frame(self.notebook, style='Modern.TFrame')
        self.notebook.add(image_frame, text="🎨 Traitement Images")

        # Intégrer directement l'interface ImageEnhancerApp
        try:
            from enhance_canva_like import ImageEnhancerApp
            # Créer une instance de ImageEnhancerApp mais sans appeler mainloop()
            self.image_enhancer = ImageEnhancerApp()
            # Détruire la fenêtre principale créée par ImageEnhancerApp
            self.image_enhancer.destroy()
            
            # Créer un nouveau frame pour intégrer dans notre notebook
            self.image_enhancer_frame = ttk.Frame(image_frame, style='Modern.TFrame')
            self.image_enhancer_frame.pack(fill='both', expand=True)
            
            # Recréer l'interface dans notre frame
            self._create_image_enhancer_interface()
            
        except ImportError as e:
            # Si le module n'est pas trouvé, afficher un message d'erreur
            error_label = ttk.Label(image_frame, 
                                  text=f"Erreur: Impossible d'importer enhance_canva_like.py\n{e}", 
                                  style='Modern.TLabel')
            error_label.pack(expand=True)
        except Exception as e:
            # Autre erreur
            error_label = ttk.Label(image_frame, 
                                  text=f"Erreur lors du chargement de l'interface:\n{e}", 
                                  style='Modern.TLabel')
            error_label.pack(expand=True)

    def _create_image_enhancer_interface(self):
        """Créer l'interface ImageEnhancer dans notre frame"""
        # Copier les widgets de l'ImageEnhancerApp vers notre frame
        parent_frame = self.image_enhancer_frame
        
        # Variables
        self.input_folder_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.csv_path_var = tk.StringVar()
        self.resize_width_var = tk.StringVar()
        self.preset_var = tk.StringVar(value="none")
        
        # Canva preset parameters (using existing values from code)
        self.brightness_var = tk.DoubleVar(value=1.03)
        self.contrast_var = tk.DoubleVar(value=1.06)
        self.color_var = tk.DoubleVar(value=1.08)
        self.sharpness_var = tk.DoubleVar(value=1.08)
        self.gamma_var = tk.DoubleVar(value=0.98)
        self.r_gain_var = tk.DoubleVar(value=1.02)
        self.g_gain_var = tk.DoubleVar(value=1.00)
        self.b_gain_var = tk.DoubleVar(value=0.98)
        
        # Processing state
        self._stop_requested = False
        self._worker = None
        
        # Créer les widgets
        self._create_image_widgets(parent_frame)
        
        # Initial message
        self._append_image("🎨 Bienvenue dans l'Image Enhancer avec intégration CSV!")
        self._append_image("📂 Sélectionnez le dossier parent, le dossier de sortie et le fichier CSV.")
        self._append_image("⚙️ Optionnel: spécifiez une largeur de redimensionnement.")
        self._append_image("🚀 Cliquez sur 'Démarrer le traitement' pour commencer.\n")

    def _create_image_widgets(self, parent):
        """Créer tous les widgets de l'interface image"""
        # Main title
        title_frame = ttk.Frame(parent)
        title_frame.pack(fill="x", padx=10, pady=10)
        title_label = ttk.Label(title_frame, text="🎨 Traitement d'images", 
                               font=("Arial", 16, "bold"))
        title_label.pack()
        
        # Paths section
        paths_frame = ttk.LabelFrame(parent, text="📂 Chemins")
        paths_frame.pack(fill="x", padx=10, pady=8)
        
        self._row_path_image(paths_frame, "Dossier parent (sous-dossiers):", self.input_folder_var, self.browse_input_image)
        self._row_path_image(paths_frame, "Dossier de sortie:", self.output_folder_var, self.browse_output_image)
        self._row_path_image(paths_frame, "Fichier CSV:", self.csv_path_var, self.browse_csv_image)
        
        # Options section
        options_frame = ttk.LabelFrame(parent, text="⚙️ Options")
        options_frame.pack(fill="x", padx=10, pady=8)
        
        self._row_entry_image(options_frame, "Largeur de redimensionnement:", self.resize_width_var)
        
        # Preset section
        preset_frame = ttk.LabelFrame(parent, text="🎨 Presets d'amélioration")
        preset_frame.pack(fill="x", padx=10, pady=8)
        
        # Preset selection
        preset_row = ttk.Frame(preset_frame)
        preset_row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(preset_row, text="Preset:", width=15).pack(side="left")
        preset_combo = ttk.Combobox(preset_row, textvariable=self.preset_var, 
                                   values=["none", "canva"], state="readonly", width=15)
        preset_combo.pack(side="left", padx=6)
        preset_combo.bind("<<ComboboxSelected>>", self._on_preset_change_image)
        
        # Canva preset parameters (initially hidden)
        self.canva_params_frame = ttk.LabelFrame(preset_frame, text="Paramètres Canva")
        
        # Create parameter rows
        self._row_scale_image(self.canva_params_frame, "Luminosité:", self.brightness_var, 0.5, 2.0, 0.01)
        self._row_scale_image(self.canva_params_frame, "Contraste:", self.contrast_var, 0.5, 2.0, 0.01)
        self._row_scale_image(self.canva_params_frame, "Couleur:", self.color_var, 0.5, 2.0, 0.01)
        self._row_scale_image(self.canva_params_frame, "Netteté:", self.sharpness_var, 0.5, 2.0, 0.01)
        self._row_scale_image(self.canva_params_frame, "Gamma:", self.gamma_var, 0.5, 1.5, 0.01)
        self._row_scale_image(self.canva_params_frame, "Gain Rouge:", self.r_gain_var, 0.5, 1.5, 0.01)
        self._row_scale_image(self.canva_params_frame, "Gain Vert:", self.g_gain_var, 0.5, 1.5, 0.01)
        self._row_scale_image(self.canva_params_frame, "Gain Bleu:", self.b_gain_var, 0.5, 1.5, 0.01)
        
        # Control buttons
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill="x", padx=10, pady=8)
        
        self.start_btn_image = ttk.Button(control_frame, text="🚀 Démarrer le traitement", 
                                   command=self.start_processing_image, style='Success.TButton')
        self.start_btn_image.pack(side="left", padx=5)
        
        self.stop_btn_image = ttk.Button(control_frame, text="⏹️ Arrêter", 
                                  command=self.stop_processing_image, state="disabled", style='Danger.TButton')
        self.stop_btn_image.pack(side="left", padx=5)
        
        # Log section
        log_frame = ttk.LabelFrame(parent, text="📋 Journal")
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)
        
        self.log_text_image = tk.Text(log_frame, height=15, bg="#1f1f1f", fg="#eaeaea", 
                               insertbackground="#eaeaea", borderwidth=0, highlightthickness=0)
        self.log_text_image.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Now that log_text is created, we can call _on_preset_change
        self._on_preset_change_image()

    def _row_path_image(self, parent, label, var, command):
        """Create a row with label, entry and browse button"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(row, text=label, width=25).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="Parcourir…", command=command).pack(side="left")
    
    def _row_entry_image(self, parent, label, var, placeholder=""):
        """Create a row with label and entry"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=4)
        
        ttk.Label(row, text=label, width=25).pack(side="left")
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side="left", fill="x", expand=True, padx=6)
        if placeholder:
            entry.insert(0, placeholder)
        return entry
    
    def _row_scale_image(self, parent, label, var, from_, to_, resolution=0.01):
        """Create a row with label and scale"""
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=3)
        
        ttk.Label(row, text=label, width=15).pack(side="left")
        val_lbl = ttk.Label(row, text=str(var.get()))
        val_lbl.pack(side="right")
        
        scale = ttk.Scale(
            row, from_=from_, to=to_, variable=var, orient="horizontal",
            command=lambda v: val_lbl.config(
                text=f"{float(v):.2f}" if resolution != 1 else f"{int(float(v))}"
            )
        )
        scale.pack(side="left", fill="x", expand=True, padx=6)
        return scale
    
    def _on_preset_change_image(self, event=None):
        """Handle preset selection change"""
        preset = self.preset_var.get()
        if preset == "canva":
            self.canva_params_frame.pack(fill="x", padx=8, pady=4)
            self._append_image("🎨 Preset 'canva' sélectionné - Paramètres ajustables affichés")
        else:
            self.canva_params_frame.pack_forget()
            self._append_image("🎨 Preset 'none' sélectionné - Aucune amélioration appliquée")
    
    def _append_image(self, text):
        """Append text to image log"""
        self.log_text_image.insert("end", text + "\n")
        self.log_text_image.see("end")
        self.update_idletasks()
    
    def browse_input_image(self):
        """Browse for input parent folder"""
        folder = filedialog.askdirectory(title="Choisir le dossier parent (contenant les sous-dossiers)")
        if folder:
            self.input_folder_var.set(folder)
            self._append_image(f"📂 Dossier parent sélectionné: {folder}")
    
    def browse_output_image(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if folder:
            self.output_folder_var.set(folder)
            self._append_image(f"📁 Dossier de sortie sélectionné: {folder}")
    
    def browse_csv_image(self):
        """Browse for CSV file"""
        file_path = filedialog.askopenfilename(
            title="Choisir le fichier CSV",
            filetypes=[("Fichiers CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if file_path:
            self.csv_path_var.set(file_path)
            self._append_image(f"📄 Fichier CSV sélectionné: {file_path}")
            
            # Load and preview CSV data
            try:
                from enhance_canva_like import read_csv_data
                csv_data = read_csv_data(file_path)
                if csv_data:
                    self._append_image(f"✅ CSV chargé: {len(csv_data)} entrées trouvées")
                    self._append_image("📋 Aperçu des données CSV:")
                    sample = list(csv_data.items())[:3]  # Show first 3 entries
                    for sku_orig, row in sample:
                        sku_kyopa = row.get('sku kyopa', '')
                        title = row.get('title', '')
                        tags = row.get('tags', '')
                        self._append_image(f"  • '{sku_orig}' → SKU: '{sku_kyopa}'")
                        self._append_image(f"    Titre: '{title}'")
                        self._append_image(f"    Tags: '{tags}'")
                else:
                    self._append_image("⚠️ Aucune donnée valide trouvée dans le CSV")
            except Exception as e:
                self._append_image(f"⚠️ Erreur lors du chargement du CSV: {e}")
    
    def start_processing_image(self):
        """Start the image processing"""
        input_folder = self.input_folder_var.get().strip()
        output_folder = self.output_folder_var.get().strip()
        csv_path = self.csv_path_var.get().strip()
        resize_width_str = self.resize_width_var.get().strip()
        
        # Validation
        if not input_folder or not output_folder or not csv_path:
            messagebox.showerror("Erreur", "Veuillez sélectionner tous les chemins requis:\n• Dossier parent\n• Dossier de sortie\n• Fichier CSV")
            return
        
        if not os.path.isdir(input_folder):
            messagebox.showerror("Erreur", f"Le dossier parent '{input_folder}' n'existe pas ou n'est pas un dossier.")
            return
        
        if not Path(csv_path).exists():
            messagebox.showerror("Erreur", f"Le fichier CSV '{csv_path}' n'existe pas.")
            return
        
        # Parse resize width
        resize_width = None
        if resize_width_str:
            try:
                resize_width = int(resize_width_str)
            except ValueError:
                messagebox.showerror("Erreur", "La largeur de redimensionnement doit être un nombre entier.")
                return
        
        # Update UI
        self.start_btn_image.config(state="disabled")
        self.stop_btn_image.config(state="normal")
        self._stop_requested = False
        
        self._append_image("\n" + "="*60)
        self._append_image("🚀 DÉMARRAGE DU TRAITEMENT")
        self._append_image("="*60)
        self._append_image(f"📂 Dossier parent: {input_folder}")
        self._append_image(f"📁 Dossier de sortie: {output_folder}")
        self._append_image(f"📄 Fichier CSV: {csv_path}")
        if resize_width:
            self._append_image(f"📏 Largeur de redimensionnement: {resize_width}px")
        self._append_image("")
        
        # Get preset selection
        preset = self.preset_var.get()
        
        # Start processing in background thread
        def worker():
            try:
                self._run_processing_image(input_folder, output_folder, csv_path, resize_width, preset)
            except Exception as e:
                self._append_image(f"❌ Erreur inattendue: {e}")
            finally:
                self.start_btn_image.config(state="normal")
                self.stop_btn_image.config(state="disabled")
                self._append_image("\n🎉 Traitement terminé!")
        
        self._worker = Thread(target=worker, daemon=True)
        self._worker.start()
    
    def _run_processing_image(self, input_folder, output_folder, csv_path, resize_width, preset):
        """Run the actual image processing"""
        try:
            from enhance_canva_like import (
                read_csv_data, convert_and_enhance, remove_metadata_from_folder,
                set_metadata_with_exiftool, clean_filename
            )
            
            # Load CSV data
            self._append_image("📄 Chargement des données CSV...")
            csv_data = read_csv_data(csv_path)
            if not csv_data:
                self._append_image("❌ Erreur: Aucune donnée valide trouvée dans le fichier CSV")
                return
            
            self._append_image(f"✅ CSV chargé: {len(csv_data)} entrées trouvées")
            self._append_image(f"🎨 Preset sélectionné: {preset}")
            
            if preset == "canva":
                self._append_image("📊 Paramètres Canva:")
                self._append_image(f"  - Luminosité: {self.brightness_var.get():.2f}")
                self._append_image(f"  - Contraste: {self.contrast_var.get():.2f}")
                self._append_image(f"  - Couleur: {self.color_var.get():.2f}")
                self._append_image(f"  - Netteté: {self.sharpness_var.get():.2f}")
                self._append_image(f"  - Gamma: {self.gamma_var.get():.2f}")
                self._append_image(f"  - Gain Rouge: {self.r_gain_var.get():.2f}")
                self._append_image(f"  - Gain Vert: {self.g_gain_var.get():.2f}")
                self._append_image(f"  - Gain Bleu: {self.b_gain_var.get():.2f}")
            
            input_path = Path(input_folder)
            output_path = Path(output_folder)
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Find all subfolders
            subfolders = [f for f in input_path.iterdir() if f.is_dir()]
            
            if not subfolders:
                self._append_image("❌ Aucun sous-dossier trouvé dans le dossier parent!")
                return
            
            self._append_image(f"📂 {len(subfolders)} sous-dossiers trouvés:")
            for i, subfolder in enumerate(subfolders, 1):
                self._append_image(f"  {i}. {subfolder.name}")
            self._append_image("")
            
            # Process each subfolder
            total_processed = 0
            total_failed = 0
            processed_folders = []
            failed_folders = []
            csv_matched_folders = []
            csv_unmatched_folders = []
            
            for i, subfolder in enumerate(subfolders, 1):
                if self._stop_requested:
                    self._append_image("⏹️ Arrêt demandé par l'utilisateur")
                    break
                    
                self._append_image(f"🔄 Traitement du sous-dossier {i}/{len(subfolders)}: {subfolder.name}")
                
                # Look for CSV match
                subfolder_name_lower = subfolder.name.strip().lower()
                csv_match = csv_data.get(subfolder_name_lower)
                
                if csv_match:
                    self._append_image(f"✅ Correspondance CSV trouvée pour '{subfolder.name}'")
                    csv_title = csv_match.get('title', '').strip()
                    csv_sku_kyopa = csv_match.get('sku kyopa', '').strip()
                    csv_tags = csv_match.get('tags', '').strip()
                    
                    self._append_image(f"📋 Données CSV:")
                    self._append_image(f"  - Titre: '{csv_title}'")
                    self._append_image(f"  - SKU Kyopa: '{csv_sku_kyopa}'")
                    self._append_image(f"  - Tags: '{csv_tags}'")
                    
                    output_folder_name = csv_sku_kyopa if csv_sku_kyopa else subfolder.name
                    csv_matched_folders.append(subfolder.name)
                else:
                    self._append_image(f"⚠️ Aucune correspondance CSV trouvée pour '{subfolder.name}' - utilisation du nom original")
                    output_folder_name = subfolder.name
                    csv_title = ""
                    csv_tags = ""
                    csv_unmatched_folders.append(subfolder.name)
                
                # Create output subfolder
                output_subfolder = output_path / output_folder_name
                output_subfolder.mkdir(parents=True, exist_ok=True)
                self._append_image(f"📁 Dossier de sortie: {output_subfolder}")
                
                # Step 1: Convert and enhance
                self._append_image("🔄 Étape 1: Conversion et amélioration des images...")
                
                # Prepare canva parameters if preset is canva
                canva_params = None
                if preset == "canva":
                    canva_params = {
                        'brightness': self.brightness_var.get(),
                        'contrast': self.contrast_var.get(),
                        'color': self.color_var.get(),
                        'sharpness': self.sharpness_var.get(),
                        'gamma': self.gamma_var.get(),
                        'r_gain': self.r_gain_var.get(),
                        'g_gain': self.g_gain_var.get(),
                        'b_gain': self.b_gain_var.get()
                    }
                
                enhanced_folder = convert_and_enhance(str(subfolder), str(output_subfolder), resize_width, preset=preset, canva_params=canva_params)
                
                if enhanced_folder and Path(enhanced_folder).exists():
                    self._append_image(f"✅ Étape 1 terminée pour '{subfolder.name}'!")
                    
                    # Step 2: Remove metadata
                    self._append_image("🗑️ Étape 2: Suppression des métadonnées...")
                    result = remove_metadata_from_folder(enhanced_folder, self._append_image)
                    
                    if result["success"]:
                        self._append_image(f"✅ Suppression des métadonnées terminée!")
                        self._append_image(f"🗑️ Métadonnées supprimées de {result['processed_count']} fichiers")
                        
                        # Step 3: Apply CSV metadata and rename
                        if csv_match and csv_title:
                            self._append_image("📄 Étape 3: Application des métadonnées CSV et renommage...")
                            
                            # Process tags
                            if csv_tags:
                                tags = [t.strip() for t in (csv_tags.split(",") if ',' in csv_tags else csv_tags.split()) if t.strip()]
                            else:
                                tags = []
                            
                            # Clean title for filename
                            safe_title = clean_filename(csv_title)
                            
                            # Find all JPEG files and rename them
                            jpeg_files = list(Path(enhanced_folder).glob("*.jpg")) + list(Path(enhanced_folder).glob("*.jpeg"))
                            
                            processed_images = 0
                            for j, jpeg_file in enumerate(jpeg_files, 1):
                                try:
                                    # Create new filename with sequential number
                                    new_filename = f"{safe_title}_{j}.jpg"
                                    new_path = Path(enhanced_folder) / new_filename
                                    
                                    # Rename the file
                                    jpeg_file.rename(new_path)
                                    self._append_image(f"📝 Renommé: {jpeg_file.name} -> {new_filename}")
                                    
                                    # Set metadata using ExifTool
                                    success = set_metadata_with_exiftool(new_path, csv_title, tags, "5", self._append_image)
                                    
                                    if success:
                                        self._append_image(f"✅ Métadonnées appliquées: {new_filename}")
                                        processed_images += 1
                                    else:
                                        self._append_image(f"⚠️ Échec de l'application des métadonnées: {new_filename}")
                                        
                                except Exception as e:
                                    self._append_image(f"❌ Erreur lors du traitement de {jpeg_file.name}: {e}")
                            
                            self._append_image(f"✅ Métadonnées CSV appliquées à {processed_images} images")
                        else:
                            self._append_image(f"ℹ️ Aucune donnée CSV à appliquer pour '{subfolder.name}'")
                        
                        total_processed += result['processed_count']
                        processed_folders.append(subfolder.name)
                    else:
                        self._append_image(f"⚠️ Problème avec la suppression des métadonnées pour '{subfolder.name}', mais la conversion a réussi.")
                        total_failed += 1
                        failed_folders.append(subfolder.name)
                else:
                    self._append_image(f"❌ Étape 1 échouée pour '{subfolder.name}'!")
                    total_failed += 1
                    failed_folders.append(subfolder.name)
                
                self._append_image("-" * 40)
            
            # Final summary
            self._append_image("\n🎉 TRAITEMENT PAR LOTS TERMINÉ!")
            self._append_image("📊 Résumé:")
            self._append_image(f"  📂 Total des sous-dossiers traités: {len(subfolders)}")
            self._append_image(f"  ✅ Traités avec succès: {len(processed_folders)}")
            self._append_image(f"  ❌ Échoués: {len(failed_folders)}")
            self._append_image(f"  🖼️ Total des images traitées: {total_processed}")
            self._append_image(f"  📄 Dossiers avec correspondance CSV: {len(csv_matched_folders)}")
            self._append_image(f"  ⚠️ Dossiers sans correspondance CSV: {len(csv_unmatched_folders)}")
            
            if csv_matched_folders:
                self._append_image(f"\n✅ Dossiers avec correspondance CSV:")
                for folder in csv_matched_folders:
                    self._append_image(f"  - {folder}")
            
            if csv_unmatched_folders:
                self._append_image(f"\n⚠️ Dossiers sans correspondance CSV:")
                for folder in csv_unmatched_folders:
                    self._append_image(f"  - {folder}")
                    
        except Exception as e:
            self._append_image(f"❌ Erreur lors du traitement: {e}")
    
    def stop_processing_image(self):
        """Stop the image processing"""
        self._stop_requested = True
        self._append_image("⏹️ Arrêt demandé... Le traitement s'arrêtera après le fichier en cours.")
        messagebox.showinfo("Info", "Le traitement s'arrêtera après avoir terminé le fichier en cours.")

    def open_video_processor(self):
        """Lancer l'outil de traitement vidéo (EXE si dispo, sinon script)."""
        try:
            base = Path(__file__).parent
            candidates = [
                base / "dist" / "BatchVideoProcessor" / "BatchVideoProcessor.exe",
                base / "dist" / "BatchVideoProcessor.exe",
                base / "build" / "BatchVideoProcessor" / "BatchVideoProcessor.exe",
                base / "batch_video_processor.py",
                base / "batchprocessorCopy.py",
            ]
            target = next((p for p in candidates if p.exists()), None)
            if not target:
                messagebox.showerror("Introuvable", 
                                     "Aucun exécutable/script de traitement vidéo trouvé (dist/build ou .py).")
                return
            # Lancer détaché
            if target.suffix.lower() == ".exe":
                subprocess.Popen([str(target)])
            else:
                subprocess.Popen([sys.executable, str(target)])
            self.update_status("Traitement Vidéos lancé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le traitement vidéo:\n{e}")
    
    def update_status(self, message):
        """Mettre à jour la barre de statut"""
        self.status_bar.config(text=message)
        self.update_idletasks()
    
    def log_message(self, message, log_widget):
        """Ajouter un message au log"""
        log_widget.config(state='normal')
        log_widget.insert('end', f"{message}\n")
        log_widget.see('end')
        log_widget.config(state='disabled')
        self.update_idletasks()
    
    def auto_connect(self):
        """Connexion automatique au démarrage"""
        def connect_worker():
            try:
                self.update_status("Connexion à Google...")
                self.sheets_manager.connect()
                self.drive_manager.connect()
                
                self.connection_status.config(text="✅ Connecté à Google")
                self.update_status("Prêt - Connecté à Google")
                
                # Charger les données
                self.refresh_data()
                # Rafraîchir le dashboard (si disponible)
                try:
                    if hasattr(self, 'dashboard_tab'):
                        self.dashboard_tab.refresh_stats()
                except Exception:
                    pass
                
            except Exception as e:
                self.connection_status.config(text="❌ Erreur de connexion")
                self.update_status(f"Erreur: {e}")
                messagebox.showerror("Erreur de connexion", 
                                   f"Impossible de se connecter à Google:\n{e}")
        
        threading.Thread(target=connect_worker, daemon=True).start()
    
    def refresh_data(self):
        """Actualiser les données du spreadsheet"""
        def refresh_worker():
            try:
                self.update_status("Chargement des données...")
                
                # Récupérer les données
                data = self.sheets_manager.get_worksheet_data(WORKSHEET_NAME)
                
                if not data:
                    self.update_status("Aucune donnée trouvée")
                    return
                
                self.worksheet_data = data
                self.filtered_data = data.copy()
                
                # Mettre à jour l'affichage
                self.update_table_display()
                
                self.update_status(f"✅ {len(data)-1} lignes chargées")
                
            except Exception as e:
                self.update_status(f"Erreur: {e}")
                messagebox.showerror("Erreur", f"Impossible de charger les données:\n{e}")
        
        threading.Thread(target=refresh_worker, daemon=True).start()
    
    def update_table_display(self):
        """Mettre à jour l'affichage du tableau"""
        if not self.filtered_data:
            return
        
        # Effacer le contenu existant
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.filtered_data:
            return
        
        # Configuration des colonnes
        if len(self.filtered_data) > 0:
            headers = self.filtered_data[0]
            self.tree['columns'] = list(range(len(headers)))
            self.tree['show'] = 'headings'
            
            # Configuration des en-têtes
            for i, header in enumerate(headers):
                self.tree.heading(i, text=header)
                self.tree.column(i, width=140,minwidth=140, anchor='w')
            
            # Ajout des données
            for row in self.filtered_data[1:]:
                # Compléter la ligne si nécessaire
                padded_row = row + [''] * (len(headers) - len(row))
                self.tree.insert('', 'end', values=padded_row)
    
    def filter_errors(self):
        """Filtrer les lignes avec statut 'erreur'"""
        if not self.worksheet_data:
            messagebox.showwarning("Attention", "Aucune donnée chargée")
            return
        
        try:
            headers = self.worksheet_data[0]
            # Trouver la colonne Status (insensible à la casse)
            status_col = None
            for i, header in enumerate(headers):
                if 'status' in header.lower():
                    status_col = i
                    break
            
            if status_col is None:
                messagebox.showwarning("Attention", "Colonne 'Status' non trouvée")
                return
            
            # Filtrer les lignes
            self.filtered_data = [headers]
            for row in self.worksheet_data[1:]:
                if len(row) > status_col and 'erreur' in str(row[status_col]).lower():
                    self.filtered_data.append(row)
            
            self.update_table_display()
            self.update_status(f"Filtré: {len(self.filtered_data)-1} lignes avec erreurs")
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du filtrage: {e}")
    
    def show_all(self):
        """Afficher toutes les données"""
        self.filtered_data = self.worksheet_data.copy()
        self.update_table_display()
        self.update_status("Toutes les données affichées")
    
    def on_search(self, event=None):
        """Recherche dans les données"""
        if not self.worksheet_data:
            return
        
        search_term = self.search_var.get().lower()
        if not search_term:
            self.show_all()
            return
        
        headers = self.worksheet_data[0]
        self.filtered_data = [headers]
        
        for row in self.worksheet_data[1:]:
            # Rechercher dans toutes les colonnes
            if any(search_term in str(cell).lower() for cell in row):
                self.filtered_data.append(row)
        
        self.update_table_display()
        self.update_status(f"Recherche: {len(self.filtered_data)-1} résultats")
    
    def on_row_double_click(self, event):
        """Gérer le double-clic sur une ligne"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = self.tree.item(selection[0])
        values = item['values']
        
        if not values:
            return
        
        # Créer une fenêtre de détails
        self.show_row_details(values)
    
    def show_row_details(self, row_values):
        """Afficher les détails d'une ligne"""
        if not self.worksheet_data:
            return
        
        headers = self.worksheet_data[0]
        
        # Fenêtre de détails
        detail_window = tk.Toplevel(self)
        detail_window.title("Détails de l'entrée")
        detail_window.geometry("600x400")
        detail_window.configure(bg='#2b2b2b')
        
        # Frame principal
        main_frame = ttk.Frame(detail_window, style='Modern.TFrame')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Titre
        ttk.Label(main_frame, text="Détails de l'entrée", 
                 font=('Arial', 14, 'bold'), style='Modern.TLabel').pack(pady=(0, 20))
        
        # Scrollable frame pour les détails
        canvas = tk.Canvas(main_frame, bg='#2b2b2b')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Modern.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Afficher les données
        for i, (header, value) in enumerate(zip(headers, row_values)):
            frame = ttk.Frame(scrollable_frame, style='Modern.TFrame')
            frame.pack(fill='x', pady=5)
            
            ttk.Label(frame, text=f"{header}:", font=('Arial', 10, 'bold'), 
                     style='Modern.TLabel', width=20).pack(side='left', anchor='n')
            
            # Gestion spéciale pour les URLs
            if 'url' in header.lower() and value:
                link_frame = ttk.Frame(frame, style='Modern.TFrame')
                link_frame.pack(side='right', fill='x', expand=True)
                
                ttk.Label(link_frame, text=value, style='Modern.TLabel', 
                         foreground='#4a90e2').pack(side='top', anchor='w')
                
                ttk.Button(link_frame, text="Ouvrir", 
                          command=lambda url=value: webbrowser.open(url),
                          style='Modern.TButton').pack(side='top', anchor='w', pady=2)
            else:
                ttk.Label(frame, text=str(value), style='Modern.TLabel',
                         wraplength=400).pack(side='right', fill='x', expand=True)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Boutons d'action
        button_frame = ttk.Frame(detail_window, style='Modern.TFrame')
        button_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(button_frame, text="Fermer", command=detail_window.destroy,
                  style='Modern.TButton').pack(side='right')
    
    def select_drive_folder(self):
        """Sélectionner un dossier Google Drive de destination (sous-dossiers Etsy seulement)"""
        def select_worker():
            try:
                self.update_status("Récupération des sous-dossiers Etsy...")
                folders = self.drive_manager.list_etsy_subfolders()
                
                if not folders:
                    messagebox.showinfo("Info", "Aucun sous-dossier trouvé dans Photos Etsy Kyopadeco Shop")
                    return
                
                # Créer une fenêtre de sélection
                self.show_folder_selection_dialog(folders)
                
            except Exception as e:
                self.update_status(f"Erreur: {e}")
                messagebox.showerror("Erreur", f"Impossible de récupérer les sous-dossiers Etsy: {e}")
        
        threading.Thread(target=select_worker, daemon=True).start()
    
    def show_folder_selection_dialog(self, folders):
        """Afficher la boîte de dialogue de sélection de dossier"""
        dialog = tk.Toplevel(self)
        dialog.title("Sélectionner un dossier Google Drive")
        dialog.geometry("500x400")
        dialog.configure(bg='#2b2b2b')
        
        # Liste des dossiers
        frame = ttk.Frame(dialog, style='Modern.TFrame')
        frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        ttk.Label(frame, text="Choisissez un dossier de destination:", 
                 style='Modern.TLabel').pack(anchor='w', pady=(0, 10))
        
        listbox = tk.Listbox(frame, height=15)
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # Ajouter les dossiers
        for folder in folders:
            listbox.insert('end', f"{folder['name']} (ID: {folder['id']})")
        
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Boutons
        button_frame = ttk.Frame(dialog, style='Modern.TFrame')
        button_frame.pack(fill='x', padx=20, pady=10)
        
        def select_folder():
            selection = listbox.curselection()
            if selection:
                selected_folder = folders[selection[0]]
                self.selected_drive_folder = selected_folder
                self.drive_folder_var.set(f"{selected_folder['name']} (ID: {selected_folder['id']})")
                dialog.destroy()
            else:
                messagebox.showwarning("Attention", "Veuillez sélectionner un dossier")
        
        ttk.Button(button_frame, text="Sélectionner", command=select_folder,
                  style='Success.TButton').pack(side='right', padx=5)
        ttk.Button(button_frame, text="Annuler", command=dialog.destroy,
                  style='Modern.TButton').pack(side='right')
    
    def select_parent_folder(self):
        """Sélectionner le dossier parent contenant les sous-dossiers"""
        folder = filedialog.askdirectory(title="Choisir le dossier parent contenant les sous-dossiers")
        if folder:
            self.selected_parent_folder = Path(folder)
            self.parent_folder_var.set(str(self.selected_parent_folder))
            self.refresh_subfolders()
    
    def refresh_subfolders(self):
        """Actualiser la liste des sous-dossiers"""
        self.subfolders_listbox.delete(0, 'end')
        self.detected_subfolders.clear()
        
        if not self.selected_parent_folder or not self.selected_parent_folder.exists():
            return
        
        try:
            for item in self.selected_parent_folder.iterdir():
                if item.is_dir():
                    self.detected_subfolders.append(item)
                    self.subfolders_listbox.insert('end', item.name)
            
            self.update_status(f"{len(self.detected_subfolders)} sous-dossiers détectés")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la lecture du dossier: {e}")
    
    def modify_subfolders(self):
        """Modifier les noms des sous-dossiers"""
        if not self.detected_subfolders:
            messagebox.showwarning("Attention", "Aucun sous-dossier détecté")
            return
        
        # Créer une fenêtre de modification
        modify_window = tk.Toplevel(self)
        modify_window.title("Modifier les noms des sous-dossiers")
        modify_window.geometry("600x500")
        modify_window.configure(bg='#2b2b2b')
        
        main_frame = ttk.Frame(modify_window, style='Modern.TFrame')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        ttk.Label(main_frame, text="Modifier les noms des sous-dossiers:", 
                 font=('Arial', 12, 'bold'), style='Modern.TLabel').pack(pady=(0, 20))
        
        # Frame scrollable pour les entrées
        canvas = tk.Canvas(main_frame, bg='#2b2b2b')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='Modern.TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Variables pour stocker les nouveaux noms
        new_names = []
        
        for i, subfolder in enumerate(self.detected_subfolders):
            frame = ttk.Frame(scrollable_frame, style='Modern.TFrame')
            frame.pack(fill='x', pady=5)
            
            ttk.Label(frame, text=f"Dossier {i+1}:", style='Modern.TLabel', width=15).pack(side='left')
            
            name_var = tk.StringVar(value=subfolder.name)
            new_names.append(name_var)
            
            ttk.Entry(frame, textvariable=name_var, width=40).pack(side='left', padx=10)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Boutons
        button_frame = ttk.Frame(modify_window, style='Modern.TFrame')
        button_frame.pack(fill='x', padx=20, pady=10)
        
        def apply_changes():
            try:
                for i, (subfolder, new_name_var) in enumerate(zip(self.detected_subfolders, new_names)):
                    new_name = new_name_var.get().strip()
                    if new_name and new_name != subfolder.name:
                        new_path = subfolder.parent / new_name
                        if not new_path.exists():
                            subfolder.rename(new_path)
                            self.detected_subfolders[i] = new_path
                
                self.refresh_subfolders()
                modify_window.destroy()
                messagebox.showinfo("Succès", "Modifications appliquées avec succès")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la modification: {e}")
        
        ttk.Button(button_frame, text="Appliquer", command=apply_changes,
                  style='Success.TButton').pack(side='right', padx=5)
        ttk.Button(button_frame, text="Annuler", command=modify_window.destroy,
                  style='Modern.TButton').pack(side='right')
    
    def start_subfolder_upload(self):
        """Démarrer l'upload des sous-dossiers"""
        if not self.detected_subfolders:
            messagebox.showwarning("Attention", "Aucun sous-dossier détecté")
            return
        
        if not self.selected_drive_folder:
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de destination dans Google Drive")
            return
        
        def upload_worker():
            try:
                self.upload_progress.start()
                self.log_message("🚀 Début de l'upload des sous-dossiers...", self.upload_log)
                
                # Upload tous les sous-dossiers
                uploaded_folders = self.drive_manager.upload_subfolders_only(
                    self.selected_parent_folder,
                    self.selected_drive_folder['id'],
                    lambda msg: self.log_message(msg, self.upload_log)
                )
                
                # Mettre à jour Google Sheets pour chaque dossier uploadé
                for folder_info in uploaded_folders:
                    self.update_sheets_with_url(folder_info['name'], folder_info['url'])
                    self.log_message(f"📋 Mis à jour dans Sheets: {folder_info['name']}", self.upload_log)
                
                self.log_message(f"🎉 Upload terminé! {len(uploaded_folders)} sous-dossiers uploadés", self.upload_log)
                messagebox.showinfo("Succès", f"{len(uploaded_folders)} sous-dossiers uploadés avec succès!")
                
            except Exception as e:
                self.log_message(f"❌ Erreur: {e}", self.upload_log)
                messagebox.showerror("Erreur d'upload", f"Erreur lors de l'upload: {e}")
            finally:
                self.upload_progress.stop()
        
        threading.Thread(target=upload_worker, daemon=True).start()
    
    def update_sheets_with_url(self, folder_name, folder_url):
        """Mettre à jour Google Sheets avec l'URL du dossier dans la colonne 'Drive Folder URL Kyopa'"""
        try:
            if not self.worksheet_data:
                self.log_message("⚠️ Aucune donnée de feuille chargée", self.upload_log)
                return
            
            headers = self.worksheet_data[0]
            
            # Trouver les colonnes
            sku_kyopa_col = None
            drive_url_kyopa_col = None
            
            for i, header in enumerate(headers):
                header_clean = str(header).strip().lower()
                if 'sku kyopa' in header_clean:
                    sku_kyopa_col = i
                elif 'drive folder url kyopa' in header_clean:
                    drive_url_kyopa_col = i
            
            if sku_kyopa_col is None:
                self.log_message(f"⚠️ Colonne 'sku kyopa' non trouvée pour {folder_name}", self.upload_log)
                return
            
            if drive_url_kyopa_col is None:
                self.log_message(f"⚠️ Colonne 'Drive Folder URL Kyopa' non trouvée pour {folder_name}", self.upload_log)
                return
            
            # Chercher la ligne correspondante par nom de dossier
            for row_index, row in enumerate(self.worksheet_data[1:], start=2):
                if len(row) > sku_kyopa_col:
                    sku_value = str(row[sku_kyopa_col]).strip()
                    if sku_value.lower() == folder_name.lower():
                        # Mettre à jour la cellule Drive Folder URL Kyopa
                        cell_range = f"{chr(65 + drive_url_kyopa_col)}{row_index}"
                        self.sheets_manager.update_cell(WORKSHEET_NAME, cell_range, folder_url)
                        self.log_message(f"✅ Lien mis à jour pour {folder_name} dans Drive Folder URL Kyopa", self.upload_log)
                        
                        # Mettre à jour les données localement
                        if len(self.worksheet_data[row_index-1]) <= drive_url_kyopa_col:
                            self.worksheet_data[row_index-1].extend([''] * (drive_url_kyopa_col + 1 - len(self.worksheet_data[row_index-1])))
                        self.worksheet_data[row_index-1][drive_url_kyopa_col] = folder_url
                        return
            
            self.log_message(f"⚠️ Aucune ligne trouvée avec sku kyopa = '{folder_name}'", self.upload_log)
            
        except Exception as e:
            self.log_message(f"❌ Erreur mise à jour Sheets: {e}", self.upload_log)
    
    def select_download_folder(self):
        """Sélectionner le dossier de téléchargement"""
        folder = filedialog.askdirectory(title="Choisir le dossier de téléchargement")
        if folder:
            self.download_path_var.set(folder)
    
    def download_error_folders(self):
        """Télécharger les dossiers avec statut erreur depuis la colonne 'Drive Folder URL'"""
        if not self.download_path_var.get():
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de téléchargement")
            return
        
        def download_worker():
            try:
                self.download_progress.start()
                
                if not self.worksheet_data:
                    self.log_message("❌ Aucune donnée chargée", self.download_log)
                    return
                
                headers = self.worksheet_data[0]
                
                # Chercher exactement la colonne "Drive Folder URL"
                drive_url_col = None
                status_col = None
                
                for i, header in enumerate(headers):
                    header_clean = str(header).strip().lower()
                    if header_clean == "drive folder url":
                        drive_url_col = i
                    elif "status" in header_clean:
                        status_col = i
                
                if drive_url_col is None:
                    self.log_message("❌ Colonne 'Drive Folder URL' non trouvée", self.download_log)
                    return
                if status_col is None:
                    self.log_message("❌ Colonne 'Status' non trouvée", self.download_log)
                    return
                
                # Log pour vérification
                self.log_message(f"ℹ️ Utilisation de la colonne '{headers[drive_url_col]}' (index {drive_url_col})", self.download_log)
                
                download_path = Path(self.download_path_var.get())
                error_count = 0
                success_count = 0
                
                for row in self.worksheet_data[1:]:
                    if len(row) <= max(status_col, drive_url_col):
                        continue
                        
                    status = str(row[status_col]).lower()
                    drive_folder_url = str(row[drive_url_col]).strip()
                    
                    if 'erreur' in status and drive_folder_url:
                        error_count += 1
                        try:
                            self.log_message(f"Téléchargement depuis: {drive_folder_url}", self.download_log)
                            
                            success = self.drive_manager.download_folder_by_url(
                                drive_folder_url,
                                download_path,
                                lambda msg: self.log_message(msg, self.download_log)
                            )
                            
                            if success:
                                success_count += 1
                                self.log_message(f"✅ Téléchargement réussi: {drive_folder_url}", self.download_log)
                            
                        except Exception as e:
                            self.log_message(f"❌ Erreur: {e}", self.download_log)
                
                self.log_message(f"🎉 Terminé: {success_count}/{error_count} téléchargements réussis", self.download_log)
                
            except Exception as e:
                self.log_message(f"❌ Erreur générale: {e}", self.download_log)
            finally:
                self.download_progress.stop()
        
        threading.Thread(target=download_worker, daemon=True).start()
    
    def download_by_drive_folder_url(self):
        """Télécharger les dossiers depuis 'Drive Folder URL' quand 'Drive Folder URL Kyopa' est null"""
        if not self.download_path_var.get():
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de téléchargement")
            return
        
        def download_worker():
            try:
                self.download_progress.start()
                
                if not self.worksheet_data:
                    self.log_message("❌ Aucune donnée chargée", self.download_log)
                    return
                
                headers = self.worksheet_data[0]
                
                # Chercher les colonnes
                drive_folder_url_col = None
                drive_folder_url_kyopa_col = None
                
                for i, header in enumerate(headers):
                    header_clean = str(header).strip().lower()
                    if header_clean == "drive folder url":
                        drive_folder_url_col = i
                    elif "drive folder url kyopa" in header_clean:
                        drive_folder_url_kyopa_col = i
                
                if drive_folder_url_col is None:
                    self.log_message("❌ Colonne 'Drive Folder URL' non trouvée", self.download_log)
                    return
                
                if drive_folder_url_kyopa_col is None:
                    self.log_message("❌ Colonne 'Drive Folder URL Kyopa' non trouvée", self.download_log)
                    return
                
                # Log pour vérification
                self.log_message(f"ℹ️ Recherche lignes où '{headers[drive_folder_url_kyopa_col]}' est vide et '{headers[drive_folder_url_col]}' contient un lien", self.download_log)
                
                download_path = Path(self.download_path_var.get())
                candidates_count = 0
                success_count = 0
                
                for row_index, row in enumerate(self.worksheet_data[1:], start=2):
                    if len(row) <= max(drive_folder_url_col, drive_folder_url_kyopa_col):
                        continue
                    
                    # Vérifier que Drive Folder URL Kyopa est vide/null
                    kyopa_url = str(row[drive_folder_url_kyopa_col]).strip() if len(row) > drive_folder_url_kyopa_col else ""
                    drive_folder_url = str(row[drive_folder_url_col]).strip() if len(row) > drive_folder_url_col else ""
                    
                    # Condition: Drive Folder URL Kyopa est vide ET Drive Folder URL contient un lien
                    if (not kyopa_url or kyopa_url.lower() in ['', 'null', 'none']) and drive_folder_url and 'drive.google.com' in drive_folder_url:
                        candidates_count += 1
                        try:
                            self.log_message(f"Téléchargement depuis Drive Folder URL: {drive_folder_url}", self.download_log)
                            
                            success = self.drive_manager.download_folder_by_url(
                                drive_folder_url,
                                download_path,
                                lambda msg: self.log_message(msg, self.download_log)
                            )
                            
                            if success:
                                success_count += 1
                                self.log_message(f"✅ Téléchargement réussi: {drive_folder_url}", self.download_log)
                            
                        except Exception as e:
                            self.log_message(f"❌ Erreur ligne {row_index}: {e}", self.download_log)
                
                self.log_message(f"🎉 Terminé: {success_count}/{candidates_count} téléchargements réussis", self.download_log)
                
                if candidates_count == 0:
                    self.log_message("ℹ️ Aucune ligne trouvée avec Drive Folder URL Kyopa vide et Drive Folder URL rempli", self.download_log)
                
            except Exception as e:
                self.log_message(f"❌ Erreur générale: {e}", self.download_log)
            finally:
                self.download_progress.stop()
        
        threading.Thread(target=download_worker, daemon=True).start()
    
   
    def reconnect(self):
        """Reconnecter à Google"""
        # Supprimer le token pour forcer une nouvelle authentification
        if os.path.exists(self.creds_manager.token_file):
            os.remove(self.creds_manager.token_file)
        
        self.auto_connect()
    
    def open_sheets(self):
        """Ouvrir Google Sheets"""
        if self.sheets_manager.spreadsheet_id:
            url = f"https://docs.google.com/spreadsheets/d/{self.sheets_manager.spreadsheet_id}"
            webbrowser.open(url)
        else:
            messagebox.showwarning("Attention", "ID du spreadsheet non trouvé")

def main():
    """Fonction principale"""
    if not GOOGLE_AVAILABLE:
        messagebox.showerror("Erreur", 
                           "Librairies Google non disponibles.\n"
                           "Installez avec: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return
    
    app = ModernGoogleDriveApp()
    app.mainloop()

if __name__ == "__main__":
    main()