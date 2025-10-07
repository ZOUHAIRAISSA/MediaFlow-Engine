import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import threading


class DashboardTab(ttk.Frame):
    """Tableau de bord pour la feuille 'stock Etsy Listing'.

    Affiche des cartes de statistiques cliquables et permet d'ouvrir
    une vue tabulaire des lignes filtrées ou de toutes les lignes.
    """

    def __init__(self, parent, sheets_manager, update_status_callback, stock_worksheet_name="stock Etsy Listing", drive_manager=None):
        super().__init__(parent)
        self.sheets_manager = sheets_manager
        self.update_status = update_status_callback
        self.stock_worksheet_name = stock_worksheet_name
        self.drive_manager = drive_manager

        # Données en cache pour éviter des appels multiples inutiles
        self._last_data = None

        # Couleurs des cartes (modernes)
        self.colors = {
            "inserer": "#2ecc71",   # vert
            "erreur": "#e74c3c",    # rouge
            "supprimer": "#95a5a6", # gris
            "attente": "#f39c12",   # orange
        }

        self._build_ui()

    def _build_ui(self):
        # En-tête
        header = ttk.Label(self, text="📈 Dashboard - stock Etsy Listing", font=('Arial', 14, 'bold'))
        header.pack(pady=(12, 6))

        # Cartes
        cards_frame = ttk.Frame(self)
        cards_frame.pack(fill='x', padx=12, pady=8)

        self.inserer_var = tk.StringVar(value="-")
        self.erreur_var = tk.StringVar(value="-")
        self.supprimer_var = tk.StringVar(value="-")
        self.attente_var = tk.StringVar(value="-")

        self._create_card(cards_frame, "Insérer", self.inserer_var, self.colors["inserer"],
                          lambda: self._show_filtered_embedded(lambda s: "inser" in s))
        self._create_card(cards_frame, "Erreur", self.erreur_var, self.colors["erreur"],
                          lambda: self._show_filtered_embedded(lambda s: "erreur" in s))
        self._create_card(cards_frame, "Supprimer", self.supprimer_var, self.colors["supprimer"],
                          lambda: self._show_filtered_embedded(lambda s: "supprim" in s))
        self._create_card(cards_frame, "En attente", self.attente_var, self.colors["attente"],
                          lambda: self._show_filtered_embedded(lambda s: "attent" in s))

        # Actions
        action_frame = ttk.Frame(self)
        action_frame.pack(fill='x', padx=12, pady=(0, 8))

        ttk.Button(action_frame, text="🔄 Actualiser statistiques", command=self.refresh_stats).pack(side='left')
        ttk.Button(action_frame, text="📥 Télécharger dossier origin", style='Warning.TButton', command=self._download_selected_origin).pack(side='left', padx=8)
        ttk.Button(action_frame, text="📥 Télécharger dossier Kyopa", style='Success.TButton', command=self._download_selected_kyopa).pack(side='left')
        ttk.Button(action_frame, text="👁️ Voir tout (stock)", command=self._show_all_embedded).pack(side='right')

        # Tableau intégré (aperçu / résultats)
        self.table_container = ttk.LabelFrame(self, text="Aperçu des données (stock)")
        self.table_container.pack(fill='both', expand=True, padx=12, pady=(0, 12))

        self.table_tree = ttk.Treeview(self.table_container, height=12)
        self.table_vsb = ttk.Scrollbar(self.table_container, orient='vertical', command=self._on_vsb)
        self.table_hsb = ttk.Scrollbar(self.table_container, orient='horizontal', command=self.table_tree.xview)
        self.table_tree.configure(yscrollcommand=self.table_vsb.set, xscrollcommand=self.table_hsb.set)

        self.table_tree.grid(row=0, column=0, sticky='nsew')
        self.table_vsb.grid(row=0, column=1, sticky='ns')
        self.table_hsb.grid(row=1, column=0, sticky='ew')

        self.table_container.grid_rowconfigure(0, weight=1)
        self.table_container.grid_columnconfigure(0, weight=1)

        # Lazy loading state
        self._current_table_headers = []
        self._current_table_rows = []
        self._table_inserted_count = 0
        self._table_batch_size = 200

        # Scrolling triggers
        self.table_tree.bind('<MouseWheel>', self._on_mousewheel)
        self.table_tree.bind('<KeyRelease-Up>', self._on_key_nav)
        self.table_tree.bind('<KeyRelease-Down>', self._on_key_nav)
        self.table_tree.bind('<Double-1>', self._on_table_double_click)

    def _create_card(self, parent, title, value_var, color, on_click):
        # Carte colorée (conteneur TK pour gérer le bg)
        outer = tk.Frame(parent, bg=color, bd=0, highlightthickness=0)
        outer.pack(side='left', fill='x', expand=True, padx=6)

        inner = tk.Frame(outer, bg="#2b2b2b")
        inner.pack(fill='both', expand=True, padx=1, pady=1)

        title_lbl = tk.Label(inner, text=title, bg="#2b2b2b", fg="#ffffff", font=('Arial', 10, 'bold'))
        title_lbl.pack(anchor='w', padx=12, pady=(10, 0))

        value_lbl = tk.Label(inner, textvariable=value_var, bg="#2b2b2b", fg="#ffffff", font=('Arial', 28, 'bold'))
        value_lbl.pack(anchor='w', padx=12, pady=(2, 10))

        btn = ttk.Button(inner, text="Voir", command=on_click)
        btn.pack(anchor='e', padx=12, pady=(0, 12))

    def refresh_stats(self):
        """Rafraîchir les statistiques (threadé)."""
        def worker():
            try:
                self.update_status("Chargement statistiques (stock Etsy Listing)...")
                data = self.sheets_manager.get_worksheet_data(self.stock_worksheet_name)
                self._last_data = data

                if not data:
                    self._set_counts(0, 0, 0, 0)
                    self.update_status("Aucune donnée 'stock Etsy Listing'")
                    return

                headers = data[0] if data else []
                status_col = None
                for i, h in enumerate(headers):
                    if 'status' in str(h).lower():
                        status_col = i
                        break

                if status_col is None:
                    self._set_counts(0, 0, 0, 0)
                    self.update_status("Colonne 'Status' introuvable dans 'stock Etsy Listing'")
                    return

                inserer = erreur = supprimer = attente = 0
                for row in data[1:]:
                    if len(row) <= status_col:
                        continue
                    s = str(row[status_col]).strip().lower()
                    if 'inser' in s:
                        inserer += 1
                    if 'erreur' in s:
                        erreur += 1
                    if 'supprim' in s:
                        supprimer += 1
                    if 'attent' in s:
                        attente += 1

                self._set_counts(inserer, erreur, supprimer, attente)
                # Aperçu: afficher les 10 premières lignes sous les boutons
                headers = data[0]
                rows = data[1:]
                # Indices de feuille correspondants (2..)
                indices = list(range(2, 2 + len(rows)))
                self.after(0, lambda: self._set_table_dataset(headers, rows, indices, preview=True))
                self.update_status("Statistiques (stock) mises à jour")
            except Exception as e:
                self.update_status(f"Erreur statistiques: {e}")
                self._set_counts('-', '-', '-', '-')

        threading.Thread(target=worker, daemon=True).start()

    def _set_counts(self, inserer, erreur, supprimer, attente):
        self.inserer_var.set(str(inserer))
        self.erreur_var.set(str(erreur))
        self.supprimer_var.set(str(supprimer))
        self.attente_var.set(str(attente))

    def _show_filtered_embedded(self, predicate):
        """Afficher dans le tableau intégré les lignes filtrées selon predicate(status_str)."""
        if not self._last_data:
            messagebox.showinfo("Info", "Veuillez d'abord actualiser les statistiques.")
            return

        headers = self._last_data[0]
        status_col = None
        for i, h in enumerate(headers):
            if 'status' in str(h).lower():
                status_col = i
                break
        if status_col is None:
            messagebox.showwarning("Attention", "Colonne 'Status' introuvable.")
            return

        rows = []
        indices = []
        for idx, row in enumerate(self._last_data[1:], start=2):
            if len(row) > status_col and predicate(str(row[status_col]).strip().lower()):
                rows.append(row)
                indices.append(idx)

        # Utiliser lazy loading si beaucoup de lignes
        self._set_table_dataset(headers, rows, indices, preview=False)
        self.update_status("Filtre appliqué")

    def _show_all_embedded(self):
        """Afficher toutes les lignes (lazy loading) dans le tableau intégré."""
        def worker():
            try:
                self.update_status("Chargement (tout le stock)...")
                data = self._last_data or self.sheets_manager.get_worksheet_data(self.stock_worksheet_name)
                self._last_data = data
                if not data:
                    self.update_status("Aucune donnée")
                    return
                headers = data[0]
                rows = data[1:]
                indices = list(range(2, 2 + len(rows)))
                self.after(0, lambda: self._set_table_dataset(headers, rows, indices, preview=False))
                self.update_status("Affichage complet du stock")
            except Exception as e:
                self.update_status(f"Erreur chargement: {e}")

        threading.Thread(target=worker, daemon=True).start()

    # ========= Table helpers (embedded with lazy loading) =========
    def _configure_table_headers(self, headers):
        self.table_tree.delete(*self.table_tree.get_children())
        self.table_tree['columns'] = list(range(len(headers)))
        self.table_tree['show'] = 'headings'
        for i, h in enumerate(headers):
            self.table_tree.heading(i, text=str(h))
            self.table_tree.column(i, width=160, minwidth=120, anchor='w')

    def _set_table_dataset(self, headers, rows, row_indices, preview=False):
        self._current_table_headers = headers or []
        self._current_table_rows = rows or []
        self._current_row_indices = row_indices or []
        self._table_inserted_count = 0
        self._table_batch_size = 10 if preview else 200
        self._configure_table_headers(self._current_table_headers)
        self._insert_next_batch()

    def _insert_next_batch(self):
        if not self._current_table_rows:
            return
        start = self._table_inserted_count
        end = min(start + self._table_batch_size, len(self._current_table_rows))
        headers_len = len(self._current_table_headers)
        for row in self._current_table_rows[start:end]:
            padded = row + [''] * (headers_len - len(row))
            self.table_tree.insert('', 'end', values=padded)
        self._table_inserted_count = end

    def _maybe_load_more(self):
        if self._table_inserted_count >= len(self._current_table_rows):
            return
        top, bottom = self.table_tree.yview()
        if bottom >= 0.95:
            self._insert_next_batch()

    def _on_vsb(self, *args):
        self.table_tree.yview(*args)
        self._maybe_load_more()

    def _on_mousewheel(self, event):
        # Laisser Treeview gérer le scroll puis vérifier
        self.after(10, self._maybe_load_more)

    def _on_key_nav(self, event):
        self.after(10, self._maybe_load_more)

    # ========= Row details & editing =========
    def _on_table_double_click(self, event=None):
        selection = self.table_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        values = self.table_tree.item(item_id, 'values')
        if not values:
            return
        # Determine index in current table
        index_in_table = self.table_tree.index(item_id)
        if index_in_table >= len(self._current_row_indices):
            return
        sheet_row_number = self._current_row_indices[index_in_table]
        headers = self._current_table_headers
        self._open_edit_window(headers, list(values), sheet_row_number)

    def _open_edit_window(self, headers, row_values, sheet_row_number):
        win = tk.Toplevel(self)
        win.title(f"Détails (ligne {sheet_row_number})")
        win.geometry("1100x800")
        win.configure(bg="#2b2b2b")

        frame = ttk.Frame(win)
        frame.pack(fill='both', expand=True, padx=12, pady=12)

        canvas = tk.Canvas(frame, bg="#2b2b2b")
        vsb = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set)

        entries = []
        for i, (h, v) in enumerate(zip(headers, row_values)):
            rowf = ttk.Frame(inner)
            rowf.pack(fill='x', padx=6, pady=4)
            ttk.Label(rowf, text=str(h), width=24).pack(side='left')
            value_str = str(v)
            is_long = ('\n' in value_str) or (len(value_str) > 120)
            if is_long:
                text_frame = ttk.Frame(rowf)
                text_frame.pack(side='left', fill='both', expand=True)
                txt = tk.Text(text_frame, height=min(12, max(6, len(value_str) // 120)), wrap='word',
                              bg="#1f1f1f", fg="#eaeaea", insertbackground="#eaeaea", borderwidth=0, highlightthickness=0)
                sv = ttk.Scrollbar(text_frame, orient='vertical', command=txt.yview)
                txt.configure(yscrollcommand=sv.set)
                txt.pack(side='left', fill='both', expand=True)
                sv.pack(side='right', fill='y')
                txt.insert('1.0', value_str)
                entries.append((i, 'text', txt))
            else:
                ent = ttk.Entry(rowf)
                ent.insert(0, value_str)
                ent.pack(side='left', fill='x', expand=True)
                entries.append((i, 'entry', ent))

        def on_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        inner.bind('<Configure>', on_configure)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        btns = ttk.Frame(win)
        btns.pack(fill='x', padx=12, pady=(6, 8))

        def save_changes():
            try:
                if not self.sheets_manager:
                    messagebox.showerror("Erreur", "Sheets manager non disponible")
                    return
                # Update changed cells only
                for col_index, kind, widget in entries:
                    if kind == 'entry':
                        new_val = widget.get()
                    else:
                        new_val = widget.get('1.0', 'end-1c')
                    old_val = row_values[col_index] if col_index < len(row_values) else ""
                    if new_val != old_val:
                        a1_col = self._col_index_to_a1(col_index)
                        self.sheets_manager.update_cell(self.stock_worksheet_name, f"{a1_col}{sheet_row_number}", new_val)
                        # Reflect in caches
                        # Update _last_data as well
                        if self._last_data and sheet_row_number - 1 < len(self._last_data):
                            row = self._last_data[sheet_row_number - 1]
                            # Ensure row length
                            if len(row) <= col_index:
                                row.extend([''] * (col_index + 1 - len(row)))
                            row[col_index] = new_val
                # Refresh current table row visually
                self.refresh_stats()
                messagebox.showinfo("Succès", "Modifications enregistrées")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Échec de l'enregistrement: {e}")

        ttk.Button(btns, text="💾 Enregistrer", command=save_changes).pack(side='right')
        ttk.Button(btns, text="Fermer", command=win.destroy).pack(side='right', padx=6)

    def _col_index_to_a1(self, idx):
        # idx is 0-based
        result = ''
        n = idx + 1
        while n:
            n, rem = divmod(n - 1, 26)
            result = chr(65 + rem) + result
        return result

    # ========= Downloads =========
    def _download_selected_origin(self):
        self._download_selected(using_kyopa=False)

    def _download_selected_kyopa(self):
        self._download_selected(using_kyopa=True)

    def _download_selected(self, using_kyopa):
        try:
            if not self.drive_manager:
                messagebox.showerror("Erreur", "Drive manager non disponible")
                return
            selection = self.table_tree.selection()
            if not selection:
                messagebox.showwarning("Attention", "Veuillez sélectionner une ligne dans le tableau")
                return
            item_id = selection[0]
            index_in_table = self.table_tree.index(item_id)
            if index_in_table >= len(self._current_table_rows):
                return
            headers = [str(h).strip().lower() for h in self._current_table_headers]
            row = list(self._current_table_rows[index_in_table])

            # Trouver colonnes
            origin_col = self._find_header(headers, ["drive origin url", "origin url", "drive folder url"])  # fallback
            kyopa_col = self._find_header(headers, ["kyopa dossiers link", "kyopa folder url", "drive folder url kyopa", "kyopa url", "kyopa link"])

            target_col = kyopa_col if using_kyopa else origin_col
            if target_col is None:
                messagebox.showwarning("Attention", "Colonne URL introuvable pour ce téléchargement")
                return

            url = str(row[target_col]).strip() if target_col < len(row) else ""
            if not url:
                messagebox.showwarning("Attention", "URL vide dans la ligne sélectionnée")
                return

            dest = filedialog.askdirectory(title="Choisir le dossier de téléchargement")
            if not dest:
                return

            self.update_status("Téléchargement en cours...")
            ok = self.drive_manager.download_folder_by_url(url, Path(dest), lambda m: None)
            if ok:
                self.update_status("Téléchargement terminé")
                messagebox.showinfo("Succès", "Téléchargement terminé")
        except Exception as e:
            self.update_status(f"Erreur téléchargement: {e}")
            messagebox.showerror("Erreur", f"Échec du téléchargement: {e}")

    def _find_header(self, headers_lower, candidates):
        # Exact/substring tolerant match
        for i, h in enumerate(headers_lower):
            for c in candidates:
                if c == h or c in h:
                    return i
        return None


