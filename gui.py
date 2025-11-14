#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Application de conversion PDF vers Word et Excel avec interface graphique Tkinter

Cette application permet de convertir des fichiers PDF vers les formats Word (.docx) et Excel (.xlsx)
en utilisant l'API REST Adobe PDF Services, avec une interface graphique conviviale.

Auteur: Mass
Date: 13 novembre 2025
"""

import os, sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Chemin vers le fichier JSON empaqueté ou local
json_path = resource_path("pdfservices-api-credentials.json")


import os
import json
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import requests


class PDFConverter:
    """Classe pour convertir des fichiers PDF vers Word et Excel via l'API Adobe PDF Services."""

    # Endpoints de l'API
    BASE_URL = "https://pdf-services.adobe.io"
    TOKEN_ENDPOINT = f"{BASE_URL}/token"
    ASSETS_ENDPOINT = f"{BASE_URL}/assets"
    EXPORT_PDF_ENDPOINT = f"{BASE_URL}/operation/exportpdf"
    
    # Formats de conversion supportés
    FORMATS = {
        "word": "docx",
        "excel": "xlsx"
    }

    def __init__(self, credentials_path):
        """
        Initialise le convertisseur avec les identifiants de l'API.
        
        Args:
            credentials_path (str): Chemin vers le fichier d'identifiants JSON
        """
        self.credentials = self._load_credentials(credentials_path)
        self.access_token = None
        self.token_expiry = 0
    
    def _load_credentials(self, credentials_path):
        """
        Charge les identifiants depuis le fichier JSON.
        
        Args:
            credentials_path (str): Chemin vers le fichier d'identifiants
            
        Returns:
            dict: Identifiants chargés
        """
        try:
            with open(credentials_path) as f:
                return json.load(f)
        except (IOError, json.JSONDecodeError) as e:
            raise Exception(f"Erreur lors du chargement des identifiants: {e}")
    
    def _get_access_token(self):
        """
        Obtient un token d'accès pour l'API.
        
        Returns:
            str: Token d'accès
        """
        # Vérifier si le token existant est encore valide
        if self.access_token and time.time() < self.token_expiry:
            return self.access_token
        
        # Préparer les données pour la requête de token
        client_id = self.credentials["client_credentials"]["client_id"]
        client_secret = self.credentials["client_credentials"]["client_secret"]
        
        data = {
            "client_id": client_id,
            "client_secret": client_secret
        }
        
        try:
            response = requests.post(self.TOKEN_ENDPOINT, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data["access_token"]
            # Définir l'expiration du token (généralement 24h, mais on prend une marge)
            self.token_expiry = time.time() + int(token_data.get("expires_in", 86000))
            
            return self.access_token
        except requests.exceptions.RequestException as e:
            raise Exception(f"Erreur lors de l'obtention du token: {e}")
    
    def _get_headers(self):
        """
        Prépare les en-têtes pour les requêtes API.
        
        Returns:
            dict: En-têtes HTTP
        """
        token = self._get_access_token()
        client_id = self.credentials["client_credentials"]["client_id"]
        
        return {
            "Authorization": f"Bearer {token}",
            "x-api-key": client_id,
            "Content-Type": "application/json"
        }
    
    def _upload_file(self, file_path, callback=None):
        """
        Télécharge un fichier PDF vers le stockage temporaire d'Adobe.
        
        Args:
            file_path (str): Chemin vers le fichier PDF à télécharger
            callback (function): Fonction de rappel pour les mises à jour de statut
            
        Returns:
            str: ID de l'asset téléchargé
        """
        if callback:
            callback("Téléchargement du fichier...")
        
        # 1. Obtenir une URL pré-signée pour le téléchargement
        headers = self._get_headers()
        payload = {"mediaType": "application/pdf"}
        
        try:
            # Étape 1: Obtenir l'URL pré-signée et l'assetID
            response = requests.post(self.ASSETS_ENDPOINT, headers=headers, json=payload)
            response.raise_for_status()
            
            upload_data = response.json()
            asset_id = upload_data.get("assetID")
            upload_uri = upload_data.get("uploadUri")
            
            if not asset_id or not upload_uri:
                raise Exception("Impossible d'obtenir l'assetID ou l'uploadUri")
            
            if callback:
                callback(f"Asset ID obtenu: {asset_id}")
            
            # Étape 2: Télécharger le fichier PDF vers l'URL pré-signée
            with open(file_path, 'rb') as file:
                file_content = file.read()
            
            upload_response = requests.put(upload_uri, data=file_content, headers={"Content-Type": "application/pdf"})
            upload_response.raise_for_status()
            
            if callback:
                callback("Fichier téléchargé avec succès vers le stockage Adobe.")
            return asset_id
            
        except requests.exceptions.RequestException as e:
            raise Exception(f"Erreur lors du téléchargement du fichier: {e}")
    
    def _export_pdf(self, asset_id, target_format, callback=None):
        """
        Convertit un PDF vers le format spécifié.
        
        Args:
            asset_id (str): ID de l'asset PDF
            target_format (str): Format cible (docx, xlsx)
            callback (function): Fonction de rappel pour les mises à jour de statut
            
        Returns:
            str: ID de l'asset converti
        """
        if callback:
            callback(f"Conversion du PDF vers {target_format}...")
        
        headers = self._get_headers()
        payload = {
            "assetID": asset_id,
            "targetFormat": target_format
        }
        
        try:
            response = requests.post(self.EXPORT_PDF_ENDPOINT, headers=headers, json=payload)
            response.raise_for_status()
            
            # L'API renvoie un code 201 avec l'URL de statut dans l'en-tête Location
            status_url = response.headers.get("Location")
            if not status_url:
                raise Exception("Aucune URL de statut reçue")
            
            if callback:
                callback("Opération de conversion initiée...")
            
            # Attendre que la conversion soit terminée
            output_asset_id = self._wait_for_completion(status_url, callback=callback)
            
            if callback:
                callback(f"Conversion terminée. Asset ID du résultat: {output_asset_id}")
            
            return output_asset_id
            
        except requests.exceptions.RequestException as e:
            raise Exception(f"Erreur lors de la conversion du PDF: {e}")
    
    def _wait_for_completion(self, status_url, max_retries=30, delay=2, callback=None):
        """
        Attend que l'opération de conversion soit terminée.
        
        Args:
            status_url (str): URL de statut de l'opération
            max_retries (int): Nombre maximum de tentatives
            delay (int): Délai entre les tentatives en secondes
            callback (function): Fonction de rappel pour les mises à jour de statut
            
        Returns:
            str: ID de l'asset résultant
        """
        headers = self._get_headers()
        
        for i in range(max_retries):
            try:
                if callback:
                    callback(f"Vérification du statut de la conversion (tentative {i+1}/{max_retries})...")
                
                response = requests.get(status_url, headers=headers)
                response.raise_for_status()
                
                status_data = response.json()
                status = status_data.get("status")
                
                if status == "done":
                    return status_data.get("asset", {}).get("assetID")
                elif status == "failed":
                    error = status_data.get("error", {})
                    raise Exception(f"La conversion a échoué: {error.get('message', 'Erreur inconnue')}")
                
                if callback:
                    callback(f"Statut actuel: {status}. Attente de {delay} secondes...")
                
                # Attendre avant la prochaine vérification
                time.sleep(delay)
                
            except requests.exceptions.RequestException as e:
                if callback:
                    callback(f"Erreur lors de la vérification du statut: {e}. Nouvelle tentative...")
                time.sleep(delay)
        
        raise Exception(f"Délai d'attente dépassé après {max_retries} tentatives")
    
    def _download_result(self, asset_id, output_path, callback=None):
        """
        Télécharge le fichier converti.
        
        Args:
            asset_id (str): ID de l'asset à télécharger
            output_path (str): Chemin où sauvegarder le fichier
            callback (function): Fonction de rappel pour les mises à jour de statut
            
        Returns:
            str: Chemin du fichier téléchargé
        """
        if callback:
            callback(f"Téléchargement du résultat vers {output_path}...")
        
        headers = self._get_headers()
        
        try:
            # Obtenir l'URL de téléchargement
            response = requests.get(f"{self.ASSETS_ENDPOINT}/{asset_id}", headers=headers)
            response.raise_for_status()
            
            download_uri = response.json().get("downloadUri")
            if not download_uri:
                raise Exception("Aucune URL de téléchargement reçue")
            
            # Télécharger le fichier
            download_response = requests.get(download_uri)
            download_response.raise_for_status()
            
            # Sauvegarder le fichier
            with open(output_path, 'wb') as f:
                f.write(download_response.content)
            
            if callback:
                callback(f"Fichier téléchargé avec succès: {output_path}")
            
            return output_path
            
        except requests.exceptions.RequestException as e:
            raise Exception(f"Erreur lors du téléchargement du résultat: {e}")
    
    def convert(self, pdf_path, output_format, output_path=None, callback=None):
        """
        Convertit un fichier PDF vers Word ou Excel.
        
        Args:
            pdf_path (str): Chemin vers le fichier PDF à convertir
            output_format (str): Format de sortie ('word' ou 'excel')
            output_path (str, optional): Chemin de sortie pour le fichier converti
            callback (function): Fonction de rappel pour les mises à jour de statut
            
        Returns:
            str: Chemin du fichier converti
        """
        # Vérifier que le format est supporté
        if output_format not in self.FORMATS:
            raise ValueError(f"Format non supporté: {output_format}. Formats supportés: {', '.join(self.FORMATS.keys())}")
        
        # Vérifier que le fichier PDF existe
        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"Le fichier PDF n'existe pas: {pdf_path}")
        
        # Déterminer le chemin de sortie si non spécifié
        if not output_path:
            pdf_name = Path(pdf_path).stem
            extension = self.FORMATS[output_format]
            output_dir = os.path.dirname(pdf_path)
            output_path = os.path.join(output_dir, f"{pdf_name}.{extension}")
        
        try:
            # Télécharger le fichier PDF
            asset_id = self._upload_file(pdf_path, callback)
            
            # Convertir le PDF
            target_format = self.FORMATS[output_format]
            result_asset_id = self._export_pdf(asset_id, target_format, callback)
            
            # Télécharger le résultat
            return self._download_result(result_asset_id, output_path, callback)
            
        except Exception as e:
            if callback:
                callback(f"Erreur lors de la conversion: {e}")
            raise


class PDFConverterApp:
    """Application Tkinter pour la conversion de PDF vers Word et Excel."""
    
    def __init__(self, root):
        """
        Initialise l'application Tkinter.
        
        Args:
            root: Fenêtre principale Tkinter
        """
        self.root = root
        self.root.title("Convertisseur PDF vers Word/Excel")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.output_format = tk.StringVar(value="word")
        self.credentials_path = tk.StringVar(value="pdfservices-api-credentials.json")
        self.status_text = tk.StringVar(value="Prêt")
        
        # Créer l'interface
        self._create_widgets()
        
        # Centrer la fenêtre
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self):
        """Crée les widgets de l'interface."""
        # Style
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")
        style.configure("TLabel", padding=6)
        style.configure("TRadiobutton", padding=5)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titre
        title_label = ttk.Label(main_frame, text="Convertisseur PDF vers Word/Excel", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Sélection du fichier PDF
        pdf_frame = ttk.LabelFrame(main_frame, text="Fichier PDF source", padding="10")
        pdf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        pdf_entry = ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=50)
        pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        pdf_button = ttk.Button(pdf_frame, text="Parcourir...", command=self._browse_pdf)
        pdf_button.pack(side=tk.RIGHT, padx=5)
        
        # Format de sortie
        format_frame = ttk.LabelFrame(main_frame, text="Format de sortie", padding="10")
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        word_radio = ttk.Radiobutton(format_frame, text="Word (.docx)", variable=self.output_format, value="word")
        word_radio.pack(anchor=tk.W, padx=5, pady=2)
        
        excel_radio = ttk.Radiobutton(format_frame, text="Excel (.xlsx)", variable=self.output_format, value="excel")
        excel_radio.pack(anchor=tk.W, padx=5, pady=2)
        
        # Fichier de sortie
        output_frame = ttk.LabelFrame(main_frame, text="Fichier de sortie (optionnel)", padding="10")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        output_button = ttk.Button(output_frame, text="Parcourir...", command=self._browse_output)
        output_button.pack(side=tk.RIGHT, padx=5)
        
        # Fichier d'identifiants
        creds_frame = ttk.LabelFrame(main_frame, text="Fichier d'identifiants API", padding="10")
        creds_frame.pack(fill=tk.X, padx=5, pady=5)
        
        creds_entry = ttk.Entry(creds_frame, textvariable=self.credentials_path, width=50)
        creds_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        creds_button = ttk.Button(creds_frame, text="Parcourir...", command=self._browse_credentials)
        creds_button.pack(side=tk.RIGHT, padx=5)
        
        # Bouton de conversion
        convert_button = ttk.Button(main_frame, text="Convertir", command=self._start_conversion)
        convert_button.pack(pady=10)
        
        # Zone de statut
        status_frame = ttk.LabelFrame(main_frame, text="Statut", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.status_area = tk.Text(status_frame, wrap=tk.WORD, height=10)
        self.status_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar pour la zone de statut
        scrollbar = ttk.Scrollbar(self.status_area, command=self.status_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_area.config(yscrollcommand=scrollbar.set)
        
        # Barre de statut
        status_bar = ttk.Label(self.root, textvariable=self.status_text, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _browse_pdf(self):
        """Ouvre une boîte de dialogue pour sélectionner un fichier PDF."""
        filename = filedialog.askopenfilename(
            title="Sélectionner un fichier PDF",
            filetypes=[("Fichiers PDF", "*.pdf"), ("Tous les fichiers", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            # Suggérer un chemin de sortie basé sur le fichier PDF
            if not self.output_path.get():
                pdf_name = Path(filename).stem
                output_dir = os.path.dirname(filename)
                extension = "docx" if self.output_format.get() == "word" else "xlsx"
                self.output_path.set(os.path.join(output_dir, f"{pdf_name}.{extension}"))
    
    def _browse_output(self):
        """Ouvre une boîte de dialogue pour sélectionner le fichier de sortie."""
        extension = "docx" if self.output_format.get() == "word" else "xlsx"
        filename = filedialog.asksaveasfilename(
            title="Enregistrer sous",
            defaultextension=f".{extension}",
            filetypes=[
                ("Fichiers Word", "*.docx") if extension == "docx" else ("Fichiers Excel", "*.xlsx"),
                ("Tous les fichiers", "*.*")
            ]
        )
        if filename:
            self.output_path.set(filename)
    
    def _browse_credentials(self):
        """Ouvre une boîte de dialogue pour sélectionner le fichier d'identifiants."""
        filename = filedialog.askopenfilename(
            title="Sélectionner le fichier d'identifiants",
            filetypes=[("Fichiers JSON", "*.json"), ("Tous les fichiers", "*.*")]
        )
        if filename:
            self.credentials_path.set(filename)
    
    def _update_status(self, message):
        """
        Met à jour la zone de statut avec un nouveau message.
        
        Args:
            message (str): Message à afficher
        """
        self.status_area.insert(tk.END, f"{message}\n")
        self.status_area.see(tk.END)
        self.status_text.set(message)
        self.root.update_idletasks()
    
    def _start_conversion(self):
        """Démarre le processus de conversion dans un thread séparé."""
        # Vérifier que les champs obligatoires sont remplis
        if not self.pdf_path.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier PDF source.")
            return
        
        if not os.path.isfile(self.pdf_path.get()):
            messagebox.showerror("Erreur", f"Le fichier PDF n'existe pas: {self.pdf_path.get()}")
            return
        
        if not os.path.isfile(self.credentials_path.get()):
            messagebox.showerror("Erreur", f"Le fichier d'identifiants n'existe pas: {self.credentials_path.get()}")
            return
        
        # Effacer la zone de statut
        self.status_area.delete(1.0, tk.END)
        
        # Démarrer la conversion dans un thread séparé
        threading.Thread(target=self._convert_pdf, daemon=True).start()
    
    def _convert_pdf(self):
        """Effectue la conversion du PDF dans un thread séparé."""
        try:
            self._update_status("Initialisation de la conversion...")
            
            # Créer le convertisseur
            converter = PDFConverter(self.credentials_path.get())
            
            # Convertir le PDF
            output_path = converter.convert(
                self.pdf_path.get(),
                self.output_format.get(),
                self.output_path.get() if self.output_path.get() else None,
                self._update_status
            )
            
            # Afficher un message de succès
            self._update_status(f"Conversion réussie! Fichier sauvegardé: {output_path}")
            messagebox.showinfo("Succès", f"Conversion réussie!\nFichier sauvegardé: {output_path}")
            
            # Ouvrir le dossier contenant le fichier
            os.system(f'explorer "{os.path.dirname(output_path)}"' if os.name == 'nt' else f'xdg-open "{os.path.dirname(output_path)}"')
            
        except Exception as e:
            self._update_status(f"Erreur: {e}")
            messagebox.showerror("Erreur", f"Une erreur est survenue lors de la conversion:\n{e}")


def main():
    """Fonction principale pour démarrer l'application."""
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
