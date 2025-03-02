import msal
import requests
import pandas as pd
import os
from datetime import datetime

# Définition des paramètres de connexion
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
tenant_id = 'YOUR_TENANT_ID'
authority = f'https://login.microsoftonline.com/{tenant_id}'

# Scopes nécessaires pour Defender for Endpoint
scopes = [
    'https://api.securitycenter.microsoft.com/.default'  # Scope pour Defender for Endpoint
]

class DefenderConnector:
    def __init__(self, client_id, client_secret, tenant_id, scopes):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f'https://login.microsoftonline.com/{tenant_id}'
        self.scopes = scopes
        self.access_token = None
        self.headers = None
        self.api_base_url = 'https://api.securitycenter.microsoft.com/api'
    
    def authenticate(self):
        """Authentification et obtention du token d'accès"""
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret,
        )
        
        result = app.acquire_token_for_client(scopes=self.scopes)
        
        if 'access_token' in result:
            self.access_token = result['access_token']
            self.headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            print("Authentification réussie!")
            return True
        else:
            print("Erreur d'authentification:")
            print(f"Error: {result.get('error')}")
            print(f"Description: {result.get('error_description')}")
            print(f"Correlation ID: {result.get('correlation_id')}")
            return False
    
    def get_device_groups(self):
        """Récupère tous les groupes de machines disponibles"""
        if not self.headers:
            if not self.authenticate():
                return None
        
        url = f"{self.api_base_url}/machineGroups"
        response = requests.get(url, headers=self.headers)
        
        if response.status_code == 200:
            groups = response.json().get('value', [])
            print(f"Nombre de groupes trouvés: {len(groups)}")
            return groups
        else:
            print(f"Erreur lors de la récupération des groupes: {response.status_code}")
            try:
                print(response.json())
            except:
                print("Impossible d'afficher le contenu de la réponse")
            return None
    
    def get_machines_by_group(self, group_id, group_name=None):
        """Récupère les machines appartenant à un groupe spécifique"""
        if not self.headers:
            if not self.authenticate():
                return None
        
        # URL pour récupérer les machines d'un groupe spécifique
        url = f"{self.api_base_url}/machines?$filter=machineGroups/any(g:g/id eq '{group_id}')"
        
        # Si vous avez besoin de données supplémentaires, ajoutez $expand
        # url += "&$expand=lastIpAddress,lastExternalIpAddress,osPlatform"
        
        response = requests.get(url, headers=self.headers)
        
        if response.status_code == 200:
            machines = response.json().get('value', [])
            group_info = group_name if group_name else group_id
            print(f"Nombre de machines dans le groupe '{group_info}': {len(machines)}")
            return machines
        else:
            print(f"Erreur lors de la récupération des machines pour le groupe {group_id}: {response.status_code}")
            try:
                print(response.json())
            except:
                print("Impossible d'afficher le contenu de la réponse")
            return None
    
    def export_machines_to_excel(self, machines, group_name, output_folder="./defender_exports"):
        """Exporte les données des machines vers un fichier Excel"""
        if not machines:
            print(f"Aucune machine à exporter pour le groupe '{group_name}'.")
            return False
        
        # Créer le dossier de destination s'il n'existe pas
        os.makedirs(output_folder, exist_ok=True)
        
        # Créer un DataFrame pandas avec les données des machines
        df = pd.DataFrame(machines)
        
        # Formater le nom du fichier avec le nom du groupe et la date
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{output_folder}/defender_machines_{group_name.replace(' ', '_')}_{timestamp}.xlsx"
        
        # Exporter vers Excel
        try:
            df.to_excel(filename, index=False, engine='openpyxl')
            print(f"Données exportées vers {filename}")
            return filename
        except Exception as e:
            print(f"Erreur lors de l'exportation vers Excel: {str(e)}")
            
            # Si Excel échoue, essayer d'exporter en CSV comme solution de secours
            try:
                csv_filename = filename.replace('.xlsx', '.csv')
                df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
                print(f"Données exportées vers {csv_filename} (format CSV)")
                return csv_filename
            except Exception as csv_err:
                print(f"Erreur lors de l'exportation en CSV: {str(csv_err)}")
                return False
    
    def export_all_machine_groups(self, group_names=None):
        """Exporte les machines de tous les groupes ou des groupes spécifiés"""
        # Récupérer tous les groupes disponibles
        all_groups = self.get_device_groups()
        
        if not all_groups:
            print("Impossible de récupérer les groupes. Arrêt du processus.")
            return False
        
        # Filtrer les groupes si une liste de noms est fournie
        if group_names:
            filtered_groups = [g for g in all_groups if g.get('name') in group_names]
            if not filtered_groups:
                print(f"Aucun des groupes spécifiés n'a été trouvé. Groupes disponibles: {[g.get('name') for g in all_groups]}")
                return False
            groups_to_process = filtered_groups
        else:
            groups_to_process = all_groups
        
        # Créer un dossier avec horodatage pour cette session d'exportation
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_folder = f"./defender_exports/export_session_{timestamp}"
        os.makedirs(export_folder, exist_ok=True)
        
        # Exporter les machines pour chaque groupe
        exported_files = []
        for group in groups_to_process:
            group_id = group.get('id')
            group_name = group.get('name', 'unnamed_group')
            
            print(f"\nTraitement du groupe: {group_name}")
            machines = self.get_machines_by_group(group_id, group_name)
            
            if machines:
                export_file = self.export_machines_to_excel(machines, group_name, export_folder)
                if export_file:
                    exported_files.append((group_name, export_file))
        
        print("\nRécapitulatif des exports:")
        for group_name, file_path in exported_files:
            print(f"- Groupe '{group_name}': {file_path}")
        
        return exported_files

def main():
    # Spécifiez ici les groupes que vous souhaitez extraire
    # Si None, tous les groupes seront exportés
    groups_to_extract = [
        "SST",
        "Finance",
        "IT",
        "HR"
    ]
    
    # Initialiser le connecteur Defender
    connector = DefenderConnector(client_id, client_secret, tenant_id, scopes)
    
    # Authentification
    if connector.authenticate():
        # Exporter les machines pour les groupes spécifiés
        connector.export_all_machine_groups(groups_to_extract)

if __name__ == "__main__":
    main()