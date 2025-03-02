import msal
import requests
import os
from datetime import datetime

# Définition des paramètres de connexion
client_id = 'YOUR_CLIENT_ID'
client_secret = 'YOUR_CLIENT_SECRET'
tenant_id = 'YOUR_TENANT_ID'
authority = f'https://login.microsoftonline.com/{tenant_id}'

# Les scopes dont nous avons besoin
scopes = [
    'https://graph.microsoft.com/.default'  # Scope général pour MS Graph
]

class MicrosoftGraphConnector:
    def __init__(self, client_id, client_secret, tenant_id, scopes):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f'https://login.microsoftonline.com/{tenant_id}'
        self.scopes = scopes
        self.access_token = None
        self.headers = None
    
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
            return True
        else:
            print("Erreur d'authentification:")
            print(f"Error: {result.get('error')}")
            print(f"Description: {result.get('error_description')}")
            print(f"Correlation ID: {result.get('correlation_id')}")
            return False
    
    def get_sharepoint_site_id(self, site_name):
        """Récupère l'ID d'un site SharePoint spécifique"""
        if not self.headers:
            if not self.authenticate():
                return None
                
        # Récupérer tous les sites
        response = requests.get('https://graph.microsoft.com/v1.0/sites', headers=self.headers)
        
        if response.status_code == 200:
            sites = response.json().get('value', [])
            for site in sites:
                if site_name.lower() in site.get('displayName', '').lower():
                    return site['id']
            
            print(f"Site SharePoint '{site_name}' non trouvé.")
            return None
        else:
            print(f"Erreur lors de la récupération des sites: {response.status_code}")
            print(response.json())
            return None
    
    def get_sharepoint_drive_items(self, site_id, drive_id=None, folder_path=None):
        """Récupère les éléments d'un dossier SharePoint"""
        if not self.headers:
            if not self.authenticate():
                return None
                
        if not drive_id:
            # Récupérer le drive par défaut
            response = requests.get(f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives', headers=self.headers)
            
            if response.status_code == 200:
                drives = response.json().get('value', [])
                if drives:
                    drive_id = drives[0]['id']
                else:
                    print("Aucun drive trouvé dans ce site.")
                    return None
            else:
                print(f"Erreur lors de la récupération des drives: {response.status_code}")
                print(response.json())
                return None
        
        # Construire l'URL pour accéder au dossier ou à la racine du drive
        url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root'
        if folder_path:
            url += f":/'{folder_path}':/children"
        else:
            url += "/children"
            
        response = requests.get(url, headers=self.headers)
        
        if response.status_code == 200:
            return response.json().get('value', [])
        else:
            print(f"Erreur lors de la récupération des fichiers: {response.status_code}")
            print(response.json())
            return None

    def download_sharepoint_csv(self, site_id, drive_id, file_id, destination_folder="."):
        """Télécharge un fichier CSV depuis SharePoint"""
        if not self.headers:
            if not self.authenticate():
                return False
                
        # Récupérer le contenu du fichier
        response = requests.get(
            f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content',
            headers=self.headers,
            stream=True
        )
        
        if response.status_code == 200:
            # Récupérer le nom du fichier
            file_info_response = requests.get(
                f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}',
                headers=self.headers
            )
            
            if file_info_response.status_code == 200:
                file_name = file_info_response.json().get('name', f'downloaded_file_{datetime.now().strftime("%Y%m%d%H%M%S")}.csv')
                os.makedirs(destination_folder, exist_ok=True)
                filepath = os.path.join(destination_folder, file_name)
                
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                        
                print(f"Fichier téléchargé: {filepath}")
                return filepath
            else:
                print(f"Erreur lors de la récupération des informations du fichier: {file_info_response.status_code}")
                return False
        else:
            print(f"Erreur lors du téléchargement du fichier: {response.status_code}")
            return False

# Fonction principale
def main():
    # Initialiser le connecteur MS Graph
    connector = MicrosoftGraphConnector(client_id, client_secret, tenant_id, scopes)
    
    if connector.authenticate():
        print("Authentification réussie!")
        
        # Télécharger des CSV depuis SharePoint
        site_name = "YOUR_SHAREPOINT_SITE_NAME"  # Remplacez par votre nom de site
        site_id = connector.get_sharepoint_site_id(site_name)
        
        if site_id:
            print(f"ID du site SharePoint: {site_id}")
            
            # Récupérer les éléments du dossier spécifique (optionnel)
            folder_path = "Documents/Folder_With_CSV"  # Remplacez par votre chemin de dossier
            items = connector.get_sharepoint_drive_items(site_id, folder_path=folder_path)
            
            if items:
                # Télécharger tous les fichiers CSV
                csv_files = [item for item in items if item.get('name', '').endswith('.csv')]
                print(f"Nombre de fichiers CSV trouvés: {len(csv_files)}")
                
                for csv_file in csv_files:
                    connector.download_sharepoint_csv(
                        site_id, 
                        csv_file.get('parentReference', {}).get('driveId'), 
                        csv_file.get('id'), 
                        destination_folder="./downloaded_sharepoint_files"
                    )
            else:
                print("Aucun élément trouvé dans le dossier spécifié ou erreur lors de la récupération.")
        else:
            print("Impossible de trouver le site SharePoint spécifié.")

if __name__ == "__main__":
    main()