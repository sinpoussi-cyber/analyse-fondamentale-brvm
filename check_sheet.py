import gspread
import os
import json
from google.oauth2 import service_account

# --- À CONFIGURER ---
# Assurez-vous que cet ID est bien celui de votre fichier
SPREADSHEET_ID = '1EGXyg13ml8a9zr4OaUPnJN3i-rwVO2uq330yfxJXnSM'
# --------------------

print("="*60)
print("--- SCRIPT DE DIAGNOSTIC GOOGLE SHEETS ---")
print(f"Tentative de connexion au Spreadsheet ID: {SPREADSHEET_ID}")
print("="*60)

try:
    # 1. Récupérer les identifiants depuis le secret GitHub
    creds_json_str = os.environ.get('GSPREAD_SERVICE_ACCOUNT')
    if not creds_json_str:
        print("❌ ERREUR : Le secret GSPREAD_SERVICE_ACCOUNT est introuvable ou vide.")
        exit(1)
    
    creds_dict = json.loads(creds_json_str)
    
    # 2. S'authentifier auprès de l'API Google
    print("Authentification auprès de Google...")
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    print("✅ Authentification réussie.")

    # 3. Ouvrir le fichier et lister les onglets
    print(f"Ouverture du fichier Google Sheet...")
    sheet = gc.open_by_key(SPREADSHEET_ID)
    
    print("✅ Fichier ouvert avec succès.")
    
    worksheets = sheet.worksheets()
    sheet_titles = [ws.title for ws in worksheets]

    print("\n" + "-"*60)
    if not sheet_titles:
        print("⚠️ AVERTISSEMENT : Le fichier a été trouvé, mais il ne contient aucun onglet (feuille).")
    else:
        print("🎉 SUCCÈS ! Voici la liste EXACTE des onglets trouvés :")
        print(sheet_titles)
        print("\nCOMPAREZ CETTE LISTE avec les clés du dictionnaire dans `main.py`:")
        print("['ABJC', 'BICB', 'BICC', 'BNBC', 'BOAB', ...]")
    print("-"*60 + "\n")


except gspread.exceptions.SpreadsheetNotFound:
    print("\n❌ ERREUR FATALE : SpreadsheetNotFound")
    print("Le fichier avec cet ID n'a pas été trouvé. Causes possibles :")
    print("1. L'ID du Spreadsheet est incorrect.")
    print("2. Le fichier n'a PAS été partagé avec l'adresse e-mail du compte de service.")
    print(f"   (L'adresse ressemble à: ...@...iam.gserviceaccount.com)")
    exit(1)

except gspread.exceptions.APIError as e:
    print(f"\n❌ ERREUR FATALE : APIError (Code: {e.response.status_code})")
    print("Une erreur d'API est survenue. Si le code est 403 (FORBIDDEN/PERMISSION_DENIED) :")
    print("1. Le compte de service a bien été partagé avec le fichier, mais l'API 'Google Sheets API' n'est pas activée sur votre projet Google Cloud.")
    print("2. Vérifiez que le partage donne bien les droits 'Lecteur' ou 'Éditeur'.")
    exit(1)

except Exception as e:
    print(f"\n❌ ERREUR INCONNUE : {e}")
    exit(1)
