# ==============================================================================
# ANALYSEUR FINANCIER BRVM - SCRIPT FINAL V6.7 (FILTRAGE AVANCÉ - SANS QUOTA)
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. IMPORTATION DES BIBLIOTHÈQUES
# ------------------------------------------------------------------------------
import gspread
import requests
from bs4 import BeautifulSoup
import time
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import sys
from datetime import datetime
import logging
import io
import unicodedata
import urllib3
import json
from collections import defaultdict

# Imports Selenium
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Imports pour l'authentification Google
from google.oauth2 import service_account

# NOUVEAU : Import pour l'API Gemini
import google.generativeai as genai

# Désactiver les avertissements de sécurité
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ------------------------------------------------------------------------------
# 2. CONFIGURATION DU LOGGING
# ------------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ------------------------------------------------------------------------------
# 3. CLASSE PRINCIPALE DE L'ANALYSEUR
# ------------------------------------------------------------------------------
class BRVMAnalyzer:
    def __init__(self, spreadsheet_id, api_key):
        self.spreadsheet_id = spreadsheet_id
        self.api_key = api_key
        
        self.societes_mapping = {
            'SIVC': {'nom_rapport': 'AIR LIQUIDE CI', 'alternatives': ['air liquide ci']},
            'BOABF': {'nom_rapport': 'BANK OF AFRICA BF', 'alternatives': ['bank of africa bf']},
            'BOAB': {'nom_rapport': 'BANK OF AFRICA BN', 'alternatives': ['bank of africa bn']},
            'BOAC': {'nom_rapport': 'BANK OF AFRICA CI', 'alternatives': ['bank of africa ci']},
            'BOAM': {'nom_rapport': 'BANK OF AFRICA ML', 'alternatives': ['bank of africa ml']},
            'BOAN': {'nom_rapport': 'BANK OF AFRICA NG', 'alternatives': ['bank of africa ng']},
            'BOAS': {'nom_rapport': 'BANK OF AFRICA SN', 'alternatives': ['bank of africa sn']},
            'BNBC': {'nom_rapport': 'BERNABE CI', 'alternatives': ['bernabe ci']},
            'BICC': {'nom_rapport': 'BICI CI', 'alternatives': ['bici ci']},
            'CABC': {'nom_rapport': 'CABC', 'alternatives': ['cabc']},
            'CFAC': {'nom_rapport': 'CFAO MOTORS CI', 'alternatives': ['cfao motors ci']},
            'CIEC': {'nom_rapport': 'CIE CI', 'alternatives': ['cie ci']},
            'CBIBF': {'nom_rapport': 'CORIS BANK INTERNATIONAL', 'alternatives': ['coris bank international']},
            'ECOC': {'nom_rapport': 'ECOBANK COTE D\'IVOIRE', 'alternatives': ["ecobank cote d ivoire", "ecobank ci"]},
            'ETIT': {'nom_rapport': 'ECOBANK TRANS. INCORP. TG', 'alternatives': ['ecobank trans', 'ecobank tg']},
            'FTSC': {'nom_rapport': 'FILTISAC CI', 'alternatives': ['filtisac ci']},
            'NEIC': {'nom_rapport': 'NEI-CEDA CI', 'alternatives': ['nei-ceda ci']},
            'NSBC': {'nom_rapport': 'NSIA BANQUE CI', 'alternatives': ['nsia banque ci', 'nsbc']},
            'ONTBF': {'nom_rapport': 'ONATEL BF', 'alternatives': ['onatel bf']},
            'ORAC': {'nom_rapport': 'ORANGE CI', 'alternatives': ['orange ci', "cote d'ivoire telecom"]},
            'PALC': {'nom_rapport': 'PALM CI', 'alternatives': ['palm ci']},
            'SAFC': {'nom_rapport': 'SAFCA CI', 'alternatives': ['safca ci']},
            'SPHC': {'nom_rapport': 'SAPH CI', 'alternatives': ['saph ci']},
            'STAC': {'nom_rapport': 'SETAO CI', 'alternatives': ['setao ci']},
            'SGBC': {'nom_rapport': 'SOCIETE GENERALE CI', 'alternatives': ['societe generale ci', 'sgb ci']},
            'SIBC': {'nom_rapport': 'SOCIETE IVOIRIENNE DE BANQUE', 'alternatives': ['societe ivoirienne de banque', 'sib']},
            'SLBC': {'nom_rapport': 'SOLIBRA CI', 'alternatives': ['solibra ci', 'solibra']},
            'SNTS': {'nom_rapport': 'SONATEL SN', 'alternatives': ['sonatel sn', 'fctc sonatel', 'sonatel']},
            'SCRC': {'nom_rapport': 'SUCRIVOIRE CI', 'alternatives': ['sucrivoire ci', 'sucrivoire']},
            'TTLC': {'nom_rapport': 'TOTALENERGIES MARKETING CI', 'alternatives': ['totalenergies marketing ci', 'total']},
            'TTLS': {'nom_rapport': 'TOTALENERGIES MARKETING SN', 'alternatives': ['totalenergies marketing senegal', 'total senegal s.a.']},
            'UNLC': {'nom_rapport': 'UNILEVER CI', 'alternatives': ['unilever ci']},
            'UNXC': {'nom_rapport': 'UNIWAX CI', 'alternatives': ['uniwax ci']},
            'SHEC': {'nom_rapport': 'VIVO ENERGY CI', 'alternatives': ['vivo energy ci']},
        }

        self.gc = None
        self.driver = None
        self.gemini_model = None
        self.original_societes_mapping = self.societes_mapping.copy()
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})
        
    def setup_selenium(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument("--window-size=1920,1080")
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            logger.info("✅ Pilote Selenium (Chrome) démarré.")
        except Exception as e:
            logger.error(f"❌ Impossible de démarrer le pilote Selenium: {e}")
            self.driver = None

    def configure_gemini(self):
        if not self.api_key:
            logger.error("❌ Clé API Google (GOOGLE_API_KEY) non trouvée. L'analyse par IA est impossible.")
            return False
        try:
            genai.configure(api_key=self.api_key)
            self.gemini_model = genai.GenerativeModel('gemini-1.5-flash-latest')
            logger.info("✅ API Gemini configurée avec succès.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur lors de la configuration de l'API Gemini: {e}")
            return False
            
    def authenticate_google_services(self):
        logger.info("Authentification Google...")
        try:
            creds_json_str = os.environ.get('GSPREAD_SERVICE_ACCOUNT')
            if not creds_json_str:
                logger.error("❌ Secret GSPREAD_SERVICE_ACCOUNT introuvable.")
                return False
            creds_dict = json.loads(creds_json_str)
            scopes = ['https://www.googleapis.com/auth/spreadsheets']
            creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
            self.gc = gspread.authorize(creds)
            logger.info("✅ Authentification Google réussie.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur d'authentification : {e}")
            return False

    def verify_and_filter_companies(self):
        try:
            logger.info(f"Vérification des feuilles dans G-Sheet...")
            sheet = self.gc.open_by_key(self.spreadsheet_id)
            existing_sheets = [ws.title for ws in sheet.worksheets()]
            logger.info(f"Onglets trouvés : {existing_sheets}")
            symbols_to_keep = [s for s in self.original_societes_mapping if s in existing_sheets]
            self.societes_mapping = {k: v for k, v in self.original_societes_mapping.items() if k in symbols_to_keep}
            if not self.societes_mapping:
                logger.error("❌ ERREUR FATALE : Aucune société à analyser.")
                return False
            logger.info(f"✅ {len(self.societes_mapping)} sociétés seront analysées.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur lors de la vérification du G-Sheet: {e}")
            return False

    def _normalize_text(self, text):
        if not text: return ""
        text = text.replace('-', ' ')
        text = ''.join(c for c in unicodedata.normalize('NFD', str(text).lower()) if unicodedata.category(c) != 'Mn')
        text = re.sub(r'[^a-z0-9\s\.]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()
    
    def _find_all_reports(self):
        # ... (Cette fonction est inchangée)
        return {}

    def _get_symbol_from_name(self, company_name_normalized):
        # ... (Cette fonction est inchangée)
        return None

    def _extract_date_from_text(self, text):
        # ... (Cette fonction est inchangée)
        return datetime(1900, 1, 1)

    def _analyze_pdf_with_gemini(self, pdf_url):
        if not self.gemini_model:
            return "Analyse IA non disponible (API non configurée)."
        
        logger.info(f"    -> Téléchargement du PDF pour l'envoyer à Gemini...")
        uploaded_file = None
        temp_pdf_path = "temp_report.pdf"
        try:
            response = self.session.get(pdf_url, timeout=45, verify=False)
            response.raise_for_status()
            pdf_content = response.content
            if len(pdf_content) < 1024:
                return "Fichier PDF invalide ou vide."
            with open(temp_pdf_path, 'wb') as f:
                f.write(pdf_content)
            logger.info(f"    -> Envoi du fichier PDF ({os.path.getsize(temp_pdf_path)} octets) à l'API Gemini...")
            uploaded_file = genai.upload_file(
                path=temp_pdf_path,
                display_name="Rapport Financier BRVM"
            )
            prompt = """
            Tu es un analyste financier expert spécialisé dans les entreprises de la zone UEMOA cotées à la BRVM...
            """ # (Le prompt reste le même)
            
            logger.info("    -> Fichier envoyé. Génération de l'analyse...")
            response = self.gemini_model.generate_content([prompt, uploaded_file])
            
            if response.parts:
                return response.text
            elif response.prompt_feedback:
                block_reason = response.prompt_feedback.block_reason.name
                error_message = f"Analyse bloquée par l'IA. Raison : {block_reason}."
                return error_message
            else:
                 return "Erreur inconnue : L'API Gemini n'a retourné ni contenu ni feedback."

        except Exception as e:
            error_details = f"Erreur technique lors de l'analyse par l'IA : {str(e)}"
            return error_details
        finally:
            if uploaded_file:
                try:
                    logger.info(f"    -> Suppression du fichier temporaire de l'API Gemini.")
                    genai.delete_file(uploaded_file.name)
                except Exception as e:
                    logger.warning(f"    -> N'a pas pu supprimer le fichier temporaire de l'API : {e}")

            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)
                logger.info(f"    -> Suppression du fichier PDF local ({temp_pdf_path}).")

    def process_all_companies(self):
        all_reports = self._find_all_reports()
        results = {}
        if not all_reports:
            logger.error("❌ ÉCHEC FINAL : Aucun rapport n'a pu être collecté sur le site de la BRVM.")
            return {}
        logger.info(f"\n✅ COLLECTE TERMINÉE : {sum(len(r) for r in all_reports.values())} rapports trouvés au total.")
        
        # --- DÉFINITION DES CRITÈRES DE FILTRAGE ---
        date_2024_start = datetime(2024, 1, 1)
        date_2025_start = datetime(2025, 1, 1)
        keywords_financiers = ['états financiers', 'etats financiers', 'certifié', 'commissaires aux comptes', 'rapport annuel']

        for symbol, info in self.societes_mapping.items():
            logger.info(f"\n📊 Traitement des données pour {symbol} - {info['nom_rapport']}")
            
            company_reports = all_reports.get(symbol, [])
            analysis_data = {'nom': info['nom_rapport'], 'rapports_analyses': []}

            # --- BLOC DE FILTRAGE AVANCÉ ---
            reports_to_analyze = []
            for report in company_reports:
                report_date = report['date']
                title_lower = report['titre'].lower()

                # Règle 1 : Rapports de 2024 (entre 01/01/2024 et 31/12/2024)
                if date_2024_start <= report_date < date_2025_start:
                    if any(keyword in title_lower for keyword in keywords_financiers):
                        reports_to_analyze.append(report)
                
                # Règle 2 : Rapports à partir de 2025
                elif report_date >= date_2025_start:
                    reports_to_analyze.append(report)
            
            reports_to_analyze.sort(key=lambda x: x['date'], reverse=True)
            
            if not reports_to_analyze:
                analysis_data['statut'] = 'Aucun rapport pertinent trouvé selon les critères de filtrage (date/titre).'
                results[symbol] = analysis_data
                continue
            
            logger.info(f"  -> {len(reports_to_analyze)} rapport(s) pertinent(s) trouvé(s) après filtrage.")

            # --- BOUCLE D'ANALYSE SANS LIMITATION ---
            for i, report in enumerate(reports_to_analyze):
                logger.info(f"  -> Analyse IA {i+1}/{len(reports_to_analyze)}: {report['titre'][:60]}...")
                
                gemini_analysis = self._analyze_pdf_with_gemini(report['url'])
                
                analysis_data['rapports_analyses'].append({
                    'titre': report['titre'], 
                    'url': report['url'], 
                    'date': report['date'].strftime('%Y-%m-%d'),
                    'analyse_ia': gemini_analysis
                })
                
                time.sleep(3) # Pause pour ne pas surcharger l'API
            
            results[symbol] = analysis_data
        
        logger.info("\n✅ Traitement de toutes les sociétés terminé.")
        return results

    def create_word_report(self, results, output_path):
        # ... (Cette fonction est inchangée)
        pass

    def run(self):
        # ... (Cette fonction est inchangée)
        pass

if __name__ == "__main__":
    # ... (Ce bloc est inchangé)
    pass
