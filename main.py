# ==============================================================================
# ANALYSEUR FINANCIER BRVM - SCRIPT FINAL
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
import pdfplumber
import unicodedata
import urllib3
import json
from collections import defaultdict

# Imports Selenium
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Imports pour l'authentification Google
try:
    from google.colab import userdata
except ImportError:
    userdata = None
from google.oauth2 import service_account

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
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id
        # Dictionnaire des sociétés à suivre. Les clés (ex: 'ABJC') DOIVENT correspondre aux noms des onglets dans le Google Sheet.
        self.societes_mapping = {
            'ABJC': {'nom_rapport': 'SERVAIR ABIDJAN CI', 'alternatives': ['servair', 'servair abidjan', 'abjc']},
            'BICB': {'nom_rapport': 'BIIC BN', 'alternatives': ['biic', 'bicb']},
            'BICC': {'nom_rapport': 'BICI CI', 'alternatives': ['bici', 'bicc']},
            'BNBC': {'nom_rapport': 'BERNABE CI', 'alternatives': ['bernabe', 'bnbc']},
            'BOAB': {'nom_rapport': 'BANK OF AFRICA BN', 'alternatives': ['bank of africa', 'boa', 'boab', 'benin']},
            'BOABF': {'nom_rapport': 'BANK OF AFRICA BF', 'alternatives': ['bank of africa', 'boa', 'boabf', 'burkina']},
            'BOAC': {'nom_rapport': 'BANK OF AFRICA CI', 'alternatives': ['bank of africa', 'boa', 'boac', 'ivoire']},
            'BOAM': {'nom_rapport': 'BANK OF AFRICA ML', 'alternatives': ['bank of africa', 'boa', 'boam', 'mali']},
            'BOAN': {'nom_rapport': 'BANK OF AFRICA NG', 'alternatives': ['bank of africa', 'boa', 'boan', 'niger']},
            'BOAS': {'nom_rapport': 'BANK OF AFRICA SENEGAL', 'alternatives': ['bank of africa', 'boa', 'boas', 'senegal']},
            'CABC': {'nom_rapport': 'SICABLE CI', 'alternatives': ['sicable', 'cabc']},
            'CBIBF': {'nom_rapport': 'CORIS BANK INTERNATIONAL', 'alternatives': ['coris', 'cbibf']},
            'CFAC': {'nom_rapport': 'CFAO MOTORS CI', 'alternatives': ['cfao', 'cfac']},
            'CIEC': {'nom_rapport': 'CIE CI', 'alternatives': ['cie', 'ciec']},
            'ECOC': {'nom_rapport': 'ECOBANK COTE D\'IVOIRE', 'alternatives': ['ecobank', 'ecoc']},
            'ETIT': {'nom_rapport': 'ECOBANK TRANS. INCORP. TG', 'alternatives': ['ecobank', 'eti', 'etit']},
            'FTSC': {'nom_rapport': 'FILTISAC CI', 'alternatives': ['filtisac', 'ftsc']},
            'LNBB': {'nom_rapport': 'LOTERIE NATIONALE DU BENIN', 'alternatives': ['loterie', 'lnbb']},
            'NEIC': {'nom_rapport': 'NEI-CEDA CI', 'alternatives': ['nei', 'ceda', 'neic']},
            'NSBC': {'nom_rapport': 'NSIA BANQUE COTE D\'IVOIRE', 'alternatives': ['nsia', 'nsbc']},
            'NTLC': {'nom_rapport': 'NESTLE CI', 'alternatives': ['nestle', 'ntlc']},
            'ONTBF': {'nom_rapport': 'ONATEL BF', 'alternatives': ['onatel', 'ontbf']},
            'ORAC': {'nom_rapport': 'ORANGE COTE D\'IVOIRE', 'alternatives': ['orange', 'orac']},
            'ORGT': {'nom_rapport': 'ORAGROUP TOGO', 'alternatives': ['oragroup', 'orgt']},
            'PALC': {'nom_rapport': 'PALM CI', 'alternatives': ['palm', 'palmci', 'palc']},
            'PRSC': {'nom_rapport': 'TRACTAFRIC MOTORS CI', 'alternatives': ['tractafric', 'prsc']},
            'SAFC': {'nom_rapport': 'SAFCA CI', 'alternatives': ['safca', 'safc']},
            'SCRC': {'nom_rapport': 'SUCRIVOIRE', 'alternatives': ['sucrivoire', 'scrc']},
            'SDCC': {'nom_rapport': 'SODE CI', 'alternatives': ['sodeci', 'sode', 'sdcc']},
            'SDSC': {'nom_rapport': 'AFRICA GLOBAL LOGISTICS CI', 'alternatives': ['africa global logistics', 'bollore', 'sdsc']},
            'SEMC': {'nom_rapport': 'EVIOSYS PACKAGING SIEM CI', 'alternatives': ['eviosys', 'siem', 'crown', 'semc']},
            'SGBC': {'nom_rapport': 'SOCIETE GENERALE COTE D\'IVOIRE', 'alternatives': ['societe generale', 'sgbci', 'sgbc']},
            'SHEC': {'nom_rapport': 'VIVO ENERGY CI', 'alternatives': ['vivo energy', 'shell', 'shec']},
            'SIBC': {'nom_rapport': 'SOCIETE IVOIRIENNE DE BANQUE', 'alternatives': ['sib', 'sibc']},
            'SICC': {'nom_rapport': 'SICOR CI', 'alternatives': ['sicor', 'sicc']},
            'SIVC': {'nom_rapport': 'AIR LIQUIDE CI', 'alternatives': ['air liquide', 'sivc']},
            'SLBC': {'nom_rapport': 'SOLIBRA CI', 'alternatives': ['solibra', 'slbc']},
            'SMBC': {'nom_rapport': 'SMB CI', 'alternatives': ['smb', 'smbc']},
            'SNTS': {'nom_rapport': 'SONATEL SN', 'alternatives': ['sonatel', 'snts']},
            'SOGC': {'nom_rapport': 'SOGB CI', 'alternatives': ['sogb', 'sogc']},
            'SPHC': {'nom_rapport': 'SAPH CI', 'alternatives': ['saph', 'sphc']},
            'STAC': {'nom_rapport': 'SETAO CI', 'alternatives': ['setao', 'stac']},
            'STBC': {'nom_rapport': 'SITAB CI', 'alternatives': ['sitab', 'stbc']},
            'TTLC': {'nom_rapport': 'TOTALENERGIES MARKETING CI', 'alternatives': ['total', 'totalenergies', 'ttlc']},
            'TTLS': {'nom_rapport': 'TOTALENERGIES MARKETING SN', 'alternatives': ['total', 'totalenergies', 'ttls', 'senegal']},
            'UNLC': {'nom_rapport': 'UNILEVER CI', 'alternatives': ['unilever', 'unlc']},
            'UNXC': {'nom_rapport': 'UNIWAX CI', 'alternatives': ['uniwax', 'unxc']}
        }
        self.gc = None
        self.driver = None
        self.original_societes_mapping = self.societes_mapping.copy()
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})

    # ===== CORRECTION APPLIQUÉE ICI =====
    def setup_selenium(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument("--window-size=1920,1080")
        
        # NOUVELLE LIGNE : Spécifier l'emplacement du binaire Chromium pour l'environnement GitHub
        chrome_options.binary_location = '/usr/bin/chromium-browser'
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            logger.info("✅ Pilote Selenium (Chrome) démarré avec succès.")
        except Exception as e:
            logger.error(f"❌ Impossible de démarrer le pilote Selenium: {e}")
            self.driver = None

    def authenticate_google_services(self):
        logger.info("Authentification Google via le compte de service...")
        try:
            creds_json_str = None
            if userdata:
                creds_json_str = userdata.get('GSPREAD_SERVICE_ACCOUNT')
            else:
                creds_json_str = os.environ.get('GSPREAD_SERVICE_ACCOUNT')

            if not creds_json_str:
                logger.error("❌ Le secret 'GSPREAD_SERVICE_ACCOUNT' est introuvable ou vide.")
                return False
            creds_dict = json.loads(creds_json_str)
            scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
            self.gc = gspread.authorize(creds)
            logger.info("✅ Authentification Google par compte de service réussie.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur lors de l'authentification par compte de service : {e}")
            return False

    def verify_and_filter_companies(self):
        try:
            logger.info(f"Vérification des feuilles dans G-Sheet (ID: {self.spreadsheet_id})...")
            sheet = self.gc.open_by_key(self.spreadsheet_id)
            existing_sheets = [ws.title for ws in sheet.worksheets()]
            
            logger.info(f"Onglets trouvés dans le G-Sheet: {existing_sheets}")

            symbols_to_keep = [s for s in self.original_societes_mapping if s in existing_sheets]
            missing_symbols = [s for s in self.original_societes_mapping if s not in existing_sheets]
            
            if missing_symbols:
                print("\n" + "="*50 + "\n⚠️  AVERTISSEMENT : FEUILLES MANQUANTES  ⚠️")
                for symbol in missing_symbols:
                    print(f"  - {symbol} ({self.original_societes_mapping[symbol]['nom_rapport']})")
                print("L'analyse continuera uniquement pour les sociétés trouvées.\n" + "="*50 + "\n")
            
            self.societes_mapping = {k: v for k, v in self.original_societes_mapping.items() if k in symbols_to_keep}
            
            if not self.societes_mapping:
                logger.error("❌ ERREUR FATALE : Aucune société à analyser. Vérifiez que les noms des onglets de votre Google Sheet correspondent aux symboles du script (ex: 'BOAC', 'SNTS').")
                return False
            
            logger.info(f"✅ Vérification réussie. {len(self.societes_mapping)} sociétés seront analysées.")
            return True
            
        except gspread.exceptions.SpreadsheetNotFound:
            logger.error(f"❌ Erreur: Le Spreadsheet avec l'ID '{self.spreadsheet_id}' est introuvable.")
            logger.error("Veuillez vérifier que l'ID est correct et que le compte de service a les droits d'accès 'Lecteur'.")
            return False
        except Exception as e:
            logger.error(f"❌ Erreur lors de la vérification du G-Sheet: {e}")
            return False

    def _normalize_text(self, text):
        if not text: return ""
        text = ''.join(c for c in unicodedata.normalize('NFD', str(text).lower()) if unicodedata.category(c) != 'Mn')
        text = re.sub(r'[^a-z0-9\s]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()
    
    def _find_all_reports_with_selenium_wire(self):
        if not self.driver: return {}
        url = "https://www.brvm.org/fr/rapports-des-societes-cotees/all"
        companies_reports = defaultdict(list)
        try:
            logger.info(f"Navigation vers {url}...")
            self.driver.get(url)

            try:
                cookie_wait = WebDriverWait(self.driver, 5)
                cookie_button = cookie_wait.until(EC.element_to_be_clickable((By.ID, "tarteaucitronPersonalize2")))
                logger.info("Bannière de cookies trouvée. Clic sur 'Accepter'.")
                cookie_button.click()
                time.sleep(2)
            except (TimeoutException, NoSuchElementException):
                logger.info("Aucune bannière de cookies n'a été détectée.")

            wait = WebDriverWait(self.driver, 30)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.view-content")))
            logger.info("Le conteneur des rapports a été trouvé sur la page.")

            last_height = self.driver.execute_script("return document.body.scrollHeight")
            for i in range(20):
                soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                found_count = self._associate_reports_from_soup(soup, companies_reports)
                logger.info(f"Itération {i+1}: {found_count} nouveaux rapports. Total unique: {sum(len(v) for v in companies_reports.values())}")
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(3)
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    logger.info("Fin du scroll, la hauteur de la page ne change plus.")
                    break
                last_height = new_height
        
        except TimeoutException:
            logger.error("Échec : Le conteneur des rapports n'est pas apparu dans le temps imparti.")
            self._save_debug_info()
            return {}
        except Exception as e:
            logger.error(f"Erreur critique lors du scraping : {e}", exc_info=True)
            self._save_debug_info()
            return {}

        if not companies_reports:
             logger.warning("Le scraping s'est terminé mais aucun rapport n'a pu être associé. Sauvegarde des infos de débogage.")
             self._save_debug_info()
        return companies_reports

    def _save_debug_info(self):
        try:
            screenshot_path = 'debug_screenshot.png'
            html_path = 'debug_page.html'
            self.driver.save_screenshot(screenshot_path)
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(self.driver.page_source)
            logger.info(f"Infos de débogage sauvegardées : '{screenshot_path}' et '{html_path}'.")
        except Exception as e:
            logger.error(f"Impossible de sauvegarder les infos de débogage : {e}")

    def _associate_reports_from_soup(self, soup, companies_reports):
        reports_found_this_pass = 0
        potential_items = soup.select("div.view-content div.views-row")
        if not potential_items:
            logger.warning("Aucun élément 'div.views-row' trouvé dans le HTML analysé.")
        for item in potential_items:
            link_tag = item.find('a', href=lambda href: href and '.pdf' in href.lower())
            if not link_tag: continue
            item_text_normalized = self._normalize_text(item.get_text())
            for symbol, info in self.societes_mapping.items():
                if any(self._normalize_text(alt) in item_text_normalized for alt in info['alternatives']):
                    href = link_tag.get('href')
                    full_url = href if href.startswith('http') else f"https://www.brvm.org{href}"
                    if not any(r['url'] == full_url for r in companies_reports[symbol]):
                        report_data = {'titre': link_tag.get_text(strip=True), 'url': full_url, 'date': self._extract_date_from_text(item.get_text())}
                        companies_reports[symbol].append(report_data)
                        reports_found_this_pass += 1
                    break 
        return reports_found_this_pass
        
    def _extract_date_from_text(self, text):
        if not text: return datetime(1900, 1, 1)
        year_match = re.search(r'\b(20\d{2})\b', text)
        if not year_match: return datetime(1900, 1, 1)
        year = int(year_match.group(1))
        text_lower = text.lower()
        trim_match = re.search(r't(\d)|(\d)\s*er\s*trimestre', text_lower)
        if trim_match:
            trimester = int(trim_match.group(1) or trim_match.group(2))
            return datetime(year, trimester * 3, 1)
        sem_match = re.search(r's(\d)|(\d)\s*er\s*semestre', text_lower)
        if sem_match:
            semester = int(sem_match.group(1) or sem_match.group(2))
            return datetime(year, 6 if semester == 1 else 12, 1)
        if 'annuel' in text_lower or '31/12' in text or '31 dec' in text_lower: return datetime(year, 12, 31)
        return datetime(year, 6, 15)

    def _extract_financial_data_from_pdf(self, pdf_url):
        data = {'evolution_ca': 'Non trouvé', 'evolution_activites': 'Non trouvé', 'evolution_rn': 'Non trouvé'}
        try:
            logger.info(f"    -> Analyse du PDF...")
            response = self.session.get(pdf_url, timeout=45, verify=False)
            response.raise_for_status()
            with pdfplumber.open(io.BytesIO(response.content)) as pdf:
                full_text = " ".join(page.extract_text() for page in pdf.pages if page.extract_text())
            if not full_text: return data
            clean_text = re.sub(r'\s+', ' ', full_text.lower().replace('\n', ' ').replace(',', '.'))
            patterns = {
                'evolution_ca': r"chiffre d'affaires.*?(?:évolution|variation|progression de|hausse de|baisse de|[\+\-–—])\s*([\+\-–—]?\s*\d+[\d\.\s]*%)",
                'evolution_activites': r"(?:résultat des activités ordinaires|résultat d'exploitation).*?(?:évolution|variation|progression de|hausse de|baisse de|[\+\-–—])\s*([\+\-–—]?\s*\d+[\d\.\s]*%)",
                'evolution_rn': r"résultat net.*?(?:évolution|variation|progression de|hausse de|baisse de|[\+\-–—])\s*([\+\-–—]?\s*\d+[\d\.\s]*%)"
            }
            for key, pattern in patterns.items():
                match = re.search(pattern, clean_text, re.IGNORECASE)
                if match: data[key] = re.sub(r'[^\d\.\-%+]', '', match.group(1))
        except Exception as e:
            logger.warning(f"    -> Erreur lors de l'analyse du PDF: {e}")
        return data

    def process_all_companies(self):
        all_reports = self._find_all_reports_with_selenium_wire()
        results = {}
        total_reports_found = sum(len(reports) for reports in all_reports.values())
        if total_reports_found == 0:
            logger.error("❌ ÉCHEC FINAL : Aucun rapport trouvé sur le site de la BRVM.")
            return {}
        logger.info(f"✅ {total_reports_found} rapports uniques trouvés au total.")
        for symbol, info in self.societes_mapping.items():
            logger.info(f"\n📊 Traitement de {symbol} - {info['nom_rapport']}")
            company_reports = all_reports.get(symbol, [])
            if not company_reports:
                logger.warning(f"  -> Aucun rapport trouvé pour {symbol}")
                results[symbol] = {'nom': info['nom_rapport'], 'statut': 'Aucun rapport trouvé', 'rapports_analyses': []}
                continue
            company_reports.sort(key=lambda x: x['date'], reverse=True)
            reports_to_analyze = company_reports[:5]
            analysis_data = {'nom': info['nom_rapport'], 'rapports_analyses': []}
            for i, report in enumerate(reports_to_analyze):
                logger.info(f"  -> Analyse {i+1}/{len(reports_to_analyze)}: {report['titre'][:60]}...")
                financial_data = self._extract_financial_data_from_pdf(report['url'])
                analysis_data['rapports_analyses'].append({'titre': report['titre'], 'url': report['url'], 'date': report['date'].strftime('%Y-%m') if report['date'].year > 1900 else 'Date inconnue', 'donnees': financial_data})
                time.sleep(1)
            results[symbol] = analysis_data
        logger.info("\n✅ Traitement de toutes les sociétés terminé.")
        return results

    def create_word_report(self, results, output_path):
        logger.info(f"Création du rapport Word : {output_path}")
        try:
            doc = Document()
            doc.styles['Normal'].font.name = 'Calibri'
            doc.styles['Normal'].font.size = Pt(11)
            doc.add_heading('Analyse Financière des Sociétés Cotées', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('Bourse Régionale des Valeurs Mobilières (BRVM)', 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = doc.add_paragraph(f'\nRapport généré le : {datetime.now().strftime("%d %B %Y à %H:%M")}\n', style='Caption')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            total_companies = len(results)
            companies_with_reports = len([r for r in results.values() if r.get('rapports_analyses')])
            total_reports = sum(len(r.get('rapports_analyses', [])) for r in results.values())
            stats_p = doc.add_paragraph()
            stats_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            stats_run = stats_p.add_run(f'Synthèse : {companies_with_reports}/{total_companies} sociétés avec rapports trouvés • {total_reports} rapports récents analysés')
            stats_run.bold = True
            doc.add_page_break()
            for symbol, data in results.items():
                doc.add_heading(f"{symbol} - {data['nom']}", level=2)
                if not data.get('rapports_analyses'):
                    doc.add_paragraph("❌ Aucun rapport pertinent n'a été trouvé ou analysé pour cette société.")
                    continue
                table = doc.add_table(rows=1, cols=5, style='Table Grid')
                headers = ['Titre du Rapport', 'Date', 'Évol. CA', 'Évol. Activités', 'Évol. RN']
                for i, header_text in enumerate(headers):
                    run = table.rows[0].cells[i].paragraphs[0].add_run(header_text)
                    run.bold = True
                for rapport in data['rapports_analyses']:
                    row_cells = table.add_row().cells
                    row_cells[0].text = rapport['titre'][:70] + ('...' if len(rapport['titre']) > 70 else '')
                    row_cells[1].text = rapport['date']
                    donnees = rapport['donnees']
                    row_cells[2].text = donnees.get('evolution_ca', 'N/A')
                    row_cells[3].text = donnees.get('evolution_activites', 'N/A')
                    row_cells[4].text = donnees.get('evolution_rn', 'N/A')
                doc.add_paragraph()
            doc.save(output_path)
            print("\n" + "="*80 + "\n🎉 RAPPORT FINALISÉ 🎉\n" + f"📁 Fichier sauvegardé : {output_path}" + "\n" + "="*80 + "\n")
        except Exception as e:
            logger.error(f"❌ Impossible d'enregistrer le rapport Word : {e}")

    def run(self):
        try:
            logger.info("🚀 Démarrage de l'analyse BRVM...")
            self.setup_selenium()
            if not self.driver or not self.authenticate_google_services(): return
            if not self.verify_and_filter_companies(): return
            analysis_results = self.process_all_companies()
            if analysis_results and any(res.get('rapports_analyses') for res in analysis_results.values()):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                output_filename = f"Analyse_Financiere_BRVM_{timestamp}.docx"
                self.create_word_report(analysis_results, output_filename)
            else:
                logger.warning("❌ Aucun résultat d'analyse à inclure dans le rapport.")
                print("\n" + "="*60 + "\n⚠️  AUCUN RAPPORT GÉNÉRÉ\n" + "="*60)
        except Exception as e:
            logger.critical(f"❌ Une erreur critique a interrompu l'analyse: {e}", exc_info=True)
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("Navigateur Selenium fermé.")
            logger.info("🏁 Fin du processus d'analyse.")

# ------------------------------------------------------------------------------
# 4. POINT D'ENTRÉE DU SCRIPT
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    SPREADSHEET_ID = '1EGXyg13ml8a9zr4OaUPnJN3i-rwVO2uq330yfxJXnSM'
    print("="*50 + "\n      🔍 ANALYSEUR FINANCIER BRVM 🔍\n" + "="*50)
    analyzer = BRVMAnalyzer(spreadsheet_id=SPREADSHEET_ID)
    analyzer.run()
