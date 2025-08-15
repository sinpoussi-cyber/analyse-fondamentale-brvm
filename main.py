# ==============================================================================
# ANALYSEUR FINANCIER BRVM - SCRIPT FINAL V5.3 (DICTIONNAIRE AMÉLIORÉ)
# ==============================================================================

# ... [Toutes les importations et configurations restent identiques] ...
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
from selenium.common.exceptions import TimeoutException
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
        # ===== DICTIONNAIRE AMÉLIORÉ AVEC PLUS D'ALTERNATIVES =====
        self.societes_mapping = {
            'ABJC': {'nom_rapport': 'SERVAIR ABIDJAN CI', 'alternatives': ['servair abidjan', 'servair']},
            'SIVC': {'nom_rapport': 'AIR LIQUIDE CI', 'alternatives': ['air liquide ci']},
            'BOABF': {'nom_rapport': 'BANK OF AFRICA BF', 'alternatives': ['bank of africa bf', 'boa bf']},
            'BOAB': {'nom_rapport': 'BANK OF AFRICA BN', 'alternatives': ['bank of africa bn', 'boa bn']},
            'BOAC': {'nom_rapport': 'BANK OF AFRICA CI', 'alternatives': ['bank of africa ci', 'boa ci']},
            'BOAM': {'nom_rapport': 'BANK OF AFRICA ML', 'alternatives': ['bank of africa ml', 'boa ml']},
            'BOAN': {'nom_rapport': 'BANK OF AFRICA NG', 'alternatives': ['bank of africa ng', 'boa ng']},
            'BOAS': {'nom_rapport': 'BANK OF AFRICA SN', 'alternatives': ['bank of africa sn', 'boa sn']},
            'BNBC': {'nom_rapport': 'BERNABE CI', 'alternatives': ['bernabe ci']},
            'BICC': {'nom_rapport': 'BICI CI', 'alternatives': ['bici ci']},
            'CABC': {'nom_rapport': 'CABC', 'alternatives': ['cabc']},
            'CFAC': {'nom_rapport': 'CFAO MOTORS CI', 'alternatives': ['cfao motors ci']},
            'CIEC': {'nom_rapport': 'CIE CI', 'alternatives': ['cie ci']},
            'CBIBF': {'nom_rapport': 'CORIS BANK INTERNATIONAL', 'alternatives': ['coris bank international', 'coris bank']},
            'ECOC': {'nom_rapport': 'ECOBANK COTE D\'IVOIRE', 'alternatives': ["ecobank cote d ivoire", 'ecobank ci']},
            'ETIT': {'nom_rapport': 'ECOBANK TRANS. INCORP. TG', 'alternatives': ['ecobank trans', 'ecobank togo']},
            'FTSC': {'nom_rapport': 'FILTISAC CI', 'alternatives': ['filtisac ci']},
            'NEIC': {'nom_rapport': 'NEI-CEDA CI', 'alternatives': ['nei ceda ci']},
            'NSBC': {'nom_rapport': 'NSIA BANQUE CI', 'alternatives': ['nsia banque ci']},
            'ONTBF': {'nom_rapport': 'ONATEL BF', 'alternatives': ['onatel bf']},
            'ORAC': {'nom_rapport': 'ORANGE CI', 'alternatives': ['orange ci', "cote d'ivoire telecom"]},
            'PALC': {'nom_rapport': 'PALM CI', 'alternatives': ['palm ci']},
            'SAFC': {'nom_rapport': 'SAFCA CI', 'alternatives': ['safca ci']},
            'SPHC': {'nom_rapport': 'SAPH CI', 'alternatives': ['saph ci']},
            'STAC': {'nom_rapport': 'SETAO CI', 'alternatives': ['setao ci']},
            'SGBC': {'nom_rapport': 'SOCIETE GENERALE CI', 'alternatives': ['societe generale ci', 'sg ci']},
            'SIBC': {'nom_rapport': 'SOCIETE IVOIRIENNE DE BANQUE', 'alternatives': ['societe ivoirienne de banque', 'sib']},
            'SLBC': {'nom_rapport': 'SOLIBRA CI', 'alternatives': ['solibra ci']},
            'SNTS': {'nom_rapport': 'SONATEL SN', 'alternatives': ['sonatel sn', 'fctc sonatel']},
            'SCRC': {'nom_rapport': 'SUCRIVOIRE CI', 'alternatives': ['sucrivoire ci']},
            'TTLC': {'nom_rapport': 'TOTALENERGIES MARKETING CI', 'alternatives': ['totalenergies marketing ci', 'total ci']},
            'TTLS': {'nom_rapport': 'TOTALENERGIES MARKETING SN', 'alternatives': ['totalenergies marketing senegal', 'total sn']},
            'UNLC': {'nom_rapport': 'UNILEVER CI', 'alternatives': ['unilever ci']},
            'UNXC': {'nom_rapport': 'UNIWAX CI', 'alternatives': ['uniwax ci']},
            'SHEC': {'nom_rapport': 'VIVO ENERGY CI', 'alternatives': ['vivo energy ci']},
        }
        self.gc = None
        self.driver = None
        self.original_societes_mapping = self.societes_mapping.copy()
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'})

    # ... [Le reste du code est IDENTIQUE et n'a pas besoin d'être modifié]
    def setup_selenium(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.binary_location = '/usr/bin/chromium-browser'
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            logger.info("✅ Pilote Selenium (Chrome) démarré.")
        except Exception as e:
            logger.error(f"❌ Impossible de démarrer le pilote Selenium: {e}")
            self.driver = None

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
        text = ''.join(c for c in unicodedata.normalize('NFD', str(text).lower()) if unicodedata.category(c) != 'Mn')
        text = re.sub(r'[^a-z0-9\s]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()
    
    def _find_all_reports(self):
        if not self.driver: return {}
        
        main_page_url = "https://www.brvm.org/fr/rapports-societes-cotees"
        all_reports = defaultdict(list)
        
        try:
            logger.info(f"Navigation vers la page principale : {main_page_url}")
            self.driver.get(main_page_url)
            wait = WebDriverWait(self.driver, 20)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table")))
            
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            company_links = []
            table_rows = soup.select("table.views-table tbody tr")
            for row in table_rows:
                link_tag = row.find('a', href=True)
                if link_tag:
                    company_name_normalized = self._normalize_text(link_tag.text)
                    company_url = f"https://www.brvm.org{link_tag['href']}"
                    
                    symbol = self._get_symbol_from_name(company_name_normalized)
                    if symbol and symbol in self.societes_mapping:
                        company_links.append({'symbol': symbol, 'url': company_url})
            
            logger.info(f"{len(company_links)} pages de sociétés pertinentes à visiter.")

            for company in company_links:
                symbol = company['symbol']
                logger.info(f"--- Collecte pour {symbol} ---")
                
                try:
                    self.driver.get(company['url'])
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.views-table")))
                    
                    page_soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                    report_items = page_soup.select("table.views-table tbody tr")

                    if not report_items:
                        logger.warning(f"  -> Aucun rapport listé sur la page de {symbol}.")
                        continue

                    for item in report_items:
                        pdf_link_tag = item.find('a', href=lambda href: href and '.pdf' in href.lower())
                        if pdf_link_tag:
                            full_url = pdf_link_tag['href'] if pdf_link_tag['href'].startswith('http') else f"https://www.brvm.org{pdf_link_tag['href']}"
                            if not any(r['url'] == full_url for r in all_reports[symbol]):
                                report_data = {
                                    'titre': " ".join(item.get_text().split()),
                                    'url': full_url,
                                    'date': self._extract_date_from_text(item.get_text())
                                }
                                all_reports[symbol].append(report_data)
                                logger.info(f"  -> Trouvé : {report_data['titre'][:70]}...")
                    time.sleep(1)
                except TimeoutException:
                    logger.error(f"  -> Timeout sur la page de {symbol}. Passage au suivant.")
                except Exception as e:
                    logger.error(f"  -> Erreur sur la page de {symbol}: {e}. Passage au suivant.")
        
        except Exception as e:
            logger.error(f"Erreur critique lors du scraping : {e}", exc_info=True)
            return {}
            
        return all_reports

    def _get_symbol_from_name(self, company_name_normalized):
        for symbol, info in self.original_societes_mapping.items():
            for alt in info['alternatives']:
                if alt in company_name_normalized:
                    return symbol
        return None

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
        all_reports = self._find_all_reports()
        results = {}
        total_reports_found = sum(len(reports) for reports in all_reports.values())
        if total_reports_found == 0:
            logger.error("❌ ÉCHEC FINAL : Aucun rapport n'a pu être collecté sur le site de la BRVM.")
            return {}
        logger.info(f"\n✅ COLLECTE TERMINÉE : {total_reports_found} rapports uniques trouvés.")
        for symbol, info in self.societes_mapping.items():
            logger.info(f"\n📊 Traitement des données pour {symbol} - {info['nom_rapport']}")
            company_reports = all_reports.get(symbol, [])
            if not company_reports:
                logger.warning(f"  -> Aucun rapport trouvé pour {symbol} lors du traitement.")
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
            total_companies = len(self.societes_mapping)
            companies_with_reports = len([r for r in results.values() if r.get('rapports_analyses') and any(r['rapports_analyses'])])
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
    analyzer.run()```
