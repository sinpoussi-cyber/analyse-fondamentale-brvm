# ==============================================================================
# 1. INSTALLATION DES DÉPENDANCES (AVEC LA CORRECTION DE VERSION)
# ==============================================================================
print("Installation des bibliothèques et du pilote de navigateur (avec correctif de version)...")
!pip install --upgrade pip -q
# CORRECTION FINALE : Forcer une version compatible de 'blinker' pour éviter le ModuleNotFoundError
!pip install blinker==1.6.2 selenium-wire gspread google-auth-oauthlib google-auth-httplib2 beautifulsoup4 requests python-docx pandas openpyxl pdfplumber -q

# Installation de ChromeDriver
!apt-get update > /dev/null
!apt-get install -y chromium-chromedriver > /dev/null

print("✅ Toutes les dépendances sont prêtes.\n")

# ==============================================================================
# 2. IMPORTATION ET CONFIGURATION
# ==============================================================================
import gspread
import requests
from bs4 import BeautifulSoup
import time
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from datetime import datetime
import logging
import io
import pdfplumber
import unicodedata
import urllib3
import json
from collections import defaultdict

# Imports pour Selenium-wire
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options

# MODIFIÉ : Imports pour l'authentification par compte de service et gestion des secrets
from google.colab import userdata # Spécifique à Colab pour les secrets
from google.oauth2 import service_account

# Désactiver les avertissements de sécurité
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==============================================================================
# 3. CONFIGURATION DU LOGGING
# ==============================================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ==============================================================================
# 4. CLASSE PRINCIPALE DE L'ANALYSEUR (VERSION SELENIUM-WIRE)
# ==============================================================================
class BRVMAnalyzer:
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id
        # Dictionnaire robuste avec de multiples termes de recherche
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

    def setup_selenium(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            logger.info("✅ Pilote Selenium-wire (Chrome) démarré avec succès.")
        except Exception as e:
            logger.error(f"❌ Impossible de démarrer le pilote Selenium: {e}")
            self.driver = None

    # MODIFIÉ : Authentification via un compte de service (plus robuste et portable)
    def authenticate_google_services(self):
        logger.info("Authentification Google via le compte de service...")
        try:
            # Pour Colab : Récupère le contenu JSON du secret "GSPREAD_SERVICE_ACCOUNT"
            # Pour GitHub Actions : Récupère le contenu depuis la variable d'environnement
            creds_json_str = userdata.get('GSPREAD_SERVICE_ACCOUNT') if 'google.colab' in sys.modules else os.environ.get('GSPREAD_SERVICE_ACCOUNT')
            
            if not creds_json_str:
                logger.error("❌ Le secret 'GSPREAD_SERVICE_ACCOUNT' est introuvable ou vide.")
                return False

            creds_dict = json.loads(creds_json_str)
            
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
            self.gc = gspread.authorize(creds)
            
            logger.info("✅ Authentification Google par compte de service réussie.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur lors de l'authentification par compte de service : {e}")
            return False

    def verify_and_filter_companies(self):
        try:
            logger.info(f"Vérification des feuilles dans G-Sheet...")
            sheet = self.gc.open_by_key(self.spreadsheet_id)
            existing_sheets = [ws.title for ws in sheet.worksheets()]
            symbols_to_keep = [s for s in self.original_societes_mapping if s in existing_sheets]
            missing_symbols = [s for s in self.original_societes_mapping if s not in existing_sheets]
            if missing_symbols:
                print("\n" + "="*50 + "\n⚠️  AVERTISSEMENT : FEUILLES MANQUANTES  ⚠️")
                for symbol in missing_symbols:
                    print(f"  - {symbol} ({self.original_societes_mapping[symbol]['nom_rapport']})")
                print("L'analyse continuera uniquement pour les sociétés trouvées.\n" + "="*50 + "\n")
            self.societes_mapping = {k: v for k, v in self.original_societes_mapping.items() if k in symbols_to_keep}
            logger.info(f"Analyse planifiée pour {len(self.societes_mapping)} sociétés.")
            return True
        except Exception as e:
            logger.error(f"❌ Erreur vérification G-Sheet: {e}")
            return False

    def _normalize_text(self, text):
        if not text: return ""
        text = ''.join(c for c in unicodedata.normalize('NFD', str(text).lower()) if unicodedata.category(c) != 'Mn')
        text = re.sub(r'[^a-z0-9\s]', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()

    def _find_all_reports_with_selenium_wire(self):
        if not self.driver:
            logger.error("Le pilote Selenium n'est pas disponible. Arrêt de la recherche.")
            return {}

        url = "https://www.brvm.org/fr/rapports-des-societes-cotees/all"
        companies_reports = defaultdict(list)

        try:
            logger.info(f"Navigation vers {url} et interception du trafic réseau...")
            self.driver.get(url)
            self.driver.wait_for_request('/views/ajax', timeout=30)
            logger.info("Requêtes AJAX initiales interceptées.")

            last_height = self.driver.execute_script("return document.body.scrollHeight")
            for i in range(15):
                del self.driver.requests
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                try:
                    self.driver.wait_for_request('/views/ajax', timeout=5)
                    logger.info(f"Chargement de la page de rapports {i+1}...")
                except:
                    logger.info("Fin du scroll (pas de nouvelles requêtes AJAX).")
                    break
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height: break
                last_height = new_height

            logger.info("Analyse des données interceptées...")
            full_html_content = ""
            for request in self.driver.requests:
                if request.response and '/views/ajax' in request.url:
                    body_json = json.loads(request.response.body.decode('utf-8'))
                    html_content = next((cmd.get('data', '') for cmd in body_json if cmd.get('command') == 'insert'), None)
                    if html_content:
                        full_html_content += html_content

            if full_html_content:
                soup = BeautifulSoup(full_html_content, 'html.parser')
                self._associate_reports_from_soup(soup, companies_reports)

        except Exception as e:
            logger.error(f"Erreur critique lors de la recherche avec Selenium-wire: {e}", exc_info=True)

        return companies_reports

    def _associate_reports_from_soup(self, soup, companies_reports):
        reports_found_count = 0
        potential_items = soup.select("div.views-row")

        for item in potential_items:
            link_tag = item.find('a', href=lambda href: href and '.pdf' in href.lower())
            if not link_tag: continue

            item_text_normalized = self._normalize_text(item.get_text())
            for symbol, info in self.societes_mapping.items():
                if any(self._normalize_text(alt) in item_text_normalized for alt in info['alternatives']):
                    href = link_tag.get('href')
                    full_url = href if href.startswith('http') else f"https://www.brvm.org{href}"
                    report_data = {
                        'titre': link_tag.get_text(strip=True),
                        'url': full_url,
                        'date': self._extract_date_from_text(item.get_text())
                    }
                    if not any(r['url'] == full_url for r in companies_reports[symbol]):
                        companies_reports[symbol].append(report_data)
                        reports_found_count += 1
                    break
        logger.info(f"{reports_found_count} rapports pertinents ont été associés à partir de ce bloc de données.")
        return reports_found_count

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
        if 'annuel' in text_lower or '31/12' in text or '31 dec' in text_lower:
            return datetime(year, 12, 31)
        return datetime(year, 6, 15)

    def _extract_financial_data_from_pdf(self, pdf_url):
        data = {'evolution_ca': 'Non trouvé', 'evolution_activites': 'Non trouvé', 'evolution_rn': 'Non trouvé'}
        try:
            logger.info(f"    -> Téléchargement et analyse du PDF...")
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
                if match:
                    data[key] = re.sub(r'[^\d\.\-%+]', '', match.group(1))
        except Exception as e:
            logger.warning(f"    -> Erreur lors de l'analyse du PDF: {e}")
        return data

    def process_all_companies(self):
        all_reports = self._find_all_reports_with_selenium_wire()
        results = {}
        total_reports_found = sum(len(reports) for reports in all_reports.values())

        if total_reports_found == 0:
            logger.error("❌ ÉCHEC FINAL : Aucun rapport trouvé même avec la méthode d'interception. Le site est peut-être en maintenance ou sa structure a radicalement changé.")
            return {}

        logger.info(f"✅ {total_reports_found} rapports trouvés au total.")

        for symbol, info in self.societes_mapping.items():
            logger.info(f"\n📊 Traitement de {symbol} - {info['nom_rapport']}")
            company_reports = all_reports.get(symbol, [])
            if not company_reports:
                logger.warning(f"  -> Aucun rapport trouvé pour {symbol}")
                results[symbol] = {'nom': info['nom_rapport'], 'statut': 'Aucun rapport trouvé', 'rapports_analyses': []}
                continue

            company_reports.sort(key=lambda x: x['date'], reverse=True)
            reports_to_analyze = company_reports[:5]
            logger.info(f"  -> {len(reports_to_analyze)} rapports récents seront analysés pour {symbol}.")

            analysis_data = {'nom': info['nom_rapport'], 'rapports_analyses': []}
            for i, report in enumerate(reports_to_analyze):
                logger.info(f"  -> Analyse {i+1}/{len(reports_to_analyze)}: {report['titre'][:60]}...")
                financial_data = self._extract_financial_data_from_pdf(report['url'])
                analysis_data['rapports_analyses'].append({
                    'titre': report['titre'], 'url': report['url'],
                    'date': report['date'].strftime('%Y-%m') if report['date'].year > 1900 else 'Date inconnue',
                    'donnees': financial_data
                })
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
            logger.info(f"✅ Rapport Word '{os.path.basename(output_path)}' généré avec succès.")
            print("\n" + "="*80)
            print("🎉 RAPPORT FINALISÉ 🎉")
            print(f"📊 Sociétés traitées: {companies_with_reports}/{total_companies}")
            print(f"📄 Rapports analysés: {total_reports}")
            print(f"📁 Fichier sauvegardé dans le répertoire courant : {output_path}")
            print("="*80 + "\n")
        except Exception as e:
            logger.error(f"❌ Impossible d'enregistrer le rapport Word : {e}")

    def run_analysis(self):
        try:
            logger.info("🚀 Démarrage de l'analyse BRVM (méthode d'interception réseau)...")
            self.setup_selenium()
            if not self.authenticate_google_services(): return
            if not self.verify_and_filter_companies(): return

            analysis_results = self.process_all_companies()

            if analysis_results and any(res.get('rapports_analyses') for res in analysis_results.values()):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                # MODIFIÉ : Chemin de sauvegarde local et portable
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
                logger.info("Navigateur Selenium-wire fermé.")
            logger.info("🏁 Fin du processus d'analyse.")

# ==============================================================================
# 5. EXÉCUTION PRINCIPALE
# ==============================================================================
# La condition `if __name__ == "__main__"` permet d'exécuter ce bloc 
# uniquement lorsque le script est lancé directement.
if __name__ == "__main__":
    # MODIFIÉ : Utilisation du nouvel ID de votre Spreadsheet
    # Extrait de l'URL : https://docs.google.com/spreadsheets/d/1EGXyg13ml8a9zr4OaUPnJN3i-rwVO2uq330yfxJXnSM/edit
    SPREADSHEET_ID = '1EGXyg13ml8a9zr4OaUPnJN3i-rwVO2uq330yfxJXnSM'
    
    # MODIFIÉ : Importation de sys ici pour la logique d'authentification
    import sys

    print("="*80)
    print("      🔍 ANALYSEUR FINANCIER BRVM - VERSION FINALE (INTERCEPTION RÉSEAU) 🔍")
    print("="*80)

    analyzer = BRVMAnalyzer(spreadsheet_id=SPREADSHEET_ID)
    analyzer.run_analysis()
