import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import logging
import re
import random
from datetime import datetime
import os

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class CompleteDATEVScraper:
    def __init__(self, headless=False):
        """Initialize the complete DATEV scraper"""
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        self.base_url = "https://www.datev.de/kasus/Start?KammerId=BuKa&Suffix1=BuKaY&Suffix2=BuKaXY"
        self.all_contacts = set()  # Use set to avoid duplicates
        self.contact_data = []
        
    def get_all_contacts_comprehensive(self):
        """
        Comprehensive strategy to get all contacts using multiple approaches
        """
        logger.info(f"Starting comprehensive extraction to get all {target_count} contacts")
        
        strategies = [
            self._strategy_random_searches,
            self._strategy_city_based,
            self._strategy_industry_based,
            self._strategy_name_based,
            self._strategy_postal_code_based
        ]
        
        for i, strategy in enumerate(strategies, 1):
            logger.info(f"Executing strategy {i}: {strategy.__name__}")
            try:
                strategy()
                logger.info(f"Current unique contacts: {len(self.all_contacts)}")
                
                # Save progress periodically
                if len(self.contact_data) > 0:
                    self._save_progress()
                    
            except Exception as e:
                logger.error(f"Strategy {i} failed: {e}")
                continue
        
        return self.contact_data
    
    def _strategy_random_searches(self, iterations=200):
        """Strategy 1: Multiple random searches to get different result sets"""
        logger.info("Strategy 1: Random searches")
        
        for i in range(iterations):
            try:
                logger.info(f"Random search iteration {i+1}/{iterations}")
                
                # Navigate to search page
                self.driver.get(self.base_url)
                time.sleep(2)
                
                # Set to 50 results
                try:
                    radio_50 = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @value='50']")))
                    radio_50.click()
                except:
                    pass
                
                # Click search without any criteria (random results)
                search_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
                search_button.click()
                time.sleep(3)
                
                # Extract results
                results = self._extract_page_results()
                self._add_unique_results(results)
                
                # Add random delay to avoid being blocked
                time.sleep(random.uniform(2, 5))
                
            except Exception as e:
                logger.warning(f"Random search {i+1} failed: {e}")
                continue
    
    def _strategy_city_based(self):
        """Strategy 2: Search by major German cities"""
        logger.info("Strategy 2: City-based searches")
        
        # Major German cities
        cities = [
            'Berlin', 'Hamburg', 'München', 'Köln', 'Frankfurt', 'Stuttgart', 'Düsseldorf',
            'Dortmund', 'Essen', 'Leipzig', 'Bremen', 'Dresden', 'Hannover', 'Nürnberg',
            'Duisburg', 'Bochum', 'Wuppertal', 'Bielefeld', 'Bonn', 'Münster', 'Karlsruhe',
            'Mannheim', 'Augsburg', 'Wiesbaden', 'Gelsenkirchen', 'Mönchengladbach',
            'Braunschweig', 'Chemnitz', 'Kiel', 'Aachen', 'Halle', 'Magdeburg', 'Freiburg',
            'Krefeld', 'Lübeck', 'Oberhausen', 'Erfurt', 'Mainz', 'Rostock', 'Kassel',
            'Hagen', 'Potsdam', 'Saarbrücken', 'Hamm', 'Mülheim', 'Ludwigshafen',
            'Leverkusen', 'Oldenburg', 'Osnabrück', 'Solingen', 'Heidelberg', 'Herne',
            'Neuss', 'Darmstadt', 'Paderborn', 'Regensburg', 'Ingolstadt', 'Würzburg',
            'Fürth', 'Wolfsburg', 'Offenbach', 'Ulm', 'Heilbronn', 'Pforzheim',
            'Göttingen', 'Bottrop', 'Trier', 'Recklinghausen', 'Reutlingen', 'Bremerhaven',
            'Koblenz', 'Bergisch Gladbach', 'Jena', 'Remscheid', 'Erlangen', 'Moers',
            'Siegen', 'Hildesheim', 'Salzgitter'
        ]
        
        for city in cities:
            try:
                logger.info(f"Searching city: {city}")
                results = self._search_with_criteria({'city': city})
                self._add_unique_results(results)
                time.sleep(random.uniform(2, 4))
            except Exception as e:
                logger.warning(f"City search for {city} failed: {e}")
                continue
    
    def _strategy_industry_based(self):
        """Strategy 3: Search by industry specializations"""
        logger.info("Strategy 3: Industry-based searches")
        
        # Major industries from the form
        industries = [
            'Einzelhandel', 'Ärzte', 'Rechtsanwälte', 'Immobilienmakler', 'Bauunternehmen - Hochbau',
            'Gaststätten, Hotel- und Übernachtungsgewerbe', 'Handwerksbetriebe', 'Unternehmensberater',
            'Freiberufler', 'Großhandel', 'Maschinenbau', 'Kfz-Handel', 'Banken / Kreditinstitute / Bausparkassen',
            'Versicherungen', 'Immobilienverwalter', 'Steuerberater', 'Wirtschaftsprüfer', 'Import-/Exportunternehmen',
            'Softwareentwicklung', 'Medien', 'Agrarwirtschaft, Land- und Forstwirte', 'Apotheken',
            'Zahnärzte', 'Architekten', 'Ingenieure', 'Verlage', 'Fotografen', 'Bäcker / Konditor',
            'Friseure', 'Elektrohandwerk', 'Heilberufe', 'Bildungseinrichtungen'
        ]
        
        for industry in industries:
            try:
                logger.info(f"Searching industry: {industry}")
                results = self._search_with_criteria({'industries': [industry]})
                self._add_unique_results(results)
                time.sleep(random.uniform(2, 4))
            except Exception as e:
                logger.warning(f"Industry search for {industry} failed: {e}")
                continue
    
    def _strategy_name_based(self):
        """Strategy 4: Search by common German surnames"""
        logger.info("Strategy 4: Name-based searches")
        
        # Common German surnames
        surnames = [
            'Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber', 'Meyer', 'Wagner',
            'Becker', 'Schulz', 'Hoffmann', 'Koch', 'Richter', 'Klein', 'Wolf',
            'Schröder', 'Neumann', 'Schwarz', 'Zimmermann', 'Braun', 'Krüger',
            'Hofmann', 'Hartmann', 'Lange', 'Schmitt', 'Werner', 'Schmitz',
            'Krause', 'Meier', 'Lehmann', 'Schmid', 'Schulze', 'Maier', 'Köhler',
            'Herrmann', 'König', 'Walter', 'Mayer', 'Huber', 'Kaiser', 'Fuchs',
            'Peters', 'Lang', 'Scholz', 'Möller', 'Weiß', 'Jung', 'Hahn',
            'Schubert', 'Schuster', 'Winkler', 'Berger', 'Lorenz', 'Ludwig'
        ]
        
        for surname in surnames:
            try:
                logger.info(f"Searching surname: {surname}")
                results = self._search_with_criteria({'name': surname})
                self._add_unique_results(results)
                time.sleep(random.uniform(2, 4))
            except Exception as e:
                logger.warning(f"Name search for {surname} failed: {e}")
                continue
    
    def _strategy_postal_code_based(self):
        """Strategy 5: Search by postal code ranges"""
        logger.info("Strategy 5: Postal code-based searches")
        
        # German postal code ranges (first 2 digits)
        postal_ranges = []
        for i in range(1, 100):
            postal_ranges.append(f"{i:02d}")
        
        for postal_range in postal_ranges:
            try:
                logger.info(f"Searching postal range: {postal_range}xxx")
                results = self._search_with_criteria({'postal_code': postal_range})
                self._add_unique_results(results)
                time.sleep(random.uniform(2, 4))
            except Exception as e:
                logger.warning(f"Postal code search for {postal_range} failed: {e}")
                continue
    
    def _search_with_criteria(self, criteria):
        """Perform a single search with given criteria"""
        try:
            # Navigate to search page
            self.driver.get(self.base_url)
            time.sleep(2)
            
            # Set to 50 results
            try:
                radio_50 = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='radio' and @value='50']")))
                radio_50.click()
            except:
                pass
            
            # Fill form with criteria
            self._fill_search_form(criteria)
            
            # Submit search
            search_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
            search_button.click()
            time.sleep(3)
            
            # Extract results
            return self._extract_page_results()
            
        except Exception as e:
            logger.error(f"Search failed: {e}")
            return []
    
    def _fill_search_form(self, criteria):
        """Fill the search form with provided criteria"""
        try:
            # Name field
            if criteria.get('name'):
                name_field = self.driver.find_element(By.NAME, "Name")
                name_field.clear()
                name_field.send_keys(criteria['name'])
            
            # City field
            if criteria.get('city'):
                city_field = self.driver.find_element(By.NAME, "Ort")
                city_field.clear()
                city_field.send_keys(criteria['city'])
            
            # Postal code field  
            if criteria.get('postal_code'):
                postal_field = self.driver.find_element(By.NAME, "Postleitzahl")
                postal_field.clear()
                postal_field.send_keys(criteria['postal_code'])
            
            # Industry checkboxes
            if criteria.get('industries'):
                for industry in criteria['industries']:
                    try:
                        industry_checkbox = self.driver.find_element(
                            By.XPATH, f"//input[@type='checkbox' and @value='{industry}']"
                        )
                        if not industry_checkbox.is_selected():
                            industry_checkbox.click()
                    except NoSuchElementException:
                        continue
                        
        except Exception as e:
            logger.warning(f"Error filling form: {e}")
    
    def _extract_page_results(self):
        """Extract results from current page"""
        results = []
        
        try:
            # Wait for page to load
            time.sleep(3)
            
            # Get page text and parse advisor blocks
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
            
            # Split by advisor entries (look for patterns)
            advisor_blocks = []
            lines = body_text.split('\n')
            
            current_block = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Start of new advisor block
                if (line.startswith(('Herrn', 'Frau')) or 
                    'Steuerberater' in line or 
                    'Steuerbevollmächtigte' in line or
                    line.endswith('GmbH') or line.endswith('mbH')):
                    
                    if current_block and len('\n'.join(current_block)) > 50:
                        advisor_blocks.append('\n'.join(current_block))
                    current_block = [line]
                else:
                    current_block.append(line)
            
            # Don't forget the last block
            if current_block and len('\n'.join(current_block)) > 50:
                advisor_blocks.append('\n'.join(current_block))
            
            # Parse each block
            for block in advisor_blocks:
                if ('Steuerberater' in block or 'Steuerbevollmächtigte' in block):
                    advisor_data = self._parse_advisor_block(block)
                    if advisor_data:
                        results.append(advisor_data)
                        
        except Exception as e:
            logger.error(f"Error extracting results: {e}")
            
        return results
    
    def _parse_advisor_block(self, block_text):
        """Parse individual advisor text block into structured data"""
        try:
            lines = [line.strip() for line in block_text.split('\n') if line.strip()]
            
            if len(lines) < 3:
                return None
            
            advisor_data = {
                'title': '',
                'name': '',
                'profession': '',
                'company': '',
                'address': '',
                'city': '',
                'postal_code': '',
                'phone': '',
                'fax': '',
                'mobile': '',
                'email': '',
                'website': '',
                'chamber': '',
                'unique_id': '',  # For deduplication
                'full_text': block_text
            }
            
            for i, line in enumerate(lines):
                # Title detection
                if line.startswith(('Herrn', 'Frau')) and not advisor_data['title']:
                    advisor_data['title'] = line
                
                # Profession detection
                elif any(prof in line for prof in ['Steuerberater', 'Steuerbevollmächtigte', 'Wirtschaftsprüfer']):
                    advisor_data['profession'] = line
                    # Previous line might be name if not company
                    if i > 0 and not any(suffix in lines[i-1] for suffix in ['GmbH', 'mbH', 'AG']):
                        advisor_data['name'] = lines[i-1]
                
                # Company name
                elif any(suffix in line for suffix in ['GmbH', 'mbH', 'AG', 'Partnerschaft', 'PartG']):
                    advisor_data['company'] = line
                
                # Address
                elif re.match(r'^[A-Za-zäöüÄÖÜß\s\-\.]+\s+\d+', line):
                    advisor_data['address'] = line
                
                # Postal code and city
                elif re.match(r'^\d{5}\s+[A-Za-zäöüÄÖÜß\s\-]+', line):
                    parts = line.split(' ', 1)
                    advisor_data['postal_code'] = parts[0]
                    if len(parts) > 1:
                        advisor_data['city'] = parts[1]
                
                # Contact info
                elif line.startswith('Tel.:'):
                    advisor_data['phone'] = line.replace('Tel.:', '').strip()
                elif line.startswith('Fax:'):
                    advisor_data['fax'] = line.replace('Fax:', '').strip()
                elif line.startswith('Mobil:'):
                    advisor_data['mobile'] = line.replace('Mobil:', '').strip()
                elif line.startswith('Email:'):
                    advisor_data['email'] = line.replace('Email:', '').strip()
                elif line.startswith('Internet:'):
                    advisor_data['website'] = line.replace('Internet:', '').strip()
                elif 'Steuerberaterkammer' in line:
                    advisor_data['chamber'] = line.replace('Zuständige Berufskammer:', '').strip()
            
            # Create unique ID for deduplication
            name_part = advisor_data.get('name', advisor_data.get('company', ''))
            address_part = advisor_data.get('address', '')
            advisor_data['unique_id'] = f"{name_part}_{address_part}".replace(' ', '_').lower()
            
            return advisor_data if name_part else None
            
        except Exception as e:
            logger.warning(f"Error parsing block: {e}")
            return None
    
    def _add_unique_results(self, results):
        """Add results to collection, avoiding duplicates"""
        for result in results:
            unique_id = result.get('unique_id', '')
            if unique_id and unique_id not in self.all_contacts:
                self.all_contacts.add(unique_id)
                self.contact_data.append(result)
    
    def _save_progress(self):
        """Save current progress to file"""
        if self.contact_data:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'datev_progress_{timestamp}.xlsx'
            
            df = pd.DataFrame(self.contact_data)
            df = df.drop_duplicates(subset=['unique_id'], keep='first')
            df.to_excel(filename, index=False)
            
            logger.info(f"Progress saved: {len(df)} unique contacts in {filename}")
    
    def save_final_results(self, filename='datev_all_contacts_final.xlsx'):
        """Save final results to Excel"""
        if not self.contact_data:
            logger.warning("No data to save")
            return
        
        df = pd.DataFrame(self.contact_data)
        df = df.drop_duplicates(subset=['unique_id'], keep='first')
        
        # Clean and organize columns
        column_order = ['name', 'title', 'profession', 'company', 'address', 'postal_code', 'city', 
                       'phone', 'fax', 'mobile', 'email', 'website', 'chamber']
        
        existing_columns = [col for col in column_order if col in df.columns]
        df_final = df[existing_columns + ['full_text']]
        
        df_final.to_excel(filename, index=False)
        logger.info(f"Final results saved: {len(df_final)} unique contacts in {filename}")
        
        return len(df_final)
    
    def close(self):
        """Close the browser"""
        self.driver.quit()

def main():
    """Main function to extract all DATEV contacts"""
    scraper = CompleteDATEVScraper(headless=False)
    
    try:
        logger.info("Starting comprehensive DATEV contact extraction")
        logger.info("Target: All available DATEV contacts")
        
        # Get all contacts using multiple strategies
        scraper.get_all_contacts_comprehensive()
        
        # Save final results
        final_count = scraper.save_final_results('datev_all_27k_contacts.xlsx')
        
        print(f"\n{'='*60}")
        print(f"EXTRACTION COMPLETED!")
        print(f"Total unique contacts extracted: {final_count}")
        print(f"All available contacts extracted.")
        print(f"Results saved to: datev_all_contacts.xlsx")
        print(f"{'='*60}")
        
    except KeyboardInterrupt:
        logger.info("Extraction interrupted by user")
        scraper.save_final_results('datev_partial_contacts.xlsx')
        
    except Exception as e:
        logger.error(f"Extraction failed: {e}")
        scraper.save_final_results('datev_error_recovery.xlsx')
        
    finally:
        scraper.close()

if __name__ == "__main__":
    main()
