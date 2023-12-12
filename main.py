import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import os
import gspread
import numpy as np
import pandas as pd
import gspread_dataframe as gd
import requests
from simple_salesforce import Salesforce
from io import StringIO
from datetime import datetime, timedelta

# Connexion aux SF
sf = Salesforce(username=os.environ["A"], password=os.environ["B"], security_token=os.environ["C"])
sf_gsl = Salesforce(username=os.environ["D"], password=os.environ["E"], security_token=os.environ["F"])

# Connexion au Gsheet
cred = {
  "type": "service_account",
  "project_id": os.environ["JSON_PROJECT_ID"],
  "private_key_id": os.environ["JSON_PRIVATE_KEY_ID"],
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDdJbJrJZwR0DGf\npbMpvMskK2d9Zgt/hYKPsm3tmW57I1/Bmqex0EyF7T2X7wNZYvownW4t/1SKI4dm\nX9EyN0qAWTGV8DYT6ikAV2F21e95VXTnmWTIAkbNqjcW3aqwlFb1vGlsGBwo4CUD\nV8I4aQCv9VwqIMlj2Pit4U8v5WEpD+CqKankYCVxblLECT2O4MpOYbAgahcLVVbz\nzPco4Z3ChVIc/oMZf3Tf+iQJjdXt4hlCVlGQwyOQj1/MGS4C+QgPTmgREtonANQZ\nzkxGukAShxcJmbW8t6JoIBflwS8C6Bkrq2wS7Da8IZGMeYyrbDz5YxF3tlMn5DQz\nd6JZ7W5nAgMBAAECggEAUUFYnSXkgmcM1Kv9eIh124RXMiwcbW6Q2lulKHgHsBb/\nSBObKipu84aH6xtXD7DeXJ57rUrztQ344hSyuNC3/xDGt2DbfdW72vRAS3mv09Ui\nbzZdYV/0w/yW4T/xR5A5o3DnC2CQeDRjZScJpdBvADgW9SO224kNVlZC0kZlvL/5\nmgQx0N5AhLuJZtRRSef08zJOrZwfBK+q+6AO45y8TbtI6nea3zScEEiUaqI+CUjX\n2o3jbLekHcTTnhGMY2FzOE3THqv0MOEiLyXv4xuYo2r3Xadql1meroz2G2Tm4YaN\nj7GsYIWeWRAp830QcXRq7f2rLE1qli5ZCFNEXD0MgQKBgQD75SWKGdbmtkc5qrc9\nXTGtNQRNDocpHuD5qZk38F7DRjEWelUtDQVy4QukJlPHi+8aoYSliUAUwvuquJ8F\nQrR2/u/e67X0ALu5j8XL0YNA8vwhVqVlgBLhRix72Mefsn/SMNDZZwJHbrV1fFjl\nk+S5p/nYWeg5by7Awx569EB+lQKBgQDgwEbY+dK2np4wCYZ5qgQs+EnOO8lHukoD\nW29XFKqoXXXXjj507zmXHwf27FdkKRAR0Dc+M8pSGuCT4DjmXBokRD0TIk+ZzvPD\ncfZ7Qjuibeuhv43+8T6haD/Dw+KSekd1LjtmwzI9ONsnFKVJ5J4stjBP+U/7vuhR\nsCzLCc6GCwKBgF/im0aVjXNnQXeXH4dxWT6Yptl6RUMG6RbAU32yty/YIUlwKcor\nYb6YIelLWark/pCBmU+2DjmY+1nCS24iNTXy13Zg/XMzcEIzk1SBnVf05rr+E5mu\nhgFQyBAgteR5eySUxntrNbfhUZu2SNSjVnbBlV6g4sAyLXbdD1Y4cfB9AoGAdOPr\ntQtxiONEOI2rr/k5xL25fRZhH/oGZmqWpL6t53T8RgjH+P82f592//h5lzE4F1uo\nb6R19G/gH2i9jymVuwj2js4IgmE9LIhH/mE7LMZoh65dxptnzICwsTteynnoUkyi\nPlcE+QxqBpBZSu4pSe3TgLSU4cSvhVTQZyUJkXMCgYAXawCUlbMbINf+HibpkA7i\n50FENfo6e/fZ/8tIR4I/HYOCQwFdLWlEq0QtFqXM0HoNMj1OPucoAl6vT3GFXmhY\nblClsBV3A8KlXO44HRealeKtp/W4nSW6WlvSRBa3xXiKsMmNlPoMDtja5aITlpdQ\nnYB0qGa7j/RhpwHS15Jikg==\n-----END PRIVATE KEY-----\n",
  "client_email": os.environ["JSON_CLIENT_EMAIL"],
  "client_id": os.environ["JSON_CLIENT_ID"],
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": os.environ["JSON_CLIENT_X509_CERT_URL"],
  "universe_domain": "googleapis.com"
}

sa = gspread.service_account_from_dict(cred)

#id_fichier = input("Entrez l'ID du Gsheet :")
id_fichier = "1NcyzXlV0iPYAro6_HwDnPjT3oI_7ikMrTpzhk6hKNuo"
fichier_gsheet = sa.open_by_key(id_fichier)

# Fonctions extraire rapports
def extraire_rapport(report_id):
    sf_org = "https://meilleursagents.my.salesforce.com/"
    export_params = '?isdtp=p1&export=1&enc=UTF-8&xf=csv'
    sf_report_url = sf_org + report_id + export_params
    response = requests.get(sf_report_url, headers=sf.headers, cookies={'sid': sf.session_id})
    new_report = response.content.decode('utf-8')
    report_df = pd.read_csv(StringIO(new_report))
    report_df = report_df.iloc[:-5]
    return report_df


def extraire_rapport_gsl(report_id):
    sf_org = "https://windu.my.salesforce.com/"
    export_params = '?isdtp=p1&export=1&enc=UTF-8&xf=csv'
    sf_report_url = sf_org + report_id + export_params
    response = requests.get(sf_report_url, headers=sf_gsl.headers, cookies={'sid': sf_gsl.session_id})
    new_report = response.content.decode('utf-8')
    report_df = pd.read_csv(StringIO(new_report))
    return report_df


# Fonctions intéragir avec Gsheet
def copier_coller_sur_nouvel_onglet(nom_nouvel_onglet, datafrayme):
    fichier_gsheet.add_worksheet(title=nom_nouvel_onglet, rows=1, cols=1) #nom_nouvel_onglet à renseigner en string
    feuille = fichier_gsheet.worksheet(nom_nouvel_onglet)
    gd.set_with_dataframe(feuille, datafrayme)


def maj_onglet(nom_onglet, datafrayme):
    feuille = fichier_gsheet.worksheet(nom_onglet)
    feuille.clear()
    gd.set_with_dataframe(feuille, datafrayme)



class Browser:
    browser, service = None, None

    def __init__(self, driver: str):
        cwd = os.getcwd()
        download_dir = os.path.join(cwd, "telechargementsTemp")
        # cleaner le contenu de telechargementsTemp au cas où :
        for filename in os.listdir(download_dir):
            file_path = os.path.join(download_dir, filename)
            os.remove(file_path)
        # setup le chromedriver
        self.service = Service(driver)
        op = webdriver.ChromeOptions()
        p = {"download.default_directory" : download_dir}
        op.add_experimental_option("prefs", p)
        op.add_argument("--headless=new")
        self.browser = webdriver.Chrome(service=self.service, options=op)

    def open_page(self, url: str):
        self.browser.get(url)

    def close_browser(self):
        self.browser.close()

    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        time.sleep(0.1)

    def click_button(self, by: By, value: str):
        button = self.browser.find_element(by=by, value=value)
        button.click()
        time.sleep(0.5)

    def login_diabolo3(self, username: str, password: str):
        self.add_input(by=By.ID, value="login", text=username)
        self.add_input(by=By.ID, value="password", text=password)
        self.click_button(by=By.XPATH, value="/html/body/dbl-root/dbl-login-page/dbl-branded-screen/article/ng-scrollbar/div/div/div/div/div/div[1]/form/dbl-auth-button-group/dbl-auth-button/button")

    def scraping_diabolo(self):
        #Performance des agents sur les appels entrants - MOIS EN COURS
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/div[1]/dbl-filters-panel/div/dbl-date-range-filter/dbl-filter-container/form/div/div/div')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-menu__item-btn[data-test-id="data-range-period-btn-CURRENT_MONTH"]')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-footer__btn[data-test-id="data-range-apply-btn"]')
        time.sleep(7)
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/header/button')
        self.click_button(by=By.CLASS_NAME, value='report-container__action-btn[data-test-id="cld-export-csv-btn"]')
        time.sleep(1)

        #Synthèse files - MOIS EN COURS
        self.click_button(by=By.CLASS_NAME, value='report-nav-item__btn--report[data-test-id="cld-nav-link-queues-voice"]')
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/div[1]/dbl-filters-panel/div/dbl-date-range-filter/dbl-filter-container/form/div/div/div')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-menu__item-btn[data-test-id="data-range-period-btn-CURRENT_MONTH"]')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-footer__btn[data-test-id="data-range-apply-btn"]')
        time.sleep(7)
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/header/button')
        self.click_button(by=By.CLASS_NAME, value='report-container__action-btn[data-test-id="cld-export-csv-btn"]')
        time.sleep(1)

        #Productivité Agents - MOIS EN COURS
        self.click_button(by=By.CLASS_NAME, value='report-nav-item__btn--report[data-test-id="cld-nav-link-agent-productivity"]')
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/div[1]/dbl-filters-panel/div/dbl-date-range-filter/dbl-filter-container/form/div/div/div')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-menu__item-btn[data-test-id="data-range-period-btn-CURRENT_MONTH"]')
        self.click_button(by=By.CLASS_NAME, value='dbl-range-footer__btn[data-test-id="data-range-apply-btn"]')
        time.sleep(7)
        self.click_button(by=By.XPATH, value='/html/body/dbl-root/cld-cold-statistics-container/ng-component/div/div/dbl-regular-report-container/div/header/button')
        self.click_button(by=By.CLASS_NAME, value='report-container__action-btn[data-test-id="cld-export-csv-btn"]')
        time.sleep(1)


if __name__ == '__main__':

# # Performance des agents sur les appels entrants - Synthèse files - Productivité Agents - MOIS EN COURS GSL
    browser = Browser(ChromeDriverManager().install())
    browser.open_page("https://fr4.engage.diabolocom.com/app/login")
    time.sleep(5)
    browser.login_diabolo3(os.environ["LOGINDIABOGSL"], os.environ["PWDDIABOGSL"])
    time.sleep(3)
    browser.open_page("https://fr4.engage.diabolocom.com/app/statistics-reports/reports/agent-performance-inbound/view")
    time.sleep(3)
    browser.scraping_diabolo()
    time.sleep(2)
# # Performance des agents sur les appels entrants - Synthèse files - Productivité Agents - MOIS EN COURS MA
    browser.open_page("https://fr3.engage.diabolocom.com/app/login")
    time.sleep(5)
    browser.login_diabolo3(os.environ["LOGINDIABOMA"], os.environ["PWDDIABOMA"])
    time.sleep(3)
    browser.open_page("https://fr3.engage.diabolocom.com/app/statistics-reports/reports/agent-performance-inbound/view")
    time.sleep(5)
    browser.scraping_diabolo()
    time.sleep(2)
    browser.close_browser()

# # Copy-Paste dans les onglets
    qs_par_collab_gsl = pd.read_csv('telechargementsTemp/AGENTS_PERFORMANCE_INBOUND_CALLS.csv', delimiter=";")
    maj_onglet("Perf. des agents sur les appels entrants - GSL", qs_par_collab_gsl)
    os.remove("telechargementsTemp/AGENTS_PERFORMANCE_INBOUND_CALLS.csv")

    synthese_files_gsl = pd.read_csv('telechargementsTemp/QUEUES_SYNTHESIS.csv', delimiter=";")
    maj_onglet("Synthèse files - GSL", synthese_files_gsl)
    os.remove("telechargementsTemp/QUEUES_SYNTHESIS.csv")

    productivite_agents_gsl = pd.read_csv('telechargementsTemp/AGENTS_PRODUCTIVITY.csv', delimiter=";")
    maj_onglet("Productivité Agents - GSL", productivite_agents_gsl)
    os.remove("telechargementsTemp/AGENTS_PRODUCTIVITY.csv")

    qs_par_collab_ma = pd.read_csv('telechargementsTemp/AGENTS_PERFORMANCE_INBOUND_CALLS (1).csv', delimiter=";")
    maj_onglet("Perf. des agents sur les appels entrants - MA", qs_par_collab_ma)
    os.remove("telechargementsTemp/AGENTS_PERFORMANCE_INBOUND_CALLS (1).csv")

    synthese_files_ma = pd.read_csv('telechargementsTemp/QUEUES_SYNTHESIS (1).csv', delimiter=";")
    maj_onglet("Synthèse files - MA", synthese_files_ma)
    os.remove("telechargementsTemp/QUEUES_SYNTHESIS (1).csv")

    productivite_agents_ma = pd.read_csv('telechargementsTemp/AGENTS_PRODUCTIVITY (1).csv', delimiter=";")
    maj_onglet("Productivité Agents - MA", productivite_agents_ma)
    os.remove("telechargementsTemp/AGENTS_PRODUCTIVITY (1).csv")

# tous les appels du mois sur SF
maj_onglet("SF GSL - Appels", extraire_rapport_gsl("00OSa000000Di85MAC"))
maj_onglet("SF MA - Appels", extraire_rapport("00O7S000000hAW8UAM"))

# mise à jour le
now = datetime.now() + timedelta(hours=1)
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
onglet_principal = fichier_gsheet.worksheet("Synthèse Mois en Cours")
onglet_principal.update("A1", "Dernière mise à jour Stats Appels (= Diabolocom) et OneShots (=Salesforce) : " + dt_string)
