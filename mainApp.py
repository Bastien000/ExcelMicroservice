# scraper_app.py
import excelparser
import os
import glob
import time
from flask import Flask, Response, jsonify, request, send_file, after_this_request
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import excelCreator;
import test

# --- Konfigurace ---
EMAIL = "****"
PASSWORD = "****"


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "temp_excel_downloads")


EXPECTED_FILE_EXTENSION = ".xlsx" 

app = Flask(__name__)

if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
    print(f"Vytvořen adresář pro stahování: {DOWNLOAD_DIR}")

def perform_scraping_and_download_excel():
   
    print("Zahajuji scraping a stahování Excelu...")
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True 
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = None
    downloaded_file_path = None
    error_message = None

    try:
        print(f"Čistím adresář pro stahování: {DOWNLOAD_DIR}")
        for f_path in glob.glob(os.path.join(DOWNLOAD_DIR, f"*{EXPECTED_FILE_EXTENSION}")) + \
                      glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload")): 
            try:
                os.remove(f_path)
                print(f"Odstraněn starý soubor: {f_path}")
            except OSError as e:
                print(f"Chyba při odstraňování souboru {f_path}: {e}")


        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        wait = WebDriverWait(driver, 20) 

        print("Otevírám stránku https://live.zavodny.cz")
        driver.get("https://live.zavodny.cz")
        
        print("Vyplňuji e-mail...")
        wait.until(EC.visibility_of_element_located((By.ID, "login-username"))).send_keys(EMAIL)
        
        print("Vyplňuji heslo...")
        wait.until(EC.visibility_of_element_located((By.ID, "login-password"))).send_keys(PASSWORD)
        
        print("Klikám na 'Přihlásit se'...")
        wait.until(EC.element_to_be_clickable((By.ID, "btn-fblogin"))).click()
        
        # Čekání na přesměrování nebo načtení dashboardu (přizpůsobit podle potřeby)
        time.sleep(3) # Ponecháno, ale lepší by bylo čekat na konkrétní prvek
        if "login" in driver.current_url.lower(): # Jednoduchá kontrola
             raise Exception("Přihlášení se pravděpodobně nezdařilo, stále na přihlašovací stránce.")
        print("Přihlášení úspěšné.")

        print("Proklikávám se k exportu...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-content/app-content-comp/ul/li[1]/nav/ul/li[10]/a"))).click()
        time.sleep(2)

       
        count =0
        while True:
            try:
                driver.execute_script("window.scrollBy(100,0);")
                time.sleep(1)
                print("Delete table")
                driver.find_element(By.XPATH, "/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/table/thead/tr[1]/th[2]/button/i").click()
                time.sleep(1)
                count = 0
                driver.refresh()
            except:
                count = count+1
                if(count==5):
                    break
                
        
        time.sleep(1)
        print("Čekám na dropdown menu...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select"))).click()

        options_to_click = [1,2,3,4,5,6,7,8, 1,2,3,4,5,6,7,8]
        for option_text in options_to_click:
            print("Čekám na dropdown menu...")
            wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select"))).click()
            # VP1
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[12]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[13]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            # VP3
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[11]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[6]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            
            # VP2
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[5]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[7]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[8]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[9]/option[{option_text}]"))).click()
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            print(f"Klikám na {option_text}...")
            wait.until(EC.element_to_be_clickable((By.XPATH, f"/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[2]/fieldset/select/optgroup[10]/option[{option_text}]"))).click()
            
            print(f"Čekám na načtení tabulky pro {option_text}...")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(1)
            
        
        
        print("Proklikávání dokončeno.")
        time.sleep(1)

        export_button_xpath = "/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[1]/div[2]/button[2]/text()"
        print(f"Hledám tlačítko exportu: {export_button_xpath}")
        opak = 0
        clicked = False
        while not clicked:
            try:
                element = driver.find_element(By.XPATH, "/html/body/app-content/app-content-comp/ul/li[2]/ng-component/div[2]/div/div/div[1]/div[2]/button[2]/i")
                element.click()
                clicked = True
                if(opak==10):
                    break
            except WebDriverException:
                try:
                    driver.execute_script("window.scrollTo({ left: document.body.scrollWidth, behavior: 'smooth' });")
                    driver.execute_script("arguments[0].scrollIntoView({inline: 'end'});", driver.find_element(By.XPATH, export_button_xpath))
                    driver.execute_script("window.scrollBy(5000,0);")
                except WebDriverException:
                    print("re scrolling")
                    
                opak += 1
                time.sleep(1)

        print("Kliknuto na tlačítko pro stažení Excelu.")
        
        
       
        print(f"Čekám na stažení souboru s příponou {EXPECTED_FILE_EXTENSION} do adresáře {DOWNLOAD_DIR}...")
        timeout_download = 60  
        start_time_download = time.time()
        file_found = False
        while time.time() - start_time_download < timeout_download:
           
            downloaded_files = [
                f for f in glob.glob(os.path.join(DOWNLOAD_DIR, f"*{EXPECTED_FILE_EXTENSION}"))
                if not f.endswith(".crdownload")
            ]
            
            if downloaded_files:
             
                latest_file = max(downloaded_files, key=os.path.getmtime) 
                time.sleep(2) 
                if os.path.getsize(latest_file) > 0:
                    downloaded_file_path = latest_file
                    file_found = True
                    print(f"Soubor úspěšně stažen: {downloaded_file_path}")
                    break
                else:
                    print(f"Nalezen soubor {latest_file}, ale má nulovou velikost. Čekám...")
            
          
            crdownload_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
            if not crdownload_files and not downloaded_files:
               
                 print("Nenalezen žádný stahovaný soubor (.crdownload) ani cílový soubor. Čekám...")

            time.sleep(1)

        if not file_found:
            error_message = f"Soubor nebyl stažen v časovém limitu {timeout_download}s."
         
            print(f"Obsah adresáře {DOWNLOAD_DIR}: {os.listdir(DOWNLOAD_DIR)}")


    except Exception as e:
        print(f"Chyba ve funkci perform_scraping_and_download_excel: {e}")
        error_message = str(e)
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            driver.quit()
            print("WebDriver ukončen.")
    
    return downloaded_file_path, error_message

@app.route('/getData', methods=['GET']) 
def get_excel_file_endpoint():
    print(f"Endpoint /getData zavolán.")
    downloaded_file_path, error = perform_scraping_and_download_excel()

    if error or not downloaded_file_path:
        print(f"Chyba při scrapingu nebo stahování: {error}")
        return jsonify({"error": error or "Nepodařilo se stáhnout soubor."}), 500

    if not os.path.exists(downloaded_file_path):
        print(f"Stažený soubor nebyl nalezen na cestě: {downloaded_file_path}")
        return jsonify({"error": "Stažený soubor nebyl nalezen na serveru po dokončení scrapingu."}), 500

    try:
        
        @after_this_request
        def remove_file_after_send(response):
            try:
                if os.path.exists(downloaded_file_path):
                    os.remove(downloaded_file_path)
                    print(f"Dočasný soubor {downloaded_file_path} byl úspěšně smazán.")
                else:
                    print(f"Soubor {downloaded_file_path} již neexistoval při pokusu o smazání.")
            except Exception as e_remove:
                print(f"Chyba při mazání dočasného souboru {downloaded_file_path}: {e_remove}")
            return response
        
       
        if EXPECTED_FILE_EXTENSION.lower() == ".xlsx":
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif EXPECTED_FILE_EXTENSION.lower() == ".xls":
            mimetype = 'application/vnd.ms-excel'
        else:
            mimetype = 'application/octet-stream' # Obecný binární soubor

       # print(f"Odesílám soubor: {downloaded_file_path} s MIME typem {mimetype}")

       
        
        #send_file(
         #           downloaded_file_path,
          #          as_attachment=True, # Zajistí, že prohlížeč nabídne stažení
           #         download_name=os.path.basename(downloaded_file_path) # Použije původní název souboru
                    # mimetype=mimetype # send_file by měl MIME typ odvodit, ale explicitní je jistější
             #   )
        time.sleep(2)
        rt = excelparser.RunParser()
        
        # Convert all numpy types to native Python types for JSON serialization
        import numpy as np

        def convert_types(obj):
            if isinstance(obj, dict):
                return {k: convert_types(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [convert_types(i) for i in obj]
            elif isinstance(obj, (np.integer,)):
                return int(obj)
            elif isinstance(obj, (np.floating,)):
                return float(obj)
            else:
                return obj

        safe_rt = convert_types(dict(rt))
        return jsonify(safe_rt)
     
    
            

    except Exception as e_send:
        print(f"Chyba při odesílání souboru: {str(e_send)}")
        
        if os.path.exists(downloaded_file_path):
            try:
                os.remove(downloaded_file_path)
                print(f"Dočasný soubor {downloaded_file_path} byl smazán po chybě odeslání.")
            except Exception as e_critical_remove:
                 print(f"Kritická chyba při mazání souboru {downloaded_file_path} po chybě odeslání: {e_critical_remove}")
        return jsonify({"error": f"Chyba při odesílání souboru: {str(e_send)}"}), 500





@app.route('/getExcel', methods=['POST'])
def get_excel():
    data = request.get_json()
    excelCreator.createExcel(data)
    file_path = os.path.join(BASE_DIR, "temp_excel_downloads/tenant_data.xlsx")
    @after_this_request
    def remove_file_after_send(response):
        try:
            if os.path.exists("temp_excel_downloads/tenant_data.xlsx"):
                os.remove("temp_excel_downloads/tenant_data.xlsx")
                print(f"Dočasný soubor temp_excel_downloads/tenant_data.xlsx byl úspěšně smazán.")
            else:
                print(f"Soubor temp_excel_downloads/tenant_data.xlsx již neexistoval při pokusu o smazání.")
        except Exception as e_remove:
            print(f"Chyba při mazání dočasného souboru temp_excel_downloads/tenant_data.xlsx: {e_remove}")
        return response
    if os.path.exists(file_path):
        return send_file(
            file_path,
            as_attachment=True,
            download_name="tenant_data.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        return jsonify({"error": "Soubor tenant_data.xlsx nebyl nalezen."}), 404
    
    
  
       


if __name__ == '__main__':
   
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)
        print(f"Vytvořen adresář pro stahování při startu: {DOWNLOAD_DIR}")
    
    app.run(debug=True, host='0.0.0.0', port=5000)