from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, TimeoutException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager # type: ignore
import time
import pandas as pd
import os
import tkinter as tk
import requests
from tkinter import Tk, filedialog, messagebox
from RecaptchaSolver import RecaptchaSolver
from bs4 import BeautifulSoup

# Chrome settings
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)

# For Windows development, the ChromeDriver will be automatically downloaded
print("Starting Chrome...")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
print("Chrome started!")
wait = WebDriverWait(driver, 20)

recaptchaSolver = RecaptchaSolver(driver)

def show_warning():
    root = tk.Tk()
    root.title("Warning")

    label = tk.Label(root, text="The bot protection has been detected!", padx=10, pady=10)
    label.pack()

    button = tk.Button(root, text="Close", command=root.destroy)
    button.pack()

    root.mainloop()

try:
    start = time.time()
    
    # Use tkinter to open a file picker
    root = Tk()
    root.withdraw()
    root.call('wm', 'attributes', '.', '-topmost', True)

    excel_file_path = filedialog.askopenfilename(
        title="Choose an Excel or OpenOffice file",
        filetypes=[("Spreadsheet Files", "*.xlsx *.xls *.ods")]
    )

    if not excel_file_path:
        raise FileNotFoundError("No file selected. The program will close.")

    print(f"File path: {excel_file_path}")
    if excel_file_path.endswith('.ods'):
        try:
            import odf
            df = pd.read_excel(excel_file_path, engine='odf')
        except ImportError:
            print("Missing optional dependency 'odfpy'. Install it with 'pip install odfpy'.")
            os.system('pip install odfpy')
            import odf
            df = pd.read_excel(excel_file_path, engine='odf')
    else:
        df = pd.read_excel(excel_file_path)
        df['Check'] = None

    print("Columns in the DataFrame:", df.columns)

    # Go to Keepa.com and log in once
    driver.get("https://www.keepa.com/")

    time.sleep(5)

    try:
        driver.find_element(By.ID, "popup3")
        print("Captcha is present")
        recaptchaSolver.solve()
        time.sleep(2)

    except Exception as e:
        print(f"Captcha is not present")

    input("Please log in manually and press Enter when finished...")
    

    try:
        print("Waiting for an element to appear after login...")
        search_icon = wait.until(EC.visibility_of_element_located((By.ID, "menuSearch")))
        print("Login successful, the script will continue.")
    except TimeoutException:
        print("Login failed or the element could not be found.")
        driver.quit()
        exit()

    for index, row in df[df.iloc[:, 3].isna()].iterrows():

        if pd.isna(row['ASIN']):  # or np.isnan(row['ASIN']) if ASIN is a number  
            break

        try:
            driver.find_element(By.CLASS_NAME, "popup3")
            print("Captcha is present")
            recaptchaSolver.solve()
            time.sleep(2)

        except Exception as e:
            print(f"Captcha is not present")

        try:
            print(f"Processing row {index}...")
                    
            for attempt in range(3):
                try:
                    search_icon.click()
                    break  # Exit the loop if the click is successful
                except Exception:
                    time.sleep(1)
            
     
            # if driver.find_element(By.ID, "networkBar"):
            #     print("An error occurred - trying reloading Keepa.")
            #     driver.refresh()
            #     time.sleep(2)
            # else:
            #     print(f"NetworkBar is not present")

            try:
                driver.find_element(By.CLASS_NAME, "popup3")
                print("Captcha is present")
                recaptchaSolver.solve()
                time.sleep(2)

            except Exception as e:
                print(f"Captcha is not present")

            value_b2 = row['ASIN']
            print(f"Value in the second column, row {index}: {value_b2}")

            search_field = wait.until(EC.visibility_of_element_located((By.NAME, "search")))
            search_field.clear()
            search_field.send_keys(value_b2)
            search_field.send_keys(Keys.RETURN)
            time.sleep(1)

            try:
                driver.find_element(By.ID, "searchPage")
                print(f"The product of {value_b2} is not available")
                # continue
            except Exception:
                print(f"The product of {value_b2} is available")

            price_watch_tab = wait.until(EC.element_to_be_clickable((By.ID, "tabTrack")))
            price_watch_tab.click()
            time.sleep(1)

            try:
                driver.find_element(By.CLASS_NAME, "popup3")
                print("Captcha is present")
                recaptchaSolver.solve()
                time.sleep(2)

            except Exception as e:
                print(f"Captcha is not present")

            try: 
                driver.find_element(By.ID, "updateTracking")
                print(f"You are already tracking this product.")
                continue
            except Exception:
                value_c2 = row['DE Preis']
                print(f"Value DE Price, row {index}: {value_c2}")

            try:
                price_threshold_field = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "mdc-text-field__input")))
                price_threshold_field.clear()
                price_threshold_field.send_keys(str(value_c2))
                time.sleep(1)
            except Exception:
                print("The element cannot be found")

            multi_shop_toggle = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "tracking__tab-bar-toggle-hint")))
            multi_shop_toggle.click()
            time.sleep(1)

            try:
                driver.find_element(By.CLASS_NAME, "popup3")
                print("Captcha is present")
                recaptchaSolver.solve()
                time.sleep(2)

            except Exception as e:
                print(f"Captcha is not present")

            france_checkbox = wait.until(EC.presence_of_element_located((By.ID, "multilocale-4-checkbox")))
            italy_checkbox = wait.until(EC.presence_of_element_located((By.ID, "multilocale-8-checkbox")))
            spain_checkbox = wait.until(EC.presence_of_element_located((By.ID, "multilocale-9-checkbox")))

            if not france_checkbox.is_selected():
                france_checkbox.click()
            if not italy_checkbox.is_selected():
                italy_checkbox.click()
            if not spain_checkbox.is_selected():
                spain_checkbox.click()

            try:
                activate_domains_button = wait.until(EC.element_to_be_clickable((By.ID, "multilocale__submit")))
                activate_domains_button.click()
                time.sleep(1)

            except Exception:
                close_icon = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "shareChartOverlay-close")))
                driver.find_element(By.CLASS_NAME, "mdc-typography--caption").getText()
                print("Keepa can track this product in multiple Amazon locales for you. We can not guarantee this product will be shippable to you from all locales.")
                for attempt in range(3):
                    try:
                        close_icon.click()
                        break  # Exit the loop if the click is successful
                    except Exception:
                        time.sleep(1)

            submit_tracking_button = wait.until(EC.element_to_be_clickable((By.ID, "submitTracking")))
           
            try:
                driver.find_element(By.CLASS_NAME, "popup3")
                print("Captcha is present")
                recaptchaSolver.solve()
                time.sleep(2)

            except Exception as e:
                print(f"Captcha is not present")

            for attempt in range(3):
                try:
                    submit_tracking_button.click()
                    break  # Exit the loop if the click is successful
                except Exception:
                    time.sleep(1)  # Wait and retry if the click is intercepted+

            try:
                # t0 = time.time()
                driver.find_element(By.CLASS_NAME, "popup3")
                print("Captcha is present")
                # check_box = wait.until(EC.element_to_be_clickable((By.ID, "recaptcha-anchor")))
                recaptchaSolver.solve()
                # print(f"Time to solve the captcha: {time.time()-t0:.2f} seconds")
                # driver.find_element(By.ID, "recaptcha-demo-submit").click()
                time.sleep(2)

            except Exception as e:
                print(f"Captcha is not present")
            
            time.sleep(2)

            df.at[index, 'Check'] = 'x'
            print(f"Process completed for row {index}.")

        except TimeoutException:
            print(f"TimeoutException for row {index}. Skipping this row.")
            continue

    df.to_excel(excel_file_path, index=False)
    print("Updated file saved.")

except Exception as e:
    # print(f"An issue occurred: {e}")
    # html_source = driver.page_source
    # with open("saved_page.html", "w", encoding="utf-8") as file:
    #     file.write(html_source)
    # Get the page source
    html_source = driver.page_source
    soup = BeautifulSoup(html_source, "html.parser")
    print(soup)

    # Find and inline all CSS styles
    for link in soup.find_all("link", {"rel": "stylesheet"}):
        css_url = link["href"]
        if not css_url.startswith("http"):  # Convert relative URLs to absolute
            css_url = driver.current_url + css_url
        print(soup)

        try:
            response = requests.get(css_url)
            style_tag = soup.new_tag("style")
            style_tag.string = response.text
            soup.head.append(style_tag)
            link.extract()  # Remove the original <link> tag
        except:
            pass  # Skip if CSS file cannot be loaded

    print(soup)
    # Save the modified HTML
    with open("saved_page.html", "w", encoding="utf-8") as file:
        file.write(str(soup))

finally:
    stop = time.time()
    duration = stop - start
    print(f"The scraping took {duration} seconds.")
    pass

input("Script finished. Press Enter to close the console...")