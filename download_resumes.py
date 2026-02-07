import openpyxl
import json
import os
import time
import re
import random
from playwright.sync_api import sync_playwright

# Configuration
DOWNLOAD_DIR = os.path.join(os.getcwd(), 'resumes')
MIN_DELAY = 5   # Minimum seconds to wait
MAX_DELAY = 12  # Maximum seconds to wait

def get_profile_data(excel_file):
    """Extracts names and profile URLs from Excel."""
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active
    
    # Identify columns
    header = {cell.value: i for i, cell in enumerate(sheet[1])}
    col_profile = header.get('Candidate profile')
    col_name = header.get('Name')
    
    if col_profile is None or col_name is None:
        raise ValueError("Could not find 'Candidate profile' or 'Name' columns")

    profiles = []
    # Starting from row 2
    for row in sheet.iter_rows(min_row=2, values_only=False):
        name_cell = row[col_name]
        profile_cell = row[col_profile]
        
        url = None
        if profile_cell.hyperlink:
            url = profile_cell.hyperlink.target
        elif profile_cell.value and 'http' in str(profile_cell.value):
             url = profile_cell.value
             
        if url:
            profiles.append({
                'name': name_cell.value,
                'url': url
            })
            
    print(f"Found {len(profiles)} profiles with URLs.")
    return profiles

def clean_filename(name):
    return re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')

def random_sleep(min_seconds=1, max_seconds=3):
    sleep_time = random.uniform(min_seconds, max_seconds)
    print(f"Sleeping for {sleep_time:.2f} seconds...")
    time.sleep(sleep_time)

def run(excel_file, cookies):
    # Create download directory if not exists
    if not os.path.exists(DOWNLOAD_DIR):
        os.makedirs(DOWNLOAD_DIR)

    # Fix cookie sameSite values
    for cookie in cookies:
        same_site = cookie.get('sameSite')
        if same_site in ['unspecified', 'no_restriction']:
            cookie['sameSite'] = 'Lax'
        elif same_site:
            cookie['sameSite'] = same_site.capitalize() # lax -> Lax, strict -> Strict

    profiles = get_profile_data(excel_file)
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        
        # Add cookies
        try:
            context.add_cookies(cookies)
            print("Cookies added successfully.")
        except Exception as e:
            print(f"Error adding cookies: {e}")
            
        page = context.new_page()
        
        for i, profile in enumerate(profiles):
            name = profile['name']
            url = profile['url']
            
            print(f"Processing ({i+1}/{len(profiles)}): {name}")
            
            # Check if file already exists to avoid re-downloading
            safe_name = clean_filename(str(name))
            # Simple check for existing files starting with the name
            existing = [f for f in os.listdir(DOWNLOAD_DIR) if f.startswith(safe_name)]
            if existing:
                 print(f"Skipping {name}, already downloaded: {existing[0]}")
                 continue

            try:
                page.goto(url, timeout=60000)
                random_sleep(3, 6) # Wait for page load with jitter
                
                # Check for login requirement
                if "login" in page.url:
                    print(f"Redirected to login for {name}. Cookies might be invalid or expired.")
                    continue

                # Locate download button
                try:
                    with page.expect_download(timeout=10000) as download_info:
                        if page.get_by_role("button", name="Download").is_visible():
                            page.get_by_role("button", name="Download").click()
                        elif page.locator("text=Download Resume").is_visible():
                            page.locator("text=Download Resume").click()
                        elif page.locator(".download-resume").is_visible():
                            page.locator(".download-resume").click() 
                        else:
                             page.get_by_text("Download").first.click()
                             
                    download = download_info.value
                    
                    file_name = f"{safe_name}_{int(time.time())}.pdf"
                    save_path = os.path.join(DOWNLOAD_DIR, file_name)
                    
                    download.save_as(save_path)
                    print(f"Downloaded: {file_name}")
                    
                except Exception as e:
                    print(f"Download failed for {name}: {e}")
                    
            except Exception as e:
                print(f"Error navigating to {url}: {e}")
            
            # Rate limiting with jitter
            random_sleep(MIN_DELAY, MAX_DELAY)
            
        browser.close()

if __name__ == "__main__":
    # Configuration
    EXCEL_FILE = 'Founders-Office_20260206154224_160.xlsx'
    
    # Default cookies (should be updated if expired)
    # Paste your JSON cookies here (true/false will be automatically converted)
    RAW_COOKIES = '''[
        {"domain":".naukri.com","expirationDate":1773052151.381563,"hostOnly":false,"httpOnly":true,"name":"_did","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"207e8fe2a8"},
        {"domain":".naukri.com","expirationDate":1773052151.381893,"hostOnly":false,"httpOnly":true,"name":"_odur","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"7f14503371"},
        {"domain":".naukri.com","expirationDate":1773052151.382437,"hostOnly":false,"httpOnly":true,"name":"kycEligibleCookie124691914","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"false"},
        {"domain":".naukri.com","expirationDate":1773052151.382534,"hostOnly":false,"httpOnly":true,"name":"UNPC","path":"/","sameSite":"lax","secure":true,"session":false,"storeId":"0","value":"124691914"},
        {"domain":".naukri.com","expirationDate":1773052151.382554,"hostOnly":false,"httpOnly":true,"name":"UNCC","path":"/","sameSite":"lax","secure":true,"session":false,"storeId":"0","value":"124922930"},
        {"domain":".naukri.com","expirationDate":1801996151.382634,"hostOnly":false,"httpOnly":true,"name":"loginMethod","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"naukriLauncher"},
        {"domain":".naukri.com","expirationDate":1801996151.382653,"hostOnly":false,"httpOnly":true,"name":"loginPreference","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"secureLoginMandatory"},
        {"domain":".naukri.com","expirationDate":1775644151.382859,"hostOnly":false,"httpOnly":false,"name":"secureloginenabled","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"true"},
        {"domain":".naukri.com","expirationDate":1801996151.382973,"hostOnly":false,"httpOnly":false,"name":"lastLoggedInUser","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"hello@humanbit.ai"},
        {"domain":".naukri.com","expirationDate":1773052151.382001,"hostOnly":false,"httpOnly":true,"name":"_t_ds","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"190fe9731770301747-16190fe973-0190fe973"},
        {"domain":".naukri.com","expirationDate":1801999515.371526,"hostOnly":false,"httpOnly":true,"name":"J","path":"/","sameSite":"lax","secure":true,"session":false,"storeId":"0","value":"0"},
        {"domain":".naukri.com","expirationDate":1771064952.563301,"hostOnly":false,"httpOnly":true,"name":"test","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"naukri.com"},
        {"domain":".naukri.com","expirationDate":1778099081,"hostOnly":false,"httpOnly":false,"name":"_gcl_au","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"1.1.687357093.1770323081"},
        {"domain":".naukri.com","expirationDate":1778236131,"hostOnly":false,"httpOnly":false,"name":"_fbp","path":"/","sameSite":"lax","secure":false,"session":false,"storeId":"0","value":"fb.1.1770323081004.331369503913054433"},
        {"domain":".naukri.com","expirationDate":1770467236.966416,"hostOnly":false,"httpOnly":false,"name":"bm_mi","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"66A4C43EA49366AC7EF1B5AECBFC34DD~YAAQNkYDF9kM9iCcAQAAcj+kNx5n0krOL4QD8XueUpzs8txtO2pVgsn0zn+HxS54j0XtXfdR7W1n/Y2sVzrVUadfAEwINxEQEqY6hG1fyiW1UvatyPLPuMfYi+zgaRfovvn0EZuubu/dJfMMIwsB3QohH6W3ZugpFAcm3dHdqY/l3+sBwtoxfceFR0OLj4vagD6KtwSUqGHiEmm/nlixRJO2cnsZEUQtBontB71ZKarJzrYOP019xRF8xv1nKcL2tGi0R4NR9cp2ckniyVfd1SQYPBMyR55FiqjkUoOJ2HQ2Lgd2e0M6Op0ryB822SFWc0npOyM6DE/358GX/ND1R/5v~1"},
        {"domain":".naukri.com","expirationDate":1770467236.797133,"hostOnly":false,"httpOnly":true,"name":"ak_bmsc","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"32F0C3469A420CD94486B8FD04C89F41~000000000000000000000000000000~YAAQNkYDFwYN9iCcAQAAuUKkNx6vLK/KWHw4BMU1mYcKQXzf+m11TxnZipJ4dIplbVNNKuYsE0UutJN2dD8q97XUYbqwVsbfPpHfv9EvO5yhsAc2l4e5sDKWhW5jjbxSQsJQISuplOoOe+HvnrgrkCTYtmEthg0OV+ZomKk9DVayX5tvp2YYb7dZAvfuzX8IzCEEnjFME6clVe/NcqxI2yHhOnWJeKEgsItUDiDLhwa2lLOrkk3Z9APMdRP5LrzfRlVM+FygursaTT52apzuYmqA/xIeEoImOs8gaga1kJyDcf+pVHjzsNd8TdiI1kmrWZ7K9VhoWIZwmJVJ8wTDiuIgDW1OhiwuEpMcGOTC/bxTtzuhiyHA5yl/gAfY1gQ68ARYWSUeGdqfUaibKls+mW+QH159s2Mdzo+UsT564IGoB6zQHTwpZqbxS8TldMiBWheN5cbZ9lokiTvcJ+xl/PMadRdF4FsK2B05WX6Dmg63DtK4o6ECDjq0OArSOqZVrRw+SxoyDQu03rj/ekyOGBALGztMfUx5nMCobh4y"},
        {"domain":".naukri.com","expirationDate":1778236040,"hostOnly":false,"httpOnly":false,"name":"_gcl_gs","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"2.1.k1$i1770460040$u93512962"},
        {"domain":".naukri.com","expirationDate":1778236097,"hostOnly":false,"httpOnly":false,"name":"_gcl_aw","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"GCL.1770460097.Cj0KCQiA4pvMBhDYARIsAGfgwvzUumRUljNFxRdvcqIXs1JbgKkRcB2ZvZgQ6z0CwzQQSqpvsNOuXm8aAnilEALw_wcB"},
        {"domain":".naukri.com","expirationDate":1778236097,"hostOnly":false,"httpOnly":false,"name":"_gcl_dc","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"GCL.1770460097.Cj0KCQiA4pvMBhDYARIsAGfgwvzUumRUljNFxRdvcqIXs1JbgKkRcB2ZvZgQ6z0CwzQQSqpvsNOuXm8aAnilEALw_wcB"},
        {"domain":".naukri.com","expirationDate":1770546508.345292,"hostOnly":false,"httpOnly":true,"name":"_t_s","path":"/","sameSite":"no_restriction","secure":true,"session":false,"storeId":"0","value":"seo"},
        {"domain":".naukri.com","expirationDate":1770546508.345322,"hostOnly":false,"httpOnly":true,"name":"_t_sd","path":"/","sameSite":"no_restriction","secure":true,"session":false,"storeId":"0","value":"google"},
        {"domain":".naukri.com","hostOnly":false,"httpOnly":true,"name":"_t_r","path":"/","sameSite":"no_restriction","secure":true,"session":true,"storeId":"0","value":"1030%2F%2F"},{"domain":".naukri.com","hostOnly":false,"httpOnly":true,"name":"persona","path":"/","sameSite":"no_restriction","secure":true,"session":true,"storeId":"0","value":"default"},
        {"domain":".naukri.com","expirationDate":1770467323.055956,"hostOnly":false,"httpOnly":false,"name":"SnippedURL","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"https%3A%2F%2Frecruit.naukri.com%2F"},{"domain":".naukri.com","expirationDate":1805020123.456708,"hostOnly":false,"httpOnly":false,"name":"_ga_K2YBNZVRLL","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"GS2.1.s1770460037$o3$g1$t1770460123$j34$l0$h0"},
        {"domain":".naukri.com","hostOnly":false,"httpOnly":true,"name":"4cd30c17163a8dcddf808f8343a98b751s7","path":"/","sameSite":"unspecified","secure":true,"session":true,"storeId":"0","value":"a"},{"domain":".naukri.com","expirationDate":1770546551.382251,"hostOnly":false,"httpOnly":false,"name":"bs_rnd","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"K767b83K"},
        {"domain":".naukri.com","expirationDate":1770546551.382377,"hostOnly":false,"httpOnly":true,"name":"c89e57bb00ace2e24dfcc2d8ba377c871s7","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"v0%7C8h0MYZagg3Ho91AbtCDZe%2BPvVEfynKtaC78iurXzQYZSI9NyrnuI9wLodXrH7Q9AVUV%2B17o%2F%2Bj0Esb1AX5Fe0AtUkL0LNANJHSQPmtT6kp1PY5%2FZGNqVGc28zFAM1NZzDd8piE%2F6r%2FYjS%2BowpAoMusgxbteBuTOH53GZ%2BfIixrY%3D"},
        {"domain":".naukri.com","expirationDate":1770463751.382481,"hostOnly":false,"httpOnly":false,"name":"pvId","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"1"},
        {"domain":".naukri.com","expirationDate":1773052151.382493,"hostOnly":false,"httpOnly":true,"name":"ACCESS","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"1770460151350"},
        {"domain":".naukri.com","expirationDate":1773052151.382521,"hostOnly":false,"httpOnly":true,"name":"UNID","path":"/","sameSite":"lax","secure":false,"session":false,"storeId":"0","value":"vyDgpBK4Kxaf9LLKxtwq044oc1ludUflQdo8Py7M"},
        {"domain":".naukri.com","expirationDate":1770546551.382957,"hostOnly":false,"httpOnly":true,"name":"dashboard","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"1"},
        {"domain":".naukri.com","expirationDate":1770546551.382994,"hostOnly":false,"httpOnly":false,"name":"encId","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"a9a421a10082c87d06a7960f179a1719595f0d5848110012036"},
        {"domain":".naukri.com","expirationDate":1805020152.421332,"hostOnly":false,"httpOnly":false,"name":"_ga","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"GA1.2.2019136079.1770301749"},
        {"domain":".naukri.com","expirationDate":1770546552,"hostOnly":false,"httpOnly":false,"name":"_gid","path":"/","sameSite":"unspecified","secure":false,"session":false,"storeId":"0","value":"GA1.2.383760221.1770460152"},
        {"domain":".naukri.com","expirationDate":1770463633,"hostOnly":false,"httpOnly":false,"name":"HOWTORT","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"ul=1770463512271&r=https%3A%2F%2Fhiring.naukri.com%2Fhiring%2F260126002073%2Fapply%2F69817198a309101aa5bb29b6%3Fsrc%3DexcelDownload&hd=1770463513074"},
        {"domain":".naukri.com","expirationDate":1770467236.722077,"hostOnly":false,"httpOnly":false,"name":"bm_sv","path":"/","sameSite":"unspecified","secure":true,"session":false,"storeId":"0","value":"E7E43C63B2C904112BD4F4D4CABDB446~YAAQLUYDF77anwqcAQAAlEzZNx6UnOUR+lmBZ4sScJV6w63fDx0LlA87B29QIfbiN1+GiGC2JfeZ9AlQFmt81JwqzZLDzxIepDZ9OYLdD8v2tNMKuHTuX/Bm03iEPWDzk+plFGRESqgsZ8ZBEL0uNW7rHoQVBY6Jn6rTZ7r6t2/Y6VEyWlQS8TbaWRXmVUQmPyxrDHo5kwZUMfJTUKAgMMS3ACH4WxNe7syVfsvw1qTy2xdjzspIM//89DlsPcc7kg==~1"}
    ]'''
    
    COOKIES = json.loads(RAW_COOKIES)
    
    try:
        run(EXCEL_FILE, COOKIES)
    except Exception as e:
        print(f"Fatal error: {e}")
