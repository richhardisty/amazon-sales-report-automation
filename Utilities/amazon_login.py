# ──────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Company-specific
# account name replaced with an environment
# variable. Credentials load from a .env file
# via python-dotenv. See README.md for context.
# ──────────────────────────────────────────────

import os
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyotp
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

# Load .env from the parent directory
env_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '.env'))
load_dotenv(dotenv_path=env_path)

# Credentials — set these in your .env file
EMAIL           = os.getenv("EMAIL")
AMAZON_PASSWORD = os.getenv("AMAZON_PASSWORD")
AMAZON_KEY      = os.getenv("AMAZON_KEY")

# The account name shown in Vendor Central's account selector after login.
# Set VENDOR_ACCOUNT_NAME in your .env to match your account's display name.
VENDOR_ACCOUNT_NAME = os.getenv("VENDOR_ACCOUNT_NAME", "Your Vendor Account Name")


def amazon_login(driver, log_message):
    """
    Performs login to Amazon Vendor Central with email, password, and TOTP 2FA.

    Args:
        driver:      Selenium WebDriver instance
        log_message: Callable that accepts a string — used for timestamped logging

    Raises:
        ValueError:  If required environment variables are missing
        Exception:   If login fails at any step
    """
    if not all([EMAIL, AMAZON_PASSWORD, AMAZON_KEY]):
        log_message("Missing required environment variables (EMAIL, AMAZON_PASSWORD, AMAZON_KEY)")
        raise ValueError("Required credentials not set in .env")

    log_message("Starting Amazon login")

    while True:
        try:
            # ── Email ──────────────────────────────────────────────────────
            driver.find_element(By.ID, "ap_email").send_keys(EMAIL)
            log_message("Entered Amazon email")

            try:
                password_field = driver.find_element(By.ID, "ap_password")
            except Exception:
                password_field = None

            if not password_field:
                driver.find_element(By.ID, "continue").click()
                log_message("Clicked 'Continue' button")
                time.sleep(1)

            # ── Password ───────────────────────────────────────────────────
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ap_password"))
            ).send_keys(AMAZON_PASSWORD)
            log_message("Entered Amazon password")

            driver.find_element(By.ID, "signInSubmit").click()
            time.sleep(3)
            log_message("Clicked Amazon sign-in button")

            # ── TOTP 2FA ───────────────────────────────────────────────────
            totp = pyotp.TOTP(AMAZON_KEY)
            driver.find_element(By.ID, "auth-mfa-otpcode").send_keys(totp.now())
            log_message("Entered Amazon 2FA code")
            driver.find_element(By.ID, "auth-signin-button").click()
            time.sleep(3)
            log_message("Clicked Amazon 2FA submit button")

            # ── Account selector ───────────────────────────────────────────
            # Vendor Central shows an account picker if multiple accounts are
            # associated with the credential. VENDOR_ACCOUNT_NAME must match
            # the display name shown in the picker exactly.
            driver.find_element(
                By.XPATH,
                f'//button/span[text()="{VENDOR_ACCOUNT_NAME}"]'
            ).click()

            # The "Select account" button uses a Shadow DOM web component
            driver.execute_script("""
                const btn = document.querySelector('kat-button[label="Select account"]');
                if (btn) {
                    const inner = btn.shadowRoot.querySelector('button');
                    if (inner) inner.click();
                }
            """)
            log_message("Clicked 'Select account' button via Shadow DOM")
            time.sleep(3)
            log_message("Account selected")

            # ── Dismiss tour dialog if present ─────────────────────────────
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        '//button[contains(@class, "take-tour-dialog-content-ctas-tertiary")'
                        ' and contains(text(), "Maybe later")]'
                    ))
                ).click()
                log_message("Clicked 'Maybe later' button")
            except Exception:
                log_message("'Maybe later' button not found, continuing")

            # ── Check for error page ───────────────────────────────────────
            try:
                err = driver.find_element(
                    By.XPATH, '//li[contains(text(), "Error, please try again.")]'
                )
                if err.is_displayed():
                    log_message("Error page detected — restarting login")
                    driver.refresh()
                    continue
            except Exception:
                log_message("No error message found, continuing")

            log_message("Amazon login successful")
            break

        except Exception as e:
            log_message(f"Error during Amazon login: {e}")
            driver.refresh()
            continue


if __name__ == "__main__":
    """
    Standalone test mode — initializes a Chrome WebDriver, logs in,
    and pauses for inspection. Useful for verifying credentials independently.
    """

    def standalone_log_message(message):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {message}")

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(
        service=ChromeService(ChromeDriverManager().install()),
        options=chrome_options
    )

    try:
        driver.get("https://vendorcentral.amazon.com")
        standalone_log_message("Navigated to Amazon Vendor Central")
        amazon_login(driver, standalone_log_message)
        standalone_log_message("Standalone Amazon login completed successfully")
        input("Press Enter to close the browser...")
    except Exception as e:
        standalone_log_message(f"Standalone Amazon login failed: {e}")
        raise
    finally:
        driver.quit()
        standalone_log_message("Closed WebDriver")
