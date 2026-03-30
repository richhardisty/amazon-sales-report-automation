# ──────────────────────────────────────────────
# DEMO VERSION
# Sanitized for portfolio use. Credentials load
# from a .env file via python-dotenv.
# See README.md for required environment variables.
# ──────────────────────────────────────────────

import os
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyotp

# Load .env from the parent directory
env_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '.env'))
load_dotenv(dotenv_path=env_path)

# Credentials — set these in your .env file
LONG_EMAIL        = os.getenv("LONG_EMAIL")
NETSUITE_PASSWORD = os.getenv("NETSUITE_PASSWORD")
NETSUITE_KEY      = os.getenv("NETSUITE_KEY")

# NetSuite 2FA field IDs can vary by account configuration.
# Update these in your .env if your instance uses different element IDs.
NETSUITE_2FA_INPUT_ID  = os.getenv("NETSUITE_2FA_INPUT_ID",  "uif60_input")
NETSUITE_2FA_SUBMIT_ID = os.getenv("NETSUITE_2FA_SUBMIT_ID", "uif76")


def netsuite_login(driver, log_message):
    """
    Performs login to NetSuite with email, password, and TOTP 2FA.

    Args:
        driver:      Selenium WebDriver instance
        log_message: Callable that accepts a string — used for timestamped logging

    Raises:
        ValueError:  If required environment variables are missing
        Exception:   If login fails at any step; saves a screenshot on failure
    """
    if not all([LONG_EMAIL, NETSUITE_PASSWORD, NETSUITE_KEY]):
        log_message("Missing required environment variables (LONG_EMAIL, NETSUITE_PASSWORD, NETSUITE_KEY)")
        raise ValueError("Required credentials not set in .env")

    log_message("Starting NetSuite login")

    try:
        # ── Email ──────────────────────────────────────────────────────────
        log_message("Waiting for email field...")
        email_field = WebDriverWait(driver, 30).until(
            EC.visibility_of_element_located((By.ID, "email"))
        )
        email_field.clear()
        email_field.send_keys(LONG_EMAIL)
        log_message("Entered NetSuite email")

        # ── Password ───────────────────────────────────────────────────────
        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(NETSUITE_PASSWORD)
        log_message("Entered NetSuite password")

        # ── Submit ─────────────────────────────────────────────────────────
        driver.find_element(By.ID, "login-submit").click()
        log_message("Clicked NetSuite login button")

        # ── TOTP 2FA ───────────────────────────────────────────────────────
        log_message("Waiting for 2FA input field...")
        twofa_field = WebDriverWait(driver, 25).until(
            EC.visibility_of_element_located((By.ID, NETSUITE_2FA_INPUT_ID))
        )
        totp = pyotp.TOTP(NETSUITE_KEY)
        code = totp.now()
        twofa_field.clear()
        twofa_field.send_keys(code)
        log_message(f"Entered NetSuite 2FA code: {code}")

        # ── Submit 2FA ─────────────────────────────────────────────────────
        driver.find_element(By.ID, NETSUITE_2FA_SUBMIT_ID).click()
        log_message("Clicked 2FA submit button")

        # ── Confirm redirect away from login page ──────────────────────────
        time.sleep(4)
        WebDriverWait(driver, 20).until(
            lambda d: "app/login" not in d.current_url
        )
        log_message("Login appears successful (redirected away from login page)")

    except Exception as e:
        log_message(f"Error during NetSuite login: {e}")
        driver.save_screenshot("netsuite_login_error.png")
        raise
