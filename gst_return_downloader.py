"""
GST Return Downloader (GSTR-1 / GSTR-3B)
=========================================

WHAT THIS SCRIPT DOES
----------------------
1. Reads Username / Password from an Excel file you provide.
2. Opens Chrome and goes straight to the GST portal login page, filling in
   your Username and Password automatically.
3. Pauses ONLY for the CAPTCHA (and OTP, if the portal asks for one) —
   you solve that in the browser window and click Login yourself. The
   script detects the successful login on its own and carries on, so
   there is nothing to press back here.
   Automatically solving CAPTCHAs is unreliable and against portal terms,
   so this is the one manual step left.
4. Everything else is asked BEFORE the browser opens: one inclusive month
   range (e.g. "April 2022 to July 2024") and whether you want GSTR-1,
   GSTR-3B, or both. The correct Financial Year and Quarter are calculated
   automatically for every month in the range.
5. Downloads the selected return(s) as PDF into a per-company folder
   created next to this script file:
       <Company Name>/gstr 1/GSTR1_of_April_2025.pdf
       <Company Name>/gstr 3b/GSTR3B_of_April_2025.pdf
   A return that was already downloaded is skipped, so re-running after
   a partial failure only fetches what is still missing.

BEFORE YOU RUN THIS
--------------------
1. Install Python 3.9+.
2. Install the required libraries:
       pip install selenium openpyxl webdriver-manager
3. Install Google Chrome.
4. Prepare an Excel file with these column headers in row 1:
       Username | Password | Company Name   (Company Name is optional)
   One row per GST login. Header matching ignores case and stray spaces.
5. Put this script in the folder where you want the downloaded PDFs to
   end up. The PDFs always land next to this script FILE, no matter which
   directory you launch it from.

IMPORTANT NOTES / LIMITATIONS
------------------------------
- The GST portal's HTML structure changes periodically. Element locators
  below (search for "SELECTOR:" comments) match the portal's commonly
  documented layout, but if the portal has changed, a step may fail with
  a "could not find element" error. If that happens, right-click the
  failing field/button on the actual portal page -> Inspect -> copy its
  id/name -> tell me, and I'll patch that selector.
- SECURITY: the Excel file holds plain-text GST passwords. Keep it OUT
  of any cloud-synced folder (OneDrive, Dropbox, Google Drive) — files
  put there are uploaded to the provider and stay recoverable from
  version history even after you delete them. This script warns you at
  startup if it detects that situation. Delete the file when you're done.
- SECURITY: when a portal step fails, a screenshot and the page's HTML
  are written to a "_debug" folder to help diagnose it. Those files can
  contain your GSTIN, taxpayer details and session identifiers, so treat
  them as sensitive and delete them once the problem is sorted.
"""

import os
import shutil
import time
import glob
import traceback
import zipfile
import re
from datetime import datetime

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

try:
    from webdriver_manager.chrome import ChromeDriverManager
    HAVE_WDM = True
except ImportError:
    HAVE_WDM = False

# ---------------------------------------------------------------------------
# CONSTANTS
# ---------------------------------------------------------------------------
GST_LOGIN_URL = "https://services.gst.gov.in/services/login"
COL_USERNAME = "Username"
COL_PASSWORD = "Password"

COL_COMPANY = "Company Name"

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Text the portal shows when the username/password itself was rejected.
# Spotting any of these lets us stop immediately instead of sitting out
# the full CAPTCHA wait for a login that can never succeed.
INVALID_LOGIN_MARKERS = (
    "invalid username or password",
    "enter valid username",
    "your username or password is incorrect",
    "invalid credentials",
)

# Folders that sync to a third-party cloud. Credentials sitting in one of
# these get uploaded off this machine, so we warn about it at startup.
SYNCED_FOLDER_MARKERS = ("onedrive", "dropbox", "google drive", "googledrive", "icloud")

# Windows refuses these as file/folder names regardless of extension.
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


# ---------------------------------------------------------------------------
# Helper: resolve the Excel path the user typed, falling back to the
# script's own folder (so a bare filename, or even an empty answer, works).
# ---------------------------------------------------------------------------
def find_excel_path(user_input, script_dir):
    user_input = user_input.strip('"').strip()

    # If they typed a full/relative path and it exists as-is, use it.
    if user_input and os.path.isfile(user_input):
        return user_input

    # If they typed just a filename, look for it next to the script.
    if user_input:
        candidate = os.path.join(script_dir, user_input)
        if os.path.isfile(candidate):
            return candidate

    # Nothing typed, or still not found: auto-detect .xlsx file(s) in the
    # script's own folder.
    xlsx_files = [
        f for f in glob.glob(os.path.join(script_dir, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")  # ignore Excel lock files
    ]
    if len(xlsx_files) == 1:
        print(f"Using Excel file found in this script's folder: {xlsx_files[0]}")
        return xlsx_files[0]
    elif len(xlsx_files) > 1:
        print("\nMultiple Excel files found in this script's folder:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {os.path.basename(f)}")
        while True:
            choice = input("Enter the number of the file to use: ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(xlsx_files):
                return xlsx_files[int(choice) - 1]
            print("Invalid choice, try again.")

    # Give up — return what they typed (possibly empty) so the caller's
    # FileNotFoundError fires with a clear message.
    return user_input


# ---------------------------------------------------------------------------
# Helper: if a portal step fails, save a screenshot + the page's HTML so the
# exact problem can be diagnosed and fixed from those files, instead of a
# bare traceback.
# ---------------------------------------------------------------------------
def save_debug_snapshot(driver, output_dir, step_name):
    debug_dir = os.path.join(output_dir, "_debug")
    os.makedirs(debug_dir, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(debug_dir, f"{step_name}_{stamp}.png")
    html_path = os.path.join(debug_dir, f"{step_name}_{stamp}.html")
    try:
        driver.save_screenshot(screenshot_path)
    except Exception:
        screenshot_path = None
    try:
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        html_path = None
    return screenshot_path, html_path


class ReturnNotAvailable(Exception):
    """The requested return genuinely does not exist for this period.

    Raised for a period the portal simply does not offer -- e.g. GSTR-1
    for the first two months of a quarter under QRMP, where only the
    quarter-end month has one. This is a normal outcome, not a flake, so
    the caller records it and moves on INSTEAD of burning two more retries
    (each with its own 30s waits) on something that can never succeed.
    """


def run_step(driver, output_dir, step_name, func, *args, save_snapshot=True, **kwargs):
    """Run a portal-interaction function; on failure, save a debug snapshot
    and re-raise with a clear message pointing at the saved files.

    save_snapshot=False skips writing the screenshot/HTML -- used for
    non-final retry attempts, which would otherwise leave up to nine
    snapshots per failed period sitting in the output folder.
    """
    try:
        return func(driver, *args, **kwargs)
    except ReturnNotAvailable:
        # Expected absence, not a malfunction -- nothing to diagnose.
        raise
    except Exception as e:
        if not save_snapshot:
            raise RuntimeError(
                f"{step_name}: {type(e).__name__}: {e}"
            ) from e
        screenshot_path, html_path = save_debug_snapshot(driver, output_dir, step_name)
        msg = (
            f"\nSomething went wrong during: {step_name}\n"
            f"Original error: {type(e).__name__}: {e}\n"
        )
        if screenshot_path:
            msg += f"A screenshot of the browser at the moment of failure was saved to:\n  {screenshot_path}\n"
        if html_path:
            msg += f"The page's HTML at that moment was saved to:\n  {html_path}\n"
        msg += (
            "Send me those two files (or just the screenshot) and I can pinpoint the "
            "exact selector that needs to change."
        )
        raise RuntimeError(msg) from e


# ---------------------------------------------------------------------------
# STEP 1: Read credentials from Excel
# ---------------------------------------------------------------------------
def _cell_str(value):
    """Turn an Excel cell into the string the portal actually expects.

    Excel stores an all-digit password as a NUMBER, so openpyxl hands it
    back as a float: a password of 123456 arrives as 123456.0, and str()
    would turn that into "123456.0", which then fails to log in with no
    visible explanation. Collapse whole-number floats back to plain digits
    before using them.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def warn_if_cloud_synced(excel_path):
    """Credentials in a synced folder get uploaded to the provider and
    linger in version history even after deletion, so say so plainly."""
    lowered = os.path.abspath(excel_path).lower()
    for marker in SYNCED_FOLDER_MARKERS:
        if marker in lowered:
            print("\n" + "!" * 70)
            print("SECURITY WARNING")
            print(f"Your credentials file sits inside a cloud-synced folder ({marker}).")
            print("Plain-text GST passwords stored there are uploaded off this machine and")
            print("stay recoverable from version history even after the file is deleted.")
            print("Recommended: move it to a local-only folder, and change any password")
            print("that has already been kept there.")
            print("!" * 70 + "\n")
            return True
    return False


def read_credentials_from_excel(excel_path):
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(
            f"No file found at:\n  {excel_path}\n"
            f"Double-check the path (right-click the file itself in File Explorer "
            f"-> Copy as path)."
        )

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except openpyxl.utils.exceptions.InvalidFileException:
        raise ValueError(
            "\nThis file is not a valid .xlsx file, even if it's named .xlsx.\n"
            "FIX: Open it in Excel -> File > Save As -> choose 'Excel Workbook (*.xlsx)'\n"
            "as the type, save as a new file, and point this script at that file.\n"
        )
    except zipfile.BadZipFile:
        raise ValueError(
            "\nThe file appears to be corrupted or isn't really an Excel file.\n"
            "Try re-saving it from Excel as 'Excel Workbook (*.xlsx)' and try again.\n"
        )

    try:
        ws = wb.active
        headers = [cell.value for cell in ws[1]]

        # Match headers case-insensitively and ignore stray spaces, so
        # "username " or "PASSWORD" work as well as the exact wording.
        normalized = {}
        for i, h in enumerate(headers):
            if h is None:
                continue
            key = str(h).strip().lower()
            normalized.setdefault(key, i)

        user_idx = normalized.get(COL_USERNAME.lower())
        pass_idx = normalized.get(COL_PASSWORD.lower())
        company_idx = normalized.get(COL_COMPANY.lower())
        if user_idx is None or pass_idx is None:
            raise ValueError(
                f"Excel file must have column headers '{COL_USERNAME}' and '{COL_PASSWORD}' "
                f"in row 1. Found: {headers}"
            )

        entries = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            username = _cell_str(row[user_idx]) if user_idx < len(row) else ""
            password = _cell_str(row[pass_idx]) if pass_idx < len(row) else ""
            if not username or not password:
                continue
            company = ""
            if company_idx is not None and company_idx < len(row):
                company = _cell_str(row[company_idx])
            entries.append({
                "username": username,
                "password": password,
                "company": company,
            })
    finally:
        # Release the handle on the credentials workbook rather than
        # leaving it open for the rest of the run.
        try:
            wb.close()
        except Exception:
            pass

    return entries


def choose_account(entries):
    if len(entries) == 1:
        return entries[0]

    print("\nMultiple GST accounts found in the Excel file:")
    for i, e in enumerate(entries, 1):
        label = f"  {i}. Username: {e['username']}"
        if e.get("company"):
            label += f"  ({e['company']})"
        print(label)
    while True:
        choice = input("Enter the number of the account to use: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(entries):
            return entries[int(choice) - 1]
        print("Invalid choice, try again.")


# ---------------------------------------------------------------------------
# STEP 2: Set up Chrome — downloads land in a temp subfolder of the CURRENT
#          working directory (the folder this script is run from), and get
#          moved/renamed into that same folder once complete.
# ---------------------------------------------------------------------------
def create_driver(download_dir):
    os.makedirs(download_dir, exist_ok=True)
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        # directory_upgrade makes Chrome honour the directory above even if
        # a previous profile had a different one saved.
        "download.directory_upgrade": True,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        # Recent Chrome versions silently block some downloads as
        # "insecure"; keeping Safe Browsing on with these two off is what
        # lets a legitimate portal PDF through without a prompt.
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")

    try:
        if HAVE_WDM:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
        else:
            driver = webdriver.Chrome(options=options)
    except Exception as e:
        raise RuntimeError(
            "\nCould not start Chrome.\n"
            f"Original error: {type(e).__name__}: {e}\n\n"
            "Common causes:\n"
            "  - Google Chrome isn't installed, or was updated to a version the\n"
            "    installed chromedriver doesn't match.\n"
            "  - No internet connection, so the driver could not be downloaded.\n"
            "FIX: make sure Chrome is installed and you're online, then run:\n"
            "       pip install --upgrade webdriver-manager selenium\n"
        ) from e

    # Don't let a hung portal page block forever with no way out.
    try:
        driver.set_page_load_timeout(120)
    except Exception:
        pass

    return driver


# ---------------------------------------------------------------------------
# STEP 3: Login — fills username/password immediately, pauses only for
#          CAPTCHA/OTP.
# ---------------------------------------------------------------------------
def _visible_login_error(driver):
    """Return the text of a VISIBLE credential-rejection banner, else None.

    Deliberately checks rendered elements rather than driver.page_source:
    this is an Angular page, and its error containers exist in the DOM
    (hidden) before any error has actually happened. Matching against raw
    page source would therefore report a rejected password the instant the
    login page loaded, every single time.
    """
    selectors = (
        ".alert-danger", ".alert-error", ".err", ".error",
        "[class*='err']", "[role='alert']",
    )
    for sel in selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, sel)
        except Exception:
            continue
        for el in elements:
            try:
                if not el.is_displayed():
                    continue
                text = (el.text or "").strip()
            except Exception:
                continue
            if not text:
                continue
            lowered = text.lower()
            for marker in INVALID_LOGIN_MARKERS:
                if marker in lowered:
                    return text
    return None


def login(driver, username, password):
    wait = WebDriverWait(driver, 30)

    # The window may have been minimized earlier in the run (see
    # process_account). A CAPTCHA you can't see is a CAPTCHA you can't
    # solve, so always bring the window back before asking for one --
    # this also covers the mid-run re-login after a session expiry.
    try:
        driver.maximize_window()
    except Exception:
        pass

    driver.get(GST_LOGIN_URL)

    # SELECTOR: username field
    user_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
    user_field.clear()
    user_field.send_keys(username)

    # SELECTOR: password field
    pass_field = driver.find_element(By.ID, "user_pass")
    pass_field.clear()
    pass_field.send_keys(password)

    print("\a")  # audible bell, in case you've switched to other work
    print("\n" + "=" * 70)
    print(f"Username and password are filled in for: {username}")
    print("ACTION NEEDED: Just type the CAPTCHA shown in the Chrome window and")
    print("click Login yourself (and complete OTP too, if the portal asks for it).")
    print("The script will detect the login automatically and continue —")
    print("no need to come back and press Enter.")
    print("=" * 70)

    max_wait_seconds = 300  # 5 minutes to solve CAPTCHA/OTP
    start_time = time.time()
    last_reminder = 0
    while time.time() - start_time < max_wait_seconds:
        try:
            current_url = driver.current_url.lower()
        except Exception:
            # Browser window closed by the user mid-login.
            raise RuntimeError(
                "The Chrome window was closed before the login completed."
            )

        # A rejected username/password leaves us sitting on the login page,
        # which is indistinguishable from "still typing the CAPTCHA" by URL
        # alone. Without this check the script waits out the full 5 minutes
        # and then blames its own detection logic.
        rejected = _visible_login_error(driver)
        if rejected:
            raise RuntimeError(
                f"\nThe GST portal rejected the sign-in for username '{username}'.\n"
                f"The portal said: \"{rejected}\"\n\n"
                "FIX: check that row of your Excel file. Watch out for:\n"
                "  - a password that is all digits (Excel may have mangled it),\n"
                "  - stray leading/trailing spaces,\n"
                "  - a password changed on the portal but not updated in Excel.\n"
            )

        if "login" not in current_url:
            print("Login detected — waiting for the dashboard to fully load...")
            # Don't trust the URL change alone: wait for an element that only
            # appears once the session is genuinely authenticated, so we
            # don't race ahead and hit "Access Denied" on the next navigation.
            try:
                WebDriverWait(driver, 45).until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//*[contains(translate(.,"
                        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                        "'welcome') or contains(@title,'Return Dashboard')]"
                    ))
                )
            except TimeoutException:
                pass  # the overlay/readiness checks in navigation remain the fallback
            wait_for_no_overlay(driver, timeout=10)
            return
        elapsed = time.time() - start_time
        if elapsed - last_reminder >= 30:
            print(f"(Still waiting for login... {int(elapsed)}s elapsed)")
            last_reminder = elapsed
        time.sleep(1)

    raise RuntimeError(
        "Waited 5 minutes but never detected a successful login. "
        "If you were logged in and this still triggered, the portal's post-login "
        "URL may not have changed the way this script expects — let me know and "
        "I'll adjust the detection logic."
    )


# ---------------------------------------------------------------------------
# STEP 4: Period input — one inclusive start/end range. FY and Quarter are
#          derived automatically for each month.
# ---------------------------------------------------------------------------


MONTH_TOKEN_TO_NUMBER = {}
for _month_number, _month_name in enumerate(MONTH_NAMES, 1):
    MONTH_TOKEN_TO_NUMBER[_month_name.lower()] = _month_number
    MONTH_TOKEN_TO_NUMBER[_month_name[:3].lower()] = _month_number
    MONTH_TOKEN_TO_NUMBER[str(_month_number)] = _month_number
    MONTH_TOKEN_TO_NUMBER[f"{_month_number:02d}"] = _month_number

_MONTH_TOKEN_PATTERN = (
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
    r"nov(?:ember)?|dec(?:ember)?|0?[1-9]|1[0-2])"
)
_MONTH_YEAR_PATTERN = rf"({_MONTH_TOKEN_PATTERN})[\s./\\-]+(\d{{4}})"
_RANGE_PATTERN = re.compile(
    rf"^\s*{_MONTH_YEAR_PATTERN}\s*(?:to|through|until|[-–—])\s*"
    rf"{_MONTH_YEAR_PATTERN}\s*$",
    re.IGNORECASE,
)
_SINGLE_PERIOD_PATTERN = re.compile(
    rf"^\s*{_MONTH_YEAR_PATTERN}\s*$",
    re.IGNORECASE,
)


def financial_year_for_month(month_number, calendar_year):
    """Return the GST FY label for a calendar month, e.g. Jan 2024 -> 2023-24."""
    fy_start = calendar_year if month_number >= 4 else calendar_year - 1
    return f"{fy_start}-{(fy_start + 1) % 100:02d}"


def parse_period_range(value):
    """Parse one month or an inclusive month range.

    Accepted examples include:
      April 2022 to July 2024
      Apr-2022 - Jul-2024
      04/2022 to 07/2024
      04\\2022 to 07\\2024

    Returns chronological tuples of (financial_year, month_name, year).
    """
    value = value.strip()
    match = _RANGE_PATTERN.fullmatch(value)
    if match:
        start_month_token, start_year, end_month_token, end_year = match.groups()
    else:
        match = _SINGLE_PERIOD_PATTERN.fullmatch(value)
        if not match:
            raise ValueError(
                "Enter a month range like 'April 2022 to July 2024'. "
                "Month numbers also work, for example '04/2022 to 07/2024'."
            )
        start_month_token, start_year = match.groups()
        end_month_token, end_year = start_month_token, start_year

    start_month = MONTH_TOKEN_TO_NUMBER[start_month_token.lower()]
    end_month = MONTH_TOKEN_TO_NUMBER[end_month_token.lower()]
    start_year = int(start_year)
    end_year = int(end_year)
    start_index = start_year * 12 + start_month - 1
    end_index = end_year * 12 + end_month - 1

    # GST was introduced in July 2017, so no GST return can predate it.
    gst_start_index = 2017 * 12 + 7 - 1
    if start_index < gst_start_index:
        raise ValueError("The start month cannot be earlier than July 2017 (GST introduction).")
    if end_index < start_index:
        raise ValueError("The ending month must be the same as or later than the starting month.")

    periods = []
    for month_index in range(start_index, end_index + 1):
        calendar_year, zero_based_month = divmod(month_index, 12)
        month_number = zero_based_month + 1
        periods.append((
            financial_year_for_month(month_number, calendar_year),
            MONTH_NAMES[month_number - 1],
            calendar_year,
        ))
    return periods


def get_periods_from_user():
    """Ask once for the complete range; never ask separately for FY/Quarter."""
    while True:
        period_input = input(
            "\nEnter the return period (e.g. April 2022 to July 2024): "
        ).strip()
        try:
            periods = parse_period_range(period_input)
            break
        except ValueError as e:
            print(f"Invalid period: {e}")

    _, first_month, first_year = periods[0]
    _, last_month, last_year = periods[-1]
    financial_years = list(dict.fromkeys(fy for fy, _, _ in periods))
    print(
        f"-> Inclusive range: {first_month} {first_year} to {last_month} {last_year} "
        f"({len(periods)} month{'s' if len(periods) != 1 else ''})"
    )
    print(f"-> Financial Year(s) detected automatically: {', '.join(financial_years)}")
    print("-> Quarter will be detected automatically for each month.")
    return periods


# ---------------------------------------------------------------------------
# STEP 5: Navigate to the File Returns page and select FY / Quarter / Period
#          — selectors below are confirmed from an actual saved page:
#            FY dropdown:      <select name="fin">
#            Quarter dropdown: <select name="quarter">  (only present for
#                               quarterly/QRMP filers — selecting it re-
#                               populates the Period dropdown)
#            Period dropdown:  <select name="mon">
#            Search button:    <button class="btn btn-primary srchbtn">
# ---------------------------------------------------------------------------
MONTH_TO_QUARTER = {
    "April": 1, "May": 1, "June": 1,
    "July": 2, "August": 2, "September": 2,
    "October": 3, "November": 3, "December": 3,
    "January": 4, "February": 4, "March": 4,
}


def wait_for_no_overlay(driver, timeout=15):
    """Wait for any loading/dimmer overlay to finish clearing.

    This portal shows a semi-transparent '.dimmer-holder' overlay while an
    AJAX call is in flight (e.g. right after clicking VIEW, or after a
    search). Clicking a button underneath it too early -- even while it's
    already fading out (mid-opacity-transition) -- gets intercepted by the
    overlay instead of reaching the real button. This is what caused
    'ElementClickInterceptedException ... dimmer-holder' failures.
    """
    end_time = time.time() + timeout
    while time.time() < end_time:
        blocking = False
        try:
            overlays = driver.find_elements(
                By.CSS_SELECTOR, ".dimmer-holder, .loading-overlay, .blockUI, .dimmer.active"
            )
            for ov in overlays:
                try:
                    if not ov.is_displayed():
                        continue
                    style = (ov.get_attribute("style") or "").replace(" ", "").lower()
                    if "display:none" in style:
                        continue
                    blocking = True
                    break
                except Exception:
                    continue
        except Exception:
            return
        if not blocking:
            return
        time.sleep(0.1)


def safe_click(
    driver, element=None, timeout=15, locator=None, scroll_into_view=False
):
    """Click an element, waiting out any loading overlay first, and
    re-finding it when Angular replaces the DOM node between detection and
    click. ``locator`` should be supplied for portal controls that can be
    re-rendered; older call sites may still pass a stable element directly.
    """
    end_time = time.time() + timeout
    last_error = None
    while time.time() < end_time:
        wait_for_no_overlay(driver, timeout=max(0.1, end_time - time.time()))
        if locator is not None:
            try:
                element = WebDriverWait(
                    driver,
                    max(0.1, end_time - time.time()),
                    poll_frequency=0.1,
                    ignored_exceptions=(StaleElementReferenceException,),
                ).until(EC.element_to_be_clickable(locator))
            except TimeoutException as e:
                last_error = e
                break
        if scroll_into_view:
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", element
                )
            except StaleElementReferenceException as e:
                last_error = e
                if locator is not None:
                    continue
                raise
        try:
            element.click()
            return
        except StaleElementReferenceException as e:
            last_error = e
            if locator is not None:
                continue
            raise
        except Exception as e:
            last_error = e

        # A late notification can appear in the tiny interval between the
        # readiness check and the click. Clear it before the immediate retry.
        close_any_notification(driver, timeout=0)
        wait_for_no_overlay(driver, timeout=min(5, max(0.1, end_time - time.time())))
        if locator is not None:
            continue
        try:
            driver.execute_script("arguments[0].click();", element)
            return
        except Exception as e:
            last_error = e
            break

    if last_error is not None:
        raise last_error
    raise TimeoutException("The element could not be clicked before the timeout.")


def find_first_visible(driver, locators, timeout=10, return_locator=False):
    """Return the first element matching any locator that is actually
    VISIBLE and enabled, or None if the timeout expires.

    EC.presence_of_element_located() returns the first DOM match whether or
    not it is rendered, and this portal keeps several hidden duplicates of
    its nav links (one set for PAN-based login, alternate menus, etc).
    Grabbing a hidden duplicate makes hover/click silently do nothing, so
    every menu lookup goes through here instead.
    """
    end_time = time.time() + timeout
    while True:
        for by, value in locators:
            try:
                candidates = driver.find_elements(by, value)
            except Exception:
                continue
            for el in candidates:
                try:
                    if el.is_displayed() and el.is_enabled():
                        return (el, (by, value)) if return_locator else el
                except Exception:
                    continue
        if time.time() >= end_time:
            return None
        time.sleep(0.1)


def wait_for_file_returns_page(driver, timeout=30):
    """Wait for the File Returns controls instead of sleeping blindly."""
    wait_for_no_overlay(driver, timeout=min(timeout, 15))
    WebDriverWait(driver, timeout, poll_frequency=0.1).until(
        EC.presence_of_element_located((By.NAME, "fin"))
    )


def open_file_returns_page(driver, username, password, already_on_dashboard=False):
    """Get to the Dashboard, then click Return Dashboard.

    IMPORTANT: we deliberately never use driver.get() to jump straight to the
    dashboard URL. Loading a URL directly (like pasting it into the address
    bar) sends no Referer header, and this portal's dashboard route rejects
    that as "Access Denied" even with a fully valid session. Clicking the
    real in-page "Dashboard" nav link, like a person would, sends a proper
    Referer and works reliably instead.

    already_on_dashboard=True skips navigation entirely for the very first
    period, since right after login we're already sitting on the dashboard.
    """
    wait = WebDriverWait(driver, 30)

    if not already_on_dashboard:
        # SELECTOR: top-nav "Dashboard" link (confirmed from saved HTML) —
        # there are two similarly-named links on the page (one for PAN-based
        # login, hidden), so we specifically target the visible one pointing
        # at services.gst.gov.in.
        dash_locator = (
            By.XPATH,
            "//a[contains(@href,'services.gst.gov.in/services/auth/dashboard')"
            " and not(contains(@class,'ng-hide'))]"
        )
        dash_nav = wait.until(EC.element_to_be_clickable(dash_locator))
        safe_click(driver, dash_nav, locator=dash_locator)

    page_lower = driver.page_source.lower()
    if "access denied" in page_lower or "session is expired" in page_lower:
        # The window was very likely minimized by process_account at this
        # point, so ring the bell and say so loudly -- login() restores the
        # window itself before asking for the CAPTCHA.
        print("\a")
        print("\n" + "!" * 70)
        print("Your GST portal session has expired and needs a fresh login.")
        print("The Chrome window has been restored for you —")
        print("please solve the CAPTCHA there to continue.")
        print("!" * 70)
        login(driver, username, password)
        # login() lands us back on the dashboard via the server's own
        # redirect (not a direct URL load), so no further navigation needed.

    navigate_to_return_dashboard(driver)


def navigate_to_return_dashboard(driver):
    """From the main GST Dashboard (the 'Ledger Balance' page you land on
    right after login), get onto the Returns / File Returns page.

    Two routes are tried, in order:
    1. The dashboard's direct 'File Returns' button/link, if present.
    2. The standard top-nav route on this portal: Services -> Returns ->
       Returns Dashboard. This is the route needed for the plain "Ledger
       Balance" Dashboard screen, which has no direct button of its own —
       this was the reason the script previously got stuck on that page.
    """
    lower = lower_text_xpath()

    # If a previous attempt on this page left a popup/modal open (e.g. the
    # GSTR-3B summary popup, whose backdrop blocks clicks anywhere else on
    # the page), clear it first -- otherwise every click below silently
    # fails and this function times out for a reason that has nothing to
    # do with the Services menu itself.
    # Only clear a popup that is already visible. Waiting here for a popup
    # that may never appear used to add three seconds to every return.
    close_any_notification(driver, timeout=0)

    # --- Route 1: direct link/tile, if present ---
    direct_locators = [
        # Confirmed in the supplied debug HTML. The visible dashboard button
        # currently says "File Returns" and points to this returns host.
        (By.XPATH, "//a[contains(@href,'return.gst.gov.in/returns/auth/dashboard')" 
                   " and .//*[contains(normalize-space(.),'File Returns')]]"),
        (By.XPATH, f"//a[contains({lower},'file returns')]"),
        (By.XPATH, "//span[@title='Return Dashboard']"),
        (By.XPATH, f"//button[contains({lower},'return dashboard')]"),
        (By.XPATH, f"//a[contains({lower},'return dashboard')]"),
    ]
    direct_match = find_first_visible(
        driver, direct_locators, timeout=0.5, return_locator=True
    )
    if direct_match is not None:
        direct_link, direct_locator = direct_match
        safe_click(
            driver, direct_link, locator=direct_locator, scroll_into_view=True
        )
        wait_for_file_returns_page(driver)
        return

    # --- Route 2: Services -> Returns (tab) -> Returns Dashboard ---
    # This portal's "Services" nav item opens a mega-menu with a row of
    # TABS (Registration, Ledgers, Returns, Payments, ...). Hovering the
    # "Returns" tab reveals its own column of links, one of which is
    # "Returns Dashboard". This needs continuous hovering (Services, then
    # the Returns tab) rather than clicks, or the mega-menu collapses.
    exact = lambda t: f"normalize-space({lower_text_xpath()})='{t.lower()}'"

    services_menu = find_first_visible(
        driver, [(By.XPATH, f"//a[{exact('Services')}]")], timeout=15
    )
    if services_menu is None:
        raise RuntimeError(
            "Could not find a visible 'Services' menu on the Dashboard page. "
            "If your account's menu is worded differently, send me a "
            "screenshot and I'll match the exact text."
        )

    # The CSS opens this menu when its parent <li class="dropdown"> is
    # hovered, not reliably when only the child <a> is targeted (as the
    # supplied debug HTML confirms).
    try:
        services_hover_target = services_menu.find_element(
            By.XPATH, "./ancestor::li[contains(@class,'dropdown')][1]"
        )
    except Exception:
        services_hover_target = services_menu
    ActionChains(driver).move_to_element(services_hover_target).perform()
    returns_tab = find_first_visible(driver, [
        (By.XPATH, f"//a[{exact('Returns')}]"),
        (By.XPATH, f"//li[{exact('Returns')}]"),
    ], timeout=10)
    if returns_tab is None:
        # Hover menus are unreliable in minimized Chrome windows. The same
        # destination is also present as a real anchor in the DOM; clicking
        # that element still follows the portal's link and supplies a proper
        # Referer (unlike driver.get(), which this portal rejects).
        try:
            returns_dashboard_link = driver.find_element(
                By.XPATH,
                "//a[contains(@href,'return.gst.gov.in/returns/auth/dashboard')]",
            )
            driver.execute_script("arguments[0].click();", returns_dashboard_link)
            wait_for_file_returns_page(driver)
            return
        except Exception as e:
            raise RuntimeError(
                "Found the 'Services' menu but could not open the GST "
                "Returns Dashboard through either the visible File Returns "
                "button or the portal's Returns link."
            ) from e

    ActionChains(driver).move_to_element(returns_tab).perform()

    # "Returns Dashboard" should now be visible in the revealed column.
    returns_dashboard_link = find_first_visible(driver, [
        (By.XPATH, f"//a[{exact('Returns Dashboard')}]"),
        (By.XPATH, f"//*[{exact('Returns Dashboard')}]"),
    ], timeout=8)

    if returns_dashboard_link is None:
        raise RuntimeError(
            "Could not find 'Returns Dashboard' after hovering Services > "
            "Returns on the Dashboard page. If the menu wording is "
            "different on your account, send me a screenshot and I'll "
            "match the exact text."
        )

    # Click through the open menu; fall back to a plain/JS click if needed.
    try:
        ActionChains(driver).move_to_element(returns_dashboard_link) \
            .click(returns_dashboard_link).perform()
    except Exception:
        try:
            returns_dashboard_link.click()
        except Exception:
            driver.execute_script("arguments[0].click();", returns_dashboard_link)
    wait_for_file_returns_page(driver)


def _dropdown_option_texts(driver, name):
    """Return the current option texts of a <select name="..."> if it
    exists right now, else None (element not present)."""
    try:
        el = driver.find_element(By.NAME, name)
        return [o.text for o in Select(el).options]
    except (NoSuchElementException, StaleElementReferenceException):
        return None


def force_select_and_fire_change(
    driver, select_element, visible_text, name=None, timeout=10
):
    """Select an option AND explicitly dispatch a 'change' event via JS.

    Selenium's Select.select_by_visible_text() only fires the browser's
    native 'change' event when the clicked option differs from whatever
    was already selected. This portal's Quarter/Period dropdowns are
    CASCADING — picking FY re-populates Quarter, picking Quarter
    re-populates Period — driven entirely by that 'change' event.

    When a parent really does need to be applied, explicitly firing the
    event also handles the case where its visible value happens to match the
    existing selection. Callers skip this function entirely when both the
    parent selection and its dependent dropdown are already ready.
    """
    end_time = time.time() + timeout
    while True:
        try:
            Select(select_element).select_by_visible_text(visible_text)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                select_element,
            )
            return
        except StaleElementReferenceException:
            if name is None or time.time() >= end_time:
                raise
            select_element = WebDriverWait(
                driver,
                max(0.1, end_time - time.time()),
                poll_frequency=0.1,
                ignored_exceptions=(StaleElementReferenceException,),
            ).until(EC.presence_of_element_located((By.NAME, name)))


def wait_for_dropdown_refresh(
    driver, name, previous_option_texts, timeout=15,
    expected_option_text=None, unchanged_ready_after=1.0,
):
    """Wait until a cascading dropdown is ready after an AJAX refresh.

    The FY dropdown can reload Quarter with exactly the same labels it had
    before (Quarter 1..4), and re-selecting the same quarter can likewise
    reload Period with the same three month labels. Waiting only for the
    option text list to *change* therefore burned the full 15-second timeout
    even though the dropdown was already usable.

    A genuinely changed option list is accepted immediately. If the labels
    remain identical, accept the populated dropdown after a short settling
    window, but only when it contains the option the caller needs. This keeps
    the safety check without adding a false 15-second pause.
    """
    started_at = time.time()

    def _ready(_):
        current = _dropdown_option_texts(driver, name)
        if current is None:
            return False
        readable = [text.strip() for text in current if text and text.strip()]
        if expected_option_text is not None:
            expected = expected_option_text.strip().lower()
            if not any(expected in text.lower() for text in readable):
                return False
        if previous_option_texts is None:
            return len(readable) > 1
        if current != previous_option_texts and readable:
            return True
        return (
            len(readable) > 1
            and time.time() - started_at >= unchanged_ready_after
        )

    try:
        WebDriverWait(driver, timeout, poll_frequency=0.1).until(_ready)
    except TimeoutException:
        pass  # caller will raise its own clear error if options are still wrong


def verify_selection(driver, name, expected_text, label):
    """Read a dropdown back and confirm it really holds what we just set.

    The cascading dropdowns on this page are repopulated by AJAX, and
    wait_for_dropdown_refresh() deliberately gives up quietly on timeout.
    Without a read-back, a refresh that never landed leaves the old value
    selected and the run carries on downloading the wrong period. Checking
    here turns that silent wrong answer into a loud, retryable error.
    """
    end_time = time.time() + 5
    while True:
        try:
            element = driver.find_element(By.NAME, name)
            actual = Select(element).first_selected_option.text.strip()
            break
        except NoSuchElementException:
            return  # dropdown genuinely absent (e.g. monthly filer)
        except StaleElementReferenceException:
            if time.time() >= end_time:
                raise
            time.sleep(0.1)
    if actual.replace(" ", "").lower() != expected_text.strip().replace(" ", "").lower():
        raise RuntimeError(
            f"{label} did not take the value it was set to: asked for "
            f"'{expected_text}' but the dropdown still shows '{actual}'. "
            f"The portal's cascading dropdown most likely had not finished "
            f"reloading. This is retryable."
        )


def confirm_period_after_search(driver, financial_year, month_name):
    """Sanity-check that the search results shown are for the period asked
    for, so a PDF can never be saved under a month it does not belong to.

    The dropdowns persist on the results page and are re-rendered from the
    server's response, so reading them back here reflects what the portal
    actually searched for — not merely what we clicked.
    """
    verify_selection(driver, "fin", financial_year, "Financial Year (after search)")

    end_time = time.time() + 5
    while True:
        try:
            element = driver.find_element(By.NAME, "mon")
            actual = Select(element).first_selected_option.text.strip()
            break
        except NoSuchElementException:
            return
        except StaleElementReferenceException:
            if time.time() >= end_time:
                raise
            time.sleep(0.1)
    if month_name.lower() not in actual.lower():
        raise RuntimeError(
            f"After searching, the Period dropdown shows '{actual}' but "
            f"'{month_name}' was requested. Refusing to download, because the "
            f"file would be saved under the wrong month. This is retryable."
        )


def select_period_and_search(driver, financial_year, month_name):
    wait = WebDriverWait(driver, 30)
    quarter_num = MONTH_TO_QUARTER[month_name]

    # Snapshot what Quarter / Period looked like BEFORE we touch anything,
    # so we can reliably detect once they've actually refreshed.
    quarter_before = _dropdown_option_texts(driver, "quarter")
    mon_before = _dropdown_option_texts(driver, "mon")

    # --- Financial Year dropdown (name="fin") ---
    fy_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "fin")))
    fy_select = Select(fy_dropdown)
    year_start = financial_year.split("-")[0]
    year_end = str(int(year_start) + 1)
    fy_variants = [financial_year, f"{year_start}-{year_end}"]
    fy_matched = None
    for option in fy_select.options:
        opt_norm = option.text.strip().replace(" ", "")
        if any(v.replace(" ", "") == opt_norm for v in fy_variants):
            fy_matched = option.text
            break
    if fy_matched is None:
        raise RuntimeError(
            f"Could not find financial year '{financial_year}' in the FY dropdown. "
            f"Available options were: {[o.text for o in fy_select.options]}"
        )
    selected_fy = fy_select.first_selected_option.text.strip()
    fy_is_already_selected = (
        selected_fy.replace(" ", "").lower()
        == fy_matched.strip().replace(" ", "").lower()
    )

    # Returning to this page for the next month usually preserves the FY and
    # its populated child dropdown. In that case, forcing the same selection
    # only launches a redundant AJAX request and used to trigger a false
    # 15-second refresh timeout. Reuse the ready state instead.
    if quarter_before is not None:
        expected_child_text = f"quarter {quarter_num}"
        child_is_ready = any(
            expected_child_text in text.lower()
            for text in quarter_before if text
        )
        child_name = "quarter"
        child_before = quarter_before
    else:
        expected_child_text = month_name
        child_is_ready = any(
            month_name.lower() in text.lower()
            for text in (mon_before or []) if text
        )
        child_name = "mon"
        child_before = mon_before

    if not (fy_is_already_selected and child_is_ready):
        force_select_and_fire_change(driver, fy_dropdown, fy_matched, name="fin")
        wait_for_dropdown_refresh(
            driver,
            child_name,
            child_before,
            expected_option_text=expected_child_text,
        )

    verify_selection(driver, "fin", fy_matched, "Financial Year")

    # --- Quarter dropdown (name="quarter") — only present for quarterly
    #     filers. Selecting it re-populates the Period dropdown, so it must
    #     be set (if present) before reading Period options. Re-find the
    #     element fresh here rather than reusing an earlier reference, since
    #     the AJAX refresh may have replaced the DOM node (stale element).
    #
    #     CRITICAL: only the PRESENCE check belongs inside a
    #     NoSuchElementException handler. Selenium's
    #     Select.select_by_visible_text() ALSO raises NoSuchElementException
    #     when the option text is not found, so wrapping the whole block
    #     (as this code once did) made a genuine failure to set the quarter
    #     look identical to "this account files monthly". The period then
    #     kept its previous value and the WRONG MONTH was downloaded under
    #     the right filename — silently. ---
    try:
        quarter_dropdown = driver.find_element(By.NAME, "quarter")
    except NoSuchElementException:
        quarter_dropdown = None  # monthly filer — no Quarter dropdown to set

    if quarter_dropdown is not None:
        quarter_select = Select(quarter_dropdown)
        quarter_matched = None
        for option in quarter_select.options:
            if f"quarter {quarter_num}" in option.text.lower():
                quarter_matched = option.text
                break
        if quarter_matched is None:
            raise RuntimeError(
                f"Could not find the quarter for {month_name} in the Quarter dropdown. "
                f"Available options were: {[o.text for o in quarter_select.options]}"
            )
        mon_before_quarter = _dropdown_option_texts(driver, "mon")
        selected_quarter = quarter_select.first_selected_option.text.strip()
        quarter_is_already_selected = (
            selected_quarter.replace(" ", "").lower()
            == quarter_matched.strip().replace(" ", "").lower()
        )
        period_is_ready = any(
            month_name.lower() in text.lower()
            for text in (mon_before_quarter or []) if text
        )

        if not (quarter_is_already_selected and period_is_ready):
            force_select_and_fire_change(
                driver, quarter_dropdown, quarter_matched, name="quarter"
            )
            wait_for_dropdown_refresh(
                driver,
                "mon",
                mon_before_quarter,
                expected_option_text=month_name,
            )
        verify_selection(driver, "quarter", quarter_matched, "Quarter")

    # --- Period / month dropdown (name="mon") — re-find fresh for the same
    #     stale-element reason as above. ---
    period_dropdown = wait.until(EC.presence_of_element_located((By.NAME, "mon")))
    select_obj = Select(period_dropdown)
    available = [o.text for o in select_obj.options]
    matched = None
    for option in select_obj.options:
        if month_name.lower() in option.text.lower():
            matched = option.text
            break
    if matched is None:
        # Distinguish two very different situations that both land here:
        #
        #  a) The dropdown is populated and simply has no such month. That
        #     is a real, expected absence (the normal case for QRMP filers,
        #     where GSTR-1 exists only for the quarter-end month) — skip it
        #     without burning two more retries.
        #
        #  b) The dropdown came back empty or unreadable. That is a
        #     breakage or a refresh that never landed, NOT evidence the
        #     period doesn't exist. Treat it as retryable, or a genuine
        #     fault would be silently written off as "not available".
        readable = [t for t in available if t and t.strip()]
        if len(readable) <= 1:
            raise RuntimeError(
                f"The Period dropdown came back empty or unreadable while looking "
                f"for '{month_name}' (options seen: {available}). The cascading "
                f"refresh most likely hadn't finished. This is retryable."
            )
        raise ReturnNotAvailable(
            f"The portal offers no '{month_name}' period for this account "
            f"in FY {financial_year}. Available options were: {readable}. "
            f"(For QRMP/quarterly filers this is expected — GSTR-1 is filed "
            f"only for the last month of each quarter.)"
        )
    force_select_and_fire_change(driver, period_dropdown, matched, name="mon")
    wait_for_no_overlay(driver, timeout=10)
    verify_selection(driver, "mon", matched, "Period")

    # --- Search button — re-find fresh too, in case the button itself sits
    #     inside a panel that got re-rendered by the cascading updates. ---
    search_locator = (By.CSS_SELECTOR, "button.srchbtn")
    search_btn = wait.until(EC.element_to_be_clickable(search_locator))
    safe_click(driver, search_btn, locator=search_locator)
    wait_for_no_overlay(driver, timeout=20)

    # Confirm the results on screen really are for the period requested,
    # BEFORE anything gets downloaded and named after that period.
    confirm_period_after_search(driver, fy_matched, month_name)


# ---------------------------------------------------------------------------
# STEP 6: Download a given return
# ---------------------------------------------------------------------------
def close_extra_windows(driver, main_handle):
    """Some portal actions (e.g. downloading a PDF) can pop open a new
    browser tab/window. If that happens and it's left open, Selenium's
    'current' window can end up pointing at that stray tab instead of the
    portal, which then breaks every step for the NEXT month (this is a
    likely cause of the script silently stopping after April). Close any
    extra tabs and make sure we're back on the main portal tab before
    moving on."""
    try:
        handles = driver.window_handles
    except Exception:
        return
    if len(handles) <= 1:
        return
    for handle in handles:
        if handle != main_handle:
            try:
                driver.switch_to.window(handle)
                driver.close()
            except Exception:
                pass
    try:
        driver.switch_to.window(main_handle)
    except Exception:
        pass


INCOMPLETE_DOWNLOAD_SUFFIXES = (".crdownload", ".tmp", ".part", ".download", ".partial")


def wait_for_new_download(
    download_dir, before_files, timeout=90, stable_checks_needed=3, poll_interval=0.2
):
    """Wait for a new file to appear in download_dir and finish downloading.

    Chrome (often together with antivirus scanning on the download) writes
    a file under a temporary name first -- usually ending in .crdownload,
    but sometimes .tmp -- and only renames it to its real name once fully
    written. Just checking the extension isn't quite enough on its own
    (there's a brief window where the rename has happened but the file is
    still being flushed to disk), so we also require the file's size to
    stay the same across a couple of checks in a row before trusting it's
    actually done.
    """
    end_time = time.time() + timeout
    stable_count = 0
    last_size = None
    last_candidate = None
    while time.time() < end_time:
        current_files = set(glob.glob(os.path.join(download_dir, "*")))
        new_files = current_files - before_files
        finished = [
            f for f in new_files
            if not f.lower().endswith(INCOMPLETE_DOWNLOAD_SUFFIXES)
        ]
        if finished:
            # More than one new file can show up (a stray portal artefact,
            # or a leftover from a prior attempt that only just finished
            # writing). The newest one is the download we just triggered.
            candidate = max(finished, key=lambda p: os.path.getmtime(p))
            try:
                size = os.path.getsize(candidate)
            except OSError:
                size = None
            if candidate != last_candidate:
                # A different file than last check (e.g. it just got
                # renamed away from .crdownload) -- reset the stability
                # count and start timing this one.
                last_candidate = candidate
                stable_count = 0
                last_size = size
            elif size is not None and size > 0 and size == last_size:
                stable_count += 1
            else:
                stable_count = 0
                last_size = size
            if stable_count >= stable_checks_needed:
                return candidate
        else:
            last_candidate = None
            stable_count = 0
        time.sleep(poll_interval)
    raise TimeoutException(
        "Download did not complete within the expected time (the file "
        "stayed in an incomplete/temporary state)."
    )


def verify_downloaded_pdf(path):
    """Reject an error/HTML response before it is named and kept as a PDF."""
    try:
        size = os.path.getsize(path)
        with open(path, "rb") as downloaded_file:
            header = downloaded_file.read(1024)
    except OSError as e:
        raise RuntimeError(f"Could not read the downloaded file for verification: {e}") from e

    if size < 100 or b"%PDF-" not in header:
        raise RuntimeError(
            "The portal download completed, but its contents are not a valid PDF "
            f"(received {os.path.basename(path)}, {size} bytes). It will not be "
            "saved under a return filename; this attempt is retryable."
        )


def lower_text_xpath(node_expr="."):
    """Build the XPath snippet that lowercases an element's full text content,
    since XPath's contains() is case-sensitive and this portal's buttons are
    in ALL CAPS (e.g. 'VIEW', 'DOWNLOAD (PDF)')."""
    return (f"translate(normalize-space({node_expr}),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')")


def _find_visible_notification_close(driver):
    """Return a close control for a notification that is visible right now."""
    lower = lower_text_xpath()
    close_locators = [
        (By.CSS_SELECTOR, "button[data-ng-click*='showStatusTable']"),
        (By.XPATH, "//button[@aria-label='Close']"),
        (By.XPATH, f"//button[{lower}='close']"),
        (By.CSS_SELECTOR, "button.close"),
        (By.XPATH, f"//button[contains({lower},'close')]"),
        (By.CSS_SELECTOR, ".alert .close, .alert [class*='close']"),
    ]
    for by, value in close_locators:
        try:
            candidates = driver.find_elements(by, value)
        except Exception:
            continue
        for element in candidates:
            try:
                if element.is_displayed() and element.is_enabled():
                    return element
            except Exception:
                continue
    return None


def close_any_notification(driver, timeout=0, max_popups=3):
    """The GSTR-3B view page can show one or more notification popups on
    load — e.g. the 'System generated summary for GSTR-3B' modal, plus
    separate dismissible alert banners. These sit on top of the page and
    block the Download button, so close whatever shows up before moving on.

    IMPORTANT: this page also has many OTHER close/alert-dismiss buttons in
    the DOM that are simply not visible right now (leftover from other
    alerts that aren't currently shown). A plain
    WebDriverWait(...).until(element_to_be_clickable(locator)) only ever
    looks at the FIRST element matching a locator — if that happens to be
    one of the hidden ones, it times out without ever reaching the real,
    visible button. So instead we fetch ALL matches for each locator and
    pick the first one that's actually visible and enabled.
    """
    end_time = time.time() + timeout
    popups_closed = 0
    while popups_closed < max_popups:
        # Always check once, even when timeout=0. This makes clearing an
        # already-open popup instantaneous without waiting for a hypothetical
        # popup that may never arrive.
        el = _find_visible_notification_close(driver)
        if el is None:
            if time.time() >= end_time:
                break
            time.sleep(0.1)
            continue
        try:
            el.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", el)
            except Exception:
                break
        popups_closed += 1
        try:
            WebDriverWait(driver, 3, poll_frequency=0.1).until(
                lambda _: not el.is_displayed()
            )
        except Exception:
            pass
    return popups_closed


def wait_for_action_while_closing_notifications(driver, actions, timeout=30):
    """Wait for the next page action, closing any popup encountered en route.

    ``actions`` is an ordered list of ``(name, locator)`` pairs. Returning as
    soon as a real action is available avoids the old 25+5 second popup waits
    while still handling notifications that render late.
    """
    end_time = time.time() + timeout
    while time.time() < end_time:
        close_button = _find_visible_notification_close(driver)
        if close_button is not None:
            try:
                close_button.click()
            except Exception:
                driver.execute_script("arguments[0].click();", close_button)
            wait_for_no_overlay(driver, timeout=5)
            try:
                WebDriverWait(driver, 3, poll_frequency=0.1).until(
                    lambda _: not close_button.is_displayed()
                )
            except Exception:
                pass
            continue

        for action_name, (by, value) in actions:
            try:
                candidates = driver.find_elements(by, value)
            except Exception:
                continue
            for element in candidates:
                try:
                    if element.is_displayed() and element.is_enabled():
                        return action_name, element, (by, value)
                except Exception:
                    continue
        time.sleep(0.1)
    raise TimeoutException(
        f"None of the expected actions became available: "
        f"{', '.join(name for name, _ in actions)}"
    )


def download_return(driver, return_type, download_dir, output_dir, month_name, year, main_handle=None):
    wait = WebDriverWait(driver, 30)
    lower = lower_text_xpath()

    if return_type == "GSTR1":
        tile_exact_label = "GSTR-1"
        display_name = "GSTR1"
    else:
        tile_exact_label = "GSTR-3B"
        display_name = "GSTR3B"

    # --- Step 1: click "VIEW" inside the correct tile (exact label match,
    #     so 'GSTR-1' doesn't also match 'GSTR-1A' or similar tiles) ---
    tile_view_xpath = (
        f"//*[normalize-space(text())='{tile_exact_label}']"
        "/ancestor::div[contains(@class,'panel') or contains(@class,'tile') "
        "or contains(@class,'col-md') or contains(@class,'col-sm')][1]"
        f"//button[contains({lower},'view') and not(contains({lower},'summary'))]"
    )
    try:
        view_locator = (By.XPATH, tile_view_xpath)
        view_btn = wait.until(EC.element_to_be_clickable(view_locator))
    except TimeoutException:
        # No VIEW button on this tile. Either the return was never filed
        # for this period (normal, skip it) or the portal's markup changed
        # (a real fault worth retrying and diagnosing).
        if looks_not_filed(driver):
            raise ReturnNotAvailable(
                f"{tile_exact_label} does not appear to have been filed for "
                f"{month_name} {year}, so there is nothing to download."
            )
        raise
    safe_click(driver, view_btn, locator=view_locator)
    wait_for_no_overlay(driver, timeout=15)

    if return_type == "GSTR1":
        # --- Step 2: follow "VIEW SUMMARY" when this page has it. Waiting
        #     for SUMMARY and DOWNLOAD together avoids a pointless optional-
        #     step timeout on the direct-download layout. ---
        try:
            next_action, action_button, action_locator = wait_for_action_while_closing_notifications(
                driver,
                [
                    ("view summary", (
                        By.XPATH, f"//button[contains({lower},'view summary')]"
                    )),
                    ("download PDF", (
                        By.XPATH,
                        f"//button[contains({lower},'download') and contains({lower},'pdf')]",
                    )),
                ],
                timeout=30,
            )
            if next_action == "view summary":
                safe_click(driver, action_button, locator=action_locator)
                wait_for_no_overlay(driver, timeout=15)
                _, download_btn, download_locator = wait_for_action_while_closing_notifications(
                    driver,
                    [("download PDF", (
                        By.XPATH,
                        f"//button[contains({lower},'download') and contains({lower},'pdf')]",
                    ))],
                    timeout=30,
                )
            else:
                download_btn = action_button
                download_locator = action_locator
        except TimeoutException:
            if looks_not_filed(driver):
                raise ReturnNotAvailable(
                    f"GSTR-1 for {month_name} {year} has no filed summary to "
                    f"download (the portal reports nothing filed for this period)."
                )
            raise
    else:
        # GSTR-3B goes straight from VIEW to its own summary/payment page
        # (no separate "View Summary" click), and sometimes shows a
        # notification popup on load that has to be closed first. The
        # button here reads "DOWNLOAD FILED GSTR-3B", not "Download (PDF)".
        try:
            _, download_btn, download_locator = wait_for_action_while_closing_notifications(
                driver,
                [("download filed GSTR-3B", (
                    By.XPATH, f"//button[contains({lower},'download filed')]"
                ))],
                timeout=30,
            )
        except TimeoutException:
            if looks_not_filed(driver):
                raise ReturnNotAvailable(
                    f"GSTR-3B for {month_name} {year} has not been filed, so "
                    f"there is no filed return to download."
                )
            raise

    # Scroll the button into view first -- it's often below the fold (as in
    # the screenshot), and a click on an off-screen/partially-covered
    # element can silently fail or get intercepted by something else on
    # the page (e.g. the sticky footer or the loading dimmer).
    wait_for_no_overlay(driver, timeout=10)

    before_files = set(glob.glob(os.path.join(download_dir, "*")))
    safe_click(
        driver,
        download_btn,
        locator=download_locator,
        scroll_into_view=True,
    )

    print(f"Downloading {display_name} of {month_name} {year} ... waiting for it to finish.")
    downloaded_path = wait_for_new_download(download_dir, before_files)
    verify_downloaded_pdf(downloaded_path)

    if main_handle is not None:
        close_extra_windows(driver, main_handle)

    # The contents have been verified as a real PDF before the file receives
    # its final return name.
    final_path = return_pdf_path(return_type, output_dir, month_name, year)

    # Never clobber an existing return silently. process_account skips
    # periods that are already on disk, so reaching here with the file
    # present means something unexpected -- keep both and say so.
    if os.path.exists(final_path):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        root, ext = os.path.splitext(final_path)
        final_path = f"{root}__redownloaded_{stamp}{ext}"
        print(f"NOTE: a file for this return already existed; saving the new "
              f"copy alongside it instead of overwriting.")

    shutil.move(downloaded_path, final_path)
    print(f"Saved: {final_path}")


def detect_company_name(driver, fallback):
    """Best-effort scrape of the taxpayer's legal name from the portal, so
    downloaded returns can be saved under a folder named after the company
    -- useful once more than one company's credentials are in the Excel
    file. Confirmed from real portal HTML at two spots:
      1. The top-right account dropdown, which carries the FULL legal name
         in its title attribute (the visible text itself is truncated):
         <a class="dropdown-toggle lang-dpwn" title="FULL LEGAL NAME">
      2. A "Legal Name - ..." line shown on some pages (e.g. File Returns).
    Falls back to whatever the caller passes in (e.g. the GST username) if
    neither is found, so this never blocks the run.
    """
    try:
        el = driver.find_element(By.CSS_SELECTOR, "a.dropdown-toggle.lang-dpwn")
        title = (el.get_attribute("title") or "").strip()
        if title:
            return title
    except Exception:
        pass

    try:
        el = driver.find_element(
            By.XPATH, "//p[contains(@class,'reg') and contains(.,'Legal Name')]"
        )
        text = el.text.strip()
        if "Legal Name" in text:
            name = text.split("-", 1)[-1].strip()
            if name:
                return name
    except Exception:
        pass

    return fallback


def sanitize_folder_name(name):
    """Strip characters that aren't safe in a Windows/Mac/Linux folder name."""
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    # Control characters are illegal in Windows paths too.
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    # Windows silently drops trailing dots and spaces, so a folder created
    # as "Acme Ltd." is really "Acme Ltd" — which then never matches on the
    # next run and quietly makes a second folder.
    name = name.rstrip(". ")
    # CON, PRN, AUX, COM1... are reserved regardless of extension.
    if name.upper() in WINDOWS_RESERVED_NAMES:
        name = f"{name}_"
    # Leave room for the filenames that go inside it.
    if len(name) > 120:
        name = name[:120].rstrip(". ")
    return name or "Unknown_Company"


def return_pdf_path(return_type, output_dir, month_name, year):
    """The final resting path for one return, in one place so the
    already-downloaded check and the save both agree on the name."""
    display_name = "GSTR1" if return_type == "GSTR1" else "GSTR3B"
    return os.path.join(output_dir, f"{display_name}_of_{month_name}_{year}.pdf")


NOT_FILED_MARKERS = (
    "no records found",
    "return is not filed",
    "not filed",
    "no summary to display",
    "you have not filed",
    "return has not been filed",
)


def looks_not_filed(driver):
    """True if the page is telling us this return simply was never filed.

    Distinguishes "nothing to download because it doesn't exist" from
    "the button moved and the script is broken" — only the latter is worth
    retrying twice and dumping debug snapshots for.
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        return False
    return any(marker in body for marker in NOT_FILED_MARKERS)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def process_account(account, periods, return_types, base_output_dir, download_dir):
    """Log in for one company, detect its name, and download everything
    requested into a subfolder named after that company. Returns the list
    of failures (each a short description string) for this account."""
    failures = []
    skipped = []
    driver = None
    try:
        # Inside the try so that a failure to start Chrome, or to read the
        # window handle, still reaches the driver.quit() in `finally`
        # instead of orphaning a Chrome process.
        driver = create_driver(download_dir)
        main_handle = driver.current_window_handle

        login(driver, account["username"], account["password"])

        # From here on everything is fully automated (no more CAPTCHA/OTP
        # or questions to answer), so minimize the window and let it run
        # in the background while you do other things. If your system
        # doesn't support this for some reason, it's harmless — the run
        # just continues with the window visible.
        try:
            driver.minimize_window()
            print("Browser minimized — the rest of this run happens in the "
                  "background. Feel free to switch to other work.")
        except Exception:
            pass

        # The Excel sheet's optional "Company Name" column wins: it is what
        # the user actually chose to call this company, and it stays stable
        # even when the portal's own markup shifts.
        company_name = account.get("company") or detect_company_name(
            driver, fallback=account["username"]
        )
        company_folder = sanitize_folder_name(company_name)
        company_output_dir = os.path.join(base_output_dir, company_folder)
        os.makedirs(company_output_dir, exist_ok=True)

        # Each return type gets its own subfolder inside the company
        # folder, e.g. "<Company>/gstr 1/..." and "<Company>/gstr 3b/...".
        # Only the types actually requested get a folder — no more empty
        # "gstr 3b" directories after a GSTR-1-only run.
        all_return_type_dirs = {
            "GSTR1": os.path.join(company_output_dir, "gstr 1"),
            "GSTR3B": os.path.join(company_output_dir, "gstr 3b"),
        }
        return_type_dirs = {r: all_return_type_dirs[r] for r in return_types}
        for rtype_dir in return_type_dirs.values():
            os.makedirs(rtype_dir, exist_ok=True)

        print(f"\nCompany: {company_name}")
        print(f"Returns for this company will be saved in: {company_output_dir}")
        for rtype in return_types:
            label = "GSTR-1 " if rtype == "GSTR1" else "GSTR-3B"
            print(f"  - {label} -> {return_type_dirs[rtype]}")
        print("")

        # True ONLY while we are still sitting on the dashboard that login()
        # left us on. It is consumed by the first navigation and never set
        # again.
        #
        # This used to be a `first_navigation_done` flag that got reset to
        # False after any failure — which then made the NEXT period believe
        # it was on the dashboard when it was actually on whatever broken
        # page the failure left behind, so one bad month reliably knocked
        # over the month after it.
        fresh_from_login = True
        max_attempts = 3  # 1 initial try + 2 retries for a one-off flake
        for financial_year, month_name, year in periods:
            print(f"\n--- Processing {month_name} {year} (FY {financial_year}) ---")
            for rtype in return_types:

                # Already have this one? Don't spend a portal round-trip on
                # it. This is what makes re-running after a partial failure
                # cheap, and it is also why download_return should never
                # find an existing file to overwrite.
                target_path = return_pdf_path(
                    rtype, return_type_dirs[rtype], month_name, year
                )
                if os.path.isfile(target_path):
                    print(f"Already downloaded, skipping: {os.path.basename(target_path)}")
                    continue

                succeeded = False
                last_error = None
                for attempt in range(1, max_attempts + 1):
                    # Only the final attempt writes a debug snapshot —
                    # otherwise one failing period leaves up to nine
                    # screenshots plus nine full page dumps behind.
                    is_last = (attempt == max_attempts)
                    try:
                        # Always make sure we're on the single main tab
                        # before starting a period — a stray tab left open
                        # from a previous download is a common reason the
                        # script would otherwise get stuck.
                        close_extra_windows(driver, main_handle)

                        # The "we're already on the dashboard" shortcut is
                        # only ever valid straight after login. Consume it
                        # here so no later iteration can wrongly reuse it.
                        use_shortcut = fresh_from_login
                        fresh_from_login = False

                        open_wrapped = lambda d, s=use_shortcut: open_file_returns_page(
                            d, account["username"], account["password"], already_on_dashboard=s
                        )
                        run_step(driver, company_output_dir,
                                 f"open_file_returns_page_{month_name}_{year}_{rtype}",
                                 open_wrapped, save_snapshot=is_last)

                        select_wrapped = lambda d: select_period_and_search(d, financial_year, month_name)
                        run_step(driver, company_output_dir,
                                 f"select_period_{month_name}_{year}_{rtype}",
                                 select_wrapped, save_snapshot=is_last)

                        run_step(driver, company_output_dir,
                                 f"download_{rtype}_{month_name}_{year}",
                                 download_return, rtype, download_dir,
                                 return_type_dirs[rtype], month_name, year, main_handle,
                                 save_snapshot=is_last)

                        succeeded = True
                        break

                    except ReturnNotAvailable as e:
                        # Nothing to retry: the portal has no such return.
                        print(f"Skipping {rtype} for {month_name} {year} — {e}")
                        skipped.append(
                            f"{company_name}: {rtype} of {month_name} {year} (not available/not filed)"
                        )
                        succeeded = True  # not a failure; don't report it as one
                        break

                    except Exception as e:
                        last_error = e
                        if attempt < max_attempts:
                            print(f"\n--- Attempt {attempt} for {rtype} {month_name} {year} "
                                  f"didn't go through, retrying ({attempt + 1}/{max_attempts})... ---")
                            print(str(e).splitlines()[0] if str(e) else "")
                        continue

                if not succeeded:
                    # Don't let one bad month/return type kill the whole run
                    # — log it and keep going so the remaining months still
                    # get attempted.
                    print(f"\n--- Could not download {rtype} for {month_name} {year} "
                          f"after {max_attempts} attempts ---")
                    print(str(last_error))
                    print("Continuing with the remaining periods...\n")
                    failures.append(f"{company_name}: {rtype} of {month_name} {year}")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

    return failures, skipped


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    print("=== GST Return Downloader ===\n")

    # Everything downloads into the same folder this script FILE is saved in
    # (not wherever the terminal happens to be pointed).
    output_dir = os.path.dirname(os.path.abspath(__file__))
    download_dir = os.path.join(output_dir, "_chrome_downloads_temp")
    print(f"Downloaded returns will be saved in: {output_dir}\n"
          f"(each company gets its own subfolder, named after the company)\n")

    excel_input = input(
        "Paste the full path to your Excel file with GST credentials "
        "(or just press Enter to auto-use the .xlsx file in this script's folder): "
    )
    excel_path = find_excel_path(excel_input, output_dir)
    if not excel_path:
        raise FileNotFoundError(
            "No Excel file was given and none could be found next to this script.\n"
            "Put your credentials .xlsx in the same folder as this script, or paste "
            "its full path when asked (right-click the file -> Copy as path)."
        )
    warn_if_cloud_synced(excel_path)
    entries = read_credentials_from_excel(excel_path)
    if not entries:
        print("No valid credential rows found in the Excel file. Exiting.")
        return

    # If the Excel has more than one company, offer to run all of them in
    # this single session (each still needs its own CAPTCHA to log in, one
    # after another), or just process one as before.
    if len(entries) > 1:
        print(f"\n{len(entries)} companies found in the Excel file.")
        print("  1. Download for ALL of them in this run (one login/CAPTCHA per company, one after another)")
        print("  2. Download for just ONE of them")
        while True:
            batch_choice = input("Enter 1 or 2: ").strip()
            if batch_choice in ("1", "2"):
                break
            print("Invalid choice, try again.")
        accounts_to_process = entries if batch_choice == "1" else [choose_account(entries)]
    else:
        accounts_to_process = entries

    # --- Ask everything up front: one inclusive month range and which
    #     return(s) to download. FY and Quarter are derived for each month,
    #     so the user is never asked for them separately. These choices apply
    #     to every company being processed in this run. ---
    periods = get_periods_from_user()

    print("\nWhich return(s) do you want to download?")
    print("  1. GSTR-1 only")
    print("  2. GSTR-3B only")
    print("  3. Both")
    while True:
        r_choice = input("Enter 1, 2, or 3: ").strip()
        if r_choice in ("1", "2", "3"):
            break
        print("Invalid choice, try again.")

    return_types = []
    if r_choice in ("1", "3"):
        return_types.append("GSTR1")
    if r_choice in ("2", "3"):
        return_types.append("GSTR3B")

    all_failures = []
    all_skipped = []
    for i, account in enumerate(accounts_to_process, 1):
        if len(accounts_to_process) > 1:
            print(f"\n{'=' * 70}\nCompany {i} of {len(accounts_to_process)} "
                  f"— username: {account['username']}\n{'=' * 70}")
        # Isolate each company. Without this, a single login timeout or
        # browser crash on company 2 of 8 aborted the whole run and every
        # remaining company was silently never attempted — after the user
        # had already sat through the earlier CAPTCHAs.
        try:
            failures, skipped = process_account(
                account, periods, return_types, output_dir, download_dir
            )
            all_failures.extend(failures)
            all_skipped.extend(skipped)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            label = account.get("company") or account["username"]
            print(f"\n--- Company '{label}' could not be processed at all ---")
            print(f"{type(e).__name__}: {e}")
            print("Moving on to the next company...\n")
            all_failures.append(f"{label}: entire company failed ({type(e).__name__})")

    try:
        if os.path.isdir(download_dir) and not os.listdir(download_dir):
            os.rmdir(download_dir)
    except OSError:
        pass

    print("\n" + "=" * 70)
    if all_skipped:
        print("\nNot available on the portal (nothing filed for these periods):")
        for s in all_skipped:
            print(f"  - {s}")

    if all_failures:
        print("\nFinished, but these could not be downloaded "
              "(see messages above / each company's _debug folder):")
        for f in all_failures:
            print(f"  - {f}")
        print("\nRe-running the script will retry only the missing ones — "
              "anything already downloaded is skipped.")
    elif all_skipped:
        print("\nEverything that was available has been downloaded.")
    else:
        print("\nAll requested returns downloaded successfully.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nStopped at your request. Anything already downloaded has been kept.")
    except (FileNotFoundError, ValueError, RuntimeError) as e:
        print("\n--- ERROR ---")
        print(e)
    except Exception as e:
        # Anything else (a WebDriver/Chrome-version mismatch, a permission
        # error writing the PDF, ...) used to dump a raw traceback at a
        # non-technical user. Show the useful line, keep the traceback
        # available for diagnosis.
        print("\n--- UNEXPECTED ERROR ---")
        print(f"{type(e).__name__}: {e}")
        print("\nFull technical details:")
        traceback.print_exc()
    finally:
        try:
            input("\nPress Enter to close this window...")
        except (EOFError, KeyboardInterrupt):
            pass
