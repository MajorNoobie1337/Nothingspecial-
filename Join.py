"""
IBM Cloud Invite Automation - Playwright + Edge
-------------------------------------------------
1. Parses your .olm export to extract IBM Cloud invite links
2. For each link opens Edge and:
   - Clicks "J'accepte le produit"
   - Clicks "Rejoindre un compte"

Requirements:
    pip install playwright
    python -m playwright install msedge

Usage:
    python email_automation.py
"""

import zipfile
import re
import time
import logging
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ── CONFIG ────────────────────────────────────────────────────────────────────
OLM_FILE = "your_emails.olm"       # <- Change to your actual .olm file path
WAIT_TIMEOUT = 20000                # milliseconds (20 seconds)
DELAY_BETWEEN_EMAILS = 4            # seconds pause between each invite
LOG_FILE = "automation_log.txt"
# ─────────────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

# IBM Cloud invite URL pattern
IBM_INVITE_PATTERN = (
    r'https://cloud\.ibm\.com/registration/accept-invite-start\?token=[A-Za-z0-9\-_]+'
)


def extract_ibm_invite_links(olm_path: str) -> list[str]:
    """Unzip the .olm archive and find all IBM Cloud invite URLs."""
    links = []

    try:
        with zipfile.ZipFile(olm_path, 'r') as z:
            xml_files = [n for n in z.namelist() if n.endswith('.xml')]
            log.info(f"Found {len(xml_files)} XML files inside '{olm_path}'")

            for xml_file in xml_files:
                with z.open(xml_file) as f:
                    try:
                        content = f.read().decode('utf-8', errors='replace')
                        found = re.findall(IBM_INVITE_PATTERN, content)
                        for link in found:
                            link = link.rstrip('"\'>)')
                            if link not in links:
                                links.append(link)
                    except Exception as e:
                        log.warning(f"Could not read {xml_file}: {e}")

    except zipfile.BadZipFile:
        log.error(f"'{olm_path}' is not a valid .olm/ZIP file.")
        return []

    log.info(f"Extracted {len(links)} unique IBM Cloud invite links")
    return links


def process_invite(page, url: str, index: int, total: int) -> bool:
    """
    Full flow for one IBM Cloud invite:
      1. Load the invite URL
      2. Click "J'accepte le produit"
      3. Click "Rejoindre un compte"
    """
    log.info(f"\n[{index}/{total}] {url}")

    try:
        page.goto(url, wait_until="networkidle", timeout=WAIT_TIMEOUT)
        time.sleep(2)  # Extra settle time for IBM Cloud

        # ── STEP 1: Accept the product ("J'accepte le produit") ───────────
        accepte_selectors = [
            "text=J'accepte le produit",
            "text=accepte le produit",
            "label:has-text('accepte')",
            "input[type='checkbox']",
            "[id*='accept']",
            "[class*='accept']",
        ]

        clicked = False
        for selector in accepte_selectors:
            try:
                page.wait_for_selector(selector, timeout=5000)
                page.click(selector)
                clicked = True
                log.info(f"[{index}] OK - Clicked \"J'accepte le produit\"")
                break
            except PlaywrightTimeout:
                continue

        if not clicked:
            log.error(f"[{index}] FAIL - Could not find \"J'accepte le produit\"")
            return False

        time.sleep(2)

        # ── STEP 2: Submit ("Rejoindre un compte") ─────────────────────────
        rejoindre_selectors = [
            "text=Rejoindre un compte",
            "button:has-text('Rejoindre')",
            "a:has-text('Rejoindre')",
            "button[type='submit']",
            "[id*='rejoindre']",
            "[class*='rejoindre']",
        ]

        clicked = False
        for selector in rejoindre_selectors:
            try:
                page.wait_for_selector(selector, timeout=5000)
                page.click(selector)
                clicked = True
                log.info(f"[{index}] OK - Clicked 'Rejoindre un compte'")
                break
            except PlaywrightTimeout:
                continue

        if not clicked:
            log.error(f"[{index}] FAIL - Could not find 'Rejoindre un compte'")
            return False

        time.sleep(3)  # Wait for confirmation
        log.info(f"[{index}] SUCCESS - Invite accepted")
        return True

    except Exception as e:
        log.error(f"[{index}] Unexpected error: {e}")
        return False


def main():
    # 1. Extract links
    links = extract_ibm_invite_links(OLM_FILE)

    if not links:
        log.error("No IBM Cloud invite links found.")
        log.error("Check OLM_FILE path and that emails contain IBM Cloud URLs.")
        return

    total = len(links)
    log.info(f"Starting automation for {total} invites...\n")

    # 2. Results tracker
    results = {"success": 0, "failed": 0, "failed_urls": []}

    # 3. Launch Edge via Playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="msedge",
            headless=False,        # Show the browser
            args=["--start-maximized"]
        )
        context = browser.new_context(
            locale="fr-FR",        # French UI to match button labels
            viewport={"width": 1280, "height": 800}
        )
        page = context.new_page()

        try:
            for i, url in enumerate(links, start=1):
                success = process_invite(page, url, i, total)
                if success:
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    results["failed_urls"].append(url)
                time.sleep(DELAY_BETWEEN_EMAILS)

        finally:
            browser.close()

    # 4. Summary
    log.info(f"\n{'='*60}")
    log.info(f"DONE: {results['success']}/{total} succeeded, {results['failed']} failed")
    if results["failed_urls"]:
        log.info("Failed URLs:")
        for u in results["failed_urls"]:
            log.info(f"  - {u}")
    log.info(f"Log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
