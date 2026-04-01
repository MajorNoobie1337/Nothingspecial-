"""
IBM Cloud Invite Automation - Playwright + Edge
-------------------------------------------------
1. Parses your .olm export to extract IBM Cloud invite links
2. For each link opens Edge and:
   - Ticks the checkbox "J'accepte le produit"
   - Clicks the "Rejoindre un compte" button (unlocked after checkbox)

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
import random

# ── CONFIG ────────────────────────────────────────────────────────────────────
OLM_FILE = "your_emails.olm"       # <- Change to your actual .olm file path
WAIT_TIMEOUT = 20000                # milliseconds (20 seconds)
DELAY_BETWEEN_EMAILS = 4            # seconds pause between each invite
LOG_FILE = "automation_log.txt"
PROXY_USER = "your_username"        # <- Your corporate proxy username
PROXY_PASS = "your_password"        # <- Your corporate proxy password
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

IBM_INVITE_PATTERN = (
    r'https://cloud\.ibm\.com/registration/accept-invite-start\?token=[A-Za-z0-9\-_]+'
)


def extract_ibm_invite_links(olm_path: str) -> list:
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
    log.info(f"
[{index}/{total}] {url}")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=WAIT_TIMEOUT)
        time.sleep(4)

        # Scroll down to make sure checkbox is visible
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)

        # STEP 1: Tick checkbox - exact selector from screenshot
        try:
            page.wait_for_selector("input[type='checkbox']", timeout=8000)
            page.click("input[type='checkbox']")
            log.info(f"[{index}] OK - Ticked checkbox")
        except PlaywrightTimeout:
            log.error(f"[{index}] FAIL - Checkbox not found")
            return False

        time.sleep(2)

        # STEP 2: Click Rejoindre un compte (now enabled)
        try:
            page.wait_for_selector("button:has-text('Rejoindre')", state="visible", timeout=8000)
            btn = page.query_selector("button:has-text('Rejoindre')")
            if btn and not btn.is_disabled():
                btn.click()
                log.info(f"[{index}] OK - Clicked Rejoindre un compte")
            else:
                log.error(f"[{index}] FAIL - Rejoindre button still disabled")
                return False
        except PlaywrightTimeout:
            log.error(f"[{index}] FAIL - Rejoindre button not found")
            return False

        time.sleep(3)
        log.info(f"[{index}] SUCCESS")
        return True

    except Exception as e:
        log.error(f"[{index}] Error: {e}")
        return False


        time.sleep(2)  # Give page time to enable the Rejoindre button

        # ── STEP 2: Click "Rejoindre un compte" (enabled after checkbox) ──
        # From screenshot: greyed button at bottom, becomes blue after checkbox
        rejoindre_selectors = [
            "button:has-text('Rejoindre un compte')",
            "button:has-text('Rejoindre')",
            "[role='button']:has-text('Rejoindre')",
            "text=Rejoindre un compte",
            "button[type='submit']",
        ]

        joined = False
        for selector in rejoindre_selectors:
            try:
                # Wait for it to be visible and not disabled
                page.wait_for_selector(selector, state="visible", timeout=8000)
                btn = page.query_selector(selector)
                if btn and not btn.is_disabled():
                    btn.click()
                    joined = True
                    log.info(f"[{index}] OK - Clicked 'Rejoindre un compte'")
                    break
            except PlaywrightTimeout:
                continue
            except Exception:
                continue

        if not joined:
            log.error(f"[{index}] FAIL - 'Rejoindre un compte' not clickable")
            return False

        time.sleep(3)
        log.info(f"[{index}] SUCCESS - Invite accepted")
        return True

    except Exception as e:
        log.error(f"[{index}] Unexpected error: {e}")
        return False


def main():
    # 1. Extract links from OLM
    links = extract_ibm_invite_links(OLM_FILE)

    if not links:
        log.error("No IBM Cloud invite links found.")
        log.error("Check OLM_FILE path and that emails contain IBM Cloud URLs.")
        return

    total = len(links)
    log.info(f"Starting automation for {total} invites...\n")

    results = {"success": 0, "failed": 0, "failed_urls": []}

    # 2. Launch Edge via Playwright
    with sync_playwright() as p:
        import os
        edge_profile = os.path.expanduser(
            "~/Library/Application Support/Microsoft Edge"
        )
        context = p.chromium.launch_persistent_context(
            user_data_dir=edge_profile,
            channel="msedge",
            headless=False,
            args=["--start-maximized"],
            locale="fr-FR",
            viewport={"width": 1280, "height": 800},
            proxy={
                "server": "http://vip-svc-sra-prx.fr.xcp.net.intra:3132",
                "username": PROXY_USER,
                "password": PROXY_PASS,
            }
        )
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        page = context.new_page()

        # Pause here so you can log in to IBM Cloud manually first
        page.goto('https://cloud.ibm.com')
        input('>>> Log in to IBM Cloud in the browser, then press ENTER here to start automation...')


        try:
            for i, url in enumerate(links, start=1):
                success = process_invite(page, url, i, total)
                if success:
                    results["success"] += 1
                else:
                    results["failed"] += 1
                    results["failed_urls"].append(url)
                time.sleep(DELAY_BETWEEN_EMAILS + random.uniform(1, 3))
        finally:
            context.close()

    # 3. Summary
    log.info(f"\n{'='*60}")
    log.info(f"DONE: {results['success']}/{total} succeeded, {results['failed']} failed")
    if results["failed_urls"]:
        log.info("Failed URLs:")
        for u in results["failed_urls"]:
            log.info(f"  - {u}")
    log.info(f"Log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
