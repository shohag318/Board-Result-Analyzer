"""
scraper.py - Playwright Web Automation for Bangladesh Education Board Results
Handles headless browser automation, math captcha solving, and data extraction.
"""

import asyncio
import base64
import hashlib
import json
import logging
import random
import re
import sys
import time
from typing import Any, Callable, Dict, List, Optional
import requests
import urllib3
from bs4 import BeautifulSoup
from playwright.async_api import Browser, BrowserContext, Page, async_playwright

try:
    import ddddocr
    _OCR = ddddocr.DdddOcr(show_ad=False)
    _OCR.set_ranges("0123456789")
except Exception:
    _OCR = None

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if sys.platform == "win32":
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

PORTAL_URL = "http://www.educationboardresults.gov.bd/"

# Board mapping helper (mapped to educationboardresults.gov.bd select values)
BOARD_MAP = {
    "dhaka": "dhaka",
    "barisal": "barisal",
    "barishal": "barisal",
    "chittagong": "chittagong",
    "chattogram": "chittagong",
    "comilla": "comilla",
    "cumilla": "comilla",
    "dinajpur": "dinajpur",
    "jessore": "jessore",
    "jashore": "jessore",
    "mymensingh": "mymensingh",
    "rajshahi": "rajshahi",
    "sylhet": "sylhet",
    "madrasah": "madrasah",
    "madrasa": "madrasah",
    "technical": "tec",
    "tec": "tec",
    "bteb": "tec",
    "dibs": "dibs"
}

# Exam mapping helper (mapped to educationboardresults.gov.bd select values)
EXAM_MAP = {
    "ssc": "ssc",
    "ssc_voc": "ssc_voc",
    "hsc": "hsc",
    "hsc_voc": "hsc_voc",
    "hsc_hbm": "hsc_hbm",
    "hsc_dic": "hsc_dic",
    "hsc_biz": "hsc_biz",
    "dakhil": "ssc",
    "alim": "hsc",
    "jsc": "jsc"
}


def solve_math_captcha_from_html(html_text: str) -> Optional[int]:
    """
    Finds arithmetic expression like '4 + 7', '8+3', or '9 - 4' in the captcha region.
    """
    # 1. Look for patterns like "5 + 4" or "9 - 2" near value_s input first
    soup = BeautifulSoup(html_text, "html5lib")
    value_s = soup.find("input", {"name": "value_s"})
    if value_s:
        parent_tr = value_s.find_parent("tr")
        if parent_tr:
            tr_text = parent_tr.get_text()
            match = re.search(r'(\d+)\s*([\+\-])\s*(\d+)', tr_text)
            if match:
                a, op, b = int(match.group(1)), match.group(2), int(match.group(3))
                res = a + b if op == "+" else a - b
                logger.info(f"Solved Captcha from parent row: {a} {op} {b} = {res}")
                return res

    # 2. Fallback: scan full HTML for digits with plus/minus sign near value_s
    matches = re.findall(r'(\d+)\s*([\+\-])\s*(\d+)', html_text)
    if matches:
        a, op, b = int(matches[0][0]), matches[0][1], int(matches[0][2])
        res = a + b if op == "+" else a - b
        logger.info(f"Solved Captcha from general HTML: {a} {op} {b} = {res}")
        return res

    return None


class ResultScraper:
    def __init__(self, headless: bool = True, timeout: int = 30000):
        self.headless = headless
        self.timeout = timeout
        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None

    async def start(self):
        """Initializes Playwright browser session using official Chrome/Edge to prevent AV quarantine."""
        self.playwright = await async_playwright().start()
        launch_args = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
        ]
        try:
            self.browser = await self.playwright.chromium.launch(
                channel="chrome",
                headless=self.headless,
                args=launch_args
            )
            logger.info("Playwright launched with system Google Chrome (channel='chrome')")
        except Exception as e1:
            try:
                self.browser = await self.playwright.chromium.launch(
                    channel="msedge",
                    headless=self.headless,
                    args=launch_args
                )
                logger.info("Playwright launched with system Microsoft Edge (channel='msedge')")
            except Exception as e2:
                logger.warning(f"Could not launch Chrome/Edge ({e1}, {e2}). Falling back to bundled chromium...")
                self.browser = await self.playwright.chromium.launch(
                    headless=self.headless,
                    args=launch_args
                )

        self.context = await self.browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 800}
        )

    async def stop(self):
        """Closes browser session."""
        try:
            if self.context:
                await self.context.close()
            if self.browser:
                await self.browser.close()
            if self.playwright:
                await self.playwright.stop()
        except Exception as e:
            logger.warning(f"Error closing Playwright browser: {e}")

    async def fetch_student_result(
        self,
        exam: str,
        year: str,
        board: str,
        roll: str,
        reg: str,
        max_retries: int = 3
    ) -> Dict[str, Any]:
        """
        Navigates to the portal, fills form, solves captcha, and parses student result.
        Retries up to max_retries on captcha error or temporary network glitch.
        """
        board_val = BOARD_MAP.get(board.lower(), board.lower())
        exam_val = EXAM_MAP.get(exam.lower(), exam.lower())

        for attempt in range(1, max_retries + 1):
            page: Optional[Page] = None
            alert_message = None

            try:
                page = await self.context.new_page()

                # Handle JavaScript alert popups (e.g. wrong captcha or result not found)
                async def on_dialog(dialog):
                    nonlocal alert_message
                    alert_message = dialog.message
                    logger.warning(f"Browser Dialog Alert: {alert_message}")
                    await dialog.accept()

                page.on("dialog", on_dialog)

                logger.info(f"Navigating to {PORTAL_URL} (Attempt {attempt}/{max_retries})")
                await page.goto(PORTAL_URL, wait_until="domcontentloaded", timeout=self.timeout)

                # Wait for form elements
                await page.wait_for_selector('input[name="roll"], #roll', timeout=self.timeout)

                # Select Examination
                exam_selector = 'select[name="sr"], select[name="exam"], select[name="examination"], #sr, #examination'
                if await page.query_selector(exam_selector):
                    try:
                        await page.select_option(exam_selector, value=exam_val)
                    except Exception:
                        # Try selecting by label/text if value mismatch
                        await page.select_option(exam_selector, label=exam.upper())

                # Select Year
                year_selector = 'select[name="year"], #year'
                try:
                    await page.select_option(year_selector, value=str(year), timeout=5000)
                except Exception:
                    await page.close()
                    return {
                        "Roll": roll,
                        "Registration": reg,
                        "Exam": exam.upper(),
                        "Year": year,
                        "Board": board.title(),
                        "Name": "N/A",
                        "Result": "Year Not Published",
                        "GPA": "0.00",
                        "Status": f"Results for year {year} are not yet published or available on the official board website. Turn on 'Test / Demo Mode' to preview.",
                        "Subject_Grades": {},
                        "Subject_Grades_Summary": f"Results for {year} are not available on the official website."
                    }

                # Select Board
                board_selector = 'select[name="board"], #board'
                try:
                    await page.select_option(board_selector, value=board_val)
                except Exception:
                    await page.select_option(board_selector, label=board.capitalize())

                # Enter Roll
                roll_selector = 'input[name="roll"], #roll'
                await page.fill(roll_selector, str(roll).strip())

                # Enter Registration (if provided)
                reg_selector = 'input[name="reg"], #reg'
                if reg:
                    await page.fill(reg_selector, str(reg).strip())

                # Captcha Handling: Extract and solve math problem
                html_content = await page.content()
                captcha_answer = solve_math_captcha_from_html(html_content)

                if captcha_answer is None:
                    # Alternative: directly evaluate text in the element before #value_s
                    try:
                        math_text = await page.locator("tr:has(input[name='value_s']) td:first-child").inner_text()
                        m = re.search(r'(\d+)\s*([\+\-])\s*(\d+)', math_text)
                        if m:
                            a, op, b = int(m.group(1)), m.group(2), int(m.group(3))
                            captcha_answer = a + b if op == "+" else a - b
                    except Exception:
                        pass

                if captcha_answer is None:
                    logger.error("Could not find or solve captcha. Retrying...")
                    await page.close()
                    await asyncio.sleep(1)
                    continue

                captcha_selector = 'input[name="value_s"], #value_s'
                await page.fill(captcha_selector, str(captcha_answer))

                # Submit the form and wait for navigation or alert
                submit_btn = 'input[type="submit"], input[name="button2"], #button2'
                await page.click(submit_btn)

                # Wait for alert dialog or navigation to result page
                start_wait = asyncio.get_event_loop().time()
                while asyncio.get_event_loop().time() - start_wait < 10.0:
                    if alert_message:
                        break
                    if "result.php" in page.url:
                        break
                    if await page.query_selector("table table") or await page.query_selector("td:has-text('Student Name')"):
                        break
                    await asyncio.sleep(0.3)

                if alert_message:
                    # Captcha was incorrect or student record not found
                    if "captcha" in alert_message.lower():
                        logger.warning(f"Captcha rejected by server for Roll {roll}. Retrying...")
                        await page.close()
                        continue
                    else:
                        logger.info(f"Server returned alert: '{alert_message}' for Roll: {roll}")
                        await page.close()
                        return {
                            "Roll": roll,
                            "Registration": reg,
                            "Exam": exam.upper(),
                            "Year": year,
                            "Board": board.title(),
                            "Name": "N/A",
                            "Result": "Not Found",
                            "GPA": "0.00",
                            "Status": alert_message,
                            "Subject_Grades": {},
                            "Subject_Grades_Summary": alert_message
                        }

                # Parse the response page
                page_html = await page.content()
                parsed_data = self._parse_result_html(page_html, roll, reg, exam, year, board)

                await page.close()
                return parsed_data

            except Exception as e:
                logger.warning(f"Error scraping Roll {roll} (Attempt {attempt}): {e}")
                if page:
                    try:
                        await page.close()
                    except Exception:
                        pass
                if attempt < max_retries:
                    await asyncio.sleep(1.5)

        # Fallback to HTTP session if Playwright attempts did not yield result
        logger.info(f"Trying HTTP fallback for Roll {roll}...")
        http_fallback = await asyncio.to_thread(fetch_via_http_session, exam, year, board, roll, reg)
        if http_fallback and http_fallback.get("Result") not in ["Failed / Error", "Result Not Found"]:
            return http_fallback

        return {
            "Roll": roll,
            "Registration": reg,
            "Exam": exam.upper(),
            "Year": year,
            "Board": board.title(),
            "Name": "N/A",
            "Result": "Failed / Error",
            "GPA": "0.00",
            "Status": "Connection / Captcha Timeout",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Connection / Captcha Timeout"
        }

    def _parse_result_html(
        self,
        html_str: str,
        roll: str,
        reg: str,
        exam: str,
        year: str,
        board: str
    ) -> Dict[str, Any]:
        """
        Parses the result page HTML tables to extract student details, GPA, and subject grades.
        """
        # Check for alert script inside returned HTML (e.g. <script>alert('Result not found')</script>)
        alert_match = re.search(r"alert\(['\"]([^'\"]+)['\"]\)", html_str)
        if alert_match:
            alert_msg = alert_match.group(1).strip()
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Result Not Found" if "not found" in alert_msg.lower() else "Failed / Error",
                "GPA": "0.00",
                "Status": alert_msg,
                "Subject_Grades": {},
                "Subject_Grades_Summary": alert_msg
            }

        soup = BeautifulSoup(html_str, "html5lib")

        # Check for error / not found message in page body
        text_content = soup.get_text()
        if "result not found" in text_content.lower() or "no result found" in text_content.lower():
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Result Not Found",
                "GPA": "0.00",
                "Status": "Not Found",
                "Subject_Grades": {},
                "Subject_Grades_Summary": "Result Not Found"
            }

        student_info: Dict[str, str] = {}
        subject_grades: Dict[str, str] = {}

        # Scan all tables for student details and subject breakdown
        tables = soup.find_all("table")

        for table in tables:
            rows = table.find_all("tr")
            for row in rows:
                cols = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
                
                # Check 2-column or 4-column key-value pairs (e.g. Roll No, Name, GPA)
                if len(cols) == 2:
                    k, v = cols[0].rstrip(":").strip(), cols[1].strip()
                    if k:
                        student_info[k.lower()] = v
                elif len(cols) == 4:
                    # Could be key-value pairs or side-by-side subjects
                    if cols[1] in ["A+", "A", "A-", "B", "C", "D", "F"] and cols[3] in ["A+", "A", "A-", "B", "C", "D", "F"]:
                        subject_grades[cols[0].strip()] = cols[1].strip()
                        subject_grades[cols[2].strip()] = cols[3].strip()
                    else:
                        k1, v1 = cols[0].rstrip(":").strip(), cols[1].strip()
                        k2, v2 = cols[2].rstrip(":").strip(), cols[3].strip()
                        if k1:
                            student_info[k1.lower()] = v1
                        if k2:
                            student_info[k2.lower()] = v2

                # Check 3-column subject grade rows (Code, Subject Name, Grade)
                elif len(cols) == 3:
                    c1, c2, c3 = cols[0].strip(), cols[1].strip(), cols[2].strip().upper()
                    # Check if c3 looks like a grade (A+, A, A-, B, C, D, F)
                    if c3 in ["A+", "A", "A-", "B", "C", "D", "F"]:
                        subj_name = f"{c1} {c2}".strip() if c1.isdigit() else c2
                        subject_grades[subj_name] = c3

        # Match common field variants
        name = (
            student_info.get("name")
            or student_info.get("student's name")
            or student_info.get("student name")
            or "N/A"
        )
        father_name = student_info.get("father's name") or student_info.get("father name") or ""
        mother_name = student_info.get("mother's name") or student_info.get("mother name") or ""
        gpa = student_info.get("gpa") or student_info.get("gpa (out of 5)") or "0.00"
        result_status = student_info.get("result") or student_info.get("result status") or ("PASSED" if gpa and gpa != "0.00" else "FAILED")
        group = student_info.get("group") or student_info.get("stream") or ""
        board_val = (student_info.get("board") or board).title()
        type_candidate = (
            student_info.get("type")
            or student_info.get("candidate type")
            or student_info.get("student type")
            or ""
        ).strip().upper()
        if not type_candidate or type_candidate in ["NAN", "NONE", "-"]:
            if 0 < len(subject_grades) <= 4:
                type_candidate = "IRREGULAR"
            else:
                type_candidate = "REGULAR"
        elif "IRREG" in type_candidate:
            type_candidate = "IRREGULAR"
        elif "REG" in type_candidate:
            type_candidate = "REGULAR"
        elif "IMP" in type_candidate:
            type_candidate = "IMPROVEMENT"
        elif "PRIV" in type_candidate:
            type_candidate = "PRIVATE"
        institute = student_info.get("institute") or student_info.get("institution") or ""

        # Subject Grades summary string: "Bangla: A+, English: A, Mathematics: A+"
        grades_summary = ", ".join([f"{k}: {v}" for k, v in subject_grades.items()]) if subject_grades else "N/A"

        return {
            "Roll": roll,
            "Registration": reg,
            "Exam": exam.upper(),
            "Year": year,
            "Board": board_val,
            "Name": name,
            "Father Name": father_name,
            "Mother Name": mother_name,
            "Institute": institute,
            "Group": group,
            "Type": type_candidate,
            "Result": result_status,
            "GPA": gpa,
            "Status": "Success" if name != "N/A" else "Result Found",
            "Subject_Grades": subject_grades,
            "Subject_Grades_Summary": grades_summary
        }


# ==========================================
# Simulated Demo Scraper (For Instant Testing)
# ==========================================
async def generate_simulated_result(
    roll: str,
    reg: str,
    exam: str,
    year: str,
    board: str
) -> Dict[str, Any]:
    """
    Generates a realistic mock result for testing and demonstration
    when live board servers are unreachable or for local dev.
    """
    await asyncio.sleep(random.uniform(0.6, 1.2))  # Realistic network latency simulation

    sample_names = [
        "MD. TANVIR HASAN", "NUSRAT JAHAN", "RAKIBUL ISLAM", "FATEMA AKTER",
        "SAKIB AL HASAN", "SADIA SULTANA", "ARIFUR RAHMAN", "MIM AKTER",
        "MEHEDI HASAN", "JANNATUL MAOWA", "KAMRUN NAHAR", "DIPA AKTER"
    ]

    # Diverse groups and realistic curriculum
    is_ssc = "ssc" in exam.lower() or "dakhil" in exam.lower()
    if is_ssc:
        group_curricula = {
            "SCIENCE": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("109", "MATHEMATICS"),
                ("154", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("136", "PHYSICS"), ("137", "CHEMISTRY"), ("138", "BIOLOGY"), ("126", "HIGHER MATHEMATICS")
            ],
            "BUSINESS STUDIES": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("109", "MATHEMATICS"),
                ("154", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("146", "ACCOUNTING"), ("152", "FINANCE & BANKING"), ("143", "BUSINESS ENTREPRENEURSHIP")
            ],
            "HUMANITIES": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("109", "MATHEMATICS"),
                ("154", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("140", "GEOGRAPHY & ENVIRONMENT"), ("153", "HISTORY OF BANGLADESH & WORLD CIVILIZATION"), ("141", "CIVICS & CITIZENSHIP")
            ]
        }
    else:
        group_curricula = {
            "SCIENCE": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("174", "PHYSICS"), ("175", "CHEMISTRY"), ("176", "BIOLOGY"), ("129", "HIGHER MATHEMATICS")
            ],
            "BUSINESS STUDIES": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("253", "ACCOUNTING"), ("277", "BUSINESS ORGANIZATION & MANAGEMENT"),
                ("286", "PRODUCTION MANAGEMENT & MARKETING"), ("292", "FINANCE, BANKING & INSURANCE")
            ],
            "HUMANITIES": [
                ("101", "BANGLA"), ("107", "ENGLISH"), ("275", "INFORMATION & COMMUNICATION TECHNOLOGY"),
                ("109", "ECONOMICS"), ("267", "ISLAMIC HISTORY & CULTURE"),
                ("269", "CIVICS & GOOD GOVERNANCE"), ("271", "SOCIAL WORK")
            ]
        }

    group = random.choice(["SCIENCE", "BUSINESS STUDIES", "HUMANITIES"])
    subjects = group_curricula[group]
    passing_pool = ["A+", "A", "A-", "B", "C", "D"]

    # Determine outcome: 4% Absent, 62% Pass, 19% 1-subject fail, 10% 2-subject fail, 5% 3-subject fail
    outcome_roll = random.random()
    subject_grades = {}

    if outcome_roll < 0.04:
        # Absent candidate
        selected_gpa = "0.00"
        result_status = "ABSENT"
        subject_grades = {}
    elif outcome_roll < 0.66:
        # Passed
        selected_gpa = random.choice(["5.00", "4.89", "4.72", "4.50", "4.17", "3.83", "3.50", "3.17"])
        for code, name in subjects:
            subject_grades[f"{code} {name}"] = random.choice(passing_pool)
        result_status = "PASSED"
    elif outcome_roll < 0.85:
        # 1-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_idx = random.randrange(len(subjects))
        for i, (code, name) in enumerate(subjects):
            if i == failed_idx:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)
    elif outcome_roll < 0.95:
        # 2-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_indices = set(random.sample(range(len(subjects)), 2))
        for i, (code, name) in enumerate(subjects):
            if i in failed_indices:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)
    else:
        # 3-Subject Fail
        selected_gpa = "0.00"
        result_status = "FAIL"
        failed_indices = set(random.sample(range(len(subjects)), 3))
        for i, (code, name) in enumerate(subjects):
            if i in failed_indices:
                subject_grades[f"{code} {name}"] = "F"
            else:
                subject_grades[f"{code} {name}"] = random.choice(passing_pool)

    summary = ", ".join([f"{k}: {v}" for k, v in subject_grades.items()]) if subject_grades else "ABSENT"

    candidate_type = "REGULAR" if random.random() < 0.75 else "IRREGULAR"

    return {
        "Roll": roll,
        "Registration": reg,
        "Exam": exam.upper(),
        "Year": year,
        "Board": board.title(),
        "Name": random.choice(sample_names),
        "Father Name": "MD. ABDUR RAHIM",
        "Mother Name": "MRS. ROKEYA BEGUM",
        "Group": group,
        "Type": candidate_type,
        "Result": result_status,
        "GPA": selected_gpa,
        "Status": "Success (Simulated)",
        "Subject_Grades": subject_grades,
        "Subject_Grades_Summary": summary
    }


def fetch_via_http_session(exam: str, year: str, board: str, roll: str, reg: str) -> Optional[Dict[str, Any]]:
    """Fast HTTP fallback using requests without browser overhead."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    }
    board_val = BOARD_MAP.get(board.lower(), board.lower())
    exam_val = EXAM_MAP.get(exam.lower(), exam.lower())
    try:
        s = requests.Session()
        r = s.get("http://www.educationboardresults.gov.bd/", headers=headers, timeout=10)
        if r.status_code == 200:
            captcha_ans = solve_math_captcha_from_html(r.text)
            if captcha_ans is not None:
                soup = BeautifulSoup(r.text, "html5lib")
                payload = {}
                for hidden in soup.find_all("input", {"type": "hidden"}):
                    if hidden.get("name"):
                        payload[hidden["name"]] = hidden.get("value", "")

                payload["sr"] = payload.get("sr", "3")
                payload["et"] = payload.get("et", "2")
                payload["exam"] = exam_val
                payload["year"] = str(year)
                payload["board"] = board_val
                payload["roll"] = str(roll).strip()
                payload["reg"] = str(reg).strip()
                payload["value_s"] = str(captcha_ans)
                payload["button2"] = "Submit"

                res_post = s.post("http://www.educationboardresults.gov.bd/result.php", data=payload, headers=headers, timeout=12)
                if res_post.status_code == 200:
                    dummy = ResultScraper()
                    return dummy._parse_result_html(res_post.text, roll, reg, exam, year, board)
    except Exception as e:
        logger.warning(f"HTTP fallback error: {e}")
    return None


CUMILLA_HSC_PORTAL_URL = "https://hscresult.comillaboard.gov.bd/result.php"
CUMILLA_SSC_PORTAL_URL = "https://result19.comillaboard.gov.bd/SSCRCBY26/individual/result.php"
CUMILLA_SSC_BASE_URL = "https://result19.comillaboard.gov.bd/SSCRCBY26/individual/"


def fetch_cumilla_hsc_result_sync(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """
    Directly fetches HSC result from the dedicated BISE Cumilla server:
    https://hscresult.comillaboard.gov.bd/result.php
    Requires ONLY Roll number (Registration and Captcha are not needed).
    """
    clean_roll = str(roll).strip()
    clean_reg = str(reg).strip() if reg else ""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        ),
        "Referer": "https://hscresult.comillaboard.gov.bd/"
    }

    try:
        resp = requests.post(
            CUMILLA_HSC_PORTAL_URL,
            data={"roll": clean_roll, "button2": "Submit"},
            headers=headers,
            verify=False,
            timeout=12
        )

        if resp.status_code == 200:
            scraper = ResultScraper()
            m = re.search(r'HSC Result\s*(\d{4})', resp.text, re.IGNORECASE)
            detected_year = m.group(1) if m else (str(year) if year else "2026")

            result = scraper._parse_result_html(
                html_str=resp.text,
                roll=clean_roll,
                reg=clean_reg,
                exam="HSC",
                year=detected_year,
                board="Cumilla"
            )
            return result
        else:
            return {
                "Roll": clean_roll,
                "Registration": clean_reg,
                "Exam": "HSC",
                "Year": str(year) if year else "2026",
                "Board": "Cumilla",
                "Name": "N/A",
                "Result": f"Server Error ({resp.status_code})",
                "GPA": "0.00",
                "Status": f"Cumilla Board server returned HTTP {resp.status_code}",
                "Subject_Grades": {},
                "Subject_Grades_Summary": f"Server responded with error status {resp.status_code}"
            }
    except Exception as e:
        logger.error(f"Error fetching Cumilla HSC result for roll {clean_roll}: {e}")
        return {
            "Roll": clean_roll,
            "Registration": clean_reg,
            "Exam": "HSC",
            "Year": str(year) if year else "2026",
            "Board": "Cumilla",
            "Name": "N/A",
            "Result": "Network Error",
            "GPA": "0.00",
            "Status": f"Failed to connect to {CUMILLA_HSC_PORTAL_URL}: {str(e)}",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Failed to connect to Cumilla board server"
        }


def fetch_cumilla_ssc_result_sync(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """
    Directly fetches SSC result from the dedicated BISE Cumilla server:
    https://result19.comillaboard.gov.bd/SSCRCBY26/individual/result.php
    Requires ONLY Roll number (Registration and Captcha are not needed).
    """
    clean_roll = str(roll).strip()
    clean_reg = str(reg).strip() if reg else ""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        ),
        "Referer": "https://result19.comillaboard.gov.bd/SSCRCBY26/individual/"
    }

    try:
        # First try result.php endpoint
        resp = None
        target_used = CUMILLA_SSC_PORTAL_URL
        try:
            resp = requests.post(
                CUMILLA_SSC_PORTAL_URL,
                data={"roll": clean_roll, "button2": "Submit"},
                headers=headers,
                verify=False,
                timeout=12
            )
        except Exception:
            resp = None

        # Fallback to base endpoint if result.php returned 404/405 or was unreachable
        if resp is None or resp.status_code in [404, 405]:
            target_used = CUMILLA_SSC_BASE_URL
            resp = requests.post(
                CUMILLA_SSC_BASE_URL,
                data={"roll": clean_roll, "button2": "Submit"},
                headers=headers,
                verify=False,
                timeout=12
            )

        if resp and resp.status_code == 200:
            scraper = ResultScraper()
            m = re.search(r'SSC Result\s*(\d{4})', resp.text, re.IGNORECASE)
            detected_year = m.group(1) if m else (str(year) if year else "2026")

            result = scraper._parse_result_html(
                html_str=resp.text,
                roll=clean_roll,
                reg=clean_reg,
                exam="SSC",
                year=detected_year,
                board="Cumilla"
            )
            return result
        else:
            status_code = resp.status_code if resp else "No Response"
            return {
                "Roll": clean_roll,
                "Registration": clean_reg,
                "Exam": "SSC",
                "Year": str(year) if year else "2026",
                "Board": "Cumilla",
                "Name": "N/A",
                "Result": f"Server Error ({status_code})",
                "GPA": "0.00",
                "Status": f"Cumilla Board SSC server returned HTTP {status_code}",
                "Subject_Grades": {},
                "Subject_Grades_Summary": f"Server responded with error status {status_code}"
            }
    except Exception as e:
        logger.error(f"Error fetching Cumilla SSC result for roll {clean_roll}: {e}")
        return {
            "Roll": clean_roll,
            "Registration": clean_reg,
            "Exam": "SSC",
            "Year": str(year) if year else "2026",
            "Board": "Cumilla",
            "Name": "N/A",
            "Result": "Network Error",
            "GPA": "0.00",
            "Status": f"Failed to connect to {CUMILLA_SSC_BASE_URL}: {str(e)}",
            "Subject_Grades": {},
            "Subject_Grades_Summary": "Failed to connect to Cumilla SSC board server"
        }


async def fetch_cumilla_hsc_result(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """Async wrapper for fetch_cumilla_hsc_result_sync."""
    return await asyncio.to_thread(fetch_cumilla_hsc_result_sync, roll, year, reg)


async def fetch_cumilla_ssc_result(
    roll: str,
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """Async wrapper for fetch_cumilla_ssc_result_sync."""
    return await asyncio.to_thread(fetch_cumilla_ssc_result_sync, roll, year, reg)


async def fetch_cumilla_result(
    roll: str,
    exam: str = "hsc",
    year: str = "",
    reg: str = ""
) -> Dict[str, Any]:
    """
    Directly routes Cumilla result requests to dedicated SSC or HSC portal.
    SSC: https://result19.comillaboard.gov.bd/SSCRCBY26/individual/
    HSC: https://hscresult.comillaboard.gov.bd/
    """
    clean_exam = str(exam).strip().lower()
    if "ssc" in clean_exam or "dakhil" in clean_exam:
        return await fetch_cumilla_ssc_result(roll=roll, year=year, reg=reg)
    return await fetch_cumilla_hsc_result(roll=roll, year=year, reg=reg)


class V2ResultService:
    """
    Manages direct high-speed session to the National Education Board V2 Result System
    (educationboardresults.gov.bd/v2/home and eboardresults.com/v2/home).
    Provides instant result lookups for ALL education boards in Bangladesh.
    """
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/123.0.0.0 Safari/537.36"
            )
        })
        self.authenticated = False
        self.last_auth_time = 0.0
        self.base_host = "https://www.educationboardresults.gov.bd"

    def ensure_authenticated(self) -> bool:
        now = time.time()
        if self.authenticated and (now - self.last_auth_time < 900):
            return True

        for base_host in ["https://www.educationboardresults.gov.bd", "https://eboardresults.com"]:
            try:
                challenge_url = f"{base_host}/_challenge-engine?goto=%2Fv2%2Fhome"
                r = self.session.get(challenge_url, timeout=8)
                m = re.search(r'var challenge = ({.*?});', r.text)
                if m:
                    c = json.loads(m.group(1))
                    proof_str = f"browser-proof:{c['seed']}:{c['salt']}"
                    sha = hashlib.sha256(proof_str.encode("utf-8")).hexdigest()
                    time.sleep(2.5)  # Required by verification nonce/delay
                    verify_url = f"{base_host}/_challenge-verify"
                    v = self.session.post(
                        verify_url,
                        json={"token": c["token"], "answer": f"hash:{sha}"},
                        headers={"Referer": challenge_url},
                        timeout=8
                    )
                    if v.status_code == 200:
                        self.session.get(f"{base_host}/v2/home", timeout=8)
                        self.authenticated = True
                        self.last_auth_time = now
                        self.base_host = base_host
                        logger.info(f"V2ResultService successfully authenticated via {base_host}")
                        return True
            except Exception as e:
                logger.warning(f"Error authenticating challenge on {base_host}: {e}")

        return False

    def get_captcha_image(self) -> Optional[bytes]:
        if not self.ensure_authenticated():
            return None
        try:
            t = int(time.time() * 1000)
            url = f"{self.base_host}/v2/captcha?t={t}"
            r = self.session.get(
                url,
                headers={"Referer": f"{self.base_host}/v2/home"},
                timeout=8
            )
            if r.status_code == 200 and len(r.content) > 500:
                return r.content
        except Exception as e:
            logger.warning(f"Failed to fetch captcha from {self.base_host}: {e}")
        return None

    def query_result(
        self,
        exam: str,
        year: str,
        board: str,
        roll: str,
        reg: str,
        captcha: str
    ) -> Dict[str, Any]:
        if not self.ensure_authenticated():
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Connection Error",
                "GPA": "0.00",
                "Status": "Unable to connect to national board server. Please try again.",
                "Subject_Grades": {},
                "Subject_Grades_Summary": "Failed to connect to national board server."
            }

        board_val = BOARD_MAP.get(board.lower(), board.lower())
        exam_val = EXAM_MAP.get(exam.lower(), exam.lower())

        payload = {
            "board": board_val,
            "exam": exam_val,
            "year": str(year),
            "result_type": "1",
            "roll": str(roll).strip(),
            "reg": str(reg).strip() if reg else "",
            "captcha": str(captcha).strip()
        }
        headers = {
            "Referer": f"{self.base_host}/v2/home",
            "X-Requested-With": "XMLHttpRequest"
        }

        try:
            resp = self.session.post(
                f"{self.base_host}/v2/getres",
                data=payload,
                headers=headers,
                timeout=12
            )
            if resp.status_code == 200:
                data = resp.json()
                return self._parse_v2_json(data, roll, reg, exam, year, board)
            else:
                return {
                    "Roll": roll,
                    "Registration": reg,
                    "Exam": exam.upper(),
                    "Year": year,
                    "Board": board.title(),
                    "Name": "N/A",
                    "Result": f"Server Error ({resp.status_code})",
                    "GPA": "0.00",
                    "Status": f"Server returned HTTP {resp.status_code}",
                    "Subject_Grades": {},
                    "Subject_Grades_Summary": f"Server responded with status {resp.status_code}"
                }
        except Exception as e:
            logger.error(f"Error querying V2 result for roll {roll}: {e}")
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Network Error",
                "GPA": "0.00",
                "Status": str(e),
                "Subject_Grades": {},
                "Subject_Grades_Summary": "Connection error querying board server."
            }

    def _parse_v2_json(
        self,
        data: Dict[str, Any],
        roll: str,
        reg: str,
        exam: str,
        year: str,
        board: str
    ) -> Dict[str, Any]:
        if data.get("status") != 0:
            msg = data.get("msg") or "Result not found."
            is_cap_error = "captcha" in msg.lower() or "security" in msg.lower()
            return {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Invalid Security Code" if is_cap_error else "Result Not Found",
                "GPA": "0.00",
                "Status": msg,
                "Subject_Grades": {},
                "Subject_Grades_Summary": msg
            }

        res = data.get("res", {})
        name = res.get("name") or "N/A"
        fname = res.get("fname") or ""
        mname = res.get("mname") or ""
        board_name = (res.get("board_name") or board).title()
        group = res.get("stud_group") or ""
        stud_type = (res.get("stud_type") or "REGULAR").upper()
        inst_name = res.get("inst_name") or ""

        res_detail = str(res.get("res_detail") or "")
        gpa = str(res.get("gpa") or "0.00")
        if not gpa or gpa == "0.00":
            m = re.search(r"GPA\s*=\s*([\d\.]+)", res_detail, re.IGNORECASE)
            if m:
                gpa = m.group(1)

        result_status = "PASSED" if (gpa and gpa != "0.00" and gpa != "0") or "PASS" in res_detail.upper() else "FAILED"
        if "ABSENT" in res_detail.upper():
            result_status = "ABSENT"
        elif "FAIL" in res_detail.upper():
            result_status = "FAILED"

        # Subject breakdown
        subject_grades = {}
        disp = res.get("display_details")
        if isinstance(disp, list):
            for item in disp:
                if isinstance(item, dict):
                    sname = item.get("sub_name") or item.get("name") or item.get("sub") or ""
                    sgrade = item.get("grade") or item.get("gpa") or ""
                    if sname and sgrade:
                        subject_grades[sname.strip()] = sgrade.strip()
        elif isinstance(disp, dict):
            subject_grades = disp

        summary = ", ".join([f"{k}: {v}" for k, v in subject_grades.items()]) if subject_grades else "N/A"

        return {
            "Roll": roll,
            "Registration": reg or res.get("regno", ""),
            "Exam": exam.upper(),
            "Year": year,
            "Board": board_name,
            "Name": name,
            "Father Name": fname,
            "Mother Name": mname,
            "Institute": inst_name,
            "Group": group,
            "Type": stud_type,
            "Result": result_status,
            "GPA": gpa,
            "Status": "Success",
            "Subject_Grades": subject_grades,
            "Subject_Grades_Summary": summary
        }


# Global V2 service instance
_V2_SERVICE: Optional[V2ResultService] = None


def get_v2_captcha() -> Dict[str, Any]:
    """Retrieves fresh captcha image and runs OCR prediction."""
    global _V2_SERVICE
    if _V2_SERVICE is None:
        _V2_SERVICE = V2ResultService()
    img_bytes = _V2_SERVICE.get_captcha_image()
    if not img_bytes:
        return {"success": False, "image": "", "predicted": ""}
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    data_uri = f"data:image/jpeg;base64,{b64}"
    pred = ""
    if _OCR:
        try:
            pred = _OCR.classification(img_bytes)
        except Exception:
            pass
    return {
        "success": True,
        "image": data_uri,
        "predicted": pred if len(pred) == 4 and pred.isdigit() else ""
    }


async def scrape_single_student(
    exam: str,
    year: str,
    board: str,
    roll: str,
    reg: str = "",
    captcha: str = "",
    demo_mode: bool = False
) -> Dict[str, Any]:
    """
    Fetches the result for an individual student across ALL boards in Bangladesh.
    Uses dedicated Cumilla portal for Cumilla HSC, national V2 portal for all other boards,
    or simulated results in demo mode.
    """
    if demo_mode:
        logger.info(f"Checking individual Roll: {roll} in DEMO mode.")
        return await generate_simulated_result(roll, reg, exam, year, board)

    # 1. Special Fast Dedicated Server for Cumilla Board Results (Only Roll required)
    if board.lower() in ["comilla", "cumilla"]:
        is_ssc = "ssc" in exam.lower() or "dakhil" in exam.lower()
        logger.info(f"Checking Cumilla dedicated portal for Roll: {roll} ({'SSC' if is_ssc else 'HSC'})...")
        cumilla_res = await fetch_cumilla_result(roll=roll, exam=exam, year=year, reg=reg)
        if cumilla_res and cumilla_res.get("Name") not in ["N/A", ""] and cumilla_res.get("Result") not in ["Result Not Found", "Network Error"]:
            return cumilla_res
        if is_ssc or "hsc" in exam.lower():
            return cumilla_res

    # 2. National V2 Board Result System (Dhaka, Rajshahi, Chittagong, Barisal, Dinajpur, Jessore, Sylhet, Mymensingh, Madrasah, Technical, DIBS)
    global _V2_SERVICE
    if _V2_SERVICE is None:
        _V2_SERVICE = V2ResultService()

    # If captcha provided directly, query immediately
    if captcha and str(captcha).strip():
        logger.info(f"Querying National V2 portal for {board.title()} {exam.upper()} Roll: {roll} with provided captcha.")
        return await asyncio.to_thread(_V2_SERVICE.query_result, exam, year, board, roll, reg, captcha)

    # If captcha not provided, attempt auto-solving via OCR (up to 3 tries)
    logger.info(f"Attempting auto-captcha solve for {board.title()} {exam.upper()} Roll: {roll}...")
    for attempt in range(1, 4):
        cap_bytes = await asyncio.to_thread(_V2_SERVICE.get_captcha_image)
        if cap_bytes and _OCR:
            code = _OCR.classification(cap_bytes)
            if len(code) == 4 and code.isdigit():
                logger.info(f"Auto-OCR predicted captcha: '{code}' (Attempt {attempt})")
                res = await asyncio.to_thread(_V2_SERVICE.query_result, exam, year, board, roll, reg, code)
                if res.get("Result") not in ["Invalid Security Code"]:
                    return res
        await asyncio.sleep(0.5)

    # Return clear instruction asking user to enter the 4-digit code shown on screen
    return {
        "Roll": roll,
        "Registration": reg,
        "Exam": exam.upper(),
        "Year": year,
        "Board": board.title(),
        "Name": "N/A",
        "Result": "Captcha Required",
        "GPA": "0.00",
        "Status": "Please enter the 4-digit security code shown on the screen to fetch result for " + board.title() + " Board.",
        "Subject_Grades": {},
        "Subject_Grades_Summary": "Enter the 4-digit security code shown above and click Check Single Result."
    }


async def scrape_student_batch(
    student_records: List[Dict[str, str]],
    exam: str,
    year: str,
    board: str,
    demo_mode: bool = False,
    on_progress: Optional[Callable[[int, int, Dict[str, Any]], Any]] = None
) -> List[Dict[str, Any]]:
    """
    Iterates through a list of student records, fetching results for each one.
    Calls `on_progress(current_index, total_count, student_result)` for real-time reporting.
    """
    results = []
    total = len(student_records)

    if demo_mode:
        logger.info(f"Running batch of {total} students in DEMO mode.")
        for idx, item in enumerate(student_records, start=1):
            roll = str(item.get("Roll", "")).strip()
            reg = str(item.get("Registration", "")).strip()
            result = await generate_simulated_result(roll, reg, exam, year, board)
            results.append(result)
            if on_progress:
                await on_progress(idx, total, result)
        return results

    # 1. Special Fast Dedicated Server for Cumilla Board (Only Roll required)
    if board.lower() in ["comilla", "cumilla"]:
        is_ssc = "ssc" in exam.lower() or "dakhil" in exam.lower()
        portal_name = CUMILLA_SSC_BASE_URL if is_ssc else "https://hscresult.comillaboard.gov.bd/"
        logger.info(f"Running fast batch extraction for {total} Cumilla students via {portal_name}")
        for idx, item in enumerate(student_records, start=1):
            roll = str(item.get("Roll", "")).strip()
            reg = str(item.get("Registration", "")).strip()
            logger.info(f"[{idx}/{total}] Fetching Cumilla Roll: {roll} ({'SSC' if is_ssc else 'HSC'})")

            try:
                res = await fetch_cumilla_result(roll=roll, exam=exam, year=year, reg=reg)
            except Exception as e_ind:
                logger.warning(f"Error fetching Cumilla roll {roll}: {e_ind}")
                res = {
                    "Roll": roll,
                    "Registration": reg,
                    "Exam": exam.upper(),
                    "Year": year,
                    "Board": "Cumilla",
                    "Name": "N/A",
                    "Result": "Error",
                    "GPA": "0.00",
                    "Status": str(e_ind),
                    "Subject_Grades": {},
                    "Subject_Grades_Summary": "Error fetching record from Cumilla server."
                }

            results.append(res)
            if on_progress:
                await on_progress(idx, total, res)

            # Modest delay between server requests (150ms) to ensure stability
            if idx < total:
                await asyncio.sleep(0.15)

        return results

    # 2. National V2 Board Extraction (Dhaka, Rajshahi, Chittagong, Barisal, Dinajpur, etc.)
    global _V2_SERVICE
    if _V2_SERVICE is None:
        _V2_SERVICE = V2ResultService()

    logger.info(f"Running batch extraction for {total} students ({board.title()} {exam.upper()}) via National V2 System.")

    # Pre-authenticate V2 session
    await asyncio.to_thread(_V2_SERVICE.ensure_authenticated)

    for idx, item in enumerate(student_records, start=1):
        roll = str(item.get("Roll", "")).strip()
        reg = str(item.get("Registration", "")).strip()
        logger.info(f"[{idx}/{total}] Processing Roll: {roll}, Reg: {reg}")

        res = None
        # Try auto-solving captcha for this student
        for attempt in range(1, 4):
            try:
                cap_bytes = await asyncio.to_thread(_V2_SERVICE.get_captcha_image)
                if cap_bytes and _OCR:
                    code = _OCR.classification(cap_bytes)
                    if len(code) == 4 and code.isdigit():
                        candidate_res = await asyncio.to_thread(_V2_SERVICE.query_result, exam, year, board, roll, reg, code)
                        if candidate_res.get("Result") not in ["Invalid Security Code"]:
                            res = candidate_res
                            break
            except Exception as e_cap:
                logger.warning(f"Captcha attempt {attempt} error for roll {roll}: {e_cap}")
            await asyncio.sleep(0.3)

        if not res:
            res = {
                "Roll": roll,
                "Registration": reg,
                "Exam": exam.upper(),
                "Year": year,
                "Board": board.title(),
                "Name": "N/A",
                "Result": "Verification Pending",
                "GPA": "0.00",
                "Status": "Security code verification required by official board. Please check individual roll on the single search tab or turn on Demo Mode.",
                "Subject_Grades": {},
                "Subject_Grades_Summary": "Official board security code verification needed."
            }

        results.append(res)
        if on_progress:
            await on_progress(idx, total, res)

        if idx < total:
            await asyncio.sleep(0.4)

    return results
