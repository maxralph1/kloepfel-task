import argparse
import mechanize
import re
import pathlib
import sys
import json
from bs4 import BeautifulSoup
import openpyxl

# Dictionaries to map arguments to values
schlagwortOptionen = {
    "all": 1,
    "min": 2,
    "exact": 3
}

class HandelsRegister:
    def __init__(self, args):
        self.args = args
        self.browser = mechanize.Browser()

        self.browser.set_debug_http(args.debug)
        self.browser.set_debug_responses(args.debug)
        # self.browser.set_debug_redirects(True)

        self.browser.set_handle_robots(False)
        self.browser.set_handle_equiv(True)
        self.browser.set_handle_gzip(True)
        self.browser.set_handle_refresh(False)
        self.browser.set_handle_redirect(True)
        self.browser.set_handle_referer(True)

        self.browser.addheaders = [
            (
                "User-Agent",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.5 Safari/605.1.15",
            ),
            ("Accept-Language", "en-GB,en;q=0.9"),
            ("Accept-Encoding", "gzip, deflate, br"),
            (
                "Accept",
                "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            ),
            ("Connection", "keep-alive"),
        ]

        self.cachedir = pathlib.Path("cache")
        self.cachedir.mkdir(parents=True, exist_ok=True)

    def open_startpage(self):
        self.browser.open("https://www.handelsregister.de/rp_web/welcome.xhtml", timeout=10)

    def companyname2cachename(self, companyname):
        # map a companyname to a filename, that caches the downloaded HTML, so re-running this script touches the
        # webserver less often.
        return self.cachedir / companyname

    def search_company(self):
        cachename = self.companyname2cachename(self.args.schlagwoerter)
        if self.args.force == False and cachename.exists():
            with open(cachename, "r") as f:
                html = f.read()
                print("return cached content for %s" % self.args.schlagwoerter)
        else:
            self.browser.open("https://www.handelsregister.de/rp_web/erweitertesuche.xhtml")
            if self.args.debug == True:
                print(self.browser.title())

            self.browser.select_form(name="form")

            self.browser["form:schlagwoerter"] = self.args.schlagwoerter
            so_id = schlagwortOptionen.get(self.args.schlagwortOptionen)

            self.browser["form:schlagwortOptionen"] = [str(so_id)]

            response_result = self.browser.submit()

            if self.args.debug == True:
                print(self.browser.title())

            html = response_result.read().decode("utf-8")
            with open(cachename, "w") as f:
                f.write(html)

        return get_companies_in_searchresults(html)

def parse_result(result):
    cells = []
    for cellnum, cell in enumerate(result.find_all('td')):
        cells.append(cell.text.strip())

    d = {}
    d['court'] = cells[1]  # Gericht
    d['name'] = cells[2]  # Firmenname
    d['state'] = cells[3]  # Sitz
    d['status'] = cells[4]  # Status
    d['documents'] = cells[5]  # Dokumente
    
        # Extract and print the relevant portion of HTML containing "SI"
    if 'SI' in cells[5]:
        start_index = max(cells[5].find('SI') - 100, 0)
        end_index = cells[5].find('SI') + 100
        print("---- Debug: HTML Around 'SI' ----")
        print(cells[5][start_index:end_index])
        print("---------------------------------")
        
    d['history'] = [6]  # Verlauf

    # Extract history if available
    history_cells = result.find_all('td')[8:]
    if history_cells:
        for i in range(0, len(history_cells), 2):
            event = history_cells[i].text.strip()
            date = history_cells[i + 1].text.strip()
            d['history'].append((event, date))

    return d

def get_companies_in_searchresults(html):
    soup = BeautifulSoup(html, 'html.parser')
    grid = soup.find('table', role='grid')

    results = []
    for result in grid.find_all('tr'):
        a = result.get('data-ri')
        if a is not None:
            d = parse_result(result)
            results.append(d)
    return results

def save_to_excel(companies, filepath):
    # Überprüfen, ob die Datei existiert, ansonsten eine neue erstellen
    try:
        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # Header einfügen
        sheet.append(["Firmenname", "Gericht", "Sitz", "Status", "Handelsregister-Nummer", "Dokumente", "Verlauf"])

    # Füge die Firmendaten zur Excel-Datei hinzu
    for company in companies:
        sheet.append([
            company.get("name", ""),
            company.get("court", ""),
            company.get("state", ""),
            company.get("status", ""),
            company.get("documents", ""),
           # " | ".join([f"{event[0]} am {event[1]}" for event in company.get("history", [])])
        ])

    workbook.save(filepath)

def parse_args():
    parser = argparse.ArgumentParser(description='A handelsregister CLI')
    parser.add_argument(
        "-d",
        "--debug",
        help="Enable debug mode and activate logging",
        action="store_true"
    )
    parser.add_argument(
        "-f",
        "--force",
        help="Force a fresh pull and skip the cache",
        action="store_true"
    )
    parser.add_argument(
        "-s",
        "--schlagwoerter",
        help="Search for the provided keywords",
        default="Kloepfel Consulting GmbH"
    )
    parser.add_argument(
        "-so",
        "--schlagwortOptionen",
        help="Keyword options: all=contain all keywords; min=contain at least one keyword; exact=contain the exact company name.",
        choices=["all", "min", "exact"],
        default="exact"
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Path to the output Excel file",
        default="handelsregister_result.xlsx"
    )
    args = parser.parse_args()

    if args.debug:
        import logging
        logger = logging.getLogger("mechanize")
        logger.addHandler(logging.StreamHandler(sys.stdout))
        logger.setLevel(logging.DEBUG)

    return args

if __name__ == "__main__":
    args = parse_args()
    h = HandelsRegister(args)
    h.open_startpage()
    companies = h.search_company()
    
    if companies:
        save_to_excel(companies, args.output)
        print(f"Ergebnisse wurden in der Datei {args.output} gespeichert.")
