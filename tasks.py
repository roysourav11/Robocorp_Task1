import os

import time

 

from robocorp.tasks import task

from robocorp import browser

 

from RPA.HTTP import HTTP

from RPA.Excel.Files import Files

from RPA.PDF import PDF

 

@task

def robot_spare_bin_python():

    """Insert the sales data for the week and export it as a PDF"""

    # Prefer a preinstalled browser on the runner to avoid Playwright CDN downloads that

    # fail in hosted environments with certificate/network issues.

    # Order: Chrome -> Edge -> Chromium.

    browser.configure(

        browser_engine="chrome",

        install=False,

        slowmo=1000,

    )

    try:

        open_the_intranet_website()

    except Exception:

        browser.configure(

            browser_engine="msedge",

            install=False,

            slowmo=1000,

        )

        open_the_intranet_website()

    time.sleep(int(os.getenv("BROWSER_WAIT_BEFORE_LOGIN_SECONDS", "5")))

    log_in()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))

    download_excel_file()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))

    #fill_and_submit_sales_form()

    fill_form_with_excel_data()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))

    collect_results()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))

    export_as_pdf()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))

    log_out()

    time.sleep(int(os.getenv("BROWSER_WAIT_AFTER_LOGIN_SECONDS", "5")))  

    close_browser()

 

def open_the_intranet_website():

    """Navigates to the given URL"""

    browser.goto("https://robotsparebinindustries.com/")

 

def log_in():

    """Fills in the login form and clicks the 'Log in' button"""

    page = browser.page()

    page.fill("#username", "maria")

    page.fill("#password", "thoushallnotpass")

    page.click("button:text('Log in')")

 

def fill_and_submit_sales_form(sales_rep):

    """Fills in the sales data and click the 'Submit' button"""

    page = browser.page()

 

    page.fill("#firstname", sales_rep["First Name"])

    page.fill("#lastname", sales_rep["Last Name"])

    page.select_option("#salestarget", str(sales_rep["Sales Target"]))

    page.fill("#salesresult", str(sales_rep["Sales"]))

    page.click("text=Submit")

 

def download_excel_file():

    """Downloads excel file from the given URL"""

    http = HTTP()

    http.download(

        url="https://robotsparebinindustries.com/SalesData.xlsx",

        overwrite=True,

        verify=False,

    )

 

def fill_form_with_excel_data():

    """Read data from excel and fill in the sales form"""

    excel = Files()

    excel.open_workbook("SalesData.xlsx")

    worksheet = excel.read_worksheet_as_table("data", header=True)

    excel.close_workbook()

 

    for row in worksheet:

        fill_and_submit_sales_form(row)

 

def collect_results():

    """Take a screenshot of the page"""

    page = browser.page()

    page.screenshot(path="output/sales_summary.png")

 

def log_out():

    """Presses the 'Log out' button"""

    page = browser.page()  

    page.click("text=Log out")

 

def export_as_pdf():

    """Export the data to a pdf file"""

    page = browser.page()

    sales_results_html = page.locator("#sales-results").inner_html()

 

    pdf = PDF()

    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

 

def close_browser():

    """Closes the active browser context."""

    browser.context().close()