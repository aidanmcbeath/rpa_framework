#Importing libraries
from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import time

#Initialising variables (will pass credentials in dynamically from vault at a later stage)
#Credentials
rsb_username = "maria"
rsb_password = "thoushallnotpass"

#Other
excel_url = "https://robotsparebinindustries.com/SalesData.xlsx"
workbook_path = "C:\\Users\\bot1\\Documents\\Bots\\SalesData.xlsx"

#Main process flow is contained in 'minimal_task'

@task
def minimal_task():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(slowmo=1,)
    open_the_intranet_website()
    log_in(rsb_username, rsb_password)
    download_excel_file(excel_url, workbook_path)
    fill_form_with_excel_data(workbook_path)
    collect_results()
    export_as_pdf()
    log_out()

#Defining functions to be used above

def open_the_intranet_website():
    """Navigates to the RSB intranet site"""
    browser.goto("https://robotsparebinindustries.com/")

def log_in(username, password):
    """Fills in the login fields and clicks the "Log In" button"""
    page = browser.page()
    page.fill("#username", username)
    page.fill("#password", password)
    page.click("button:text('Log in')")

def fill_and_submit_form(sales_rep):
    """Fills in the sales data and clicks the 'Submit' button"""
    page = browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.select_option('#salestarget', str(sales_rep["Sales Target"]))

    page.click("text=Submit")

def download_excel_file(excel_url, workbook_path):
    """Downloads the excel file from the URL"""
    http = HTTP()
    http.download(excel_url, overwrite=True, target_file=workbook_path)

def fill_form_with_excel_data(workbook_path):
    """Read data from excel and fill in the sales form"""
    excel = Files()
    excel.open_workbook(workbook_path)
    worksheet = excel.read_worksheet_as_table(name="data", header=True)
    excel.close_workbook()
    
    for employee in worksheet:
        fill_and_submit_form(employee)

def collect_results():
    """Takes a screenshot of the page"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the sales data table to a padf file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """presses the 'Log Out' button"""
    page = browser.page()
    time.sleep(5)
    page.click("#logout")
    time.sleep(10)