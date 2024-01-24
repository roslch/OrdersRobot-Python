# Import libraries
from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem

# Main task
@task
def robot_order_python():
    """Generate required robot orders from csv file
    """
    browser = open_robot_intranet()
    download_excel()
    process_orders_from_csv(browser)
    create_zip_and_remove_files()

# Subtasks
def open_robot_intranet():
    """open required website

    Returns:
        object: browser (Selenium)
    """
    browser = Selenium()
    browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order", maximized=True)
    
    return browser

def download_excel():
    """Download csv file from website
    """
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", "output/orders.csv", True)
    
def fill_form(browser, robot):
    """Fill out robot order form

    Args:
        browser (Selenium): current website
        robot (dict): current robot order information
    """
    browser.select_from_list_by_value("//select[@id='head']", robot["Head"])
    browser.select_radio_button("body", robot["Body"])
    browser.input_text("//label[text()='3. Legs:']/following::input[1]", robot["Legs"])
    browser.input_text("//input[@id='address']", robot["Address"])
    
def take_screenshot_and_make_pdf(browser, orderNo):
    """Create pdf based on robot screenshot and receipt information

    Args:
        browser (Selenium): current website
        orderNo (str): current order number
    """
    pdf = PDF()
    receiptInfo = False
    while not receiptInfo:
        try:
            browser.click_button("order")
            browser.wait_until_element_is_visible("//div[@id='receipt']", "2s")
            receiptInfo = browser.is_element_visible("//div[@id='receipt']")
        except:
            continue
    browser.wait_until_element_is_visible("//div[@id='robot-preview-image']", "5s")
    receipt = browser.get_element_attribute("//div[@id='receipt']", "outerHTML")
    browser.screenshot("//div[@id='robot-preview-image']", f"output/robot{orderNo}.png")
    pdf.html_to_pdf(receipt, f"output/robot{orderNo}.pdf")
    pdf.add_files_to_pdf([f"output/robot{orderNo}.png"], f"output/robot{orderNo}.pdf", True)
    browser.click_button("order-another")
    
def process_orders_from_csv(browser):
    """Process all robot requests based on the csv file

    Args:
        browser (Selenium): current website
    """
    table = Tables()
    robots = table.read_table_from_csv("output/orders.csv", True, delimiters=",")

    for robot in robots:
        orderNo = robot["Order number"]
        browser.click_button_when_visible("//button[text()='Yep']")
        fill_form(browser, robot)
        take_screenshot_and_make_pdf(browser, orderNo)
    browser.close_browser()
        
def create_zip_and_remove_files():
    """Generate zip file of all receipts and remove remaining files
    """
    files = Archive()
    lib = FileSystem()
    files.archive_folder_with_zip("output", "output/robots.zip", include="*.pdf")
    types = ('png', 'pdf', 'csv')
    _ = [lib.remove_file(file.path) for ext in types for file in lib.find_files(f"output/*.{ext}")]