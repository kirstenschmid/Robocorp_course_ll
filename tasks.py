
from robocorp.tasks import task
from RPA.Tables import Tables
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import time
from pypdf import PdfWriter, PdfReader
from  RPA.Archive import Archive
import pdfkit

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(slowmo=100)
    open_robot_order_website()
    orders = get_orders()
    for i in orders:
        fill_the_form(i)

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    return orders

def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def fill_the_form(i):  # Now 'order' will accept the argument passed when calling this function
    page = browser.page()
    close_annoying_modal()
    fill_in_order(i)
    store_receipt_and_screenshot(i)  # Make sure this function is correctly defined to accept an order too
    page.click("[id='order-another']")

def fill_in_order(i):
    page = browser.page()
    page.select_option("xpath=//select[@id='head']", (i["Head"]))
    body_value = i["Body"]
    page.dblclick(f"[id='id-body-{body_value}']")
    page.fill("xpath=/html/body/div/div/div[1]/div/div[1]/form/div[3]/input ", (i["Legs"]))
    page.fill("#address", (i["Address"]))
    page.click("[id='order']")
    error = page.locator(".alert.alert-danger[role=\'alert\']")

    if error.is_visible():
        click_button_with_retry(max_retries=100, wait_time_between_retries=1)

def click_button_with_retry(max_retries=100, wait_time_between_retries=1):
    page = browser.page()
    selector = "[id='order']"
   
    # Attempts to click a button, retrying up to max_retries times with a wait in between.
    for _ in range(max_retries):
        try:
            # Try clicking the button
            page.click(selector)
            # Wait to see if an error appears
            page.wait_for_selector('.alert.alert-danger[role=\'alert\']', state='visible', timeout=wait_time_between_retries * 1000)
            print("Error detected after clicking. Retrying...")
        except Exception as e:
            # If the element does not become visible within the timeout, it's considered a success
            print("Button clicked successfully without error.")
            return True
        time.sleep(wait_time_between_retries)
    print("Maximum retries reached without success.")
    return False

def store_receipt_and_screenshot(i):
    page = browser.page()
    page.wait_for_selector("[id='receipt']", state='visible')
    order_number = i

    # Define file paths
    screenshot_path = f"screenshots/{order_number}.png"
    pdf_path = f"pdfs/{order_number}.pdf"
    merged_pdf_path = f"./output/{order_number}_final.pdf"
    zip_path = f"./output/{order_number}.zip"
    
    # taking screenshot of the image of the robot
    page.locator("[id='robot-preview-image']").screenshot(path=screenshot_path)
    
    # stroring receipt in pdf output
    receipts_html = page.locator("[id='receipt']").inner_html()

    # using the RPA.PDF library and the html_to_pdf() function does not work to create a PDF file from html content despite it being used in the first course
    # it seems like the library does not contain any function to do this anymore. Nevertheless used this library for demonstration purposes. 

    PDF.html_to_pdf(receipts_html, pdf_path)
    merger = PdfWriter()
    merged_pdf_path = f"./output/{order_number}_final.pdf"

    merger.append([pdf_path, screenshot_path], merged_pdf_path)
    create_zip(merged_pdf_path, zip_path)

def create_zip(files, zip_path):
    archive = Archive()
    """Create a ZIP archive containing the specified files."""
    archive.archive_files(
        files=files,
        archive=zip_path,
        format="zip"
    )